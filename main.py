"""
====================================================================
 FUSION & DÉDOUBLONNAGE DE CONTACTS (avec matching FNMNS)
====================================================================

Application Streamlit permettant de :
  1. Importer jusqu'à 3 fichiers Excel de contacts hétérogènes.
  2. Importer (optionnellement) un 4e fichier "Adhérents FNMNS".
  3. Normaliser strictement chaque champ (nom, prénom, adresse,
     téléphone, email, etc.).
  4. Identifier les doublons via fuzzy matching (rapidfuzz) sur
     Nom + Prénom (accents/fautes de frappe tolérés).
  5. Fusionner ou conserver les lignes selon la distance
     géographique (geopy / Nominatim) entre les adresses.
  6. Marquer les contacts retrouvés dans la base FNMNS.
  7. Exporter un tableau final unique au format .xlsx.

Auteur : généré avec Claude (Anthropic) — à adapter selon vos besoins.
====================================================================
"""

import io
import json
import re
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
from unidecode import unidecode

try:
    import phonenumbers
    from phonenumbers import PhoneNumberFormat

    PHONENUMBERS_AVAILABLE = True
except ImportError:  # pragma: no cover
    PHONENUMBERS_AVAILABLE = False

try:
    from geopy.geocoders import Nominatim
    from geopy.distance import geodesic
    from geopy.extra.rate_limiter import RateLimiter
    from geopy.exc import GeocoderServiceError, GeocoderTimedOut

    GEOPY_AVAILABLE = True
except ImportError:  # pragma: no cover
    GEOPY_AVAILABLE = False


# ====================================================================
# 1. CONSTANTES & CONFIGURATION
# ====================================================================

# Colonnes finales attendues dans le tableau consolidé
FINAL_COLUMNS = [
    "id",
    "nom",
    "prenom",
    "adresse",
    "code_postal",
    "ville",
    "pays",
    "tel",
    "email",
    "mode_contact_pref",
    "latitude",
    "longitude",
    "tag",
    "filtres",
    "commentaire",
]

# Séparateur utilisé pour représenter les champs "liste" dans Excel
# (Excel ne stocke pas de vraies listes : on joint les valeurs avec
# ce séparateur. À la relecture, il suffit de faire .split(LIST_SEP)).
LIST_SEP = "; "

# Dictionnaire de détection des colonnes sources (noms possibles ->
# champ cible). La détection se fait en tolérant accents/majuscules.
COLUMN_CANDIDATES = {
    "nom": ["nom", "name", "lastname", "last name", "nom de famille", "surname"],
    "prenom": ["prenom", "prénom", "firstname", "first name", "prenom(s)"],
    # Colonne "Nom complet" (ex: "Arnaud JOSSET") utilisée en dernier recours
    # uniquement si aucune colonne nom/prenom séparée n'est trouvée.
    "nom_complet": [
        "nom complet",
        "nom et prenom",
        "prenom et nom",
        "nom prenom",
        "prenom nom",
        "full name",
        "contact",
        "identite",
    ],
    "adresse": ["adresse", "address", "adresse postale", "rue", "voie"],
    "code_postal": ["code postal", "cp", "codepostal", "zip", "zipcode", "postal code"],
    "ville": ["ville", "city", "commune", "localite", "localité"],
    "pays": ["pays", "country"],
    "tel": [
        "telephone",
        "téléphone",
        "tel",
        "tél",
        "phone",
        "mobile",
        "portable",
        "gsm",
        "numero de telephone",
        "numéro de téléphone",
    ],
    "email": [
        "email",
        "e-mail",
        "mail",
        "courriel",
        "adresse mail",
        "adresse email"
    ],
    "tag": [
        "tag",
        "tags",
        "categorie",
        "catégorie",
        "groupe",
        "type",
        "etiquette",
        "étiquette",
    ],
    # Colonne de préférence de canal (whatsapp/sms/tel/email), typiquement
    # présente sur les formulaires Google Forms de collecte de contacts.
    "mode_contact": [
        "mode de contact preferentiel",
        "mode de contact",
        "canal de contact prefere",
        "preference de contact",
        "canal prefere",
        "moyen de contact prefere",
    ],
}

# Empêche certains mots génériques de "voler" la colonne d'un autre champ
# (ex: "Adresse email de contact" contient le mot 'adresse' mais est bien
# une colonne email, pas une colonne adresse postale).
COLUMN_EXCLUDE_MARKERS = {
    "adresse": ["email", "mail", "courriel", "adresse ip"],
    "tel": ["email", "mail", "courriel"],
    "email": ["telephone", "téléphone", " tel ", "adresse postale"],
    "ville": ["email", "mail", "telephone", "téléphone"],
    "code_postal": ["email", "mail", "telephone", "téléphone"],
    # Une colonne "Nom et Prénom" / "Prénom et Nom" est une colonne
    # COMBINÉE : elle ne doit être captée que par le champ 'nom_complet'
    # (qui sait la découper dans le bon sens), pas directement par les
    # champs 'nom' ou 'prenom' pris isolément (sinon le nom complet
    # entier se retrouve collé dans un seul des deux champs).
    "nom": ["et prenom", "prenom et"],
    "prenom": ["et nom", "nom et"],
}

DEFAULT_FUZZY_THRESHOLD = 85  # % minimal de similarité Nom+Prénom
DEFAULT_DISTANCE_THRESHOLD_KM = 30  # seuil de distance pour fusion (Cas C)


# ====================================================================
# 2. UTILITAIRES DE NORMALISATION DE TEXTE
# ====================================================================


def strip_accents_lower(text: str) -> str:
    """Retire les accents et met en minuscule (pour comparaisons)."""
    if text is None:
        return ""
    return unidecode(str(text)).strip().lower()


def build_match_key(nom: str, prenom: str) -> str:
    """Construit la clé normalisée Nom+Prénom utilisée pour le fuzzy
    matching (identité de personne, et matching FNMNS).

    Les tirets et apostrophes sont remplacés par des espaces AVANT
    comparaison : sans cela, un prénom composé écrit avec un tiret
    dans un fichier ('Anne-Frédérique') et avec un espace dans un
    autre ('Anne Frédérique') sont vus comme des mots totalement
    différents par `rapidfuzz.token_sort_ratio` (qui découpe sur les
    espaces), ce qui fait chuter le score de similarité très en
    dessous du seuil et empêche à tort la fusion des deux lignes."""
    nom_norm = re.sub(r"[-'’]", " ", strip_accents_lower(nom))
    prenom_norm = re.sub(r"[-'’]", " ", strip_accents_lower(prenom))
    nom_norm = re.sub(r"\s+", " ", nom_norm).strip()
    prenom_norm = re.sub(r"\s+", " ", prenom_norm).strip()
    return f"{nom_norm} {prenom_norm}".strip()


def slugify(text: str) -> str:
    """Transforme un texte en identifiant 'slug' : minuscule, sans
    accent, sans espace ni caractère spécial (séparateur '-')."""
    text = unidecode(str(text or "")).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    return text or "inconnu"


def title_case_fr(text: str) -> str:
    """Formatte un texte en 'Nom Propre' (Title Case), en gérant
    correctement les tirets et apostrophes (ex: 'jean-pierre' ->
    'Jean-Pierre', "d'artagnan" -> "D'Artagnan")."""
    if not text:
        return ""
    text = str(text).strip()
    text = re.sub(r"\s+", " ", text)
    return text.title()


def is_blank(value) -> bool:
    """Vérifie si une valeur est vide/NaN/None."""
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip() == "" or str(value).strip().lower() == "nan"


# ====================================================================
# 3. DÉTECTION AUTOMATIQUE DES COLONNES (fichiers hétérogènes)
# ====================================================================


def detect_columns(df: pd.DataFrame) -> dict:
    """Associe à chaque champ cible (nom, prenom, adresse, ...) le nom
    de la colonne source correspondante dans le DataFrame.

    Robuste à :
    - la casse / les accents / les espaces superflus ;
    - les noms de colonnes "composés" sans espace (ex: 'adressecontact',
      'Code postal contact') via une recherche de sous-chaîne ;
    - les intitulés longs/bruités (ex: formulaires Google Forms) via
      `rapidfuzz.fuzz.token_set_ratio`, qui ignore les mots en trop.

    En cas de colonnes concurrentes pour un même champ, une colonne déjà
    "nettoyée/formatée" (contenant 'nettoye', 'formate'...) est préférée.

    Retourne un dict {champ_cible: nom_colonne_source_ou_None}.
    """
    normalized_cols = {col: strip_accents_lower(col) for col in df.columns}
    CLEAN_MARKERS = ("nettoye", "formate", "clean")
    mapping = {}

    # Les candidats trop courts (ex: 'nom', 'cp', 'tel') sont ambigus en
    # recherche de sous-chaîne (ils matcheraient 'Nom Officiel EPCI',
    # 'Coordonnées...', etc.). On ne les utilise donc que pour une
    # correspondance EXACTE ; seuls les candidats suffisamment longs et
    # spécifiques sont éligibles à la recherche par sous-chaîne/floue.
    MIN_LEN_FOR_FUZZY = 5

    for target_field, candidates in COLUMN_CANDIDATES.items():
        normalized_candidates = [strip_accents_lower(c) for c in candidates]
        loose_candidates = [
            c for c in normalized_candidates if len(c) >= MIN_LEN_FOR_FUZZY
        ]
        excluded_markers = COLUMN_EXCLUDE_MARKERS.get(target_field, [])
        scored = []

        for col, norm_col in normalized_cols.items():
            if any(marker in norm_col for marker in excluded_markers):
                continue  # colonne exclue pour ce champ (mot-clé ambigu)

            best_score = 0.0

            # 1) correspondance exacte (autorisée même pour candidats courts)
            if norm_col in normalized_candidates:
                best_score = 100.0
            else:
                # 2) correspondance par sous-chaîne (gère les noms composés)
                for cand in loose_candidates:
                    if cand in norm_col:
                        ratio = len(cand) / max(len(norm_col), 1)
                        best_score = max(best_score, 60 + 35 * ratio)
                    elif norm_col in cand:
                        ratio = len(norm_col) / max(len(cand), 1)
                        best_score = max(best_score, 60 + 35 * ratio)

                # 3) correspondance floue (mots superflus/ordre différent),
                #    utilisée seulement si aucune sous-chaîne n'a matché,
                #    pour éviter qu'un candidat court ne survalorise un
                #    intitulé sans rapport (ex: 'commune' dans 'Coordonnées').
                if best_score == 0 and loose_candidates:
                    fuzzy_score = max(
                        (
                            fuzz.token_set_ratio(norm_col, cand)
                            for cand in loose_candidates
                        ),
                        default=0,
                    )
                    best_score = fuzzy_score if fuzzy_score >= 80 else 0

            if best_score > 0:
                if any(marker in norm_col for marker in CLEAN_MARKERS):
                    best_score += 12  # bonus pour colonnes déjà nettoyées/formatées
                scored.append((best_score, -len(col), col))

        if scored:
            scored.sort(reverse=True)
            mapping[target_field] = scored[0][2]
        else:
            mapping[target_field] = None

    return mapping


# ====================================================================
# 4. FORMATAGE STRICT DES CHAMPS (selon le cahier des charges)
# ====================================================================


def format_code_postal(raw) -> str:
    """Formatte un code postal en 'XX XXX' (ex: '75 001')."""
    if is_blank(raw):
        return ""
    text = str(raw).strip()

    # Piège fréquent en Corse : le CODE INSEE de la commune (ex: '2A004',
    # '2B134' — département '2A'/'2B' + 3 chiffres) est parfois saisi par
    # erreur à la place du code postal. Ce n'est PAS un code postal valide
    # (les vrais CP corses sont numériques, ex: '20000' Ajaccio). On ne
    # tente pas de deviner un CP à partir de ça (on produirait une valeur
    # fausse mais plausible, ex: '02 004' -> laisserait croire au
    # département 02 Aisne) : on retourne vide plutôt qu'une donnée
    # erronée silencieuse.
    if re.fullmatch(r"\d[ABab]\d{3}", text):
        return ""

    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    digits = digits.zfill(5)[:5] if len(digits) <= 5 else digits
    if len(digits) == 5:
        return f"{digits[:2]} {digits[2:]}"
    # Cas générique (codes postaux étrangers de longueur différente)
    return f"{digits[:2]} {digits[2:]}" if len(digits) > 2 else digits


def split_multi_values(raw) -> list:
    """Sépare une cellule pouvant contenir plusieurs valeurs
    (téléphones ou emails) via les séparateurs usuels."""
    if is_blank(raw):
        return []
    parts = re.split(r"[;,/|\n]+", str(raw))
    return [p.strip() for p in parts if p.strip()]


def split_multi_phones(raw) -> list:
    """Comme `split_multi_values`, mais spécifique aux numéros de
    téléphone : reconnaît EN PLUS le séparateur textuel ' ou ' (ex:
    '06 22 97 12 53 ou 06 78 38 47 35'), fréquent en saisie manuelle
    mais absent des séparateurs habituels (';', ',', '/'). Restreint
    aux numéros de téléphone (pas utilisé pour email/tag) car 'ou'
    pourrait sinon couper du texte libre à tort."""
    if is_blank(raw):
        return []
    text = re.sub(r"\s+(?:ou|et)\s+", ";", str(raw), flags=re.IGNORECASE)
    return split_multi_values(text)


def format_phone_number(raw: str, default_region: str = "FR") -> Optional[str]:
    """Convertit un numéro de téléphone en format international
    uniforme '+CC(C) X XX XX XX XX', en s'appuyant sur la bibliothèque
    `phonenumbers` pour identifier correctement l'indicatif pays et le
    numéro national, quelle que soit la longueur de l'indicatif (1, 2
    ou 3 chiffres, ex: +33 France, +376 Andorre). Le regroupement final
    (1er chiffre puis paires) est toujours appliqué nous-mêmes plutôt
    que la convention propre à chaque pays, afin de garder un format
    cohérent dans tout le fichier de sortie. `default_region` sert
    quand le numéro est fourni sans indicatif (hypothèse : France).

    Retourne None si le numéro n'est pas exploitable (trop court,
    caractères non numériques, etc.).
    """
    if is_blank(raw):
        return None
    if not PHONENUMBERS_AVAILABLE:
        # Repli si la librairie `phonenumbers` n'est pas installée :
        # on applique une conversion manuelle basique (hypothèse par
        # défaut : numéro français) plutôt que de renvoyer les chiffres
        # bruts tels quels (ex: éviter '+0664645848', invalide).
        return _format_phone_fallback(raw, default_region)

    text = str(raw).strip()
    try:
        parsed = phonenumbers.parse(text, default_region)
    except phonenumbers.NumberParseException:
        return None
    if not phonenumbers.is_possible_number(parsed):
        return None

    # `phonenumbers` connaît la longueur exacte de l'indicatif (1, 2 ou
    # 3 chiffres selon le pays) : on l'utilise pour isoler proprement
    # l'indicatif, puis on applique notre propre regroupement uniforme
    # '+CC(C) X XX XX XX XX' (plutôt que le format INTERNATIONAL par
    # défaut de la librairie, qui varie selon le pays - ex: tirets
    # pour les numéros américains).
    return _group_uniform(str(parsed.country_code), str(parsed.national_number))


def _group_uniform(country_code: str, national_number: str) -> str:
    """Formate '+CC(C) X XX XX XX XX' : indicatif pays (2 ou 3 chiffres
    selon le pays), puis le 1er chiffre du numéro national isolé, puis
    le reste regroupé par paires."""
    if not national_number:
        return f"+{country_code}"
    first, rest = national_number[0], national_number[1:]
    pairs = [rest[i : i + 2] for i in range(0, len(rest), 2)]
    return (
        f"+{country_code} {first} " + " ".join(pairs)
        if pairs
        else f"+{country_code} {first}"
    )


def _format_phone_fallback(raw: str, default_region: str = "FR") -> Optional[str]:
    """Formatage manuel de secours (utilisé UNIQUEMENT si le paquet
    `phonenumbers` n'est pas installé). Beaucoup plus limité que la
    bibliothèque : suppose un numéro français si aucun '+' n'est
    présent, et ne peut pas distinguer de manière fiable les
    indicatifs à 2 chiffres (ex: +41) des indicatifs à 3 chiffres
    (ex: +376) puisqu'il n'a pas la table des indicatifs. Installez
    `phonenumbers` (pip install phonenumbers) pour un résultat fiable
    quel que soit le pays.
    """
    has_plus = str(raw).strip().startswith("+")
    digits = re.sub(r"\D", "", str(raw))
    if not digits:
        return None

    if has_plus:
        # Sans base de données des indicatifs, on ne peut pas savoir de
        # façon fiable si l'indicatif fait 1, 2 ou 3 chiffres -> on
        # suppose 2 chiffres (cas le plus fréquent en Europe).
        cc, national = digits[:2], digits[2:]
    elif default_region == "FR" and digits.startswith("0") and len(digits) == 10:
        cc, national = "33", digits[1:]
    elif digits.startswith("33") and len(digits) == 11:
        cc, national = "33", digits[2:]
    else:
        # Numéro non reconnu : on le renvoie tel quel plutôt que de
        # produire un format invalide.
        return f"+{digits}"

    return _group_uniform(cc, national)


def detect_name_order_from_header(header) -> Optional[str]:
    """Analyse l'en-tête d'une colonne combinée (ex: 'Nom et Prénom' vs
    'Prénom et Nom') pour déterminer l'ordre dans lequel les valeurs
    sont rangées. Retourne 'nom_first', 'prenom_first' ou None si
    l'en-tête ne permet pas de trancher.

    Utilise des frontières de mot (\\bnom\\b) pour ne pas confondre
    'nom' avec le 'nom' contenu dans 'prenom'."""
    if is_blank(header):
        return None
    header_norm = strip_accents_lower(header)
    nom_match = re.search(r"\bnom\b", header_norm)
    prenom_match = re.search(r"\bprenom\b", header_norm)
    if nom_match and prenom_match:
        return (
            "nom_first" if nom_match.start() < prenom_match.start() else "prenom_first"
        )
    return None


def split_combined_name(full_name: str, order_hint: Optional[str] = None) -> tuple:
    """Sépare une colonne 'Nom complet' (ex: 'Arnaud JOSSET') en
    (prenom, nom), utilisée quand aucune colonne nom/prénom séparée
    n'existe.

    Ordre de priorité :
    1) `order_hint` ('nom_first' / 'prenom_first'), déduit du libellé
       de l'en-tête de colonne (ex: 'Nom et Prénom' -> nom en premier,
       'Prénom et Nom' -> prénom en premier) via
       `detect_name_order_from_header` — le plus fiable, car explicite.
    2) à défaut : le NOM est supposé en MAJUSCULES (cas fréquent dans
       les exports listes d'instructeurs/adhérents, ex: 'Arnaud JOSSET').
    3) à défaut : 'Prénom(s) Nom' (dernier mot = nom de famille).
    """
    if is_blank(full_name):
        return "", ""
    words = str(full_name).strip().split()
    if not words:
        return "", ""

    if order_hint == "prenom_first" and len(words) >= 2:
        return " ".join(words[:-1]), words[-1]
    if order_hint == "nom_first" and len(words) >= 2:
        return " ".join(words[1:]), words[0]

    upper_words = [w for w in words if w.isupper() and len(w) > 1]
    if upper_words and len(upper_words) < len(words):
        nom = " ".join(upper_words)
        prenom = " ".join(w for w in words if w not in upper_words)
        return prenom.strip(), nom.strip()
    if len(words) >= 2:
        return " ".join(words[:-1]), words[-1]
    return "", words[0]


ADDRESS_CP_VILLE_REGEX = re.compile(r"(\d{5})\s+(.+)$")

# Détecte un email "glissé" par erreur dans un champ texte libre (ex:
# une adresse postale où l'email a été collé par erreur).
EMAIL_IN_TEXT_REGEX = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")


def extract_email_from_text(text) -> tuple:
    """Si un email est détecté au milieu d'un texte libre (ex: une
    adresse contenant par erreur un email), le retire du texte et le
    retourne séparément. Retourne (texte_nettoye, email_ou_None)."""
    if is_blank(text):
        return text, None
    match = EMAIL_IN_TEXT_REGEX.search(str(text))
    if not match:
        return text, None
    email = match.group(0)
    cleaned = (str(text)[: match.start()] + str(text)[match.end() :]).strip(" ,.;:-")
    return cleaned, email


def split_multiple_addresses(raw_address: str) -> list:
    """Détecte le cas où PLUSIEURS adresses complètes distinctes sont
    enregistrées dans un même champ, séparées par '/', et les sépare
    en plusieurs chaînes.

    Règle volontairement STRICTE : on ne découpe que si CHAQUE partie
    obtenue contient elle-même un code postal à 5 chiffres (preuve
    qu'il s'agit bien de 2 adresses complètes distinctes). Sans cette
    précaution, on découperait à tort des cas très fréquents où '/'
    n'a rien à voir avec plusieurs adresses, par exemple :
      - une plage de numéros de rue : '27/31 Bd Inkermann'
      - un nom de ville contenant un '/' : 'Saint Cirgues/Couze'
      - deux éléments d'un même lieu : 'Rue Bourdelle / Les Bastides du Lac'

    Retourne une liste d'1 élément (adresse inchangée) si aucun
    découpage fiable n'est possible, ou de 2+ éléments sinon."""
    if is_blank(raw_address):
        return []
    text = str(raw_address).strip()

    if len(re.findall(r"\d{5}", text)) < 2:
        # Moins de 2 codes postaux dans le texte -> pas assez de preuve
        # qu'il s'agit de 2 adresses distinctes, on ne découpe pas.
        return [text]

    parts = [p.strip(" ,.-;") for p in text.split("/")]
    parts = [p for p in parts if p]
    if len(parts) >= 2 and all(re.search(r"\d{5}", p) for p in parts):
        return parts
    return [text]


def split_postal_codes(raw) -> list:
    """Détecte un champ CODE POSTAL contenant plusieurs codes postaux
    valides séparés par '/' (ex: '64680 / 64150' dans le fichier
    FNMNS), signe qu'une seule ligne source décrit en réalité
    PLUSIEURS lieux d'exercice pour la même personne (avec des
    colonnes Commune/Téléphone/Email également alignées en parallèle,
    voir `split_aligned`). Beaucoup plus fiable que de se baser sur la
    colonne ville (qui peut légitimement contenir un '/' dans le nom
    même d'une commune, ex: 'Saint Cirgues/Couze') : un code postal
    est purement numérique, donc sans ambiguïté.

    Retourne la liste des codes postaux si détecté (2+ éléments), ou
    une liste à 1 élément (valeur brute inchangée) sinon."""
    if is_blank(raw):
        return [raw]
    text = str(raw).strip()
    if "/" not in text:
        return [text]
    parts = [p.strip() for p in text.split("/")]
    parts = [p for p in parts if p]
    digit_parts = [re.sub(r"\D", "", p) for p in parts]
    if len(parts) >= 2 and all(4 <= len(dp) <= 5 for dp in digit_parts):
        return parts
    return [text]


def split_aligned(raw, n: int) -> list:
    """Découpe une valeur sur '/' pour l'aligner positionnellement
    avec un autre champ déjà détecté comme multi-lieux (ex: la colonne
    Commune découpée en même temps que la colonne Code postal). Si la
    valeur n'a qu'1 seul élément, elle est dupliquée `n` fois (donnée
    partagée entre les `n` lieux, ex: même nom/prénom). Si le nombre
    d'éléments obtenu ne correspond PAS à `n`, on ne tente pas
    l'appariement (risque de mélanger les mauvaises valeurs) : la
    valeur brute complète est dupliquée telle quelle sur les `n`
    lignes générées plutôt que découpée à l'aveugle."""
    if is_blank(raw):
        return [""] * n
    text = str(raw).strip()
    parts = [p.strip(" ,.-;") for p in text.split("/")]
    parts = [p for p in parts if p]
    if len(parts) == n:
        return parts
    if len(parts) == 1:
        return parts * n
    return [text] * n


def split_villes(raw) -> list:
    """Détecte un champ VILLE contenant plusieurs communes distinctes
    séparées par ' / ' AVEC ESPACES autour du slash (ex: 'MONTEILS /
    MARTIELS'), signe qu'une même ligne source décrit plusieurs lieux
    d'exercice sans qu'un code postal distinct n'ait été renseigné
    pour chacun (contrairement au cas géré par `split_postal_codes`).

    Le slash AVEC espaces est volontairement le seul déclencheur : un
    nom de commune peut légitimement contenir un '/' SANS espace
    (ex: 'Saint Cirgues/Couze'), qu'il ne faut surtout pas découper.

    Retourne la liste des communes si détecté (2+ éléments), ou une
    liste à 1 élément (valeur brute inchangée) sinon."""
    if is_blank(raw):
        return [raw]
    text = str(raw).strip()
    if not re.search(r"\s/\s", text):
        return [text]
    parts = [p.strip(" ,.-;") for p in re.split(r"\s/\s", text)]
    parts = [p for p in parts if p]
    if len(parts) >= 2:
        return parts
    return [text]


def parse_composite_address(raw_address: str) -> tuple:
    """Sépare une adresse en texte libre du type
    '69A Rue de La Liberation 68740 FESSENHEIM' en
    (rue, code_postal, ville). Si aucun code postal à 5 chiffres n'est
    détecté, retourne (texte_original, '', '')."""
    if is_blank(raw_address):
        return "", "", ""
    text = str(raw_address).strip()
    match = ADDRESS_CP_VILLE_REGEX.search(text)
    if match:
        cp = match.group(1)
        ville = match.group(2).strip(" ,.-")
        rue = text[: match.start()].strip(" ,.-")
        return rue, cp, ville
    return text, "", ""


EMAIL_REGEX = re.compile(
    r"^[A-Za-z0-9](?:[A-Za-z0-9._%+'\-]*[A-Za-z0-9])?"
    r"@[A-Za-z0-9](?:[A-Za-z0-9.\-]*[A-Za-z0-9])?\.[A-Za-z]{2,}$"
)


def format_email(raw: str) -> Optional[str]:
    """Nettoie/valide un email : minuscule, doit contenir '@' et '.'."""
    if is_blank(raw):
        return None
    candidate = str(raw).strip().lower()
    if EMAIL_REGEX.match(candidate):
        return candidate
    return None


# Ordre canonique d'affichage des modes de contact préférentiels.
MODE_CANONICAL_ORDER = ["whatsapp", "sms", "tel", "email"]

# Mots-clés (déjà normalisés unidecode+lower) associés à chaque mode.
MODE_KEYWORDS = {
    "whatsapp": ["whatsapp", "whats app"],
    "sms": ["sms", "texto", "text message"],
    "tel": ["telephone", "tel", "appel", "phone call", "call"],
    "email": ["email", "e-mail", "mail", "courriel"],
}


def parse_contact_pref(raw) -> list:
    """Extrait les modes de contact explicitement indiqués dans une
    cellule source (ex: 'Téléphone', 'SMS, WhatsApp') et les normalise
    vers l'un de : 'whatsapp', 'sms', 'tel', 'email'."""
    if is_blank(raw):
        return []
    parts = re.split(r"[;,/|\n]+", str(raw))
    found = []
    for part in parts:
        key = strip_accents_lower(part)
        for mode, keywords in MODE_KEYWORDS.items():
            if any(kw in key for kw in keywords) and mode not in found:
                found.append(mode)
    return found


def build_mode_contact_pref(explicit_modes: list, tels: list, emails: list) -> list:
    """Combine les modes explicitement déclarés par la source avec une
    déduction à partir des coordonnées réellement disponibles (si un
    téléphone existe mais qu'aucun mode 'tel/sms/whatsapp' n'est
    explicitement indiqué, on ajoute 'tel' par défaut ; idem 'email').
    Retourne la liste triée selon l'ordre canonique."""
    modes = set(explicit_modes)
    if tels and not modes & {"tel", "sms", "whatsapp"}:
        modes.add("tel")
    if emails and "email" not in modes:
        modes.add("email")
    return [m for m in MODE_CANONICAL_ORDER if m in modes]


# ====================================================================
# 5. STRUCTURE INTERNE D'UN CONTACT NORMALISÉ
# ====================================================================


@dataclass
class Contact:
    nom: str = ""
    prenom: str = ""
    adresse: str = ""
    code_postal: str = ""
    ville: str = ""
    pays: str = ""
    tel: list = field(default_factory=list)
    email: list = field(default_factory=list)
    tag: list = field(default_factory=list)
    mode_contact_pref: list = field(default_factory=list)
    filtres: str = ""
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    source_file: str = ""
    commentaire: str = ""

    @property
    def match_key(self) -> str:
        """Clé normalisée Nom+Prénom (sans accents, minuscule, tirets
        et apostrophes neutralisés) servant au fuzzy matching."""
        return build_match_key(self.nom, self.prenom)

    @property
    def full_address(self) -> str:
        """Adresse complète utilisable pour le géocodage. Retourne une
        chaîne vide si NI l'adresse, NI le code postal, NI la ville ne
        sont connus (le pays seul, ex: 'France', n'est pas assez
        précis pour être géocodé/comparé : cela produirait une
        position quasi aléatoire au centre du pays et fausserait les
        calculs de distance). Un code postal seul (sans rue ni ville)
        est déjà une information de localisation exploitable."""
        if not self.adresse and not self.ville and not self.code_postal:
            return ""
        parts = [self.adresse, self.code_postal.replace(" ", ""), self.ville, self.pays]
        return ", ".join(p for p in parts if p)

    @property
    def has_contact_method(self) -> bool:
        """Un contact est exploitable seulement s'il a au moins un
        téléphone ou un email connu."""
        return bool(self.tel) or bool(self.email)


# ====================================================================
# 6. LECTURE & NORMALISATION DES FICHIERS SOURCES
# ====================================================================


@dataclass
class ReadStats:
    """Compteurs de suivi accumulés pendant la lecture des fichiers,
    utilisés pour construire le rapport final."""

    n_lignes_lues: int = 0  # lignes non vides rencontrées dans les fichiers
    n_sans_identite: int = 0  # rejetées : ni nom ni prénom exploitable
    n_sans_contact: int = 0  # rejetées : ni téléphone ni email exploitable

    def __iadd__(self, other: "ReadStats") -> "ReadStats":
        self.n_lignes_lues += other.n_lignes_lues
        self.n_sans_identite += other.n_sans_identite
        self.n_sans_contact += other.n_sans_contact
        return self


def apply_row_filter(
    df: pd.DataFrame, filter_column: Optional[str], filter_values: Optional[list]
) -> pd.DataFrame:
    """Filtre optionnel appliqué à une feuille AVANT extraction des
    contacts : ne conserve que les lignes où `filter_column` vaut l'une
    des `filter_values` (comparaison insensible à la casse/accents).

    Exemple d'usage réel : ne garder que les personnes disponibles
    pour donner des cours particuliers dans un export d'annuaire
    officiel plus large (colonne 'DonneLeconsPart' == 'Oui').

    Si `filter_column` est vide/None, ou n'existe pas dans cette
    feuille (toutes les feuilles n'ont pas forcément les mêmes
    colonnes), la feuille est retournée inchangée -> le filtre
    s'applique uniquement là où la colonne concernée est présente."""
    if not filter_column or not filter_values:
        return df
    # Recherche de la colonne de façon tolérante (casse/accents/espaces)
    target_norm = strip_accents_lower(filter_column)
    matching_col = next(
        (col for col in df.columns if strip_accents_lower(col) == target_norm), None
    )
    if matching_col is None:
        return df  # colonne absente de cette feuille -> pas de filtre ici
    accepted_norm = {strip_accents_lower(v) for v in filter_values}
    mask = df[matching_col].apply(lambda v: strip_accents_lower(v) in accepted_norm)
    return df[mask]


def _extract_contacts_from_sheet(
    df: pd.DataFrame, source_label: str, stats: Optional[ReadStats] = None
) -> list:
    """Extrait les contacts normalisés d'une seule feuille (DataFrame).
    Si `stats` est fourni, ses compteurs sont incrémentés au passage."""
    df = df.dropna(how="all")
    if df.empty:
        return []
    if stats is not None:
        stats.n_lignes_lues += len(df)
    mapping = detect_columns(df)
    contacts = []

    for _, row in df.iterrows():

        def get(field_name):
            col = mapping.get(field_name)
            return row[col] if col and col in row else None

        nom_raw, prenom_raw = get("nom"), get("prenom")

        # Fallback : colonne "Nom complet" combinée (ex: "Arnaud JOSSET").
        # L'ordre (Nom d'abord ou Prénom d'abord) est déduit en priorité
        # du libellé de l'en-tête de colonne (ex: "Nom et Prénom" vs
        # "Prénom et Nom"), sinon une heuristique par défaut est utilisée.
        if is_blank(nom_raw) and is_blank(prenom_raw):
            combined = get("nom_complet")
            if not is_blank(combined):
                nom_complet_col = mapping.get("nom_complet")
                order_hint = detect_name_order_from_header(nom_complet_col)
                prenom_raw, nom_raw = split_combined_name(combined, order_hint)

        if is_blank(nom_raw) and is_blank(prenom_raw):
            if stats is not None:
                stats.n_sans_identite += 1
            continue  # ligne inexploitable sans identité

        # --- Adresse : colonnes séparées, ou adresse "composite" -------
        adresse_raw = get("adresse")
        cp_raw = get("code_postal")
        ville_raw = get("ville")
        tel_raw = get("tel")
        email_raw = get("email")

        # Il arrive qu'un email soit collé par erreur dans le champ
        # adresse (saisie manuelle malheureuse) : on le détecte, on le
        # retire de l'adresse et on le récupère comme email à part entière.
        email_in_address = None
        if not is_blank(adresse_raw):
            adresse_raw, email_in_address = extract_email_from_text(adresse_raw)

        # --- Détection des lignes MULTI-LIEUX -----------------------
        # Une même ligne source peut décrire plusieurs lieux d'exercice
        # pour une même personne, de deux façons différentes :
        #
        # 1) Colonnes séparées alignées en parallèle par '/' (ex:
        #    fichier FNMNS : Commune='OGEU-LES-BAINS / MOURENX',
        #    Code postal='64680 / 64150', Téléphone='A / B',
        #    Email='x / y'). Détecté via le CODE POSTAL (champ non
        #    ambigu, purement numérique) plutôt que la ville (qui peut
        #    légitimement contenir un '/' dans son nom).
        #
        # 2) Une adresse composite UNIQUE contenant 2 adresses
        #    complètes (ex: '...68740 Fessenheim / ...75000 Paris').
        #
        # Dans les deux cas, chaque lieu détecté devient une ligne à
        # part entière : le reste du pipeline (identité + comparaison
        # de distance Cas A/B/C) se charge ensuite de fusionner ou non
        # ces lignes, exactement comme pour n'importe quel doublon.
        cp_variants = split_postal_codes(cp_raw)
        if len(cp_variants) > 1:
            n = len(cp_variants)
            adresse_variants = split_aligned(adresse_raw, n)
            ville_variants = split_aligned(ville_raw, n)
            tel_variants_raw = split_aligned(tel_raw, n)
            email_variants_raw = split_aligned(email_raw, n)
        else:
            ville_variants_candidate = (
                split_villes(ville_raw) if not is_blank(ville_raw) else [ville_raw]
            )
            if len(ville_variants_candidate) > 1:
                # Plusieurs communes mais un seul code postal renseigné
                # (ex: 'MONTEILS / MARTIELS' avec CP unique '12200') :
                # on découpe quand même la ville, en dupliquant le CP
                # partagé sur chaque ligne générée (elles se
                # re-fusionneront automatiquement plus tard si leur
                # code postal identique déclenche le Cas A).
                n = len(ville_variants_candidate)
                ville_variants = ville_variants_candidate
                adresse_variants = split_aligned(adresse_raw, n)
                cp_variants = split_aligned(cp_raw, n)
                tel_variants_raw = split_aligned(tel_raw, n)
                email_variants_raw = split_aligned(email_raw, n)
            else:
                adresse_variants = (
                    split_multiple_addresses(adresse_raw)
                    if not is_blank(adresse_raw)
                    else [adresse_raw]
                )
                n = len(adresse_variants)
                cp_variants = [cp_raw] * n
                ville_variants = [ville_raw] * n
                tel_variants_raw = [tel_raw] * n
                email_variants_raw = [email_raw] * n

        # --- Coordonnées GPS : volontairement PAS lues depuis le fichier
        # source. Le géocodage est effectué par le script lui-même, une
        # fois le dédoublonnage terminé (voir resolve_duplicates), afin
        # de ne géocoder que les adresses réellement nécessaires.

        tags = split_multi_values(get("tag"))
        explicit_modes = parse_contact_pref(get("mode_contact"))

        for i in range(n):
            variant_cp, variant_ville = cp_variants[i], ville_variants[i]
            variant_rue = adresse_variants[i]
            if not is_blank(variant_rue) and (
                is_blank(variant_cp) or is_blank(variant_ville)
            ):
                rue, cp_parsed, ville_parsed = parse_composite_address(variant_rue)
                variant_rue = rue
                if is_blank(variant_cp):
                    variant_cp = cp_parsed
                if is_blank(variant_ville):
                    variant_ville = ville_parsed

            tels = [
                format_phone_number(t) for t in split_multi_phones(tel_variants_raw[i])
            ]
            emails = [
                format_email(e) for e in split_multi_values(email_variants_raw[i])
            ]
            tels = [t for t in tels if t]
            emails = [e for e in emails if e]

            # Email récupéré depuis le champ adresse (voir plus haut) :
            # rattaché à la 1ère position uniquement (il n'est pas
            # dupliqué sur toutes les lignes générées).
            if i == 0 and email_in_address:
                formatted = format_email(email_in_address)
                if formatted and formatted not in emails:
                    emails.append(formatted)

            # --- Aucun moyen de contact exploitable -> ligne ignorée --
            if not tels and not emails:
                if stats is not None:
                    stats.n_sans_contact += 1
                continue

            modes = build_mode_contact_pref(explicit_modes, tels, emails)

            contact = Contact(
                nom=str(nom_raw).strip().upper() if not is_blank(nom_raw) else "",
                prenom=title_case_fr(prenom_raw) if not is_blank(prenom_raw) else "",
                adresse=title_case_fr(variant_rue) if not is_blank(variant_rue) else "",
                code_postal=format_code_postal(variant_cp),
                ville=title_case_fr(variant_ville)
                if not is_blank(variant_ville)
                else "",
                pays=title_case_fr(get("pays"))
                if not is_blank(get("pays"))
                else "France",
                tel=tels,
                email=emails,
                tag=list(tags),
                mode_contact_pref=modes,
                source_file=source_label,
            )
            contacts.append(contact)

    return contacts


def _preformat_sheet(df: pd.DataFrame, source_label: str) -> list:
    """Variante de `_extract_contacts_from_sheet` dédiée au PRÉ-FORMATAGE
    (nettoyage) d'un fichier source, indépendamment de toute fusion.

    Différences avec l'extraction standard :
    - AUCUNE ligne n'est supprimée (ni faute d'identité, ni faute de
      moyen de contact) : le pré-formatage nettoie, il ne filtre pas.
    - Chaque ligne générée porte un `commentaire` listant les
      corrections/anomalies appliquées, pour audit et cohérence avec
      un pré-formatage réalisé via Gemini (voir PROMPT_GEMINI.md).

    Retourne une liste de dicts avec les colonnes du schéma standard :
    Nom, Prénom, Email, tel, cp, ville, adresse, Pays, tag,
    mode_contact_pref, Commentaire."""
    df = df.dropna(how="all")
    if df.empty:
        return []
    mapping = detect_columns(df)
    rows_out = []

    for _, row in df.iterrows():

        def get(field_name):
            col = mapping.get(field_name)
            return row[col] if col and col in row else None

        comments = []
        nom_raw, prenom_raw = get("nom"), get("prenom")

        if is_blank(nom_raw) and is_blank(prenom_raw):
            combined = get("nom_complet")
            if not is_blank(combined):
                nom_complet_col = mapping.get("nom_complet")
                order_hint = detect_name_order_from_header(nom_complet_col)
                prenom_raw, nom_raw = split_combined_name(combined, order_hint)
                comments.append(
                    f"Nom/Prénom déduits de la colonne combinée '{nom_complet_col}'"
                )

        adresse_raw = get("adresse")
        cp_raw = get("code_postal")
        ville_raw = get("ville")
        tel_raw = get("tel")
        email_raw = get("email")

        email_in_address = None
        if not is_blank(adresse_raw):
            adresse_raw, email_in_address = extract_email_from_text(adresse_raw)
            if email_in_address:
                comments.append("Email déplacé de l'adresse")

        cp_variants = split_postal_codes(cp_raw)
        if len(cp_variants) > 1:
            n = len(cp_variants)
            adresse_variants = split_aligned(adresse_raw, n)
            ville_variants = split_aligned(ville_raw, n)
            tel_variants_raw = split_aligned(tel_raw, n)
            email_variants_raw = split_aligned(email_raw, n)
            comments.append(f"Ligne multi-lieux détectée via code postal (x{n})")
        else:
            ville_variants_candidate = (
                split_villes(ville_raw) if not is_blank(ville_raw) else [ville_raw]
            )
            if len(ville_variants_candidate) > 1:
                n = len(ville_variants_candidate)
                ville_variants = ville_variants_candidate
                adresse_variants = split_aligned(adresse_raw, n)
                cp_variants = split_aligned(cp_raw, n)
                tel_variants_raw = split_aligned(tel_raw, n)
                email_variants_raw = split_aligned(email_raw, n)
                comments.append(f"Ligne multi-lieux détectée via ville (x{n})")
            else:
                adresse_variants = (
                    split_multiple_addresses(adresse_raw)
                    if not is_blank(adresse_raw)
                    else [adresse_raw]
                )
                n = len(adresse_variants)
                if n > 1:
                    comments.append(
                        f"Ligne multi-lieux détectée via adresse composite (x{n})"
                    )
                cp_variants = [cp_raw] * n
                ville_variants = [ville_raw] * n
                tel_variants_raw = [tel_raw] * n
                email_variants_raw = [email_raw] * n

        tags = split_multi_values(get("tag"))
        explicit_modes = parse_contact_pref(get("mode_contact"))

        for i in range(n):
            row_comments = list(comments) if i == 0 else []
            variant_cp, variant_ville = cp_variants[i], ville_variants[i]
            variant_rue = adresse_variants[i]

            if not is_blank(variant_rue) and (
                is_blank(variant_cp) or is_blank(variant_ville)
            ):
                rue, cp_parsed, ville_parsed = parse_composite_address(variant_rue)
                if cp_parsed:
                    row_comments.append(f"Adresse nettoyée (CP {cp_parsed} extrait)")
                variant_rue = rue
                if is_blank(variant_cp):
                    variant_cp = cp_parsed
                if is_blank(variant_ville):
                    variant_ville = ville_parsed

            if not is_blank(variant_cp) and re.fullmatch(
                r"\d[ABab]\d{3}", str(variant_cp).strip()
            ):
                row_comments.append(
                    f"Code postal invalide ignoré (probable code INSEE Corse) : '{variant_cp}'"
                )

            raw_tel_parts = split_multi_phones(tel_variants_raw[i])
            tels = []
            for t in raw_tel_parts:
                formatted = format_phone_number(t)
                if formatted:
                    tels.append(formatted)
                else:
                    row_comments.append(f"Téléphone invalide ignoré : '{t}'")

            raw_email_parts = split_multi_values(email_variants_raw[i])
            emails = []
            for e in raw_email_parts:
                formatted = format_email(e)
                if formatted:
                    emails.append(formatted)
                else:
                    row_comments.append(f"Email invalide ignoré : '{e}'")

            if i == 0 and email_in_address:
                formatted = format_email(email_in_address)
                if formatted and formatted not in emails:
                    emails.append(formatted)

            modes = build_mode_contact_pref(explicit_modes, tels, emails)

            rows_out.append(
                {
                    "Nom": str(nom_raw).strip().upper()
                    if not is_blank(nom_raw)
                    else "",
                    "Prénom": title_case_fr(prenom_raw)
                    if not is_blank(prenom_raw)
                    else "",
                    "Email": LIST_SEP.join(emails),
                    "tel": LIST_SEP.join(tels),
                    "cp": format_code_postal(variant_cp),
                    "ville": title_case_fr(variant_ville)
                    if not is_blank(variant_ville)
                    else "",
                    "adresse": title_case_fr(variant_rue)
                    if not is_blank(variant_rue)
                    else "",
                    "Pays": title_case_fr(get("pays"))
                    if not is_blank(get("pays"))
                    else "France",
                    "tag": LIST_SEP.join(tags),
                    "mode_contact_pref": LIST_SEP.join(modes),
                    "Commentaire": " | ".join(row_comments),
                    "_source": source_label,
                }
            )

    return rows_out


def preformat_file(uploaded_file, first_sheet_only: bool = True) -> pd.DataFrame:
    """Pré-formate UN fichier source brut (quelle que soit sa structure)
    vers le schéma standard (Nom, Prénom, Email, tel, cp, ville,
    adresse, Pays, tag, mode_contact_pref, Commentaire), sans fusion ni
    suppression de lignes. Résultat directement comparable/échangeable
    avec un pré-formatage réalisé via Gemini (voir PROMPT_GEMINI.md)."""
    file_name = getattr(uploaded_file, "name", "fichier")
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names[:1] if first_sheet_only else xl.sheet_names
    all_rows = []
    for sheet_name in sheet_names:
        df = xl.parse(sheet_name, dtype=str)
        label = f"{file_name} / {sheet_name}"
        all_rows.extend(_preformat_sheet(df, label))
    columns = [
        "Nom",
        "Prénom",
        "Email",
        "tel",
        "cp",
        "ville",
        "adresse",
        "Pays",
        "tag",
        "mode_contact_pref",
        "Commentaire",
    ]
    result = pd.DataFrame(all_rows, columns=columns + ["_source"])
    return result


def read_contacts_file(
    uploaded_file,
    first_sheet_only: bool = True,
    stats: Optional[ReadStats] = None,
    filter_column: Optional[str] = None,
    filter_values: Optional[list] = None,
) -> list:
    """Lit un fichier Excel de contacts et retourne la liste des objets
    Contact normalisés.

    Par défaut, seule la PREMIÈRE feuille du classeur est lue
    (`first_sheet_only=True`). Passez `False` pour lire toutes les
    feuilles (utile si un même fichier contient plusieurs feuilles de
    contacts distinctes, ex: une feuille "listing" et une feuille
    "réponses de formulaire"). Si `stats` est fourni, ses compteurs
    sont incrémentés au passage (pour le rapport final).

    `filter_column` / `filter_values` : filtre optionnel, appliqué à
    CHAQUE feuille avant extraction (voir `apply_row_filter`) — utile
    pour reproduire un tri manuel existant, ex: ne garder que les
    lignes où la colonne 'DonneLeconsPart' vaut 'Oui'."""
    file_name = getattr(uploaded_file, "name", "fichier")
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names[:1] if first_sheet_only else xl.sheet_names
    all_contacts = []
    for sheet_name in sheet_names:
        df = xl.parse(sheet_name, dtype=str)
        df = apply_row_filter(df, filter_column, filter_values)
        label = f"{file_name} / {sheet_name}"
        all_contacts.extend(_extract_contacts_from_sheet(df, label, stats))
    return all_contacts


def read_fnmns_members(uploaded_file, first_sheet_only: bool = True) -> list:
    """Lit le fichier optionnel des adhérents FNMNS et retourne la
    liste dédupliquée des clés normalisées 'nom prenom' pour le
    matching.

    Par défaut, seule la PREMIÈRE feuille du classeur est lue
    (`first_sheet_only=True`). Passez `False` pour lire toutes les
    feuilles."""
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names[:1] if first_sheet_only else xl.sheet_names
    keys = set()
    for sheet_name in sheet_names:
        df = xl.parse(sheet_name, dtype=str)
        df = df.dropna(how="all")
        if df.empty:
            continue
        mapping = detect_columns(df)
        nom_col, prenom_col = mapping.get("nom"), mapping.get("prenom")
        if not nom_col and not prenom_col:
            continue  # feuille non pertinente (ex: notes méthodologiques)
        for _, row in df.iterrows():
            nom = row[nom_col] if nom_col and nom_col in row else None
            prenom = row[prenom_col] if prenom_col and prenom_col in row else None
            if is_blank(nom) and is_blank(prenom):
                continue
            key = build_match_key(nom, prenom)
            if key:
                keys.add(key)
    return sorted(keys)


# ====================================================================
# 7. GÉOCODAGE (geopy / Nominatim) AVEC CACHE
# ====================================================================


class Geocoder:
    """Wrapper autour de Nominatim avec cache mémoire et gestion du
    rate-limit imposé par l'API publique OpenStreetMap (1 req/s)."""

    def __init__(self, enabled: bool = True):
        self.enabled = enabled and GEOPY_AVAILABLE
        self._cache: dict = {}
        if self.enabled:
            geolocator = Nominatim(user_agent="fnmns_contacts_app")
            self._geocode = RateLimiter(
                geolocator.geocode,
                min_delay_seconds=1,
                max_retries=2,
                error_wait_seconds=2,
            )

    def geocode(self, address: str):
        """Retourne (latitude, longitude) ou (None, None)."""
        if not self.enabled or not address:
            return None, None
        if address in self._cache:
            return self._cache[address]
        try:
            location = self._geocode(address)
            result = (
                (location.latitude, location.longitude) if location else (None, None)
            )
        except (GeocoderTimedOut, GeocoderServiceError):
            result = (None, None)
        self._cache[address] = result
        return result

    def distance_km(self, coord_a, coord_b) -> Optional[float]:
        if not GEOPY_AVAILABLE:
            return None
        if not all(coord_a) or not all(coord_b):
            return None
        try:
            return geodesic(coord_a, coord_b).km
        except Exception:
            return None


# ====================================================================
# 8. FUZZY MATCHING (identité & FNMNS)
# ====================================================================


def name_similarity(key_a: str, key_b: str) -> float:
    """Score de similarité (0-100) entre deux clés Nom+Prénom
    normalisées, tolérant fautes de frappe/accents."""
    return fuzz.token_sort_ratio(key_a, key_b)


def match_fnmns(contact: Contact, fnmns_keys: list, threshold: int) -> bool:
    """Retourne True si le contact correspond (fuzzy) à un adhérent
    FNMNS."""
    if not fnmns_keys:
        return False
    key = contact.match_key
    if not key:
        return False
    for fnmns_key in fnmns_keys:
        if name_similarity(key, fnmns_key) >= threshold:
            return True
    return False


# ====================================================================
# 9. UNION-FIND (pour le regroupement de doublons)
# ====================================================================


class UnionFind:
    def __init__(self, n: int):
        self.parent = list(range(n))

    def find(self, i: int) -> int:
        while self.parent[i] != i:
            self.parent[i] = self.parent[self.parent[i]]
            i = self.parent[i]
        return i

    def union(self, i: int, j: int):
        ri, rj = self.find(i), self.find(j)
        if ri != rj:
            self.parent[ri] = rj

    def groups(self) -> dict:
        result = {}
        for i in range(len(self.parent)):
            root = self.find(i)
            result.setdefault(root, []).append(i)
        return result


# ====================================================================
# 10. LOGIQUE DE DÉDOUBLONNAGE / FUSION (Section 6 du cahier des charges)
# ====================================================================


def merge_field_lists(contacts: list, attr: str) -> list:
    """Fusionne les listes (tel/email/tag/mode_contact_pref) de
    plusieurs contacts en éliminant les doublons, en conservant
    l'ordre d'apparition."""
    seen, merged = set(), []
    for c in contacts:
        for v in getattr(c, attr):
            norm = strip_accents_lower(v)
            if norm not in seen:
                seen.add(norm)
                merged.append(v)
    return merged


def pick_best_scalar(contacts: list, attr: str) -> str:
    """Retourne la valeur non vide la plus 'complète' (la plus longue)
    parmi les contacts pour un champ scalaire donné."""
    values = [getattr(c, attr) for c in contacts if not is_blank(getattr(c, attr))]
    if not values:
        return ""
    return max(values, key=len)


def fully_merge(contacts: list, reasons: Optional[list] = None) -> Contact:
    """Fusionne intégralement une liste de contacts en une seule ligne
    (utilisé pour les Cas A et C : même adresse/ville, ou distance
    <= seuil).

    `reasons` : liste de chaînes expliquant POURQUOI ces lignes ont été
    jugées identiques (ex: 'même code postal', 'distance 12.3 km ≤
    30 km'), utilisée pour construire le commentaire d'audit de la
    ligne fusionnée."""
    base = contacts[0]
    merged_modes = merge_field_lists(contacts, "mode_contact_pref")

    comment_parts = [f"Fusion de {len(contacts)} lignes sources"]
    if reasons:
        # Dédoublonne les raisons tout en conservant l'ordre d'apparition
        seen, uniq_reasons = set(), []
        for r in reasons:
            if r not in seen:
                seen.add(r)
                uniq_reasons.append(r)
        comment_parts.append("(" + " ; ".join(uniq_reasons) + ")")

    merged = Contact(
        nom=pick_best_scalar(contacts, "nom") or base.nom,
        prenom=pick_best_scalar(contacts, "prenom") or base.prenom,
        adresse=pick_best_scalar(contacts, "adresse"),
        code_postal=pick_best_scalar(contacts, "code_postal"),
        ville=pick_best_scalar(contacts, "ville"),
        pays=pick_best_scalar(contacts, "pays") or "France",
        tel=merge_field_lists(contacts, "tel"),
        email=merge_field_lists(contacts, "email"),
        tag=merge_field_lists(contacts, "tag"),
        mode_contact_pref=[m for m in MODE_CANONICAL_ORDER if m in merged_modes],
        filtres="fnmns" if any(c.filtres == "fnmns" for c in contacts) else "",
        latitude=next((c.latitude for c in contacts if c.latitude is not None), None),
        longitude=next(
            (c.longitude for c in contacts if c.longitude is not None), None
        ),
        source_file=" | ".join(sorted({c.source_file for c in contacts})),
        commentaire=" ".join(comment_parts),
    )
    return merged


def cross_fill_missing(contacts: list, reason: Optional[str] = None) -> list:
    """Pour le Cas B (lignes conservées séparément mais >30km) :
    complète les champs manquants (tel/email/tag/mode_contact_pref/
    filtres) de chaque ligne à partir des autres lignes du même
    cluster nom/prénom, SANS fusionner adresse/ville/coordonnées.

    `reason` : explication (ex: distance calculée) ajoutée au
    commentaire de chaque ligne concernée, pour audit."""
    union_tel = merge_field_lists(contacts, "tel")
    union_email = merge_field_lists(contacts, "email")
    union_tag = merge_field_lists(contacts, "tag")
    union_modes = merge_field_lists(contacts, "mode_contact_pref")
    any_fnmns = any(c.filtres == "fnmns" for c in contacts)
    note = (
        f"Doublon probable ({len(contacts)} lignes) conservé séparément : {reason}"
        if reason
        else f"Doublon probable ({len(contacts)} lignes) conservé séparément "
        f"(adresses distinctes)"
    )

    for c in contacts:
        if not c.tel:
            c.tel = union_tel
        if not c.email:
            c.email = union_email
        if not c.tag:
            c.tag = union_tag
        if not c.mode_contact_pref:
            c.mode_contact_pref = [m for m in MODE_CANONICAL_ORDER if m in union_modes]
        if any_fnmns:
            c.filtres = "fnmns"
        c.commentaire = (c.commentaire + " | " if c.commentaire else "") + note
    return contacts


def resolve_duplicates(
    all_contacts: list,
    geocoder: Geocoder,
    fuzzy_threshold: int,
    distance_threshold_km: float,
    progress_callback=None,
) -> list:
    """Applique l'ensemble des règles de dédoublonnage/fusion
    (Sections 4 à 6 du cahier des charges) et retourne la liste finale
    des contacts consolidés.

    Géocodage : les coordonnées ne sont PAS calculées en amont pour
    toutes les lignes. Elles ne sont récupérées que :
      1) ponctuellement, pendant le tri, pour les paires de lignes du
         même cluster nom/prénom dont il faut comparer la distance
         (décision Cas B vs Cas C) ;
      2) puis, une fois le tri terminé, pour les lignes FINALES
         consolidées qui n'ont pas encore de coordonnées.
    Le cache interne du Geocoder évite toute requête redondante entre
    ces deux étapes.
    """

    n = len(all_contacts)
    if n == 0:
        return []

    # --- Étape 1 : clustering par identité (Nom+Prénom fuzzy) -------
    uf_identity = UnionFind(n)
    for i in range(n):
        for j in range(i + 1, n):
            key_i, key_j = all_contacts[i].match_key, all_contacts[j].match_key
            if not key_i or not key_j:
                continue
            if name_similarity(key_i, key_j) >= fuzzy_threshold:
                uf_identity.union(i, j)
    identity_clusters = uf_identity.groups()

    final_contacts = []
    total_clusters = len(identity_clusters)

    for cluster_idx, (_, indices) in enumerate(identity_clusters.items()):
        cluster_contacts = [all_contacts[i] for i in indices]

        if len(cluster_contacts) == 1:
            final_contacts.append(cluster_contacts[0])
        else:
            # --- Étape 2 : sous-clustering par adresse (Cas A/B/C) ---
            m = len(cluster_contacts)
            uf_address = UnionFind(m)
            # Explique CHAQUE décision (fusion ou séparation), pour
            # construire ensuite les commentaires d'audit.
            pair_reasons = {}
            for i in range(m):
                for j in range(i + 1, m):
                    ci, cj = cluster_contacts[i], cluster_contacts[j]
                    same_adresse = ci.adresse and ci.adresse == cj.adresse
                    same_ville = (
                        ci.ville
                        and cj.ville
                        and strip_accents_lower(ci.ville)
                        == strip_accents_lower(cj.ville)
                    )
                    same_cp = ci.code_postal and ci.code_postal == cj.code_postal

                    if same_adresse or same_ville or same_cp:
                        # Cas A : fusion directe (pas besoin de géocoder).
                        # Un code postal identique est déjà un signal fort
                        # de proximité (même si l'une des deux lignes n'a
                        # pas de rue ou de ville renseignée).
                        criteres = []
                        if same_adresse:
                            criteres.append("même adresse")
                        if same_ville:
                            criteres.append("même ville")
                        if same_cp:
                            criteres.append("même code postal")
                        pair_reasons[(i, j)] = "Cas A : " + " + ".join(criteres)
                        uf_address.union(i, j)
                        continue

                    if not ci.full_address or not cj.full_address:
                        # Au moins une des deux lignes n'a AUCUNE adresse/ville
                        # connue : rien ne prouve qu'il s'agit d'un lieu
                        # différent (ex: 3 fiches "DORGANS Magali" sans
                        # adresse du tout ne doivent pas être considérées
                        # comme 3 personnes à des endroits différents).
                        # On fusionne par défaut (complète les champs
                        # manquants), plutôt que de garder séparé faute de
                        # preuve du contraire.
                        pair_reasons[(i, j)] = (
                            "Fusion par défaut : au moins une ligne sans "
                            "adresse connue (rien ne prouve un lieu différent)"
                        )
                        uf_address.union(i, j)
                        continue

                    # Les deux lignes ont une adresse/ville renseignée mais
                    # différente : on géocode (mis en cache) pour trancher
                    # entre Cas B et Cas C via la distance réelle.
                    if ci.latitude is None or ci.longitude is None:
                        ci.latitude, ci.longitude = geocoder.geocode(ci.full_address)
                    if cj.latitude is None or cj.longitude is None:
                        cj.latitude, cj.longitude = geocoder.geocode(cj.full_address)

                    dist = geocoder.distance_km(
                        (ci.latitude, ci.longitude), (cj.latitude, cj.longitude)
                    )
                    if dist is not None and dist <= distance_threshold_km:
                        # Cas C : fusion (distance <= seuil)
                        pair_reasons[(i, j)] = (
                            f"Cas C : distance {dist:.1f} km "
                            f"≤ seuil {distance_threshold_km} km"
                        )
                        uf_address.union(i, j)
                    elif dist is not None:
                        # Cas B : pas d'union -> lignes distinctes
                        pair_reasons[(i, j)] = (
                            f"Cas B : distance {dist:.1f} km "
                            f"> seuil {distance_threshold_km} km"
                        )
                    else:
                        pair_reasons[(i, j)] = (
                            "Cas B : distance non calculable (géocodage "
                            "indisponible ou adresse introuvable), lignes "
                            "conservées séparées par prudence"
                        )

            address_groups = uf_address.groups()
            sub_merged = []
            for _, sub_indices in address_groups.items():
                sub_contacts = [cluster_contacts[i] for i in sub_indices]
                if len(sub_contacts) == 1:
                    sub_merged.append(sub_contacts[0])
                else:
                    group_reasons = [
                        pair_reasons[(min(a, b), max(a, b))]
                        for idx_a, a in enumerate(sub_indices)
                        for b in sub_indices[idx_a + 1 :]
                        if (min(a, b), max(a, b)) in pair_reasons
                        and not pair_reasons[(min(a, b), max(a, b))].startswith("Cas B")
                    ]
                    sub_merged.append(fully_merge(sub_contacts, group_reasons))

            # --- Étape 3 : Cas B -> complétion croisée des champs ----
            if len(sub_merged) > 1:
                # Note : quand le cluster se scinde en plusieurs sous-groupes
                # (Cas B), on résume simplement le nombre de groupes restants
                # dans le commentaire ; le détail exact des distances entre
                # TOUTES les paires serait trop verbeux pour un cluster à
                # plus de 2 lignes.
                cas_b_examples = [
                    r for r in pair_reasons.values() if r.startswith("Cas B")
                ]
                reason_summary = cas_b_examples[0] if cas_b_examples else None
                sub_merged = cross_fill_missing(sub_merged, reason_summary)

            final_contacts.extend(sub_merged)

        if progress_callback:
            progress_callback(
                0.3 + 0.4 * (cluster_idx + 1) / max(total_clusters, 1),
                f"Dédoublonnage ({cluster_idx + 1}/{total_clusters})",
            )

    # --- Étape 4 : géocodage final, UNE FOIS le tri/la fusion terminés,
    #     et seulement sur les adresses uniques des lignes consolidées
    #     restantes (nombre bien plus faible que le total de départ).
    remaining = [
        c
        for c in final_contacts
        if c.full_address and (c.latitude is None or c.longitude is None)
    ]
    unique_addresses = sorted({c.full_address for c in remaining})
    for idx, addr in enumerate(unique_addresses):
        lat, lon = geocoder.geocode(addr)
        for c in remaining:
            if c.full_address == addr:
                c.latitude, c.longitude = lat, lon
        if progress_callback:
            progress_callback(
                0.7 + 0.28 * (idx + 1) / max(len(unique_addresses), 1),
                f"Géocodage final ({idx + 1}/{len(unique_addresses)})",
            )

    # --- Étape 5 : sécurité finale -> aucune ligne sans moyen de contact
    final_contacts = [c for c in final_contacts if c.has_contact_method]

    return final_contacts


# ====================================================================
# 11. CONSTRUCTION DE LA SORTIE FINALE (JSON = livrable principal)
# ====================================================================


def build_final_records(contacts: list) -> list:
    """Construit la liste de dictionnaires finale (structure imposée,
    Section 3 du cahier des charges), avec de VRAIES listes Python
    pour tel/email/tag/mode_contact_pref (pas de chaînes jointes) —
    c'est la source de vérité unique, utilisée à la fois pour l'export
    JSON et pour l'aperçu tableau."""
    records = []
    used_ids = set()

    for c in contacts:
        base_id = slugify(f"{c.prenom}-{c.nom}")
        contact_id = base_id
        suffix = 2
        while contact_id in used_ids:
            contact_id = f"{base_id}-{suffix}"
            suffix += 1
        used_ids.add(contact_id)

        records.append(
            {
                "id": contact_id,
                "nom": c.nom,
                "prenom": c.prenom,
                "adresse": c.adresse,
                "code_postal": c.code_postal,
                "ville": c.ville,
                "pays": c.pays,
                "tel": c.tel,
                "email": c.email,
                "mode_contact_pref": c.mode_contact_pref,
                "latitude": c.latitude,
                "longitude": c.longitude,
                "tag": c.tag,
                "filtres": c.filtres,
                "commentaire": (
                    (c.commentaire + " | " if c.commentaire else "")
                    + "Correspondance FNMNS trouvée"
                    if c.filtres == "fnmns"
                    else c.commentaire
                ),
            }
        )

    return records


def records_to_json_bytes(records: list) -> bytes:
    """Sérialise les enregistrements en JSON (livrable final).

    Le champ 'commentaire' (audit du dédoublonnage/FNMNS) est
    volontairement EXCLU du JSON : c'est une information de contrôle
    interne, utile pour vérifier le traitement dans le tableau Excel,
    mais qui n'a pas sa place dans le livrable final consommé par
    d'autres systèmes."""
    json_records = [{k: v for k, v in r.items() if k != "commentaire"} for r in records]
    return json.dumps(json_records, ensure_ascii=False, indent=2).encode("utf-8")


def records_to_display_dataframe(records: list) -> pd.DataFrame:
    """Convertit les enregistrements en DataFrame pour l'aperçu à
    l'écran uniquement (les listes y sont jointes par LIST_SEP, car
    un tableau ne peut pas afficher de vraies listes)."""
    rows = []
    for r in records:
        row = dict(r)
        for list_field in ("tel", "email", "mode_contact_pref", "tag"):
            row[list_field] = LIST_SEP.join(row[list_field])
        rows.append(row)
    return pd.DataFrame(rows, columns=FINAL_COLUMNS)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convertit un DataFrame en bytes .xlsx (export secondaire, pour
    consultation/tri rapide dans Excel)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Contacts consolidés")
    return buffer.getvalue()


# ====================================================================
# 11 bis. RAPPORT DE TRAITEMENT
# ====================================================================


def build_processing_report(
    n_lignes_totales: int, final_records: list, elapsed_seconds: float
) -> dict:
    """Construit le rapport de statistiques affiché à la fin du
    traitement (nombre de lignes traitées, complétude des adresses,
    doublons de personnes à adresses multiples, matching FNMNS,
    performance)."""

    n_finales = len(final_records)

    # Complétude adresse/CP/ville (sur les lignes finales consolidées)
    n_sans_adresse = sum(1 for r in final_records if not r["adresse"])
    n_sans_cp = sum(1 for r in final_records if not r["code_postal"])
    n_sans_ville = sum(1 for r in final_records if not r["ville"])
    n_sans_localisation_complete = sum(
        1
        for r in final_records
        if not r["adresse"] or not r["code_postal"] or not r["ville"]
    )
    n_sans_aucune_localisation = sum(
        1
        for r in final_records
        if not r["adresse"] and not r["code_postal"] and not r["ville"]
    )

    # Personnes conservées en plusieurs lignes car adresses distinctes
    # à plus de X km (Cas B du dédoublonnage) : même nom+prénom, mais
    # plusieurs lignes dans le résultat final.
    key_counts = Counter((r["nom"], r["prenom"]) for r in final_records)
    groupes_multi_adresses = {k: v for k, v in key_counts.items() if v > 1}
    n_personnes_multi_adresses = len(groupes_multi_adresses)
    n_lignes_multi_adresses = sum(groupes_multi_adresses.values())

    n_fnmns_final = sum(1 for r in final_records if r["filtres"] == "fnmns")

    temps_moyen_ms = (
        (elapsed_seconds / n_lignes_totales * 1000) if n_lignes_totales else 0.0
    )

    return {
        "lignes_totales_traitees": n_lignes_totales,
        "lignes_sans_adresse": n_sans_adresse,
        "lignes_sans_code_postal": n_sans_cp,
        "lignes_sans_ville": n_sans_ville,
        "lignes_localisation_incomplete": n_sans_localisation_complete,
        "lignes_sans_aucune_localisation": n_sans_aucune_localisation,
        "personnes_avec_adresses_multiples": n_personnes_multi_adresses,
        "lignes_liees_a_adresses_multiples": n_lignes_multi_adresses,
        "personnes_fnmns": n_fnmns_final,
        "lignes_uniques_finales": n_finales,
        "doublons_fusionnes": n_lignes_totales - n_finales,
        "temps_total_secondes": round(elapsed_seconds, 2),
        "temps_moyen_par_ligne_ms": round(temps_moyen_ms, 2),
    }


# ====================================================================
# 12. INTERFACE STREAMLIT
# ====================================================================


def main():
    st.set_page_config(page_title="Fusion & dédoublonnage de contacts", layout="wide")
    st.title("Fusion & dédoublonnage de contacts")

    if not GEOPY_AVAILABLE:
        st.warning(
            "⚠️ Le module `geopy` n'est pas installé : le géocodage et le "
            "calcul de distance seront désactivés. Installez `geopy` via "
            "requirements.txt pour activer cette fonctionnalité."
        )

    if not PHONENUMBERS_AVAILABLE:
        st.warning(
            "⚠️ Le module `phonenumbers` n'est pas installé : le formatage "
            "des numéros de téléphone utilise un repli manuel limité "
            "(indicatifs 3 chiffres non reconnus). Lancez "
            "`pip install phonenumbers` dans le Shell Replit pour un "
            "résultat fiable quel que soit le pays."
        )

    # Navigation entre les trois modes, plus compacte et proche d'une barre d'onglets.
    mode = st.segmented_control(
        "Navigation",
        [
            "🧹 Pré-formatage",
            "🔗 Fusion complète",
            "📄 Excel → JSON",
        ],
        default="🧹 Pré-formatage",
        help="Pré-formatage : nettoie un fichier source brut vers le schéma "
        "standard (Nom, Prénom, Email, tel, cp, ville, adresse, Pays...), "
        "sans fusionner ni supprimer de lignes — équivalent du "
        "pré-formatage réalisable avec Gemini (voir PROMPT_GEMINI.md). "
        "Fusion complète : à utiliser une fois vos fichiers déjà propres "
        "(pré-formatés ici ou via Gemini) — concatène, vérifie le "
        "matching FNMNS et dédoublonne. Excel → JSON : convertit "
        "directement un fichier Excel déjà propre en JSON, sans "
        "fusion ni dédoublonnage ni vérification FNMNS.",
    )

    if mode.startswith("🧹"):
        run_preformat_mode()
    elif mode.startswith("🔗"):
        run_merge_mode()
    else:
        run_excel_to_json_mode()


def run_preformat_mode():
    st.caption(
        "Dépose UN fichier Excel brut : il sera nettoyé (formats nom/prénom/"
        "téléphone/email/adresse/code postal, extraction des adresses "
        "composites et des lignes multi-lieux) et exporté dans le schéma "
        "standard, sans aucune ligne supprimée ni fusionnée."
    )
    raw_file = st.file_uploader(
        "Fichier Excel brut", type=["xlsx", "xls"], key="preformat_file"
    )
    first_sheet_only = st.checkbox(
        "Ne lire que la première feuille du classeur",
        value=True,
        help="Décochez si les données utiles sont réparties sur plusieurs feuilles.",
    )

    if st.button("🧹 Pré-formater", type="primary", disabled=not raw_file):
        with st.spinner("Nettoyage en cours..."):
            df = preformat_file(raw_file, first_sheet_only=first_sheet_only)
        st.session_state["preformat_df"] = df
        st.session_state["preformat_filename"] = getattr(raw_file, "name", "fichier")

    if "preformat_df" in st.session_state:
        df = st.session_state["preformat_df"]
        n_comments = (df["Commentaire"] != "").sum()
        st.success(
            f"✅ {len(df)} lignes générées, dont {n_comments} avec au moins une correction signalée."
        )
        st.dataframe(df.drop(columns=["_source"]), use_container_width=True)

        excel_bytes = dataframe_to_excel_bytes(df.drop(columns=["_source"]))
        source_name = st.session_state.get("preformat_filename", "fichier")
        base_name = re.sub(r"\.[^.]+$", "", source_name) or "fichier"
        preformat_download_name = f"{base_name}-formatte.xlsx"
        st.download_button(
            "⬇️ Télécharger le fichier pré-formaté (.xlsx)",
            excel_bytes,
            preformat_download_name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        if st.button("🗑️ Effacer et recommencer"):
            del st.session_state["preformat_df"]
            st.rerun()


# ====================================================================
# MODE 3 : CONVERSION SIMPLE EXCEL -> JSON (sans fusion/dédoublonnage)
# ====================================================================

# Colonnes reconnues comme des LISTES (jointes par LIST_SEP dans une
# cellule Excel, ex: "06 12 34 56 78; 06 98 76 54 32") : elles seront
# reconverties en vraies listes JSON. Comparaison insensible à la
# casse/aux accents (ex: "Téléphone" ou "tel" matchent tous les deux).
JSON_LIST_COLUMNS = {
    "tel",
    "telephone",
    "email",
    "mail",
    "tag",
    "tags",
    "mode_contact_pref",
}
JSON_FLOAT_COLUMNS = {"latitude", "longitude", "lat", "lng", "lon"}


def excel_to_json_records(uploaded_file, first_sheet_only: bool = True) -> list:
    """Convertit un fichier Excel déjà propre en une liste de dicts JSON,
    SANS fusion, dédoublonnage, ni vérification FNMNS : une simple
    conversion structurelle ligne par ligne.

    - Les colonnes reconnues comme listes (tel/email/tag/mode_contact_pref,
      voir `JSON_LIST_COLUMNS`) sont éclatées en vraies listes JSON via
      le séparateur `LIST_SEP` (ou ';'/',' si le fichier ne vient pas de
      cet outil).
    - Les colonnes latitude/longitude sont converties en nombres (ou
      `None` si vides/non numériques).
    - Toutes les autres colonnes sont conservées telles quelles (texte),
      avec `None` pour les cellules vides.
    - Les noms de colonnes du fichier source sont conservés tels quels
      (aucun renommage) : cette conversion est volontairement générique,
      pas spécifique au schéma interne de cet outil."""
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names[:1] if first_sheet_only else xl.sheet_names
    all_records = []
    for sheet_name in sheet_names:
        df = xl.parse(sheet_name, dtype=str)
        df = df.dropna(how="all")
        for _, row in df.iterrows():
            record = {}
            for col in df.columns:
                raw = row[col]
                col_norm = strip_accents_lower(str(col))
                if is_blank(raw):
                    record[col] = [] if col_norm in JSON_LIST_COLUMNS else None
                elif col_norm in JSON_LIST_COLUMNS:
                    parts = re.split(r"\s*;\s*|\s*,\s*", str(raw).strip())
                    record[col] = [p for p in parts if p]
                elif col_norm in JSON_FLOAT_COLUMNS:
                    try:
                        record[col] = float(str(raw).replace(",", "."))
                    except ValueError:
                        record[col] = None
                else:
                    record[col] = str(raw).strip()
            all_records.append(record)
    return all_records


def run_excel_to_json_mode():
    st.caption(
        "Dépose UN fichier Excel déjà propre (pré-formaté ou fusionné) : "
        "il est converti tel quel en JSON, ligne par ligne, SANS fusion, "
        "dédoublonnage ni vérification FNMNS. Les colonnes tel/email/tag/"
        "mode_contact_pref sont reconstituées en vraies listes JSON."
    )
    excel_file = st.file_uploader(
        "Fichier Excel", type=["xlsx", "xls"], key="excel_to_json_file"
    )
    first_sheet_only = st.checkbox(
        "Ne lire que la première feuille du classeur",
        value=True,
        key="excel_to_json_first_sheet",
    )

    if st.button("📄 Convertir en JSON", type="primary", disabled=not excel_file):
        with st.spinner("Conversion en cours..."):
            records = excel_to_json_records(
                excel_file, first_sheet_only=first_sheet_only
            )
        st.session_state["excel_to_json_records"] = records
        st.session_state["excel_to_json_filename"] = getattr(
            excel_file, "name", "fichier"
        )

    if "excel_to_json_records" in st.session_state:
        records = st.session_state["excel_to_json_records"]
        st.success(f"✅ {len(records)} ligne(s) converties en JSON.")
        st.dataframe(pd.DataFrame(records), use_container_width=True)

        json_bytes = json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8")
        source_name = st.session_state.get("excel_to_json_filename", "fichier")
        base_name = re.sub(r"\.[^.]+$", "", source_name) or "fichier"
        st.download_button(
            "⬇️ Télécharger le fichier JSON",
            json_bytes,
            f"{base_name}.json",
            "application/json",
            type="primary",
        )

        if st.button("🗑️ Effacer et recommencer", key="excel_to_json_clear"):
            del st.session_state["excel_to_json_records"]
            st.rerun()


def run_merge_mode():
    st.caption(
        "Importez jusqu'à 3 fichiers Excel de contacts (déjà pré-formatés, "
        "ou bruts si besoin) et, optionnellement, une base d'adhérents "
        "FNMNS pour enrichir automatiquement le fichier consolidé."
    )

    # ---- Sidebar : réglages avancés --------------------------------
    with st.sidebar:
        st.header("⚙️ Réglages")
        fuzzy_threshold = st.slider(
            "Seuil de similarité Nom+Prénom (%)",
            50,
            100,
            DEFAULT_FUZZY_THRESHOLD,
            help="Au-delà de ce score, deux lignes sont considérées comme la même personne.",
        )
        distance_threshold_km = st.slider(
            "Seuil de distance pour fusion (km)",
            1,
            200,
            DEFAULT_DISTANCE_THRESHOLD_KM,
            help="En dessous de cette distance entre deux adresses différentes, les lignes sont fusionnées (Cas C).",
        )
        enable_geocoding = st.checkbox(
            "Activer le géocodage (latitude/longitude + distance)",
            value=True,
            help="Le géocodage via OpenStreetMap est limité à ~1 requête/seconde : "
            "désactivez-le pour un traitement plus rapide si vous n'avez pas "
            "besoin des coordonnées GPS.",
        )

        st.divider()
        st.subheader("🔍 Filtre sur les fichiers de contacts")
        st.caption(
            "Optionnel : ne garder que les lignes correspondant à un critère "
            "précis (ex: uniquement les personnes disponibles pour des cours "
            "particuliers). Le filtre s'applique à CHAQUE fichier de contacts "
            "possédant la colonne indiquée ; les fichiers qui ne l'ont pas ne "
            "sont pas affectés."
        )
        filter_column = st.text_input(
            "Nom de la colonne à filtrer",
            value="",
            placeholder="ex: DonneLeconsPart",
            help="Laissez vide pour ne filtrer aucune ligne.",
        )
        filter_values_raw = st.text_input(
            "Valeur(s) à conserver (séparées par des virgules)",
            value="",
            placeholder="ex: Oui",
            help="Une ligne est conservée si la colonne ci-dessus correspond à "
            "l'une de ces valeurs (insensible à la casse/accents).",
        )
        filter_values = [v.strip() for v in filter_values_raw.split(",") if v.strip()]

    # ---- Zone 1 : fichiers contacts --------------------------------
    st.subheader("1️⃣ Fichiers clients / contacts (jusqu'à 3 fichiers Excel)")
    contact_files = st.file_uploader(
        "Déposez vos fichiers .xlsx / .xls",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="contact_files",
    )
    if contact_files and len(contact_files) > 3:
        st.error("⚠️ Merci de ne sélectionner que 3 fichiers maximum.")
        contact_files = contact_files[:3]

    # ---- Zone 2 : fichier FNMNS (optionnel) ------------------------
    st.subheader("2️⃣ Fichier Adhérents FNMNS (optionnel)")
    fnmns_file = st.file_uploader(
        "Déposez le fichier .xlsx / .xls des adhérents",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        key="fnmns_file",
    )

    st.divider()
    launch = st.button(
        "🚀 Lancer le traitement", type="primary", disabled=not contact_files
    )

    # --- Calcul : ne s'exécute QUE quand on clique sur le bouton -------
    # Le résultat est stocké dans st.session_state pour ne PAS disparaître
    # lors des prochaines relances du script (Streamlit relance TOUT le
    # script à chaque interaction : slider changé, bouton de téléchargement
    # cliqué, etc. — sans session_state, l'écran de résultat "se réinitialiserait"
    # à chaque fois).
    if launch and contact_files:
        progress_bar = st.progress(0, text="Initialisation...")
        t_start = time.time()

        def update_progress(fraction, label):
            progress_bar.progress(min(max(fraction, 0.0), 1.0), text=label)

        # --- Lecture des fichiers de contacts ------------------------
        all_contacts = []
        for idx, f in enumerate(contact_files):
            update_progress(
                0.05 + 0.15 * idx / len(contact_files), f"Lecture de {f.name}..."
            )
            all_contacts.extend(
                read_contacts_file(
                    f, filter_column=filter_column, filter_values=filter_values
                )
            )

        n_lues = len(all_contacts)

        # --- Lecture et matching FNMNS --------------------------------
        n_matched = 0
        if fnmns_file is not None:
            update_progress(0.22, "Lecture de la base FNMNS...")
            fnmns_keys = read_fnmns_members(fnmns_file)
            update_progress(0.26, "Matching FNMNS en cours...")
            for c in all_contacts:
                if match_fnmns(c, fnmns_keys, fuzzy_threshold):
                    c.filtres = "fnmns"
            n_matched = sum(1 for c in all_contacts if c.filtres == "fnmns")

        # --- Géocodage + dédoublonnage/fusion --------------------------
        geocoder = Geocoder(enabled=enable_geocoding)
        final_contacts = resolve_duplicates(
            all_contacts,
            geocoder=geocoder,
            fuzzy_threshold=fuzzy_threshold,
            distance_threshold_km=distance_threshold_km,
            progress_callback=update_progress,
        )

        # --- Construction du résultat final ------------------------------
        update_progress(0.99, "Construction du résultat final...")
        final_records = build_final_records(final_contacts)
        elapsed_seconds = time.time() - t_start
        report = build_processing_report(n_lues, final_records, elapsed_seconds)
        update_progress(1.0, "Terminé ✅")
        time.sleep(0.3)
        progress_bar.empty()

        # --- Sauvegarde en session_state : survit aux prochaines relances
        st.session_state["final_records"] = final_records
        st.session_state["n_lues"] = n_lues
        st.session_state["n_matched"] = n_matched
        st.session_state["report"] = report
        st.session_state["merged_filename"] = "liste-fnmns"

    elif launch and not contact_files:
        st.error("Merci d'importer au moins un fichier de contacts.")

    # --- Affichage : indépendant du clic, relit toujours session_state -
    # Cette section s'affiche tant qu'un résultat existe en mémoire, que
    # le dernier "rerun" ait été déclenché par le bouton de traitement,
    # un slider, ou le bouton de téléchargement.
    if "final_records" in st.session_state:
        final_records = st.session_state["final_records"]
        display_df = records_to_display_dataframe(final_records)

        st.success(
            f"✅ Traitement terminé : {st.session_state['n_lues']} lignes exploitables → "
            f"{len(final_records)} lignes consolidées dans le JSON final."
        )
        if st.session_state.get("n_matched"):
            st.info(
                f"🏷️ {st.session_state['n_matched']} contact(s) identifié(s) comme adhérent(s) FNMNS."
            )
        st.caption(
            "ℹ️ Les lignes sans aucun téléphone ni email ont été automatiquement "
            "supprimées (aucun moyen de contact exploitable)."
        )

        st.subheader("📊 Aperçu du résultat")
        st.dataframe(display_df, use_container_width=True)

        # --- Rapport de traitement ---------------------------------
        report = st.session_state.get("report")
        if report:
            st.subheader("📋 Rapport de traitement")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Lignes traitées", report["lignes_totales_traitees"])
            col2.metric("Lignes uniques finales", report["lignes_uniques_finales"])
            col3.metric("Doublons fusionnés", report["doublons_fusionnes"])
            col4.metric("Adhérents FNMNS", report["personnes_fnmns"])

            col5, col6, col7, col8 = st.columns(4)
            col5.metric(
                "Localisation incomplète", report["lignes_localisation_incomplete"]
            )
            col6.metric(
                "Sans aucune localisation", report["lignes_sans_aucune_localisation"]
            )
            col7.metric(
                "Personnes à adresses multiples",
                report["personnes_avec_adresses_multiples"],
            )
            col8.metric("Temps total", f"{report['temps_total_secondes']} s")

            with st.expander("🔍 Détail complet du rapport"):
                st.markdown(f"""
- **Lignes totales traitées** (après lecture des fichiers sources) : `{report["lignes_totales_traitees"]}`
- **Lignes sans adresse** : `{report["lignes_sans_adresse"]}`
- **Lignes sans code postal** : `{report["lignes_sans_code_postal"]}`
- **Lignes sans ville** : `{report["lignes_sans_ville"]}`
- **Lignes avec localisation incomplète** (adresse, CP ou ville manquant) : `{report["lignes_localisation_incomplete"]}`
- **Lignes sans aucune information de localisation** (ni adresse, ni CP, ni ville) : `{report["lignes_sans_aucune_localisation"]}`
- **Personnes conservées en plusieurs lignes** car adresses distinctes à plus de {distance_threshold_km} km (Cas B) : `{report["personnes_avec_adresses_multiples"]}` personnes, soit `{report["lignes_liees_a_adresses_multiples"]}` lignes au total
- **Adhérents FNMNS identifiés** : `{report["personnes_fnmns"]}`
- **Lignes uniques dans le résultat final** : `{report["lignes_uniques_finales"]}`
- **Doublons fusionnés** (lignes en moins par rapport au total traité) : `{report["doublons_fusionnes"]}`
- **Temps de traitement total** : `{report["temps_total_secondes"]} s`
- **Temps moyen par ligne traitée** : `{report["temps_moyen_par_ligne_ms"]} ms`
                """)
                rapport_json = json.dumps(report, ensure_ascii=False, indent=2).encode(
                    "utf-8"
                )
                st.download_button(
                    label="⬇️ Télécharger le rapport (.json)",
                    data=rapport_json,
                    file_name="rapport_traitement.json",
                    mime="application/json",
                    key="download_report",
                )

        json_bytes = records_to_json_bytes(final_records)
        st.download_button(
            label="⬇️ Télécharger le fichier final (.json)",
            data=json_bytes,
            file_name=f"{st.session_state.get('merged_filename', 'liste-fnmns')}.json",
            mime="application/json",
            type="primary",
            key="download_json",
        )

        with st.expander(
            "📁 Export secondaire au format .xlsx (pour consultation/tri rapide)"
        ):
            excel_bytes = dataframe_to_excel_bytes(display_df)
            st.download_button(
                label="⬇️ Télécharger la version Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"{st.session_state.get('merged_filename', 'liste-fnmns')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_xlsx",
            )
            st.caption(
                f"Dans ce fichier Excel, les champs listes (tel, email, tag, "
                f"mode_contact_pref) sont joints par `'{LIST_SEP.strip()}'` — "
                f"le fichier .json ci-dessus contient de vraies listes."
            )

        with st.expander("ℹ️ Détail des modes de contact (mode_contact_pref)"):
            st.markdown(
                "Chaque contact reçoit une liste parmi `whatsapp`, `sms`, `tel`, "
                "`email` : les valeurs explicitement indiquées dans les fichiers "
                "sources sont reprises, complétées par défaut avec `tel` si un "
                "numéro existe (sans mode explicite) et `email` si une adresse "
                "email existe (sans mode explicite)."
            )

        if st.button("🗑️ Effacer le résultat et recommencer"):
            del st.session_state["final_records"]
            st.rerun()


if __name__ == "__main__":
    main()
