"""
EBV Tool — Aggressivität config (water-only, Phase 2 build).

Encodes the parameter catalogue and synonym map required for
:mod:`pdf_parser_aggressivität` to recognise DIN 4030-1 (Beton-Wasser) and
DIN 50929-3 (Korrosion-Wasser) parameters in standard German lab reports
(AGROLAB, SGS, …). Threshold tables (XA1..XA3, N1..N7, M1..M6) are NOT
encoded here yet — those land in a follow-up module once the user confirms
the exact values against the company workbook
``2604XX_Rohdaten & Aggressivität.xlsx`` sheets ``Beton_Wasser`` and
``Korrosion_Wasser``.

Parameter ID convention (kept stable across modules):

    "pH"             — pH-Wert
    "Mg"             — Magnesium  [mg/l]
    "Ca"             — Calcium    [mg/l]
    "NH4"            — Ammonium   [mg/l]
    "SO4"            — Sulfat     [mg/l]
    "Cl"             — Chlorid    [mg/l]
    "CO2_angr"       — Kalklösende / angreifende Kohlensäure  [mg/l]
    "S2"             — Sulfid leicht freisetzbar              [mg/l]
    "KS43"           — Säurekapazität bis pH 4,3              [mmol/l]
    "Leitfaehigkeit" — Elektrische Leitfähigkeit              [µS/cm]

Lab-reported "Betonaggressivität (Angriffsgrad DIN 4030)" is captured as
``"Lab_DIN4030_assessment"`` for cross-check against the tool's own DIN
4030-1 classification.
"""
from __future__ import annotations

from typing import Dict, List, Tuple

#: Canonical Aggressivität parameter catalogue (water-only).
#: Order matters for the UMWELT echo + Beton_Wasser/Korrosion_Wasser output row.
WATER_AGGR_PARAMETERS: List[Tuple[str, str, str]] = [
    # (canonical_id, display_name, default_unit)
    ("pH",              "pH-Wert",                              "-"),
    ("Leitfaehigkeit",  "Elektrische Leitfähigkeit",            "µS/cm"),
    ("KS43",            "Säurekapazität bis pH 4,3",            "mmol/l"),
    ("Ca",              "Calcium (Ca)",                         "mg/l"),
    ("Mg",              "Magnesium (Mg)",                       "mg/l"),
    ("NH4",             "Ammonium (NH4)",                       "mg/l"),
    ("Cl",              "Chlorid (Cl)",                         "mg/l"),
    ("SO4",             "Sulfat (SO4)",                         "mg/l"),
    ("CO2_angr",        "Kalklösende Kohlensäure",              "mg/l"),
    ("S2",              "Sulfid leicht freisetzbar",            "mg/l"),
    ("Faerbung",        "Färbung",                              "-"),
    ("Geruch",          "Geruch",                               "-"),
    ("Truebung",        "Trübung",                              "-"),
    ("Oxidierbarkeit",  "Oxidierbarkeit (KMnO4-Verbrauch)",     "mg/l"),
    ("KMnO4_Index",     "KMnO4-Index (als O2)",                 "mg/l"),
    ("NO3",             "Nitrat (NO3)",                         "mg/l"),
    ("Carbonathaerte",  "Carbonathärte",                        "mg/l CaO"),
    ("Nichtcarbonathaerte", "Nichtcarbonathärte",               "mg/l CaO"),
    ("Gesamthaerte",    "Gesamthärte",                          "mg/l CaO"),
    ("Gesamthaerte_mmol","Gesamthärte (Summe Erdalkalien)",     "mmol/l"),
    ("Ca_molm3",        "Calcium (mol/m³)",                     "mol/m³"),
    ("Neutralsalze",    "Neutralsalze",                         "mol/m³"),
]

#: Fuzzy/exact synonym map (lowercase keys) -> canonical_id.
#: Both direct lab-row text and stripped variants are listed because lab
#: reports interleave (Methode-string) suffixes that we strip in the parser.
WATER_AGGR_SYNONYMS: Dict[str, str] = {
    # pH
    "ph-wert (labor)": "pH",
    "ph-wert": "pH",
    "ph wert": "pH",
    "ph": "pH",

    # Leitfähigkeit
    "leitfähigkeit bei 20 °c (labor)": "Leitfaehigkeit",
    "leitfähigkeit bei 25 °c (labor)": "Leitfaehigkeit",
    "leitfähigkeit bei 20°c": "Leitfaehigkeit",
    "leitfähigkeit bei 25°c": "Leitfaehigkeit",
    "elektrische leitfähigkeit": "Leitfaehigkeit",
    "leitfähigkeit": "Leitfaehigkeit",

    # KS 4,3
    "säurekapazität bis ph 4,3": "KS43",
    "säurekapazität ks 4,3": "KS43",
    "säurekapazität ks4,3": "KS43",
    "ks 4,3": "KS43",
    "ks4,3": "KS43",

    # Calcium
    "calcium (ca)": "Ca",
    "calcium": "Ca",
    "ca": "Ca",
    "ca2+": "Ca",

    # Magnesium
    "magnesium (mg)": "Mg",
    "magnesium": "Mg",
    "mg2+": "Mg",

    # Ammonium
    "ammonium (nh4)": "NH4",
    "ammonium": "NH4",
    "nh4+": "NH4",
    "nh4": "NH4",

    # Chlorid
    "chlorid (cl)": "Cl",
    "chlorid": "Cl",
    "cl-": "Cl",

    # Sulfat
    "sulfat (so4)": "SO4",
    "sulfat": "SO4",
    "so42-": "SO4",
    "so4": "SO4",

    # Angreifende Kohlensäure
    "kalkl. kohlensäure": "CO2_angr",
    "kalklösende kohlensäure": "CO2_angr",
    "kohlensäure angreifend": "CO2_angr",
    "co2 angreifend": "CO2_angr",
    "co2-angreifend": "CO2_angr",

    # Sulfid
    "sulfid leicht freisetzbar": "S2",
    "sulfid": "S2",
    "s2-": "S2",

    # Färbung / Geruch / Trübung (sensorisch)
    "färbung (labor)": "Faerbung",
    "färbung": "Faerbung",
    "geruch (labor)": "Geruch",
    "geruch": "Geruch",
    "trübung (labor)": "Truebung",
    "trübung": "Truebung",

    # Oxidierbarkeit + KMnO4-Index
    "oxidierbarkeit (kmno4-verbrauch)": "Oxidierbarkeit",
    "oxidierbarkeit": "Oxidierbarkeit",
    "kmno4-index (als o2)": "KMnO4_Index",
    "kmno4-index": "KMnO4_Index",
    "kmno4 index": "KMnO4_Index",

    # Nitrat
    "nitrat (no3)": "NO3",
    "nitrat": "NO3",
    "no3-": "NO3",
    "no3": "NO3",

    # Carbonat / Nichtcarbonat / Gesamthärte
    "carbonathärte": "Carbonathaerte",
    "nichtcarbonathärte": "Nichtcarbonathaerte",
    "gesamthärte (summe erdalkalien)": "Gesamthaerte_mmol",
    "gesamthärte": "Gesamthaerte",

    # Calcium mol/m³ (computed)
    "calcium mol/m³": "Ca_molm3",

    # Neutralsalze (computed)
    "neutralsalze": "Neutralsalze",
    "neutralssalze": "Neutralsalze",
}

#: Parameter rows we always ignore — non-aggressivität noise on lab reports.
WATER_AGGR_BLACKLIST: List[str] = [
    "geruchsart", "geruchsstärke",
    "temperatur",
    "marmorlöse",
]

#: Lab-reported "Betonaggressivität (Angriffsgrad DIN 4030)" capture key.
LAB_DIN4030_KEY: str = "Lab_DIN4030_assessment"


# ---------------------------------------------------------------------------
# DIN 4030-1 Beton-Wasser thresholds (encoded from sheet "Beton_Wasser",
# rows 5..11 of 2604XX_Rohdaten & Aggressivität.xlsx).
#
# Per-parameter exposure classes follow strict-interval semantics:
#   * XA0 → no attack: measured value below XA1 lower bound (or for pH,
#                       above XA1 upper bound, since pH risk goes downward).
#   * XA1, XA2, XA3 → ascending attack severity, each bounded.
#   * "Milieu unstimmig" → above XA3 upper bound (or below for pH).
#
# Each entry is a dict whose values are tuples
#     (lower, upper, lower_inclusive, upper_inclusive)
# with ``lower`` / ``upper`` possibly None to mean unbounded.
# ---------------------------------------------------------------------------

#: Sentinel meaning "unbounded on this side".
_INF = float("inf")
_NINF = float("-inf")

#: Direction of attack per parameter. ``"high"`` means higher values → worse
#: corrosion (Mg, NH4, SO4, CO2 angr). ``"low"`` means lower values → worse
#: corrosion (pH only).
DIN4030_DIRECTION: Dict[str, str] = {
    "pH": "low",
    "Mg": "high",
    "NH4": "high",
    "SO4": "high",
    "CO2_angr": "high",
}

#: Beton-Wasser threshold intervals per parameter and exposure class.
#: Each entry: (lower, upper, lower_inclusive, upper_inclusive). Unbounded
#: bounds use ``_NINF`` / ``_INF``. A value matches the class iff
#: lower < v < upper (with inclusivity flags applied).
DIN4030_THRESHOLDS: Dict[str, Dict[str, Tuple[float, float, bool, bool]]] = {
    "pH": {
        # pH: lower bound first (more acidic = more aggressive)
        # XA1: 5,5 ≤ pH ≤ 6,5
        "XA1": (5.5, 6.5, True, True),
        # XA2: 4,5 ≤ pH < 5,5
        "XA2": (4.5, 5.5, True, False),
        # XA3: 4,0 ≤ pH < 4,5
        "XA3": (4.0, 4.5, True, False),
        # Milieu unstimmig: pH < 4,0
        "Milieu unstimmig": (_NINF, 4.0, False, False),
    },
    "Mg": {
        "XA1": (300.0, 1000.0, True, True),
        "XA2": (1000.0, 3000.0, False, True),
        "XA3": (3000.0, _INF, False, False),
        # No "Milieu unstimmig" defined for Mg (XA3 = Sättigung).
    },
    "NH4": {
        "XA1": (15.0, 30.0, True, True),
        "XA2": (30.0, 60.0, False, True),
        "XA3": (60.0, 100.0, False, True),
        "Milieu unstimmig": (100.0, _INF, False, False),
    },
    "SO4": {
        "XA1": (200.0, 600.0, True, True),
        "XA2": (600.0, 3000.0, False, True),
        "XA3": (3000.0, 6000.0, False, True),
        "Milieu unstimmig": (6000.0, _INF, False, False),
    },
    "CO2_angr": {
        "XA1": (15.0, 40.0, True, True),
        "XA2": (40.0, 100.0, False, True),
        "XA3": (100.0, _INF, False, False),
        # No "Milieu unstimmig" defined for CO2_angr.
    },
    # Sulfid (S2-) carries no DIN 4030-1 threshold in this workbook (row 11
    # is "-" placeholders). Not included here; engine skips it.
}

#: Order of class severity for aggregating overall XA class per sample.
DIN4030_CLASS_HIERARCHY: List[str] = ["XA0", "XA1", "XA2", "XA3", "Milieu unstimmig"]


# ---------------------------------------------------------------------------
# DIN 50929-3 Korrosion-Wasser N/M lookup tables (encoded from sheet
# "Korrosion_Wasser" rows 7..44 of the same workbook).
#
# Convention: each rating digit Nx / Mx is computed by walking the lookup
# table for parameter x and finding the row whose interval contains the
# measured value (or whose categorical key matches). All N values apply to
# unlegierter Stahl; M values to feuerverzinkter Stahl.
#
# Interval form: list of (lower, upper, lower_inclusive, upper_inclusive,
#                          N_value, M_value).
# Categorical form: dict {category_key: (N_value, M_value)}.
# ---------------------------------------------------------------------------

#: N1/M1 — Wasserart (categorical).
N1_WASSERART: Dict[str, Tuple[float, float]] = {
    "fließend":  (0.0,  -2.0),
    "stehend":   (-1.0,  1.0),
    "Küste":     (-3.0, -3.0),
    "anaerob":   (-5.0, -5.0),
}

#: N2/M2 — Lage des Objektes (categorical).
N2_LAGE: Dict[str, Tuple[float, float]] = {
    "Unterwasserbereich":           (0.0,  0.0),
    "Wasser/Luft-Wechselbereich":   (1.0, -6.0),
    "Spritzwasserbereich":          (0.3, -2.0),
}

#: N3/M3 — c(Cl-) + 2·c(SO4²-) in mol/m³ (interval).
N3_SALT_LOAD: List[Tuple[float, float, bool, bool, float, float]] = [
    (_NINF, 1.0,   False, False,  0.0,  0.0),
    (1.0,   5.0,   True,  False, -2.0,  0.0),
    (5.0,   25.0,  True,  False, -4.0, -1.0),
    (25.0,  100.0, True,  False, -6.0, -2.0),
    (100.0, 300.0, True,  False, -7.0, -3.0),
    (300.0, _INF,  False, False, -8.0, -4.0),
]

#: N4/M4 — KS 4,3 (Säurekapazität bis pH 4,3) in mmol/l (interval).
N4_KS43: List[Tuple[float, float, bool, bool, float, float]] = [
    (_NINF, 1.0,  False, False, 1.0, -1.0),
    (1.0,   2.0,  True,  False, 2.0,  1.0),
    (2.0,   4.0,  True,  False, 3.0,  1.0),
    (4.0,   6.0,  True,  False, 4.0,  0.0),
    (6.0,   _INF, False, False, 5.0, -1.0),
]

#: N5/M5 — c(Ca²⁺) in mol/m³ (interval).
N5_CALCIUM: List[Tuple[float, float, bool, bool, float, float]] = [
    (_NINF, 0.5,  False, False, -1.0, 0.0),
    (0.5,   2.0,  True,  False,  0.0, 2.0),
    (2.0,   8.0,  True,  False,  1.0, 3.0),
    (8.0,   _INF, False, False,  2.0, 4.0),
]

#: N6/M6 — pH-Wert (interval).
N6_PH: List[Tuple[float, float, bool, bool, float, float]] = [
    (_NINF, 5.5,  False, False, -3.0, -6.0),
    (5.5,   6.5,  True,  False, -2.0, -4.0),
    (6.5,   7.0,  True,  False, -1.0, -1.0),
    (7.0,   7.5,  True,  False,  0.0,  1.0),
    (7.5,   _INF, False, False,  1.0,  1.0),
]

#: N7 — Objekt-/Wasser-Potential U_h in volts (interval). No M7 — only
#: applies to free-corrosion assessment for unlegierter Stahl.
N7_POTENTIAL: List[Tuple[float, float, bool, bool, float]] = [
    (-0.2, -0.1, True,  False, -2.0),
    (-0.1,  0.0, True,  False, -5.0),
    ( 0.0, _INF, False, False, -8.0),
]


# ---------------------------------------------------------------------------
# Bucket labels for W0/W1 and WD/WL outputs (DIN 50929-3, rows 7..10 / 14..18).
# ---------------------------------------------------------------------------

#: (W-value lower-inclusive bound, label, flaechen_label).
W_BUCKETS: List[Tuple[float, str, str]] = [
    ( 0.0, "sehr gering", "sehr gering"),
    (-4.0, "gering",      "sehr gering"),
    (-8.0, "mittel",      "gering"),
    (_NINF, "hoch",       "mittel"),
]

#: WD/WL Deckschicht-Güte buckets.
WD_BUCKETS: List[Tuple[float, str]] = [
    ( 0.0, "sehr gut"),
    (-4.0, "gut"),
    (-8.0, "befriedigend"),
    (_NINF, "nicht ausreichend"),
]

#: Corrosion-rate buckets for W (mm/a Abtragsrate, mm/a Eindringrate).
W_RATE_BUCKETS: List[Tuple[float, float, float]] = [
    ( 0.0, 0.01, 0.05),
    (-4.0, 0.02, 0.10),
    (-8.0, 0.05, 0.20),
    (_NINF, 0.10, 0.50),
]


# ---------------------------------------------------------------------------
# Unit conversions used by the Korrosion-Wasser engine.
# ---------------------------------------------------------------------------

#: Molar masses (g/mol) for converting mg/l → mol/m³ where the company
#: workbook column expects mol/m³ (N3 salt sum, N5 calcium).
MOLAR_MASS: Dict[str, float] = {
    "Cl":  35.45,   # Cl-
    "SO4": 96.06,   # SO4²-
    "Ca":  40.08,   # Ca²+
}


def mgL_to_molM3(mg_per_l: float, species: str) -> float:
    """Convert mg/l → mol/m³ for the named species.

    Args:
        mg_per_l: lab-reported concentration in milligrams per litre.
        species: canonical ID present in :data:`MOLAR_MASS` (Cl / SO4 / Ca).

    Returns:
        Equivalent molar concentration in mol/m³ (= mmol/l numerically).

    Raises:
        KeyError: if ``species`` has no molar-mass entry.
    """
    return float(mg_per_l) / MOLAR_MASS[species]
