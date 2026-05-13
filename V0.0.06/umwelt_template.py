"""
EBV Tool — UMWELT canonical parameter schema.

Master list of parameters that drives the layout of the validation
``UMWELT`` echo sheet across all three flows (EBV / Aggressivität / PAK).
Each parsed lab row is mapped to one of these entries by:

    1. Exact match on (Lab_Original_String lowercase, Lab_Unit token)
    2. Synonym match (raw label contains a known canonical substring)
    3. EBV-mapped canonical name when present in the parser output

Unmapped lab rows are appended below the canonical block under
``"Weitere Parameter (lab-spezifisch)"`` so nothing is silently dropped.

Source: ``2604XX_Rohdaten & Aggressivität.xlsx`` UMWELT sheet (rows 5..64).
Schema rows are typed (section header / parameter / blank). Reader code
is expected to walk the list in order.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional


@dataclass(frozen=True)
class UmweltRow:
    """One row in the canonical UMWELT layout.

    Attributes:
        kind: ``"section"``, ``"param"``, or ``"blank"``.
        label: display text in column A.
        unit: display text in column B (empty for sections/blanks).
        matrix: ``"Feststoff"`` / ``"Eluat"`` / ``""`` (empty for unknown).
        match_keys: lowercase substrings used to map parsed lab rows
            to this canonical entry. First match wins; order matters.
        canonical_id: optional EBV/Aggressivität canonical parameter ID
            for cross-reference (e.g. ``"PAK16"``, ``"pH"``, ``"Ca"``).
    """

    kind: str
    label: str
    unit: str = ""
    matrix: str = ""
    match_keys: tuple = ()
    canonical_id: Optional[str] = None


#: Canonical UMWELT row list — read top-to-bottom for sheet writing.
UMWELT_TEMPLATE: List[UmweltRow] = [
    # === Section: PAK16 Feststoff (individual EPA-PAK compounds) ===
    UmweltRow("section", "PAK16 Feststoff (Einzelverbindungen)"),
    UmweltRow("param", "Naphthalin",                  "mg/kg TS", "Feststoff", ("naphthalin",)),
    UmweltRow("param", "Acenaphthylen",               "mg/kg TS", "Feststoff", ("acenaphthylen",)),
    UmweltRow("param", "Acenaphthen",                 "mg/kg TS", "Feststoff", ("acenaphthen",)),
    UmweltRow("param", "Fluoren",                     "mg/kg TS", "Feststoff", ("fluoren",)),
    UmweltRow("param", "Phenanthren",                 "mg/kg TS", "Feststoff", ("phenanthren",)),
    UmweltRow("param", "Anthracen",                   "mg/kg TS", "Feststoff", ("anthracen",)),
    UmweltRow("param", "Fluoranthen",                 "mg/kg TS", "Feststoff", ("fluoranthen",)),
    UmweltRow("param", "Pyren",                       "mg/kg TS", "Feststoff", ("pyren",)),
    UmweltRow("param", "Benzo(a)anthracen",           "mg/kg TS", "Feststoff", ("benzo(a)anthracen", "benzo(a)anthrac", "benz(a)anthrac")),
    UmweltRow("param", "Chrysen",                     "mg/kg TS", "Feststoff", ("chrysen",)),
    UmweltRow("param", "Benzo(b)fluoranthen",         "mg/kg TS", "Feststoff", ("benzo(b)fluoranthen", "benzo(b)fluor")),
    UmweltRow("param", "Benzo(k)fluoranthen",         "mg/kg TS", "Feststoff", ("benzo(k)fluoranthen", "benzo(k)fluor")),
    UmweltRow("param", "Benzo(a)pyren",               "mg/kg TS", "Feststoff", ("benzo(a)pyren", "benzo[a]pyren"), "Benzo(a)pyren"),
    UmweltRow("param", "Dibenz(ah)anthracen",         "mg/kg TS", "Feststoff", ("dibenz(ah)anthracen", "dibenzo(a,h)anthracen", "dibenz(a,h)anthracen", "dibenz(ah)", "dibenzo(a,h)", "dibenz(a,h)")),
    UmweltRow("param", "Benzo(ghi)perylen",           "mg/kg TS", "Feststoff", ("benzo(ghi)perylen", "benzo(g,h,i)perylen", "benzo(ghi)", "benzo(g,h,i)")),
    UmweltRow("param", "Indeno(1,2,3-cd)pyren",       "mg/kg TS", "Feststoff", ("indeno(1,2,3-cd)pyren", "indeno[1,2,3-cd]pyren", "indeno(1,2,3", "indeno[1,2,3")),
    UmweltRow("param", "Summe PAK (16) nach EBV",     "mg/kg TS", "Feststoff", ("summe pak (16)", "summe pak 16", "pak16", "pak nach epa", "summe pak epa"), "PAK16"),

    UmweltRow("blank", ""),
    # === Section: Eluat-only parameters ===
    UmweltRow("section", "Eluat / wasserbasierte Parameter"),
    UmweltRow("param", "pH-Wert (Eluat)",             "-",        "Eluat",     ("ph-wert", "ph wert"), "pH-Wert"),
    UmweltRow("param", "Elektrische Leitfähigkeit",   "µS/cm",    "Eluat",     ("leitfähigkeit",), "Elektrische Leitfähigkeit"),
    UmweltRow("param", "Sulfat (Eluat)",              "mg/l",     "Eluat",     ("sulfat",), "Sulfat"),
    UmweltRow("param", "Phenolindex",                 "mg/l",     "Eluat",     ("phenol",), "Phenolindex"),
    UmweltRow("param", "Säurekapazität bis pH 4,3",   "mmol/l",   "Eluat",     ("säurekapazität", "ks 4,3", "ks4,3"), "KS43"),
    UmweltRow("param", "Calcium (Ca)",                "mg/l",     "Eluat",     ("calcium",), "Ca"),
    UmweltRow("param", "Magnesium (Mg)",              "mg/l",     "Eluat",     ("magnesium",), "Mg"),
    UmweltRow("param", "Ammonium (NH4)",              "mg/l",     "Eluat",     ("ammonium",), "NH4"),
    UmweltRow("param", "Chlorid (Cl)",                "mg/l",     "Eluat",     ("chlorid",), "Cl"),
    UmweltRow("param", "Kalklösende Kohlensäure",     "mg/l",     "Eluat",     ("kalklösend", "co2 angreif", "kohlensäure"), "CO2_angr"),
    UmweltRow("param", "Sulfid leicht freisetzbar",   "mg/l",     "Eluat",     ("sulfid",), "S2"),

    UmweltRow("blank", ""),
    # === Section: Probenvorbereitung + Schwermetalle (Feststoff) ===
    UmweltRow("section", "Probenvorbereitung & Schwermetalle (Feststoff)"),
    UmweltRow("param", "Siebung < 2 mm",              "",         "",          ("siebung",)),
    UmweltRow("param", "Trockenmasse (TM)",           "%",        "",          ("trockenmasse",)),
    UmweltRow("param", "Königswasseraufschluss",      "",         "",          ("königswasser", "knigswasser")),
    UmweltRow("param", "Aufschlussfaktor KÖWA",       "",         "",          ("aufschlussfaktor",)),
    UmweltRow("param", "Arsen",                       "mg/kg TS", "Feststoff", ("arsen",), "Arsen"),
    UmweltRow("param", "Blei",                        "mg/kg TS", "Feststoff", ("blei",), "Blei"),
    UmweltRow("param", "Cadmium",                     "mg/kg TS", "Feststoff", ("cadmium",), "Cadmium"),
    UmweltRow("param", "Chrom",                       "mg/kg TS", "Feststoff", ("chrom",), "Chrom, gesamt"),
    UmweltRow("param", "Kupfer",                      "mg/kg TS", "Feststoff", ("kupfer",), "Kupfer"),
    UmweltRow("param", "Nickel",                      "mg/kg TS", "Feststoff", ("nickel",), "Nickel"),
    UmweltRow("param", "Quecksilber",                 "mg/kg TS", "Feststoff", ("quecksilber",), "Quecksilber"),
    UmweltRow("param", "Thallium",                    "mg/kg TS", "Feststoff", ("thallium",), "Thallium"),
    UmweltRow("param", "Zink",                        "mg/kg TS", "Feststoff", ("zink",), "Zink"),
    UmweltRow("param", "TOC",                         "% TS",     "Feststoff", ("toc",), "TOC"),
    UmweltRow("param", "Kohlenwasserstoffe C10 - C22","mg/kg TS", "Feststoff", ("c10 - c22", "c10-c22", "c10 c22"), "Kohlenwasserstoffe (C10-C22)"),
    UmweltRow("param", "Kohlenwasserstoffe C10 - C40","mg/kg TS", "Feststoff", ("c10 - c40", "c10-c40", "c10 c40", "kw-index"), "Kohlenwasserstoffe (C10-C40)"),

    UmweltRow("blank", ""),
    # === Section: PAK Eluat (individual + sum) ===
    UmweltRow("section", "PAK Eluat (Einzelverbindungen + Summen)"),
    UmweltRow("param", "Naphthalin (Eluat)",          "µg/l",     "Eluat",     ("naphthalin",)),
    UmweltRow("param", "1-Methylnaphthalin",          "µg/l",     "Eluat",     ("1-methylnaphth", "1 methylnaphth", "methylnaphthalin")),
    UmweltRow("param", "2-Methylnaphthalin",          "µg/l",     "Eluat",     ("2-methylnaphth", "2 methylnaphth")),
    UmweltRow("param", "Summe PAK (15) nach EBV",     "µg/l",     "Eluat",     ("summe pak (15)", "summe pak 15", "pak15"), "PAK15"),
    UmweltRow("param", "Summe Naphthaline (EBV)",     "µg/l",     "Eluat",     ("summe naphthal", "summe naphtha"), "Naphthalin und Methylnaphthaline, gesamt"),

    UmweltRow("blank", ""),
    # === Section: PCB Feststoff (congeners + sum) ===
    UmweltRow("section", "PCB Feststoff (Einzelkongenere + Summe)"),
    UmweltRow("param", "PCB Nr. 28",                  "mg/kg TS", "Feststoff", ("pcb nr. 28", "pcb 28", "pcb-28")),
    UmweltRow("param", "PCB Nr. 52",                  "mg/kg TS", "Feststoff", ("pcb nr. 52", "pcb 52", "pcb-52")),
    UmweltRow("param", "PCB Nr. 101",                 "mg/kg TS", "Feststoff", ("pcb nr. 101", "pcb 101", "pcb-101")),
    UmweltRow("param", "PCB Nr. 118",                 "mg/kg TS", "Feststoff", ("pcb nr. 118", "pcb 118", "pcb-118")),
    UmweltRow("param", "PCB Nr. 138",                 "mg/kg TS", "Feststoff", ("pcb nr. 138", "pcb 138", "pcb-138")),
    UmweltRow("param", "PCB Nr. 153",                 "mg/kg TS", "Feststoff", ("pcb nr. 153", "pcb 153", "pcb-153")),
    UmweltRow("param", "PCB Nr. 180",                 "mg/kg TS", "Feststoff", ("pcb nr. 180", "pcb 180", "pcb-180")),
    UmweltRow("param", "Summe PCB nach EBV",          "mg/kg TS", "Feststoff", ("summe pcb",), "PCB6 und PCB-118"),

    UmweltRow("blank", ""),
    # === Section: weitere EBV-relevante Eluat-Parameter ===
    UmweltRow("section", "Schwermetalle (Eluat)"),
    UmweltRow("param", "Arsen (Eluat)",               "µg/l",     "Eluat",     ("arsen",), "Arsen"),
    UmweltRow("param", "Blei (Eluat)",                "µg/l",     "Eluat",     ("blei",), "Blei"),
    UmweltRow("param", "Cadmium (Eluat)",             "µg/l",     "Eluat",     ("cadmium",), "Cadmium"),
    UmweltRow("param", "Chrom (Eluat)",               "µg/l",     "Eluat",     ("chrom",), "Chrom, gesamt"),
    UmweltRow("param", "Kupfer (Eluat)",              "µg/l",     "Eluat",     ("kupfer",), "Kupfer"),
    UmweltRow("param", "Nickel (Eluat)",              "µg/l",     "Eluat",     ("nickel",), "Nickel"),
    UmweltRow("param", "Quecksilber (Eluat)",         "µg/l",     "Eluat",     ("quecksilber",), "Quecksilber"),
    UmweltRow("param", "Thallium (Eluat)",            "µg/l",     "Eluat",     ("thallium",), "Thallium"),
    UmweltRow("param", "Zink (Eluat)",                "µg/l",     "Eluat",     ("zink",), "Zink"),
    UmweltRow("param", "EOX",                         "mg/kg TS", "Feststoff", ("eox",), "EOX"),
]


def is_feststoff_unit(unit: str) -> bool:
    """Return True if ``unit`` indicates a Feststoff measurement.

    Args:
        unit: lab-reported unit string (e.g. ``"mg/kg TS"``).

    Returns:
        True for ``mg/kg``, ``mg/kg TS``, ``%``, ``Vol.-%``, ``M.-%``,
        ``M%``, ``M.-%``; False otherwise.
    """
    u = (unit or "").lower().replace(" ", "")
    return ("kg" in u) or ("vol" in u) or ("m%" in u) or ("m.-%" in u) or (u.endswith("%") and "/" not in u)


def is_eluat_unit(unit: str) -> bool:
    """Return True if ``unit`` indicates an Eluat measurement.

    Args:
        unit: lab-reported unit string (e.g. ``"mg/l"``).

    Returns:
        True for ``mg/l``, ``µg/l``, ``mmol/l``, ``mol/m³``, ``µS/cm``;
        False otherwise.
    """
    u = (unit or "").lower().replace(" ", "")
    return ("/l" in u) or ("ph" in u and u.endswith("ph")) or ("mol/m" in u) or ("s/cm" in u) or u in {"-", "--"}


def match_lab_row(
    label: str,
    unit: str,
    ebv_param: str = "",
) -> int:
    """Map a parsed lab row to a canonical UMWELT row index.

    Args:
        label: raw lab row text (Lab_Original_String).
        unit: parsed Lab_Unit.
        ebv_param: optional EBV canonical name (parser's mapping) — used
            as a fallback when the raw-text scan doesn't hit any match_key.

    Returns:
        Index into :data:`UMWELT_TEMPLATE` matching the lab row, or -1 if
        no canonical entry fits. Matrix-aware: if the parsed unit clearly
        indicates Feststoff or Eluat, prefer the canonical entry with the
        same matrix.
    """
    label_lower = (label or "").lower().strip()
    is_fest = is_feststoff_unit(unit)
    is_elu = is_eluat_unit(unit)

    matrix_pref = "Feststoff" if is_fest else ("Eluat" if is_elu else "")

    # Pass 1: EXACT label equality (case-insensitive), matrix-preferred.
    # Wins over any substring match — fixes "Benzo(k)fluoranthen" matching
    # the shorter "fluoranthen" key of the Fluoranthen row.
    for i, row in enumerate(UMWELT_TEMPLATE):
        if row.kind != "param":
            continue
        if matrix_pref and row.matrix and row.matrix != matrix_pref:
            continue
        for key in row.match_keys:
            if key and key == label_lower:
                return i

    # Pass 2: EXACT label equality, any matrix.
    for i, row in enumerate(UMWELT_TEMPLATE):
        if row.kind != "param":
            continue
        for key in row.match_keys:
            if key and key == label_lower:
                return i

    # Pass 3: substring match, matrix-preferred (fallback).
    for i, row in enumerate(UMWELT_TEMPLATE):
        if row.kind != "param":
            continue
        if matrix_pref and row.matrix and row.matrix != matrix_pref:
            continue
        for key in row.match_keys:
            if key and key in label_lower:
                return i

    # Pass 4: substring match, any matrix (last resort).
    for i, row in enumerate(UMWELT_TEMPLATE):
        if row.kind != "param":
            continue
        for key in row.match_keys:
            if key and key in label_lower:
                return i

    # Third pass: by EBV canonical name
    if ebv_param:
        ebv_lower = ebv_param.lower()
        for i, row in enumerate(UMWELT_TEMPLATE):
            if row.kind != "param":
                continue
            if row.canonical_id and row.canonical_id.lower() == ebv_lower:
                if not matrix_pref or not row.matrix or row.matrix == matrix_pref:
                    return i

    return -1
