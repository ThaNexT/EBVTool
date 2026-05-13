"""
EBV Tool — RuVA-StB 01 PAK classification config (Berlin Fassung 2018).

Source: ``230331_UmweltBerlin-EinfuehrungRuvaStb01.pdf`` — Amtsblatt für
Berlin Nr. 07/2018 S. 900, "Einführung RuVA-StB 01, Ausgabe 2001, Fassung
2005" published by Senatsverwaltung für Umwelt, Verkehr und Klimaschutz
Berlin (07.02.2018).

The Berlin Fassung supersedes the original RuVA-StB 01 Tabelle 1 with the
threshold table encoded below. The federal A1 class (PAK ≤ 10) is dropped
per ARS Nr. 29/2004.

Classification (from Tabelle 1, Abschnitt 4 amendment):

    +---------+-----------+---------------+---------------------------+
    | Klasse  | PAK16 TS  | Phenol Eluat  | Verwertung                |
    |         | [mg/kg]   | [mg/l]        |                           |
    +---------+-----------+---------------+---------------------------+
    | A       | ≤ 25      | ≤ 0,1         | Abschnitt 4.1 (4.2/4.3)   |
    | B       | > 25,     | ≤ 0,1         | kein (Entsorgung)         |
    |         | ≤ 100     |               |                           |
    | C       | > 25,     | > 0,1,        | kein (Entsorgung)         |
    |         | ≤ 100     | ≤ 50          |                           |
    +---------+-----------+---------------+---------------------------+

Per point 2 of the Berlin amendment, the following thresholds elevate the
material to gefährlicher Abfall (Abfallschlüssel 170301*), bypassing the
A/B/C scheme entirely:

* PAK16 (nach EPA) > 100 mg/kg TS
* Benzo[a]pyren    > 50  mg/kg TS
* Phenolindex      > 50  mg/l   (Eluat)

This module is the data source consumed by ``evaluator_pak.py``. Strict
type hints + verbatim values per project convention.
"""
from __future__ import annotations

from typing import Dict, List, Tuple

#: Sentinel constants for unbounded interval edges.
_INF: float = float("inf")
_NINF: float = float("-inf")

#: Canonical PAK class hierarchy (low → high severity).
PAK_CLASS_HIERARCHY: List[str] = ["A", "B", "C", "Gefährlicher Abfall"]

#: Threshold intervals for the RuVA-StB 01 Tabelle 1 (Berlin Fassung 2018).
#: Each entry is keyed by class and contains the bounded intervals for
#: ``PAK16`` (Feststoff, mg/kg TS) and ``Phenolindex`` (Eluat, mg/l).
#: Interval form: (lower, upper, lower_inclusive, upper_inclusive). Unbounded
#: bounds use ``_NINF`` / ``_INF``. Both intervals must hold for the class
#: to apply (logical AND).
RUVA_TABELLE_1: Dict[str, Dict[str, Tuple[float, float, bool, bool]]] = {
    "A": {
        "PAK16":       (_NINF,  25.0, False, True),
        "Phenolindex": (_NINF,   0.1, False, True),
    },
    "B": {
        "PAK16":       (25.0,  100.0, False, True),
        "Phenolindex": (_NINF,   0.1, False, True),
    },
    "C": {
        "PAK16":       (25.0,  100.0, False, True),
        "Phenolindex": ( 0.1,   50.0, False, True),
    },
}

#: Hazardous-waste thresholds per point 2 of the Berlin amendment.
#: Any single trigger above its threshold classifies the material as
#: "Gefährlicher Abfall" (Abfallschlüssel 170301*).
RUVA_HAZARDOUS_TRIGGERS: Dict[str, Tuple[float, str]] = {
    "PAK16":         (100.0, "mg/kg"),  # > 100 mg/kg TS PAK16
    "Benzo(a)pyren": ( 50.0, "mg/kg"),  # > 50 mg/kg Benzo[a]pyren
    "Phenolindex":   ( 50.0, "mg/l"),   # > 50 mg/l Phenolindex Eluat
}

#: Synonym map (lowercase) → canonical PAK parameter ID used in the engine.
PAK_SYNONYMS: Dict[str, str] = {
    "summe pak (16) nach ebv": "PAK16",
    "summe pak (16)": "PAK16",
    "summe pak 16": "PAK16",
    "pak 16 (epa)": "PAK16",
    "pak nach epa": "PAK16",
    "pak 16": "PAK16",
    "pak16": "PAK16",

    "benzo(a)pyren": "Benzo(a)pyren",
    "benzo[a]pyren": "Benzo(a)pyren",
    "bap": "Benzo(a)pyren",

    "phenolindex": "Phenolindex",
    "phenol-index": "Phenolindex",
    "phenol index": "Phenolindex",
    "phenole": "Phenolindex",
}

#: Reverse map: canonical ID → display label used in reports.
PAK_DISPLAY: Dict[str, str] = {
    "PAK16":         "PAK16 (nach EPA)",
    "Benzo(a)pyren": "Benzo(a)pyren",
    "Phenolindex":   "Phenolindex (Eluat)",
}

#: Hex fills for the per-class colour scheme on the PAK output cell.
PAK_CLASS_FILL_HEX: Dict[str, str] = {
    "A":                   "C6EFCE",  # green
    "B":                   "FFEB9C",  # yellow
    "C":                   "FFC7CE",  # red
    "Gefährlicher Abfall": "9C0006",  # dark red
}
