"""
EBV Tool — RuVA-StB 01 PAK evaluator (Berlin Fassung 2018).

Classifies a single Straßenaufbruch sample against the threshold table
encoded in :mod:`config_pak`. Returns a structured :class:`PakResult`
that the reporter can render verbatim.

Inputs are the canonical PAK measurements extracted at Step 1 by
:mod:`pdf_parser` (the same lab-PDF parser used for EBV — the PAK Step 1
branch reuses it for now). Per-sample required parameters:

    PAK16         Feststoff   mg/kg TS
    Benzo(a)pyren Feststoff   mg/kg TS    (only for hazardous-waste check)
    Phenolindex   Eluat       mg/l

Sample-Type: the lab reports under standard EBV format; the parser also
captures individual PAK compounds + the aggregate. The aggregate
``"Summe PAK ... nach EPA"`` is the PAK16 source.

Strict type hints + comprehensive docstrings per project convention.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from config_pak import (
    PAK_CLASS_HIERARCHY,
    RUVA_HAZARDOUS_TRIGGERS,
    RUVA_TABELLE_1,
)


@dataclass(frozen=True)
class PakResult:
    """RuVA-StB 01 (Berlin Fassung 2018) classification of one sample.

    Attributes:
        klasse: one of ``"A"`` / ``"B"`` / ``"C"`` / ``"Gefährlicher Abfall"`` /
            ``"Nicht klassifizierbar"`` (last only when required inputs are
            absent — the engine cannot decide).
        pak16_value: PAK16 (mg/kg TS), or None if not measured.
        bap_value: Benzo(a)pyren (mg/kg TS), or None.
        phenol_value: Phenolindex (mg/l Eluat), or None.
        triggers: list of parameter IDs whose value crossed a hazardous-waste
            trigger (empty when ``klasse != "Gefährlicher Abfall"``).
        driving_parameters: PAK16 and/or Phenolindex names whose interval
            placement determined the class assignment.
        verwertungsverfahren: free-text description of allowed reuse
            mechanism per RuVA-StB 01 Abschnitt 4 (e.g. ``"Abschnitt 4.1"``,
            ``"kein (Entsorgung)"``).
        notes: free-form annotations (missing inputs, edge cases).
    """

    klasse: str
    pak16_value: Optional[float]
    bap_value: Optional[float]
    phenol_value: Optional[float]
    triggers: List[str]
    driving_parameters: List[str]
    verwertungsverfahren: str
    notes: List[str] = field(default_factory=list)


def _in_interval(
    value: float,
    interval: Tuple[float, float, bool, bool],
) -> bool:
    """Return True iff ``value`` lies inside ``interval``.

    Args:
        value: numeric value to test.
        interval: ``(lower, upper, lower_inclusive, upper_inclusive)``.

    Returns:
        True if value is inside, False otherwise. Inf-bound on either
        side disables that side's check.
    """
    lo, hi, lo_inc, hi_inc = interval
    lo_ok = (value >= lo) if lo_inc else (value > lo)
    hi_ok = (value <= hi) if hi_inc else (value < hi)
    return lo_ok and hi_ok


def evaluate_pak(
    pak16: Optional[float],
    benzo_a_pyren: Optional[float],
    phenolindex: Optional[float],
) -> PakResult:
    """Classify one Straßenaufbruch sample per RuVA-StB 01 (Berlin Fassung 2018).

    Decision order:

        1. Hazardous-waste check — any of (PAK16 > 100, Benzo(a)pyren > 50,
           Phenolindex > 50) escalates to "Gefährlicher Abfall" regardless of
           other parameters.
        2. Otherwise: walk classes A → B → C; first class whose PAK16 AND
           Phenolindex intervals both match wins.
        3. If neither matches and the sample is not hazardous waste,
           return ``"Nicht klassifizierbar"`` with the relevant value
           details — happens when a required input is missing or when the
           combined values fall outside every interval (very rare given
           the table is exhaustive over the real line).

    Args:
        pak16: PAK16 (nach EPA) Feststoff concentration in mg/kg TS.
        benzo_a_pyren: Benzo(a)pyren Feststoff in mg/kg TS (only used for
            the hazardous-waste check).
        phenolindex: Phenolindex Eluat in mg/l.

    Returns:
        Populated :class:`PakResult`.
    """
    notes: List[str] = []
    triggers: List[str] = []

    # 1. Hazardous-waste triggers
    if pak16 is not None and pak16 > RUVA_HAZARDOUS_TRIGGERS["PAK16"][0]:
        triggers.append("PAK16")
    if benzo_a_pyren is not None and benzo_a_pyren > RUVA_HAZARDOUS_TRIGGERS["Benzo(a)pyren"][0]:
        triggers.append("Benzo(a)pyren")
    if phenolindex is not None and phenolindex > RUVA_HAZARDOUS_TRIGGERS["Phenolindex"][0]:
        triggers.append("Phenolindex")

    if triggers:
        return PakResult(
            klasse="Gefährlicher Abfall",
            pak16_value=pak16,
            bap_value=benzo_a_pyren,
            phenol_value=phenolindex,
            triggers=triggers,
            driving_parameters=triggers,
            verwertungsverfahren="Sonderabfallentsorgung (Abfallschlüssel 170301*)",
            notes=notes,
        )

    # 2. Sufficient inputs check
    if pak16 is None:
        notes.append("PAK16 nicht vorhanden — Klassifizierung unvollständig.")
    if phenolindex is None:
        notes.append("Phenolindex nicht vorhanden — Klassifizierung unvollständig.")
    if pak16 is None or phenolindex is None:
        return PakResult(
            klasse="Nicht klassifizierbar",
            pak16_value=pak16,
            bap_value=benzo_a_pyren,
            phenol_value=phenolindex,
            triggers=[],
            driving_parameters=[],
            verwertungsverfahren="—",
            notes=notes,
        )

    # 3. Walk classes A → B → C; first matching wins.
    verwertung_map: Dict[str, str] = {
        "A": "Abschnitt 4.1 (oder ausnahmsweise 4.2 / 4.3)",
        "B": "kein (Entsorgung)",
        "C": "kein (Entsorgung)",
    }
    for cls in ("A", "B", "C"):
        intervals = RUVA_TABELLE_1[cls]
        pak_ok = _in_interval(pak16, intervals["PAK16"])
        phen_ok = _in_interval(phenolindex, intervals["Phenolindex"])
        if pak_ok and phen_ok:
            # Driving parameter: the one that pushed us OUT of the previous
            # class. For class A everything is in-spec → both are "drivers
            # of admission". For B/C the parameter that just crossed the
            # threshold is the driver.
            if cls == "A":
                drivers = []  # nothing is "driving" — material is acceptable
            else:
                prev_intervals = RUVA_TABELLE_1[PAK_CLASS_HIERARCHY[PAK_CLASS_HIERARCHY.index(cls) - 1]]
                drivers = []
                if not _in_interval(pak16, prev_intervals["PAK16"]):
                    drivers.append("PAK16")
                if not _in_interval(phenolindex, prev_intervals["Phenolindex"]):
                    drivers.append("Phenolindex")
            return PakResult(
                klasse=cls,
                pak16_value=pak16,
                bap_value=benzo_a_pyren,
                phenol_value=phenolindex,
                triggers=[],
                driving_parameters=drivers,
                verwertungsverfahren=verwertung_map[cls],
                notes=notes,
            )

    # 4. Fallback: values cross above C's upper bound but below hazardous
    # triggers — shouldn't be reachable given thresholds, but cap to C with
    # a note for transparency.
    notes.append("Werte außerhalb der A/B/C-Intervalle — als C eingestuft.")
    return PakResult(
        klasse="C",
        pak16_value=pak16,
        bap_value=benzo_a_pyren,
        phenol_value=phenolindex,
        triggers=[],
        driving_parameters=["PAK16", "Phenolindex"],
        verwertungsverfahren=verwertung_map["C"],
        notes=notes,
    )
