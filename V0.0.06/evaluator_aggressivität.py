"""
EBV Tool — Aggressivität evaluator (water-only, Phase 3 build).

Two engines:

* :func:`evaluate_beton_wasser` — DIN 4030-1:2024-07 concrete-attack class
  per parameter and overall, returning a :class:`Din4030Result`.
* :func:`evaluate_korrosion_wasser` — DIN 50929-3:2024-05 steel-corrosion
  rating per Bewertungsziffer Nx/Mx, with derived sum-formulas W0/W1
  (unlegiert) and WD/WL (verzinkt), returning a :class:`Din50929WaterResult`.

Both engines operate on a single sample at a time. Caller is expected to
provide the canonical-ID measurement dict (output of
:mod:`pdf_parser_aggressivität` after Step 1 validation).

Strict type hints + docstrings per project convention. All threshold and
N/M tables live in :mod:`config_aggressivität`; this module contains the
classification logic only.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from config_aggressivität import (
    DIN4030_CLASS_HIERARCHY,
    DIN4030_DIRECTION,
    DIN4030_THRESHOLDS,
    N1_WASSERART,
    N2_LAGE,
    N3_SALT_LOAD,
    N4_KS43,
    N5_CALCIUM,
    N6_PH,
    N7_POTENTIAL,
    W_BUCKETS,
    WD_BUCKETS,
    W_RATE_BUCKETS,
    mgL_to_molM3,
)


# ---------------------------------------------------------------------------
# Result dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Din4030Result:
    """Concrete-attack classification per DIN 4030-1.

    Attributes:
        per_parameter_class: mapping ``canonical_id -> exposure-class label``
            (one of ``"XA0"`` / ``"XA1"`` / ``"XA2"`` / ``"XA3"`` /
            ``"Milieu unstimmig"``).
        overall_class: highest class observed across measured parameters
            (``"XA0"`` when nothing crosses the XA1 threshold).
        missing: canonical IDs that were not measured (excluded from
            overall).
        notes: free-form annotations (e.g. parameters skipped because the
            DIN table has no threshold for them).
    """

    per_parameter_class: Dict[str, str]
    overall_class: str
    missing: List[str]
    notes: List[str] = field(default_factory=list)


@dataclass(frozen=True)
class Din50929WaterResult:
    """Steel-corrosion rating per DIN 50929-3 (water exposure).

    Attributes:
        n_values: rating digits Nx for unlegierter Stahl (keys ``"N1".."N7"``).
            Missing parameters absent from the dict (engine treats as 0
            contribution to W0/W1).
        m_values: rating digits Mx for feuerverzinkter Stahl.
        W0: sum-formula Gleichung (7) for unlegierter Stahl.
        W1: sum-formula Gleichung (8) for unlegierter Stahl with
            water/luft-Wechselbereich (only when N2 is present and
            applicable).
        WD: sum-formula for verzinkter Stahl (Deckschichtgüte).
        WL: WD + M2 (verzinkt with Wechselbereich).
        class_W0: bucket label for W0 ("sehr gering" / "gering" / "mittel" / "hoch").
        class_W1: bucket label for W1.
        flaechen_class_W0: Flächenkorrosion bucket label for W0.
        flaechen_class_W1: Flächenkorrosion bucket label for W1.
        class_WD: Deckschicht-Güte bucket for WD.
        class_WL: Deckschicht-Güte bucket for WL.
        rate_W0_mm_per_a: Abtragsrate at W0 (mm/a).
        rate_W0_max_mm_per_a: max Eindringrate at W0 (mm/a).
        rate_W1_mm_per_a: Abtragsrate at W1 (mm/a).
        rate_W1_max_mm_per_a: max Eindringrate at W1 (mm/a).
        missing: canonical IDs that were not measured (engine still
            computed W with zero contribution from them).
        notes: free-form annotations.
    """

    n_values: Dict[str, float]
    m_values: Dict[str, float]
    W0: int
    W1: int
    WD: int
    WL: int
    class_W0: str
    class_W1: str
    flaechen_class_W0: str
    flaechen_class_W1: str
    class_WD: str
    class_WL: str
    rate_W0_mm_per_a: float
    rate_W0_max_mm_per_a: float
    rate_W1_mm_per_a: float
    rate_W1_max_mm_per_a: float
    missing: List[str]
    notes: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


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


def _lookup_interval_table(
    table: List[Tuple[float, float, bool, bool, float, float]],
    value: float,
) -> Tuple[float, float]:
    """Look up ``(N_value, M_value)`` for ``value`` in an N/M interval table.

    Args:
        table: list of (lower, upper, lower_inclusive, upper_inclusive,
            N_value, M_value) rows. Must be exhaustive across the real
            line.
        value: measured concentration / pH / etc.

    Returns:
        ``(N_value, M_value)`` from the matching row.

    Raises:
        ValueError: if no row matches (table not exhaustive).
    """
    for row in table:
        lo, hi, lo_inc, hi_inc, n_val, m_val = row
        if _in_interval(value, (lo, hi, lo_inc, hi_inc)):
            return float(n_val), float(m_val)
    raise ValueError(f"No matching interval for value {value!r}")


def _lookup_n7(value: float) -> Optional[float]:
    """Look up N7 (Objekt-/Wasser-Potenzial) for ``value``.

    Args:
        value: object/water potential in volts.

    Returns:
        Float N7 value, or None if value falls below the table (≤ -0.2 V
        — N7 not assessable per DIN 50929-3 footnote).
    """
    for lo, hi, lo_inc, hi_inc, n_val in N7_POTENTIAL:
        if _in_interval(value, (lo, hi, lo_inc, hi_inc)):
            return float(n_val)
    return None


def _bucket_W(w: float) -> Tuple[str, str]:
    """Bucket a W0/W1 value to (Mulden/Lochkorrosion, Flächenkorrosion) labels.

    Args:
        w: computed W0 or W1.

    Returns:
        ``(mulden_label, flaechen_label)`` per DIN 50929-3 Tabelle 7.
    """
    for threshold, mulden, flaechen in W_BUCKETS:
        if w >= threshold:
            return mulden, flaechen
    return "hoch", "mittel"


def _bucket_WD(wd: float) -> str:
    """Bucket a WD/WL value to a Deckschichtgüte label.

    Args:
        wd: computed WD or WL.

    Returns:
        Deckschichtgüte label per DIN 50929-3 Tabelle 7.
    """
    for threshold, label in WD_BUCKETS:
        if wd >= threshold:
            return label
    return "nicht ausreichend"


def _rate_W(w: float) -> Tuple[float, float]:
    """Look up corrosion-rate (Abtragsrate, max Eindringrate) per W bucket.

    Args:
        w: computed W0 or W1.

    Returns:
        ``(abtragsrate_mm_per_a, max_eindringrate_mm_per_a)`` per
        DIN 50929-3 Tabelle 8.
    """
    for threshold, rate_w, rate_max in W_RATE_BUCKETS:
        if w >= threshold:
            return float(rate_w), float(rate_max)
    return 0.10, 0.50


# ---------------------------------------------------------------------------
# Engine 1 — DIN 4030-1 Beton-Wasser
# ---------------------------------------------------------------------------


def _classify_din4030_single(
    canonical_id: str,
    value: float,
) -> str:
    """Classify a single parameter against DIN 4030-1 thresholds.

    Args:
        canonical_id: parameter ID (must be a key of
            :data:`DIN4030_THRESHOLDS`).
        value: measured value in the parameter's reference unit.

    Returns:
        Class label: ``"XA0"`` / ``"XA1"`` / ``"XA2"`` / ``"XA3"`` /
        ``"Milieu unstimmig"``.

    Notes:
        For ``"pH"`` the attack direction is inverted (lower pH → worse).
        The XA0 fallback applies when the value is *less aggressive* than
        the XA1 threshold for that parameter.
    """
    thresholds = DIN4030_THRESHOLDS[canonical_id]
    # Walk the explicit intervals; first match wins. Walk in order
    # XA1 → XA2 → XA3 → Milieu unstimmig so that increasing severity
    # is checked in the right direction.
    for cls in ("XA1", "XA2", "XA3", "Milieu unstimmig"):
        if cls in thresholds and _in_interval(value, thresholds[cls]):
            return cls
    return "XA0"


def evaluate_beton_wasser(
    measurements: Dict[str, Optional[float]],
) -> Din4030Result:
    """Run the DIN 4030-1 Beton-Wasser engine on one sample.

    Args:
        measurements: mapping canonical_id → value (or None when not
            measured). Must use the IDs from
            :data:`config_aggressivität.WATER_AGGR_PARAMETERS`.

    Returns:
        Populated :class:`Din4030Result`. Unmeasured parameters land in
        ``missing`` and do not contribute to ``overall_class``.

    Notes:
        Sulfid (S2-) carries no threshold in this company workbook and is
        always added to ``notes`` rather than ``per_parameter_class`` —
        the engine cannot classify it.
    """
    per_param: Dict[str, str] = {}
    missing: List[str] = []
    notes: List[str] = []

    for canonical_id in DIN4030_DIRECTION:
        value = measurements.get(canonical_id)
        if value is None:
            missing.append(canonical_id)
            continue
        per_param[canonical_id] = _classify_din4030_single(canonical_id, float(value))

    # Sulfid is in measurements but has no threshold — note it.
    if "S2" in measurements and measurements["S2"] is not None:
        notes.append("S²⁻ ohne Grenzwert in DIN 4030-1; vorhanden als Information.")

    # Overall = max severity over measured params
    severity = max(
        (DIN4030_CLASS_HIERARCHY.index(cls) for cls in per_param.values()),
        default=0,
    )
    overall = DIN4030_CLASS_HIERARCHY[severity]
    return Din4030Result(
        per_parameter_class=per_param,
        overall_class=overall,
        missing=missing,
        notes=notes,
    )


# ---------------------------------------------------------------------------
# Engine 2 — DIN 50929-3 Korrosion-Wasser
# ---------------------------------------------------------------------------


def evaluate_korrosion_wasser(
    measurements: Dict[str, Optional[float]],
    *,
    wasserart: str = "stehend",
    objektlage: Optional[str] = None,
    u_potential: Optional[float] = None,
) -> Din50929WaterResult:
    """Run the DIN 50929-3 Korrosion-Wasser engine on one sample.

    Computes Bewertungsziffern N1..N7 and M1..M6 from the measurement
    table, then assembles the sum-formulas:

        W0 = round(N1 + N3 + N4 + N5 + N6 + N3/N4)   (Gl. 7)
        W1 = W0 − N1 + N2·N3                          (Gl. 8)
        WD = M1 + M3 + M4 + M5 + M6
        WL = WD + M2

    Args:
        measurements: canonical-ID → measured value. Required IDs for
            full computation: ``pH``, ``KS43``, ``Ca``, ``Cl``, ``SO4``.
            Missing values are treated as zero contribution and reported
            in ``missing``.
        wasserart: one of the keys of
            :data:`config_aggressivität.N1_WASSERART`. Defaults to
            ``"stehend"`` which corresponds to most groundwater
            assessments. Override per sample as appropriate.
        objektlage: one of the keys of
            :data:`config_aggressivität.N2_LAGE`. If None, N2/M2 are
            treated as zero (no Wasser/Luft transition; W0 == W1 in that
            case, WD == WL).
        u_potential: object/water potential in volts (Cu/CuSO₄ reference).
            If provided and within the N7 table, contributes N7 to W0/W1.

    Returns:
        Populated :class:`Din50929WaterResult`.

    Raises:
        KeyError: if ``wasserart`` or ``objektlage`` are not in the
            company lookup tables.
    """
    n_values: Dict[str, float] = {}
    m_values: Dict[str, float] = {}
    missing: List[str] = []
    notes: List[str] = []

    # N1 / M1 — Wasserart (categorical, required)
    n1, m1 = N1_WASSERART[wasserart]
    n_values["N1"] = n1
    m_values["M1"] = m1

    # N2 / M2 — Lage des Objektes (categorical, optional)
    if objektlage is not None:
        n2, m2 = N2_LAGE[objektlage]
        n_values["N2"] = n2
        m_values["M2"] = m2
    else:
        n_values["N2"] = 0.0
        m_values["M2"] = 0.0
        notes.append("Objektlage nicht angegeben — N2/M2 = 0 (W0 ≡ W1, WD ≡ WL).")

    # N3 / M3 — c(Cl-) + 2·c(SO4²-) in mol/m³
    cl_mgL = measurements.get("Cl")
    so4_mgL = measurements.get("SO4")
    if cl_mgL is None and so4_mgL is None:
        missing.extend(["Cl", "SO4"])
        n_values["N3"] = 0.0
        m_values["M3"] = 0.0
    else:
        cl_mol = mgL_to_molM3(cl_mgL or 0.0, "Cl") if cl_mgL is not None else 0.0
        so4_mol = mgL_to_molM3(so4_mgL or 0.0, "SO4") if so4_mgL is not None else 0.0
        if cl_mgL is None:
            missing.append("Cl")
        if so4_mgL is None:
            missing.append("SO4")
        salt_load = cl_mol + 2.0 * so4_mol
        n3, m3 = _lookup_interval_table(N3_SALT_LOAD, salt_load)
        n_values["N3"] = n3
        m_values["M3"] = m3

    # N4 / M4 — KS 4,3 in mmol/l
    ks43 = measurements.get("KS43")
    if ks43 is None:
        missing.append("KS43")
        n_values["N4"] = 0.0
        m_values["M4"] = 0.0
    else:
        n4, m4 = _lookup_interval_table(N4_KS43, float(ks43))
        n_values["N4"] = n4
        m_values["M4"] = m4

    # N5 / M5 — c(Ca²⁺) in mol/m³
    ca_mgL = measurements.get("Ca")
    if ca_mgL is None:
        missing.append("Ca")
        n_values["N5"] = 0.0
        m_values["M5"] = 0.0
    else:
        ca_mol = mgL_to_molM3(float(ca_mgL), "Ca")
        n5, m5 = _lookup_interval_table(N5_CALCIUM, ca_mol)
        n_values["N5"] = n5
        m_values["M5"] = m5

    # N6 / M6 — pH
    ph = measurements.get("pH")
    if ph is None:
        missing.append("pH")
        n_values["N6"] = 0.0
        m_values["M6"] = 0.0
    else:
        n6, m6 = _lookup_interval_table(N6_PH, float(ph))
        n_values["N6"] = n6
        m_values["M6"] = m6

    # N7 — Object/water potential (no M7)
    if u_potential is not None:
        n7 = _lookup_n7(float(u_potential))
        if n7 is not None:
            n_values["N7"] = n7
        else:
            notes.append("U_h ≤ -0,2 V — N7 nicht abschätzbar (außerhalb Tabelle).")

    # Sum-formulas. Guard N3/N4 against division-by-zero.
    n4_safe = n_values["N4"] if n_values["N4"] != 0.0 else 1.0
    w0_raw = (
        n_values["N1"]
        + n_values["N3"]
        + n_values["N4"]
        + n_values["N5"]
        + n_values["N6"]
        + (n_values["N3"] / n4_safe)
    )
    if "N7" in n_values:
        w0_raw += n_values["N7"]

    w0_int = int(round(w0_raw))
    w1_int = int(round(w0_int - n_values["N1"] + n_values["N2"] * n_values["N3"]))

    wd_raw = m_values["M1"] + m_values["M3"] + m_values["M4"] + m_values["M5"] + m_values["M6"]
    wd_int = int(round(wd_raw))
    wl_int = int(round(wd_int + m_values["M2"]))

    class_w0, flaechen_w0 = _bucket_W(float(w0_int))
    class_w1, flaechen_w1 = _bucket_W(float(w1_int))
    class_wd = _bucket_WD(float(wd_int))
    class_wl = _bucket_WD(float(wl_int))

    rate_w0, rate_w0_max = _rate_W(float(w0_int))
    rate_w1, rate_w1_max = _rate_W(float(w1_int))

    return Din50929WaterResult(
        n_values=n_values,
        m_values=m_values,
        W0=w0_int,
        W1=w1_int,
        WD=wd_int,
        WL=wl_int,
        class_W0=class_w0,
        class_W1=class_w1,
        flaechen_class_W0=flaechen_w0,
        flaechen_class_W1=flaechen_w1,
        class_WD=class_wd,
        class_WL=class_wl,
        rate_W0_mm_per_a=rate_w0,
        rate_W0_max_mm_per_a=rate_w0_max,
        rate_W1_mm_per_a=rate_w1,
        rate_W1_max_mm_per_a=rate_w1_max,
        missing=missing,
        notes=notes,
    )
