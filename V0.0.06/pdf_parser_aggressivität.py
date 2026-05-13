"""
EBV Tool — Aggressivität PDF parser (water-only, Phase 2 build).

Extracts DIN 4030-1 (Beton-Wasser) and DIN 50929-3 (Korrosion-Wasser)
parameters from standard German lab reports (AGROLAB, SGS, …). Uses
pdfplumber for layout-preserved text extraction, then walks each line
with a prefix-match against the Aggressivität synonym table from
:mod:`config_aggressivität`.

Output schema (per row, one row per parameter recognised):

    Lab_Original_String  raw text from the lab report
    Full_Row             entire line text
    Aggr_Parameter       canonical ID from WATER_AGGR_PARAMETERS
                         (empty string if unmatched/blacklisted)
    Lab_Unit             cleaned unit string
    Lab_Operator         comparison operator: "<", ">", "< BG", or ""
    Lab_Value            numeric value as float, or None when < BG / unparseable
    Lab_Verdict_Text     filled only for the lab's own DIN 4030 verdict row

Why prefix-match instead of column-split: pdfplumber's
``extract_text(x_tolerance=2, y_tolerance=2)`` collapses tabular spacing
to single spaces, so multi-space splits fail on real lab PDFs.
"""
from __future__ import annotations

import logging
import re
from typing import List, Optional, Tuple

import pandas as pd
import pdfplumber

from config_aggressivität import (
    LAB_DIN4030_KEY,
    WATER_AGGR_BLACKLIST,
    WATER_AGGR_PARAMETERS,
    WATER_AGGR_SYNONYMS,
)

#: Canonical IDs in order (drives output row ordering).
_CANONICAL_IDS: List[str] = [p[0] for p in WATER_AGGR_PARAMETERS]

#: Synonyms sorted longest-first so e.g. "leitfähigkeit bei 25 °c (labor)"
#: wins over the shorter "leitfähigkeit" prefix.
_SYNONYMS_SORTED: List[Tuple[str, str]] = sorted(
    WATER_AGGR_SYNONYMS.items(), key=lambda kv: -len(kv[0])
)

#: Unit tokens we recognise inside the data portion of a line.
_UNIT_KEYWORDS: Tuple[str, ...] = (
    "mg/l", "µg/l", "ug/l", "µs/cm", "us/cm", "mmol/l", "mol/m³", "mol/m3",
    "°c", "°dh", "ph", "cao",
)

#: Section-header / footer / metadata line prefixes we skip outright.
_SKIP_PREFIXES: Tuple[str, ...] = (
    "sensorische", "physikalisch", "summarische", "kationen", "anionen",
    "berechnete", "hinweis", "erläuterung", "seite", "agrolab", "doc-",
    "prüfbericht", "auftrag", "analysennr", "probeneingang", "probenahme",
    "probenehmer", "kunden-probenbezeichnung", "datum", "kundennr",
    "[@", "einheit ergebnis", "die in diesem", "ag landshut",
    "hrb", "ust/vat", "de 128", "dr-pauling", "www.agrolab", "fax:",
    "ab einem wert", "stahlbeton",
)


def map_parameter_name(raw_name: str) -> Optional[str]:
    """Map a raw lab-report row label to a canonical Aggressivität parameter ID.

    Args:
        raw_name: row label text from the lab PDF.

    Returns:
        Canonical ID like ``"pH"`` / ``"Ca"`` / …, or None when the row is
        unrecognised or explicitly blacklisted.
    """
    if not raw_name or not isinstance(raw_name, str):
        return None
    s = raw_name.strip().lower()

    if s in WATER_AGGR_SYNONYMS:
        return WATER_AGGR_SYNONYMS[s]

    s_stripped = re.sub(r"\s*\([^)]*\)\s*$", "", s).strip()
    if s_stripped in WATER_AGGR_SYNONYMS:
        return WATER_AGGR_SYNONYMS[s_stripped]

    if any(b in s for b in WATER_AGGR_BLACKLIST):
        return None

    return None


def _match_synonym_prefix(line_lower: str) -> Optional[Tuple[str, int]]:
    """Find the longest synonym matching the start of ``line_lower``.

    Args:
        line_lower: lowercased line text (trimmed).

    Returns:
        Tuple ``(canonical_id, prefix_length)`` if a match is found, else None.
        ``prefix_length`` counts characters of the matched synonym; a
        following whitespace boundary is required.
    """
    for syn, cid in _SYNONYMS_SORTED:
        if line_lower.startswith(syn):
            n = len(syn)
            if len(line_lower) == n or line_lower[n].isspace():
                return cid, n
    return None


def _parse_value_token(token: str) -> Tuple[str, Optional[float]]:
    """Parse a numeric / operator-prefixed value token into (operator, float).

    Examples:
        ``"<0,04"`` → ``("<",     0.04)``
        ``"7,2"``   → ``("",      7.2)``
        ``"n.b."``  → ``("< BG",  None)``
        ``">100"``  → ``(">",     100.0)``

    Args:
        token: raw cell text.

    Returns:
        Tuple ``(operator, value)``. ``operator`` is one of
        ``""``, ``"<"``, ``">"``, ``"< BG"``. ``value`` is None when not
        quantifiable.
    """
    if token is None:
        return ("< BG", None)
    t = str(token).strip()
    if not t:
        return ("< BG", None)
    lower = t.lower()
    if lower in {"n.b.", "n.n.", "n.d.", "nb", "nn", "nd", "-", "--", "—", "–"}:
        return ("< BG", None)

    op = ""
    if t.startswith("<"):
        op = "<"
        t = t[1:].strip()
    elif t.startswith(">"):
        op = ">"
        t = t[1:].strip()

    num_match = re.search(r"[-+]?\d+(?:[.,]\d+)?", t)
    if not num_match:
        return ("< BG", None)
    try:
        return (op, float(num_match.group(0).replace(",", ".")))
    except ValueError:
        return ("< BG", None)


def _extract_unit_and_value(rest_tokens: List[str]) -> Tuple[str, str, str]:
    """Scan post-label tokens for ``(unit, value_token, string_value)``.

    The lab convention places unit before the value, but some rows have
    no explicit unit (pH-Wert, Lab verdict). The first token containing a
    digit or a comparison operator is the numeric value; if no numeric
    token is found, ``string_value`` collects the leading alphabetic
    word(s) up to the first DIN-method / unit boundary so callers can
    display textual results like ``"farblos"``, ``"klar mit Bodensatz"``,
    or ``"nicht angreifend"`` verbatim.

    Args:
        rest_tokens: whitespace-split tokens after the parameter label.

    Returns:
        ``(unit_string, value_token_string, string_value)``. Unit defaults
        to ``"-"``; ``value_token`` is ``""`` when no digit token found;
        ``string_value`` is ``""`` unless the row carries a non-numeric
        result token.
    """
    unit = "-"
    value_token = ""
    string_parts: List[str] = []
    # Stop the string capture when we hit a DIN method, a unit, an "*)"
    # footnote marker, or a numeric token.
    string_stop = ("din", "iso", "dev", "berechnung", "vdlufa")
    string_skip = {"*)", "(labor)", "visuell"}
    for t in rest_tokens:
        tl = t.lower()
        if unit == "-" and any(u in tl for u in _UNIT_KEYWORDS) and not re.match(r"^\d", t):
            unit = t
            # Once unit is found, any further string fragment is irrelevant.
            string_parts = []
            continue
        # Continuation of a multi-word unit ("mg/l CaO" / "°dH dGH" /
        # "mol/m³ Cl⁻"): if unit is already set and we see a short
        # alphabetic-only token before any numeric, glue it on.
        if unit != "-" and value_token == "" and re.match(r"^[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß0-9⁻²³⁺]*$", t) and len(t) <= 4:
            unit = f"{unit} {t}"
            continue
        if not value_token and (re.search(r"\d", t) or tl.lstrip("<>").lower() in {"n.b.", "n.n.", "n.d."}):
            value_token = t
            break
        # No digit yet, no unit yet — capture alphabetic word(s) as a
        # potential string verdict.
        if value_token == "" and unit == "-":
            # Skip footnote markers / lab tags / inline method words
            # WITHOUT halting collection (Trübung shows "*) klar mit
            # Bodensatz visuell" — "klar mit Bodensatz" is the result).
            if tl in string_skip:
                continue
            # Halt on DIN / method labels which appear at the END of
            # the row — anything beyond is bibliographic noise.
            if any(tl.startswith(s) for s in string_stop):
                break
            if re.search(r"[A-Za-zÄÖÜäöüß]", t):
                string_parts.append(t)
    string_value = " ".join(string_parts).strip()
    return unit, value_token, string_value


def _line_to_param_record(line: str) -> Optional[dict]:
    """Convert one lab-report line into a parameter record (or None).

    Args:
        line: a single line of pdfplumber-extracted text.

    Returns:
        Dict matching the parser's output schema, or None if the line is
        not a recognisable Aggressivität parameter row.
    """
    s = line.strip()
    if not s:
        return None
    lower = s.lower()

    # Cheap reject: skip noise / section headers / footer / metadata
    for prefix in _SKIP_PREFIXES:
        if lower.startswith(prefix):
            return None

    # Lab verdict line: "Betonaggressivität (Angriffsgrad DIN 4030) ..."
    if "betonaggressiv" in lower and "din 4030" in lower:
        verdict_match = re.search(
            r"(nicht\s+angreif\w*|schwach\s+angreif\w*|mäßig\s+angreif\w*|stark\s+angreif\w*|milieu\s+unstimmig|xa\s*[0-3])",
            s,
            flags=re.IGNORECASE,
        )
        verdict = verdict_match.group(1) if verdict_match else ""
        return {
            "Lab_Original_String": s[: s.find(verdict)].strip() if verdict else s,
            "Full_Row": s,
            "Aggr_Parameter": LAB_DIN4030_KEY,
            "Lab_Unit": "-",
            "Lab_Operator": "",
            "Lab_Value": None,
            "Lab_Verdict_Text": verdict,
            "Lab_Display_Override": verdict,
            "_Synonym_Matched": "",
        }

    match = _match_synonym_prefix(lower)
    if match is None:
        return None
    canonical, prefix_len = match

    rest = s[prefix_len:].strip()
    rest_tokens = rest.split()
    unit, value_token, string_value = _extract_unit_and_value(rest_tokens)
    op, val = _parse_value_token(value_token)

    # Capture the matched synonym substring so the dedup step can pick the
    # preferred variant when a lab report lists multiple temperature
    # reference points for the same canonical ID (e.g. Leitfähigkeit at
    # both 20 °C and 25 °C — engineering convention prefers 25 °C).
    synonym_matched = lower[:prefix_len].strip()

    return {
        "Lab_Original_String": s[:prefix_len].strip(),
        "Full_Row": s,
        "Aggr_Parameter": canonical,
        "Lab_Unit": unit,
        "Lab_Operator": op,
        "Lab_Value": val,
        "Lab_Verdict_Text": "",
        "Lab_Display_Override": string_value if (val is None and op == "< BG" and string_value) else "",
        "_Synonym_Matched": synonym_matched,
    }


def extract_all_data_from_pdf(pdf_path: str) -> pd.DataFrame:
    """Extract Aggressivität parameters from a lab-report PDF.

    Args:
        pdf_path: filesystem path to the PDF.

    Returns:
        DataFrame in the schema documented at module top, sorted to
        match ``WATER_AGGR_PARAMETERS`` order. Empty DataFrame on read
        failure or when no recognisable rows are present.
    """
    rows: List[dict] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
                for line in text.split("\n"):
                    rec = _line_to_param_record(line)
                    if rec is not None:
                        rows.append(rec)
    except Exception as exc:  # noqa: BLE001
        logging.error("Error reading Aggressivität PDF %s: %s", pdf_path, exc)
        return pd.DataFrame()

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df_aggr = df[df["Aggr_Parameter"] != LAB_DIN4030_KEY].copy()
    df_lab = df[df["Aggr_Parameter"] == LAB_DIN4030_KEY].copy()

    # Dedup with synonym-preference: rows whose matched synonym is in the
    # preferred set rank higher than generic / 20 °C variants. Currently the
    # only canonical ID with multiple temperature-keyed synonyms is
    # Leitfähigkeit; for DIN 4030-1 / DIN 50929-3 the 25 °C reading is the
    # convention. Other parameters fall through to the original
    # first-occurrence ordering.
    PREFERRED_SYNONYM_TOKENS: Tuple[str, ...] = ("25 °c", "25°c", "25 ° c")
    if "_Synonym_Matched" in df_aggr.columns:
        # Prefer rows whose unit string contains "cao" (mg/l CaO is the
        # master-template unit for Carbonathärte / Nichtcarbonathärte /
        # Gesamthärte) over the °dH duplicate. Otherwise honour the 25 °C
        # vs 20 °C ranking for Leitfähigkeit.
        def _pref_rank(row) -> int:
            syn = str(row.get("_Synonym_Matched") or "")
            unit_l = str(row.get("Lab_Unit") or "").lower()
            full_l = str(row.get("Full_Row") or "").lower()
            cid = str(row.get("Aggr_Parameter") or "")
            # Hardness rows: prefer mg/l CaO variant. The unit parser
            # captures "mg/l" alone (CaO becomes a separate token), so
            # detect it via Full_Row instead of the cleaned Lab_Unit.
            if cid in {"Carbonathaerte", "Nichtcarbonathaerte", "Gesamthaerte"}:
                return 0 if ("cao" in unit_l or "cao" in full_l) else 1
            # Leitfähigkeit: prefer 25 °C variant.
            return 0 if any(tok in syn for tok in PREFERRED_SYNONYM_TOKENS) else 1
        df_aggr["_pref"] = df_aggr.apply(_pref_rank, axis=1)
        # Stable sort so within equal preference, original PDF order wins.
        df_aggr = df_aggr.sort_values("_pref", kind="stable")
        df_aggr = df_aggr.drop_duplicates(subset=["Aggr_Parameter"], keep="first")
        df_aggr = df_aggr.drop(columns=["_pref"])
    else:
        df_aggr = df_aggr.drop_duplicates(subset=["Aggr_Parameter"], keep="first")

    order = {pid: i for i, pid in enumerate(_CANONICAL_IDS)}
    df_aggr["_sort"] = df_aggr["Aggr_Parameter"].map(order).fillna(99)
    df_aggr = df_aggr.sort_values("_sort").drop(columns=["_sort"])

    # Drop the internal synonym-tracking column before returning so it
    # doesn't pollute the public DataFrame schema documented in the
    # module docstring.
    for internal_col in ("_Synonym_Matched",):
        if internal_col in df_aggr.columns:
            df_aggr = df_aggr.drop(columns=[internal_col])
        if internal_col in df_lab.columns:
            df_lab = df_lab.drop(columns=[internal_col])

    return pd.concat([df_aggr, df_lab], ignore_index=True)


def extract_probenbezeichnung(pdf_path: str) -> str:
    """Best-effort Probenbezeichnung extraction from the AGROLAB-style header.

    Args:
        pdf_path: filesystem path to the PDF.

    Returns:
        Sample identifier string from the ``Kunden-Probenbezeichnung``
        header row, or ``""`` if not found.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text(x_tolerance=2, y_tolerance=2) or ""
    except Exception:  # noqa: BLE001
        return ""
    m = re.search(r"Kunden-?Probenbezeichnung\s+(\S.+?)(?:\s+(?:Einheit|$)|$)", text)
    if m:
        return m.group(1).strip()
    for line in text.splitlines():
        if line.lstrip().lower().startswith("kunden-probenbezeichnung"):
            return line.split(None, 1)[1].strip() if len(line.split(None, 1)) > 1 else ""
    return ""
