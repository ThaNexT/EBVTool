"""
EBV Tool v05 — Step 2 evaluation (dual-flow).

Two parallel flows are supported:

* **EBV flow** — loads the latest validation Excel from
  ``1_validation/EBV/``, evaluates each sample sheet against EBV Anlage 1
  Tabelle 3 (via ``evaluator.py``), reads project + per-sample metadata,
  and writes the combined reports to ``2_output/EBV/<timestamp>_Evaluation/``:

  - ``Evaluation_All_Samples_OLD_Design.xlsx`` — legacy Excel
  - ``Evaluation_All_Samples_OLD_Design.pdf``  — legacy PDF
  - ``Evaluation_All_Samples_OLD_Design.html`` — legacy HTML
  - ``Evaluation_All_Samples.pdf`` — new design, Mantelverordnung format

* **Aggressivität flow** — loads the latest validated Aggressivität
  workbook from ``1_validation/Aggressivität/``, runs the DIN engines, and
  writes the company-format output to
  ``2_output/Aggressivität/<timestamp>_Klassifizierung/``.
  *Implementation pending — placeholder in this build.*

CLI::

    python step2_auswertung.py [--flow {ebv,aggressivität,all}] [--toc_override <float>]

The ``UMWELT`` sheet introduced by Step 1 (transposed verification echo) is
ignored by the EBV evaluator just like the ``_Project`` sheet.
"""
from __future__ import annotations

import argparse
import os
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
    _BERLIN_TZ = ZoneInfo("Europe/Berlin")
except Exception:
    _BERLIN_TZ = None


def _berlin_now() -> "datetime":
    """Return current datetime in Europe/Berlin so timestamps line up
    with the user's wall clock instead of the (UTC) container clock."""
    if _BERLIN_TZ is not None:
        return datetime.now(_BERLIN_TZ)
    return datetime.now()
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from config import SOIL_TYPE_MAPPING
from evaluator import evaluate_sample
from reporter import ProjectMeta, SampleMeta, create_combined_report

#: Per-flow validation/output directory layout.
VALIDATION_ROOT: str = "1_validation"
OUTPUT_ROOT: str = "2_output"

# Backwards-compat aliases — every flow now reads/writes from the shared roots.
VALIDATION_DIR_EBV: str = VALIDATION_ROOT
VALIDATION_DIR_AGGR: str = VALIDATION_ROOT
VALIDATION_DIR_PAK: str = VALIDATION_ROOT
OUTPUT_DIR_EBV: str = OUTPUT_ROOT
OUTPUT_DIR_AGGR: str = OUTPUT_ROOT
OUTPUT_DIR_PAK: str = OUTPUT_ROOT

#: Shared output-session timestamp set lazily on first call.
_SESSION_TS_OUT: Optional[str] = None


def _get_session_dir_output() -> str:
    """Return (and lazily create) the shared output-session folder.

    All flows in a single ``step2`` invocation write into
    ``2_output/<ts>_Evaluation/`` together (Evaluation_All_Samples.pdf
    for EBV, Aggressivität.pdf/xlsx, RuVA_PAK.pdf/xlsx).

    Returns:
        Path to the timestamped session folder.
    """
    global _SESSION_TS_OUT
    if _SESSION_TS_OUT is None:
        _SESSION_TS_OUT = _berlin_now().strftime("%Y-%m-%d_%H-%M")
    session_dir = os.path.join(OUTPUT_ROOT, f"{_SESSION_TS_OUT}_Evaluation")
    os.makedirs(session_dir, exist_ok=True)
    return session_dir


def _latest_validation_session() -> Optional[str]:
    """Find the most recent ``1_validation/<ts>_Validierung/`` folder.

    Returns:
        Absolute path to the latest session, or None if none exists.
    """
    if not os.path.exists(VALIDATION_ROOT):
        return None
    subdirs = [
        os.path.join(VALIDATION_ROOT, d)
        for d in os.listdir(VALIDATION_ROOT)
        if os.path.isdir(os.path.join(VALIDATION_ROOT, d))
        and d.endswith("_Validierung")
    ]
    if not subdirs:
        return None
    return max(subdirs, key=os.path.getmtime)

PROJECT_SHEET_NAME: str = "_Project"
#: Verification-only echo sheet inserted by Step 1; never carries evaluator input.
UMWELT_SHEET_NAME: str = "UMWELT"


def _read_project_meta(excel_path: str) -> ProjectMeta:
    """Read the ``_Project`` sheet (key-value layout) into a :class:`ProjectMeta`.

    Args:
        excel_path: path to the validation workbook.

    Returns:
        A populated :class:`ProjectMeta`, or an empty one if the sheet is
        missing or malformed. A missing sheet does not raise — the new-design
        PDF still renders with blank header fields.
    """
    try:
        proj_df = pd.read_excel(excel_path, sheet_name=PROJECT_SHEET_NAME, header=0)
    except (ValueError, KeyError):
        return ProjectMeta()

    if proj_df.empty or proj_df.shape[1] < 2:
        return ProjectMeta()

    proj_df = proj_df.dropna(subset=[proj_df.columns[0]])
    fields: Dict[str, str] = {
        str(row.iloc[0]).strip(): ("" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip())
        for _, row in proj_df.iterrows()
    }

    return ProjectMeta(
        projektnummer=fields.get("Projektnummer", ""),
        bauvorhaben=fields.get("Bauvorhaben", ""),
        los=fields.get("LOS", ""),
        bauwerk=fields.get("Bauwerk", ""),
    )


def _read_sample_meta(df: pd.DataFrame, sheet_name: str) -> SampleMeta:
    """Extract per-sample metadata from a sample sheet's DataFrame.

    Args:
        df: raw per-sample DataFrame (output of ``pd.read_excel`` on a single sheet).
        sheet_name: sheet name; used as fallback for ``probenbezeichnung``.

    Returns:
        Populated :class:`SampleMeta`. Empty strings for fields whose source
        column is absent or entirely blank.
    """

    def first_non_empty(col: str) -> str:
        if col not in df.columns:
            return ""
        for v in df[col].tolist():
            if pd.notna(v) and str(v).strip() != "":
                return str(v).strip()
        return ""

    return SampleMeta(
        probenbezeichnung=first_non_empty("Probenbezeichnung") or str(sheet_name),
        petrographische_beschreibung=first_non_empty("Petrographische_Beschreibung"),
        stratigraphie=first_non_empty("Stratigraphie"),
        labor_nummer=first_non_empty("Labor_Nummer"),
    )


def _process_ebv(
    validation_dir: str = VALIDATION_DIR_EBV,
    output_dir_parent: str = OUTPUT_DIR_EBV,
    toc_override: float = -1.0,
) -> bool:
    """Run the EBV branch of Step 2.

    Loads the most recent ``Validation_All_Samples.xlsx`` under
    ``validation_dir``, evaluates each sample sheet via
    :func:`evaluate_sample`, and emits the combined report set (legacy +
    new design) to a timestamped subfolder of ``output_dir_parent``.

    Args:
        validation_dir: parent folder holding ``<timestamp>_Validierung/``
            subfolders produced by Step 1's EBV branch.
        output_dir_parent: parent folder for the new ``<timestamp>_Evaluation/``.
        toc_override: optional manual TOC value (Fn. 7). ``-1.0`` disables.

    Returns:
        True if a report set was produced; False if no validation folder
        or workbook was found.
    """
    print("STEP 2 [EBV]: Starting evaluation of validated data...")

    latest_dir = _latest_validation_session()
    if latest_dir is None:
        print(f"No validation session found under '{VALIDATION_ROOT}'. Run Step 1 first.")
        return False

    val_path = os.path.join(latest_dir, "Validation.xlsx")
    if not os.path.exists(val_path):
        print(f"File {val_path} not found — run Step 1 first.")
        return False
    _flow_prefix = "EBV_"

    out_dir = _get_session_dir_output()

    try:
        excel_tabs = pd.read_excel(val_path, sheet_name=None)
    except PermissionError:
        print("ERROR: The validation Excel file is currently open. Please close Excel and try again.")
        return False

    project_meta = _read_project_meta(val_path)

    evaluated_sheets: Dict[str, pd.DataFrame] = {}
    sample_meta_map: Dict[str, SampleMeta] = {}
    last_bodenart: str = "BM_0_Sand"

    for sheet_name, df in excel_tabs.items():
        if not sheet_name.startswith(_flow_prefix):
            continue  # other flow's sample or _Project/UMWELT

        print(f"  - Evaluating Sample: {sheet_name}")

        if "EBV_Parameter" not in df.columns:
            print(
                f"  -> WARNING: sheet '{sheet_name}' has no 'EBV_Parameter' column; skipping."
            )
            continue

        df_clean = df.dropna(subset=["EBV_Parameter"]).copy()
        df_clean = df_clean[df_clean["EBV_Parameter"] != ""]
        df_clean = df_clean.rename(
            columns={"Lab_Operator": "Operator", "Lab_Value": "Wert", "Lab_Unit": "Einheit"}
        )

        raw_soil_type = "undef"
        if "Soil_Type" in df_clean.columns and not df_clean.empty:
            extracted_soil = df_clean.iloc[0]["Soil_Type"]
            if pd.notna(extracted_soil) and str(extracted_soil).strip() != "":
                raw_soil_type = str(extracted_soil).strip().lower()

        mapped_bodenart = SOIL_TYPE_MAPPING.get(raw_soil_type, "BM_0_Sand")
        last_bodenart = mapped_bodenart

        sample_meta = _read_sample_meta(df, sheet_name)
        sample_meta_map[sheet_name] = sample_meta

        evaluated_df = evaluate_sample(
            df_clean, bodenart=mapped_bodenart, toc_override=toc_override
        )
        evaluated_sheets[sheet_name] = evaluated_df

    if not evaluated_sheets:
        print("[EBV] No evaluable sample sheets found in the validation workbook.")
        return False

    create_combined_report(
        sheet_dict=evaluated_sheets,
        output_dir=out_dir,
        original_filename="All_Samples",
        bodenart=last_bodenart,
        project_meta=project_meta,
        sample_meta_map=sample_meta_map,
    )
    print(f"\n[EBV] Evaluation completed. Reports saved to: {out_dir}")
    return True


def _process_pak(
    validation_dir: str = VALIDATION_DIR_PAK,
    output_dir_parent: str = OUTPUT_DIR_PAK,
) -> bool:
    """Run the PAK / RuVA-StB 01 branch of Step 2.

    Loads the most recent ``Validation_All_Samples.xlsx`` under
    ``validation_dir`` (produced by Step 1's PAK branch — reuses the EBV
    pdf_parser), extracts PAK16 / Benzo(a)pyren / Phenolindex per sample,
    runs :func:`evaluate_pak`, and emits a custom RuVA-StB 01 report
    (``RuVA_PAK_<projektnummer>.xlsx`` + ``.pdf``) into
    ``output_dir_parent/<timestamp>_Klassifizierung/``.

    Args:
        validation_dir: parent folder holding ``<timestamp>_Validierung/``
            subfolders produced by Step 1's PAK branch.
        output_dir_parent: parent folder for the new
            ``<timestamp>_Klassifizierung/``.

    Returns:
        True if a report was produced; False if no validation folder /
        workbook was found.
    """
    from evaluator_pak import evaluate_pak
    from reporter_pak import (
        PakProjectMeta,
        PakSampleMeta,
        create_pak_report,
    )

    print("STEP 2 [PAK]: Starting RuVA-StB 01 evaluation of validated data...")

    latest_dir = _latest_validation_session()
    if latest_dir is None:
        print(f"No validation session under '{VALIDATION_ROOT}'. Run Step 1 first.")
        return False

    val_path = os.path.join(latest_dir, "Validation.xlsx")
    if not os.path.exists(val_path):
        print(f"File {val_path} not found — run Step 1 first.")
        return False
    _flow_prefix = "PAK_"

    try:
        excel_tabs = pd.read_excel(val_path, sheet_name=None)
    except PermissionError:
        print("ERROR: The PAK validation Excel is open. Please close Excel and retry.")
        return False

    project_meta_raw = _read_project_meta(val_path)
    project = PakProjectMeta(
        projektnummer=project_meta_raw.projektnummer,
        bauvorhaben=project_meta_raw.bauvorhaben,
        los=project_meta_raw.los,
        bauwerk=project_meta_raw.bauwerk,
    )

    # Extract PAK16 / Benzo(a)pyren / Phenolindex per sample.
    # The PAK Step 1 branch reuses the EBV pdf_parser, so:
    #   - PAK16 lives under EBV_Parameter "PAK16" (Feststoff)
    #   - Benzo(a)pyren lives under EBV_Parameter "Benzo(a)pyren" (Feststoff)
    #   - Phenolindex is NOT in EBV_Parameter — it has an empty mapping in
    #     pdf_parser. We capture it from Lab_Original_String containing "Phenol".
    samples_in: List[Any] = []
    for sheet_name, df in excel_tabs.items():
        if not sheet_name.startswith(_flow_prefix):
            continue
        if "EBV_Parameter" not in df.columns:
            print(f"  -> WARNING: sheet '{sheet_name}' has no parser output; skipping.")
            continue

        def _first_non_empty(col: str) -> str:
            if col not in df.columns:
                return ""
            for v in df[col].tolist():
                if pd.notna(v) and str(v).strip() != "":
                    return str(v).strip()
            return ""

        probe = _first_non_empty("Probenbezeichnung") or str(sheet_name)
        meta = PakSampleMeta(
            probenbezeichnung=probe,
            petrographische_beschreibung=_first_non_empty("Petrographische_Beschreibung"),
            stratigraphie=_first_non_empty("Stratigraphie"),
            labor_nummer=_first_non_empty("Labor_Nummer"),
            tiefe="",
        )

        # Pull canonical values
        def _value_for_ebv(name: str, matrix: str) -> Optional[float]:
            """Look up a parsed value by EBV canonical name + matrix."""
            mask = (df["EBV_Parameter"] == name) & (df.get("Matrix", "") == matrix)
            sub = df[mask]
            if sub.empty:
                return None
            row = sub.iloc[0]
            v = row.get("Lab_Value")
            if v is None or (isinstance(v, float) and pd.isna(v)):
                # If marked < BG, treat as 0 for hazardous-trigger comparison
                op = str(row.get("Lab_Operator", "")).strip()
                return 0.0 if "<" in op else None
            try:
                return float(v)
            except (TypeError, ValueError):
                return None

        def _value_for_phenol() -> Optional[float]:
            """Look up the Phenolindex by raw lab-text scan."""
            for _, row in df.iterrows():
                label = str(row.get("Lab_Original_String", "") or "").lower()
                if "phenol" in label:
                    v = row.get("Lab_Value")
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        op = str(row.get("Lab_Operator", "")).strip()
                        return 0.0 if "<" in op else None
                    try:
                        return float(v)
                    except (TypeError, ValueError):
                        return None
            return None

        pak16 = _value_for_ebv("PAK16", "Feststoff")
        bap = _value_for_ebv("Benzo(a)pyren", "Feststoff")
        phenol = _value_for_phenol()

        print(
            f"  - Evaluating Sample: {probe}  "
            f"(PAK16={pak16}, Benzo(a)pyren={bap}, Phenolindex={phenol})"
        )
        result = evaluate_pak(pak16=pak16, benzo_a_pyren=bap, phenolindex=phenol)
        samples_in.append((meta, result))

    if not samples_in:
        print("[PAK] No evaluable sample sheets found in the validation workbook.")
        return False

    out_dir = _get_session_dir_output()

    xlsx_path, pdf_path = create_pak_report(
        samples=samples_in, output_dir=out_dir, project=project
    )

    print("\n[PAK] Results:")
    for meta, result in samples_in:
        notes = f"  ({'; '.join(result.notes)})" if result.notes else ""
        print(
            f"  {meta.probenbezeichnung:30s}  Klasse {result.klasse:22s} "
            f"→ {result.verwertungsverfahren}{notes}"
        )

    print(f"\n[PAK] Reports saved to: {out_dir}")
    print(f"  -> {xlsx_path}")
    if pdf_path:
        print(f"  -> {pdf_path}")
    else:
        print("  (PDF not generated — LibreOffice issue; see message above)")
    return True


def _process_aggressivität(
    validation_dir: str = VALIDATION_DIR_AGGR,
    output_dir_parent: str = OUTPUT_DIR_AGGR,
) -> bool:
    """Run the Aggressivität branch of Step 2.

    Loads the latest validation workbook from
    ``1_validation/Aggressivität/<timestamp>_Validierung/``, evaluates each
    sample sheet against DIN 4030-1 and DIN 50929-3 via
    :mod:`evaluator_aggressivität`, and emits a populated company
    workbook (``.xlsx`` + ``.pdf`` via LibreOffice) into
    ``output_dir_parent``.

    Args:
        validation_dir: parent folder holding ``<timestamp>_Validierung/``
            subfolders.
        output_dir_parent: parent folder for the new
            ``<timestamp>_Klassifizierung/``.

    Returns:
        True if a report was produced; False on any blocker (no
        validation folder, missing source workbook, etc.).
    """
    from reporter_aggressivität import (
        AggrProjectMeta,
        AggrSampleMeta,
        create_aggressivität_report,
    )

    print("STEP 2 [Aggressivität]: Starting evaluation of validated data...")

    latest_dir = _latest_validation_session()
    if latest_dir is None:
        print(f"No validation session under '{VALIDATION_ROOT}'. Run Step 1 first.")
        return False

    val_path = os.path.join(latest_dir, "Validation.xlsx")
    if not os.path.exists(val_path):
        print(f"File {val_path} not found — run Step 1 first.")
        return False
    _flow_prefix = "Aggr_"

    # Locate the company source workbook. From V0.0.06 onward the
    # reference file lives next to the other XLSX skeletons under
    # ``templates/`` so the user never has to drop it in the root by hand.
    # Older layouts (V0.0.05 and earlier) kept it next to the scripts —
    # we still honour that path as a fallback for backwards compatibility.
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = (
        os.path.join(here, "templates", "2604XX_Rohdaten & Aggressivität.xlsx"),
        os.path.join(here, "2604XX_Rohdaten & Aggressivität.xlsx"),
        os.path.join(here, "templates", "2604XX_Rohdaten_und_Aggressivität.xlsx"),
        os.path.join(here, "2604XX_Rohdaten_und_Aggressivität.xlsx"),
    )
    source_workbook = next((p for p in candidates if os.path.exists(p)), None)
    if source_workbook is None:
        print(
            "  -> ERROR: Aggressivität source workbook not found. Expected one of:\n    "
            + "\n    ".join(candidates)
        )
        return False

    try:
        excel_tabs = pd.read_excel(val_path, sheet_name=None)
    except PermissionError:
        print("ERROR: The Aggressivität validation Excel is open. Please close Excel and retry.")
        return False

    project_meta_raw = _read_project_meta(val_path)
    project = AggrProjectMeta(
        projektnummer=project_meta_raw.projektnummer,
        bauvorhaben=project_meta_raw.bauvorhaben,
        los=project_meta_raw.los,
        bauwerk=project_meta_raw.bauwerk,
    )

    # Build the (sample_meta, measurements, operators) list in sheet order.
    samples: List[Any] = []
    for sheet_name, df in excel_tabs.items():
        if not sheet_name.startswith(_flow_prefix):
            continue

        if "Aggr_Parameter" not in df.columns:
            print(
                f"  -> WARNING: sheet '{sheet_name}' has no 'Aggr_Parameter' column; skipping."
            )
            continue

        probe = (
            str(df["Probenbezeichnung"].iloc[0]).strip()
            if "Probenbezeichnung" in df.columns and len(df) > 0
            else sheet_name
        )

        def _first_non_empty(col: str) -> str:
            if col not in df.columns:
                return ""
            for v in df[col].tolist():
                if pd.notna(v) and str(v).strip() != "":
                    return str(v).strip()
            return ""

        meta = AggrSampleMeta(
            probenbezeichnung=probe,
            tiefe=_first_non_empty("Tiefe"),
            formation=_first_non_empty("Formation"),
            wasserart="stehend",  # default — user can extend the validation
            objektlage=None,
            u_potential=None,
            lab_din4030_verdict="",
        )

        measurements: Dict[str, Optional[float]] = {}
        operators: Dict[str, str] = {}
        for _, row in df.iterrows():
            cid = str(row.get("Aggr_Parameter", "")).strip()
            if not cid:
                continue
            if cid == "Lab_DIN4030_assessment":
                verdict_txt = str(row.get("Lab_Verdict_Text", "") or "").strip()
                if verdict_txt:
                    meta = AggrSampleMeta(
                        probenbezeichnung=meta.probenbezeichnung,
                        tiefe=meta.tiefe,
                        formation=meta.formation,
                        wasserart=meta.wasserart,
                        objektlage=meta.objektlage,
                        u_potential=meta.u_potential,
                        lab_din4030_verdict=verdict_txt,
                    )
                continue
            val = row.get("Lab_Value")
            if val is None or (isinstance(val, float) and pd.isna(val)):
                measurements[cid] = None
            else:
                try:
                    measurements[cid] = float(val)
                except (TypeError, ValueError):
                    measurements[cid] = None
            op = row.get("Lab_Operator")
            operators[cid] = str(op).strip() if pd.notna(op) else ""

        samples.append((meta, measurements, operators))
        print(f"  - Evaluating Sample: {probe} (sheet '{sheet_name}')")

    if not samples:
        print("[Aggressivität] No evaluable sample sheets found in the validation workbook.")
        return False

    out_dir = _get_session_dir_output()

    xlsx_path, pdf_path, results = create_aggressivität_report(
        samples=samples,
        output_dir=out_dir,
        project=project,
        source_workbook=source_workbook,
    )

    # Concise per-sample summary in console + cross-check vs lab verdict
    print("\n[Aggressivität] Results:")
    for (meta, _m, _o), (r4030, r50929) in zip(samples, results):
        lab = meta.lab_din4030_verdict or "(keine Lab-Aussage)"
        cross = " ✓" if (
            meta.lab_din4030_verdict
            and (
                ("nicht angreif" in lab.lower() and r4030.overall_class == "XA0")
                or (lab.upper().replace(" ", "") in {r4030.overall_class.replace(" ", "")})
                or (r4030.overall_class.upper() in lab.upper())
            )
        ) else (" ⚠ (Vergleich nicht eindeutig)" if meta.lab_din4030_verdict else "")
        print(
            f"  {meta.probenbezeichnung:30s}  DIN 4030: {r4030.overall_class:18s} "
            f"lab: {lab:25s}{cross}"
        )
        print(
            f"  {' ':30s}  DIN 50929: W0={r50929.W0:>3d} ({r50929.class_W0}), "
            f"W1={r50929.W1:>3d} ({r50929.class_W1}), "
            f"WD={r50929.WD:>3d} ({r50929.class_WD}), "
            f"WL={r50929.WL:>3d} ({r50929.class_WL})"
        )

    print(f"\n[Aggressivität] Reports saved to: {out_dir}")
    print(f"  -> {xlsx_path}")
    if pdf_path:
        print(f"  -> {pdf_path}")
    else:
        print("  (PDF not generated — LibreOffice issue; see message above)")
    return True


def main() -> None:
    """Entry point for Step 2 - dispatches to EBV / Aggressivität / PAK flow."""
    parser = argparse.ArgumentParser(
        description="EBV Tool - Step 2 (evaluation). Three-flow: EBV + Aggressivität + PAK."
    )
    parser.add_argument(
        "--flow",
        choices=["ebv", "aggressivität", "pak", "all"],
        default="all",
        help="Which flow to run. Default: all (runs all three; empty branches skipped).",
    )
    parser.add_argument(
        "--toc_override",
        type=float,
        default=-1.0,
        help="Manual TOC override for Fn 7 (e.g., 0.8). EBV flow only.",
    )
    args = parser.parse_args()

    if args.flow in ("ebv", "all"):
        _process_ebv(toc_override=args.toc_override)
    if args.flow in ("aggressivität", "all"):
        _process_aggressivität()
    if args.flow in ("pak", "all"):
        _process_pak()


if __name__ == "__main__":
    main()
