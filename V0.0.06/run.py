"""
EBV Tool V0.0.06 — unified entry point.

Replaces the two-step ``step1_extraktion.py`` → ``step2_auswertung.py``
workflow with a single command:

    python run.py [--flow {ebv,pak,aggressivität,all}]

What it does in one go:

1.  Step 1 — Extraction. Parses every PDF under ``0_input/{EBV,PAK,
    Aggressivität}/``, reads project + per-sample metadata from
    ``0_input/background_data.txt`` (optional), and writes one
    consolidated ``Validation.xlsx`` into ``1_validation/<ts>_Validierung/``.

2.  Step 2 — Auswertung. Reads the validation workbook, applies the
    Mantelverordnung Anlage 1 Tab. 3 classification, the
    Aggressivität (DIN 4030-1 / DIN 50929-3) evaluation, and the
    RuVA-StB 01 (PAK) classification. Writes the company-format report
    set into ``2_output/<ts>_Evaluation/``.

Both folders are timestamped in Europe/Berlin local time.

CLI::
    python run.py [--flow {ebv,pak,aggressivität,all}]
                  [--skip-step1] [--skip-step2]
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
from typing import List, Optional


def _run(cmd: List[str]) -> int:
    """Invoke a subprocess inheriting stdio; return its returncode."""
    print(f"\n$ {' '.join(cmd)}")
    return subprocess.call(cmd)


def _slug(text: str) -> str:
    """Make a filesystem-safe slug — keep alphanumeric + dash, collapse the rest."""
    s = re.sub(r"[^\w\-]+", "_", text.strip(), flags=re.UNICODE)
    return re.sub(r"_+", "_", s).strip("_") or "Sample"


def _bundle_folder_name(input_root: str = "0_input") -> str:
    """Derive a 'BW<bauwerk>_<short>' folder name from background_data.txt.

    Falls back to the literal ``"Bundle"`` if metadata is missing.
    """
    path = os.path.join(input_root, "background_data.txt")
    if not os.path.exists(path):
        return "Bundle"

    bauwerk = ""
    bauvorhaben = ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip("\n")
                if not line.strip() or line.lstrip().startswith("-"):
                    continue
                parts = line.split(None, 1)
                if len(parts) < 2:
                    continue
                key, val = parts[0], parts[1].strip()
                if key.lower() == "bauwerk":
                    bauwerk = val
                elif key.lower() == "bauvorhaben":
                    bauvorhaben = val
                if bauwerk and bauvorhaben:
                    break
    except OSError:
        return "Bundle"

    bw_slug = _slug(bauwerk.replace(" ", "-")) if bauwerk else ""
    short = ""
    if bauvorhaben:
        skip = {"über", "ueber", "ueber.", "ü.", "bei", "an", "auf"}
        tokens = [t for t in bauvorhaben.split() if t.lower() not in skip]
        short = _slug("_".join(tokens[-2:]) if len(tokens) >= 2 else (tokens[-1] if tokens else ""))

    if bw_slug and short:
        return f"BW{bw_slug}_{short}"
    if bw_slug:
        return f"BW{bw_slug}"
    return "Bundle"


def _latest_dir(root: str, suffix: str) -> Optional[str]:
    """Return the most recently modified ``<root>/*<suffix>`` directory, or None."""
    if not os.path.isdir(root):
        return None
    cands = [
        os.path.join(root, d)
        for d in os.listdir(root)
        if d.endswith(suffix) and os.path.isdir(os.path.join(root, d))
    ]
    return max(cands, key=os.path.getmtime) if cands else None


def _bundle_deliverables() -> Optional[str]:
    """Collect user-facing deliverables into one drag-and-drop folder inside
    the latest ``2_output/<ts>_Evaluation/``.

    Sources: ``Validation.xlsx`` (latest 1_validation folder),
    ``Evaluation_All_Samples.{pdf,xlsx}`` (new design only — skip
    ``OLD_Design``), ``Aggressivität_*.{pdf,xlsx}``, ``RuVA_PAK_*.{pdf,xlsx}``.
    Missing files (e.g. no PAK in this run) are skipped silently.

    Returns:
        Absolute path of the bundle folder, or ``None`` if no
        ``2_output`` session was found.
    """
    out_dir = _latest_dir("2_output", "_Evaluation")
    if out_dir is None:
        print("[bundle] No 2_output/*_Evaluation folder yet — skipping bundle.")
        return None

    val_dir = _latest_dir("1_validation", "_Validierung")
    bundle_dir = os.path.join(out_dir, _bundle_folder_name())
    os.makedirs(bundle_dir, exist_ok=True)

    sources: List[str] = []
    if val_dir is not None:
        v = os.path.join(val_dir, "Validation.xlsx")
        if os.path.exists(v):
            sources.append(v)

    for name in ("Evaluation_All_Samples.pdf", "Evaluation_All_Samples.xlsx"):
        p = os.path.join(out_dir, name)
        if os.path.exists(p):
            sources.append(p)

    for pat in ("Aggressivität_*.pdf", "Aggressivität_*.xlsx",
                "RuVA_PAK_*.pdf", "RuVA_PAK_*.xlsx"):
        for p in glob.glob(os.path.join(out_dir, pat)):
            sources.append(p)

    copied: List[str] = []
    for src in sources:
        dst = os.path.join(bundle_dir, os.path.basename(src))
        shutil.copy2(src, dst)
        copied.append(os.path.basename(dst))

    if copied:
        print(f"\n[bundle] Drag-and-drop folder ready: {bundle_dir}")
        for fn in copied:
            print(f"  + {fn}")
    else:
        print(f"\n[bundle] No deliverables found to bundle in {out_dir}.")
    return bundle_dir


def main() -> int:
    """Parse CLI args and chain step1 → step2 → bundle."""
    parser = argparse.ArgumentParser(
        description="EBV Tool V0.0.06 — extract & classify in one go.",
    )
    parser.add_argument(
        "--flow",
        choices=("ebv", "pak", "aggressivität", "all"),
        default="all",
        help="Which flows to ingest (default: all — empty folders are skipped).",
    )
    parser.add_argument(
        "--skip-step1",
        action="store_true",
        help="Skip extraction; re-use the most recent 1_validation folder.",
    )
    parser.add_argument(
        "--skip-step2",
        action="store_true",
        help="Stop after Step 1; do not run the classification / reporter.",
    )
    args = parser.parse_args()
    py = sys.executable or "python3"
    rc = 0
    if not args.skip_step1:
        rc = _run([py, "step1_extraktion.py", "--flow", args.flow])
        if rc != 0:
            print(f"\n!! Step 1 failed (rc={rc}). Aborting.")
            return rc
    if not args.skip_step2:
        rc = _run([py, "step2_auswertung.py"])
        if rc != 0:
            print(f"\n!! Step 2 failed (rc={rc}).")
            return rc
        # After a successful step2, gather user-facing deliverables into
        # one drag-and-drop subfolder inside the timestamped output.
        _bundle_deliverables()
    print("\n[run.py] Done. Open the latest folder under 2_output/ for the report.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
