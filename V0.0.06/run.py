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
import os
import subprocess
import sys
from typing import List


def _run(cmd: List[str]) -> int:
    """Invoke a subprocess inheriting stdio; return its returncode."""
    print(f"\n$ {' '.join(cmd)}")
    return subprocess.call(cmd)


def main() -> int:
    """Parse CLI args and chain step1 → step2."""
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
    print("\n[run.py] Done. Open the latest folder under 2_output/ for the report.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
