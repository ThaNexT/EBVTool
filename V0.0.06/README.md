# EBV Tool V0.0.06

Classifies German construction-material lab reports against:
- EBV (Ersatzbaustoffverordnung) — Anlage 1 Tab. 3 BM/BG classes.
- Aggressivität — DIN 4030-1 (Beton-Wasser) + DIN 50929-3 (Korrosion-Wasser).
- PAK / RuVA-StB 01 (Berlin Fassung 2018) for road-construction reuse.

## Quick start

1. Drop input PDFs into the right folders:
   * `0_input/EBV/` — per-sample SGS Feststoff+Eluat lab reports
   * `0_input/PAK/` — Straßenaufbruch / asphalt samples
   * `0_input/Aggressivität/` — AGROLAB water-corrosion reports
   * `0_input/background_data.txt` — optional project metadata (template inside)

2. Run one command: `python run.py` (or double-click `start_tool.bat`).

3. Open the report under `2_output/<timestamp>_Evaluation/`:
   * `Evaluation_All_Samples.xlsx` / `.pdf` — Feststoff / Eluat / Zusammenfassung Dekklaration.
   * `Aggressivität_<projektnummer>.xlsx` / `.pdf` — Beton-Wasser + Korrosion-Wasser.
   * `RuVA_PAK_<projektnummer>.xlsx` / `.pdf` — Klasse A/B/C / Gef. Abfall.

The intermediate validation workbook sits under `1_validation/<timestamp>_Validierung/Validation.xlsx`
— verification echo; no manual editing required. `background_data.txt` carries
everything that used to require manual editing between the two steps.

## CLI options

    python run.py --flow all              (default)
    python run.py --flow ebv
    python run.py --flow pak
    python run.py --flow aggressivität
    python run.py --skip-step1
    python run.py --skip-step2

## Timestamps

All folder timestamps use Europe/Berlin local time.
