# EBV Classification Tool — V0.0.05 (Three-Flow Build)

Automated classification of construction materials and groundwater samples against three German regulatory regimes:

| Flow | Regulation | Input | Output |
|---|---|---|---|
| **EBV** | Ersatzbaustoffverordnung (BGBl. I 2021 S. 2598, Anlage 1 Tabelle 3) | Lab PDFs (AGROLAB / SGS / etc.) | 3-page Mantelverordnung-format A3 PDF + legacy v04 outputs |
| **Aggressivität** | DIN 4030-1:2024-07 (Beton-Wasser) + DIN 50929-3:2024-05 (Korrosion-Wasser) | Lab PDFs (water analyses) | 1-page A3 landscape results report |
| **PAK** | RuVA-StB 01, Berlin Fassung 2018 (Amtsblatt Berlin Nr. 07/2018 S. 900) | Lab PDFs (Straßenaufbruch) | A3 landscape results report with Tabelle 1 reference + per-sample classification |

---

## What this build does (vs. v04)

* **Three parallel flows** — EBV remains correct vs. v04 baseline; Aggressivität and PAK / RuVA-StB are new.
* **PDF auto-parsing for all three flows** — drop a lab PDF in the per-flow input folder, the parser pulls the relevant parameters automatically. No more manual data entry of corrosion values.
* **UMWELT verification echo** — every Step 1 run produces a transposed "UMWELT" sheet inside the validation Excel mirroring the company workbook's UMWELT layout. Every raw lab row is preserved (individual PAK compounds, PCB congeners, heavy metals etc.) plus a mapping column showing which EBV/Aggr parameter each row aggregates to. Read-only — Step 2 ignores it.
* **EBV new-design PDF is 3 pages** — Feststoff (page 1) + Eluat (page 2) + Zusammenfassung (page 3), matching `A_4_3_1_Auswertung_Labor.pdf`.
* **Driving-parameter highlighting** — on all three EBV pages, the parameter cell whose individual class equals the sample's worst-case is filled with the same colour as the Zuordnung column, so the rationale is visible.
* **Naphthalin Fn3 fix** — the previous v05 build over-classified Naphthalin Eluat to ">BM-F3" when PAK16 Feststoff was elevated. The evaluator now (a) maps Fn3 cross-reference Naphthalin/PAK15 → PAK16 correctly, and (b) caps parameters with only-defined-`BM_0*` limits (Naphthalin, PCB6+118 Eluat) at BM-F0* rank rather than cascading to Landfill — the aggregate parameter (PAK15) governs the higher class as the regulator intends.
* **Aggressivität engines** — DIN 4030-1 (XA0/XA1/XA2/XA3/Milieu unstimmig) + DIN 50929-3 (W0, W1, WD, WL via the Gleichung-7/8 formulas confirmed against the company workbook row 49) plus lab-verdict cross-check from the input PDF.

---

## File layout

```
V0.0.05/
├── config.py                       EBV Anlage 1 Tab. 3 limits + synonym mapping
├── config_aggressivität.py         DIN 4030-1 thresholds + DIN 50929-3 N/M lookup tables (water-only)
├── config_pak.py                   RuVA-StB 01 (Berlin Fassung 2018) thresholds + hazardous-waste triggers
├── evaluator.py                    EBV per-sample classification
├── evaluator_aggressivität.py      DIN engines (Beton-Wasser + Korrosion-Wasser)
├── evaluator_pak.py                RuVA-StB 01 classifier (A/B/C/Gefährlicher Abfall)
├── pdf_parser.py                   EBV/PAK lab-PDF parser (pdfplumber + thefuzz)
├── pdf_parser_aggressivität.py     Aggressivität lab-PDF parser
├── reporter.py                     EBV unified reporter (legacy v04 + Mantelverordnung 3-page PDF)
├── reporter_aggressivität.py       Aggressivität results-only A3 landscape report
├── reporter_pak.py                 RuVA-StB 01 results report
├── step1_extraktion.py             Flow dispatcher → ingest stage
├── step2_auswertung.py             Flow dispatcher → evaluation stage
├── start_tool.bat                  Windows launcher
├── requirements.txt
├── templates/
│   └── ebv_template_skeleton.xlsx  Mantelverordnung Feststoff + Eluat + Zusammenfassung
├── 2604XX_Mantelverordnung.xlsx    Source workbook (kept for reference)
├── 2604XX_Rohdaten & Aggressivität.xlsx   Reference for Aggressivität thresholds
├── A_4_3_1_Auswertung_Labor.pdf    EBV output design target
├── 0_input/{EBV,Aggressivität,PAK}/
├── 1_validation/{EBV,Aggressivität,PAK}/<timestamp>_Validierung/
└── 2_output/{EBV,Aggressivität,PAK}/<timestamp>_(Evaluation|Klassifizierung)/
```

---

## Workflow

For each flow, drop your lab PDFs into the per-flow input folder, then run Step 1 + Step 2:

```bash
# Run all three flows end-to-end
python step1_extraktion.py --flow all
python step2_auswertung.py --flow all

# Or per-flow
python step1_extraktion.py --flow ebv
python step2_auswertung.py --flow ebv

python step1_extraktion.py --flow aggressivität
python step2_auswertung.py --flow aggressivität

python step1_extraktion.py --flow pak
python step2_auswertung.py --flow pak

# EBV-specific: override TOC value for Fn7 (heterogeneous soils)
python step2_auswertung.py --flow ebv --toc_override 0.8
```

The Windows launcher `start_tool.bat` exposes a menu with the same options.

### Between Step 1 and Step 2: review the validation Excel

`1_validation/<flow>/<timestamp>_Validierung/Validation_All_Samples.xlsx` contains:

* `_Project` sheet — fill in Bauvorhaben / LOS / Bauwerk (Projektnummer is auto-extracted from the filename).
* `UMWELT` (EBV / PAK) or `UMWELT_Aggr` (Aggressivität) — transposed parameter×sample echo for verification. **Read-only**; do not edit.
* One sheet per sample — fill in `Petrographische_Beschreibung`, `Stratigraphie`, optionally `Soil_Type` (s / l / c / undef) for the EBV flow. The Aggressivität sheets also expose `Tiefe`, `Formation` and `Sample_Type` (defaults to Wasser).

---

## Runtime dependencies

* Python 3.10+
* `pip install -r requirements.txt` (pandas, openpyxl, pdfplumber, thefuzz, reportlab)
* **LibreOffice** (`soffice` / `libreoffice` on `PATH`) for the new-design / Aggressivität / PAK PDF rendering. Without it the legacy EBV outputs still produce, but the Mantelverordnung-format PDF, Aggressivität PDF, and PAK PDF are skipped with a warning.

---

## Coverage and accepted limitations

* **EBV BM/BG only.** The other 12 EBV material classes (HOS, HS, SWS, GKOS, RC, CUM, GRS, SKG, SKA, SFA, BFA, HMVA, GS) are not yet handled. The Mantelverordnung-format PDF shows the full Tabelle 2 reference grid for these classes but no sample evaluation.
* **BM-0 soil-subtype default** — when `Soil_Type` is empty/`undef`, the evaluator picks `BM_0_Sand` (strictest). Set per sample in the validation Excel for the correct subtype.
* **Aggressivität: water samples only.** DIN 4030-1 Beton-Boden and DIN 50929-3 Korrosion-Boden engines are planned but not yet implemented. Soil-corrosion is out of scope for this build.
* **PAK: regional Fassung.** The RuVA-StB 01 thresholds encoded in `config_pak.py` are the Berlin Senatsverwaltung Fassung 2018. Other federal states usually adopt the federal ARS Nr. 40/2001 + ARS Nr. 29/2004 verbatim, but verify against your local Fassung before relying on the classification.

---

## Legal disclaimer

This tool is strictly for **preliminary automated checks**. It does **NOT** replace expert evaluation by a certified geotechnical or environmental engineer. Always cross-check the generated reports against the original laboratory documents. The author accepts no liability for incorrect classifications, parsing errors, or any operational consequences.

---
*Developed for the geotechnical engineering community. Maintained by ThaNexT.*
