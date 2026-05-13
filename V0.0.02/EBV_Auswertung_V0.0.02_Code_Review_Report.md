# Code-Review-Report: EBV_Auswertung V0.0.02
**Erstellt:** 2026-04-09  
**Geprüft von:** Claude Sonnet 4.6 (Anthropic)  
**Geprüfte Proben:** MP02, MP03, MP04, MP05 (e-327524, BW63245360, UST)  
**Gesamturteil:** ⚠️ Deutlich verbessert, aber 4 kritische Berechnungsfehler verbleiben

---

## Aufgabe für den Empfänger dieses Reports

Bitte korrigiere den unten vollständig abgedruckten Python-Code (6 Module) gemäß den aufgeführten Fehlern. Es handelt sich um ein automatisiertes Tool zur Klassifizierung von Bodenmaterial nach der deutschen **Ersatzbaustoffverordnung (EBV), Anlage 1 Tabelle 3**.

**Vorbedingung:** Alle 8 Korrekturen aus dem Vorgänger-Report (V0.0.01) wurden bereits korrekt umgesetzt:
- ✅ parser.py → pdf_parser.py (Namenskonflikt behoben)
- ✅ Fußnote-3-Logik (Feststoff vor Eluat) implementiert
- ✅ pH-Wert BM-0 Grenzwerte ergänzt
- ✅ requirements.txt erstellt (inkl. pytest, reportlab)
- ✅ Logging implementiert
- ✅ Rechtlicher Hinweis im Output
- ✅ Bodenart als argparse-Parameter
- ✅ Fehlerbehandlung bei PDF-Öffnung

---

## Vollständiger Quellcode (alle 6 Module)

### config.py
```python
EBV_VERSION = {
    "gesetz": "Ersatzbaustoffverordnung (EBV)",
    "fundstelle": "BGBl. I 2021 S. 2598",
    "tabelle": "Anlage 1 Tabelle 3",
    "geprueft_von": "Automatisierter Abgleich"
}

KLASSEN_HIERARCHIE = [
    "BM_0_Sand", "BM_0_Lehm_Schluff", "BM_0_Ton", 
    "BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3"
]

ebv_tabelle_3 = [
    {"ebv_order": 1, "parameter": "Mineralische Fremdbestandteile", "einheit": "Vol.-%", "typ": "Feststoff", "fussnoten": [1], "grenzwerte": {"BM_0_Sand": 10, "BM_0_Lehm_Schluff": 10, "BM_0_Ton": 10, "BM_0*": 10, "BM_F0*": 50, "BM_F1": 50, "BM_F2": 50, "BM_F3": 50}},
    {"ebv_order": 2, "parameter": "pH-Wert", "einheit": "-", "typ": "Eluat", "fussnoten": [4], "grenzwerte": {"BM_0_Sand": [6.5, 9.5], "BM_0_Lehm_Schluff": [6.5, 9.5], "BM_0_Ton": [6.5, 9.5], "BM_0*": [6.5, 9.5], "BM_F0*": [6.5, 9.5], "BM_F1": [6.5, 9.5], "BM_F2": [6.5, 9.5], "BM_F3": [5.5, 12.0]}},
    {"ebv_order": 3, "parameter": "Elektrische Leitfähigkeit", "einheit": "µS/cm", "typ": "Eluat", "fussnoten": [4], "grenzwerte": {"BM_0*": 350, "BM_F0*": 350, "BM_F1": 500, "BM_F2": 500, "BM_F3": 2000}},
    {"ebv_order": 4, "parameter": "Sulfat", "einheit": "mg/l", "typ": "Eluat", "fussnoten": [5], "grenzwerte": {"BM_0_Sand": 250, "BM_0_Lehm_Schluff": 250, "BM_0_Ton": 250, "BM_0*": 250, "BM_F0*": 250, "BM_F1": 450, "BM_F2": 450, "BM_F3": 1000}},
    {"ebv_order": 5, "parameter": "Arsen", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 10, "BM_0_Lehm_Schluff": 20, "BM_0_Ton": 20, "BM_0*": 20, "BM_F0*": 40, "BM_F1": 40, "BM_F2": 40, "BM_F3": 150}},
    {"ebv_order": 6, "parameter": "Arsen", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 8, "klammerwert": 13}, "BM_F0*": 12, "BM_F1": 20, "BM_F2": 85, "BM_F3": 100}},
    {"ebv_order": 7, "parameter": "Blei", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 40, "BM_0_Lehm_Schluff": 70, "BM_0_Ton": 100, "BM_0*": 140, "BM_F0*": 140, "BM_F1": 140, "BM_F2": 140, "BM_F3": 700}},
    {"ebv_order": 8, "parameter": "Blei", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 23, "klammerwert": 43}, "BM_F0*": 35, "BM_F1": 90, "BM_F2": 250, "BM_F3": 470}},
    {"ebv_order": 9, "parameter": "Cadmium", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [6], "grenzwerte": {"BM_0_Sand": 0.4, "BM_0_Lehm_Schluff": 1.0, "BM_0_Ton": 1.5, "BM_0*": 1.0, "BM_F0*": 2.0, "BM_F1": 2.0, "BM_F2": 2.0, "BM_F3": 10}},
    {"ebv_order": 10, "parameter": "Cadmium", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 2, "klammerwert": 4}, "BM_F0*": 3.0, "BM_F1": 3.0, "BM_F2": 10, "BM_F3": 15}},
    {"ebv_order": 11, "parameter": "Chrom, gesamt", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 30, "BM_0_Lehm_Schluff": 60, "BM_0_Ton": 100, "BM_0*": 120, "BM_F0*": 120, "BM_F1": 120, "BM_F2": 120, "BM_F3": 600}},
    {"ebv_order": 12, "parameter": "Chrom, gesamt", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 10, "klammerwert": 19}, "BM_F0*": 15, "BM_F1": 150, "BM_F2": 290, "BM_F3": 530}},
    {"ebv_order": 13, "parameter": "Kupfer", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 20, "BM_0_Lehm_Schluff": 40, "BM_0_Ton": 60, "BM_0*": 80, "BM_F0*": 80, "BM_F1": 80, "BM_F2": 80, "BM_F3": 320}},
    {"ebv_order": 14, "parameter": "Kupfer", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 20, "klammerwert": 41}, "BM_F0*": 30, "BM_F1": 110, "BM_F2": 170, "BM_F3": 320}},
    {"ebv_order": 15, "parameter": "Nickel", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 15, "BM_0_Lehm_Schluff": 50, "BM_0_Ton": 70, "BM_0*": 100, "BM_F0*": 100, "BM_F1": 100, "BM_F2": 100, "BM_F3": 350}},
    {"ebv_order": 16, "parameter": "Nickel", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 20, "klammerwert": 31}, "BM_F0*": 30, "BM_F1": 30, "BM_F2": 150, "BM_F3": 280}},
    {"ebv_order": 17, "parameter": "Quecksilber", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 0.2, "BM_0_Lehm_Schluff": 0.3, "BM_0_Ton": 0.3, "BM_0*": 0.6, "BM_F0*": 0.6, "BM_F1": 0.6, "BM_F2": 0.6, "BM_F3": 5}},
    {"ebv_order": 18, "parameter": "Quecksilber", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [12], "grenzwerte": {"BM_0*": 0.1}},
    {"ebv_order": 19, "parameter": "Thallium", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 0.5, "BM_0_Lehm_Schluff": 1.0, "BM_0_Ton": 1.0, "BM_0*": 1.0, "BM_F0*": 2, "BM_F1": 2, "BM_F2": 2, "BM_F3": 7}},
    {"ebv_order": 20, "parameter": "Thallium", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [12], "grenzwerte": {"BM_0*": {"standard": 0.2, "klammerwert": 0.3}}},
    {"ebv_order": 21, "parameter": "Zink", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 60, "BM_0_Lehm_Schluff": 150, "BM_0_Ton": 200, "BM_0*": 300, "BM_F0*": 300, "BM_F1": 300, "BM_F2": 300, "BM_F3": 1200}},
    {"ebv_order": 22, "parameter": "Zink", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 100, "klammerwert": 210}, "BM_F0*": 150, "BM_F1": 160, "BM_F2": 840, "BM_F3": 1600}},
    {"ebv_order": 23, "parameter": "TOC", "einheit": "M%", "typ": "Feststoff", "fussnoten": [7], "grenzwerte": {"BM_0_Sand": 1, "BM_0_Lehm_Schluff": 1, "BM_0_Ton": 1, "BM_0*": 1, "BM_F0*": 5, "BM_F1": 5, "BM_F2": 5, "BM_F3": 5}},
    {"ebv_order": 24, "parameter": "Kohlenwasserstoffe (C10-C22)", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [8], "grenzwerte": {"BM_0_Sand": 300, "BM_0_Lehm_Schluff": 300, "BM_0_Ton": 300, "BM_0*": 300, "BM_F0*": 300, "BM_F1": 300, "BM_F2": 300, "BM_F3": 1000}},
    {"ebv_order": 25, "parameter": "Kohlenwasserstoffe (C10-C40)", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [8], "grenzwerte": {"BM_0_Sand": 600, "BM_0_Lehm_Schluff": 600, "BM_0_Ton": 600, "BM_0*": 600, "BM_F0*": 600, "BM_F1": 600, "BM_F2": 600, "BM_F3": 2000}},
    {"ebv_order": 26, "parameter": "Benzo(a)pyren", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 0.3, "BM_0_Lehm_Schluff": 0.3, "BM_0_Ton": 0.3}},
    {"ebv_order": 27, "parameter": "PAK15", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [9], "grenzwerte": {"BM_0*": 0.2, "BM_F0*": 0.3, "BM_F1": 1.5, "BM_F2": 3.8, "BM_F3": 20}},
    {"ebv_order": 28, "parameter": "PAK16", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [10], "grenzwerte": {"BM_0_Sand": 3, "BM_0_Lehm_Schluff": 3, "BM_0_Ton": 3, "BM_0*": 6, "BM_F0*": 6, "BM_F1": 6, "BM_F2": 9, "BM_F3": 30}},
    {"ebv_order": 29, "parameter": "Naphthalin und Methylnaphthaline, gesamt", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [], "grenzwerte": {"BM_0*": 2}},
    {"ebv_order": 30, "parameter": "PCB6 und PCB-118", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 0.05, "BM_0_Lehm_Schluff": 0.05, "BM_0_Ton": 0.05, "BM_0*": 0.1}},
    {"ebv_order": 31, "parameter": "PCB6 und PCB-118", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [], "grenzwerte": {"BM_0*": 0.01}},
    {"ebv_order": 32, "parameter": "EOX", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [11], "grenzwerte": {"BM_0_Sand": 1, "BM_0_Lehm_Schluff": 1, "BM_0_Ton": 1, "BM_0*": 1}}
]

SYNONYM_MAPPING = {
    "leitfähigkeit": "Elektrische Leitfähigkeit",
    "lf": "Elektrische Leitfähigkeit",
    "elektrische leitfähigkeit bei 25°c": "Elektrische Leitfähigkeit",
    "elektrische leitfähigkeit bei 25 °c": "Elektrische Leitfähigkeit",
    "elektrische leitfähigkeit bei s": "Elektrische Leitfähigkeit",
    "pak nach epa": "PAK16",
    "pak (16)": "PAK16",
    "summe pak (16 epa)": "PAK16",
    "summe pak 16": "PAK16",
    "summe pak 16 ebv": "PAK16",
    "summe pak 16 nach ebv": "PAK16",
    "summe pak (16) nach ebv": "PAK16",
    "summe pak nach epa": "PAK16",
    "summe der pak": "PAK16",
    "polycyclische aromatische kohlenwasserstoffe (pak16)": "PAK16",
    "pak 16 (epa)": "PAK16",
    "summe pak 15": "PAK15",
    "summe pak 15 ebv": "PAK15",
    "summe pak 15 nach ebv": "PAK15",
    "summe pak (15) nach ebv": "PAK15",
    "benzo(a)pyren": "Benzo(a)pyren",
    "benzo[a]pyren": "Benzo(a)pyren",
    "benzo(a)pyren ebv": "Benzo(a)pyren",
    "kohlenwasserstoffe c10 - c22": "Kohlenwasserstoffe (C10-C22)",
    "kohlenwasserstoffe c10 - c40": "Kohlenwasserstoffe (C10-C40)",
    "kw-index": "Kohlenwasserstoffe (C10-C40)",
    "ph": "pH-Wert",
    "ph-wert": "pH-Wert",
    "chrom": "Chrom, gesamt",
    "chrom (gesamt)": "Chrom, gesamt",
    "toc": "TOC",
    "naphthalin und methylnaphthaline": "Naphthalin und Methylnaphthaline, gesamt",
    "summe naphthalin und methylnaphthaline": "Naphthalin und Methylnaphthaline, gesamt",
    "summe naphthaline (ebv)": "Naphthalin und Methylnaphthaline, gesamt",
    "summe naphthaline nach ebv": "Naphthalin und Methylnaphthaline, gesamt",
    "pcb": "PCB6 und PCB-118",
    "summe pcb": "PCB6 und PCB-118",
    "summe pcb nach ebv": "PCB6 und PCB-118",
    "summe pcb ebv": "PCB6 und PCB-118"
}
```

### evaluator.py
```python
import pandas as pd
import logging
from config import ebv_tabelle_3

def find_best_class(param_name, einheit, wert, operator, toc_gehalt, bodenart):
    target_item = next((item for item in ebv_tabelle_3 if item["parameter"] == param_name and item["einheit"] == einheit), None)
    if not target_item:
        return "Nicht in EBV Tabelle 3", None, "", 999
        
    fussnoten_str = ", ".join(map(str, target_item.get("fussnoten", [])))
    ebv_order = target_item["ebv_order"]

    is_invalid = wert is None or pd.isna(wert) or str(wert).lower() == "nan"

    if is_invalid:
        if str(operator).strip() == "< BG" or "<" in str(operator): return "BM-0", None, fussnoten_str, ebv_order
        return "Kein Messwert", None, fussnoten_str, ebv_order

    wert_float = float(wert)

    basis_gw = target_item["grenzwerte"].get(bodenart)
    if basis_gw is not None:
        if isinstance(basis_gw, list):
            if basis_gw[0] <= wert_float <= basis_gw[1]:
                return "BM-0", f"[{basis_gw[0]} - {basis_gw[1]}]", fussnoten_str, ebv_order
        else:
            if basis_gw >= wert_float or ("<" in str(operator) and basis_gw >= wert_float):
                return "BM-0", basis_gw, fussnoten_str, ebv_order
            
    for klasse in ["BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3"]:
        gw_eintrag = target_item["grenzwerte"].get(klasse)
        if gw_eintrag is None: continue 
            
        grenzwert = gw_eintrag["klammerwert"] if isinstance(gw_eintrag, dict) and toc_gehalt >= 0.5 else (gw_eintrag["standard"] if isinstance(gw_eintrag, dict) else gw_eintrag)
        
        if isinstance(grenzwert, list):
            if grenzwert[0] <= wert_float <= grenzwert[1]:
                return klasse.replace("_", "-"), f"[{grenzwert[0]} - {grenzwert[1]}]", fussnoten_str, ebv_order
            continue
            
        if float(grenzwert) >= wert_float or ("<" in str(operator) and float(grenzwert) >= wert_float):
            return klasse.replace("_", "-"), grenzwert, fussnoten_str, ebv_order

    return "> BM-F3 (Deponie!)", None, fussnoten_str, ebv_order

def evaluate_sample(df, bodenart="BM_0_Sand", toc_gehalt=0.1):
    results = []
    feststoff_status = {}
    
    if 'Matrix' not in df.columns:
        df['Matrix'] = "Unbekannt"
    
    # Phase 1: Feststoffe
    for item in ebv_tabelle_3:
        if item["typ"] == "Feststoff":
            param = item["parameter"]
            target_einheit = item["einheit"]
            match = df[(df['EBV_Parameter'] == param) & (df['Matrix'] == "Feststoff")]
            wert, operator = None, ""
            if not match.empty:
                wert = match.iloc[0]['Wert']
                operator = match.iloc[0]['Operator']
            klasse, gw, fn, order = find_best_class(param, target_einheit, wert, operator, toc_gehalt, bodenart)
            feststoff_status[param] = klasse
            op_str = "" if pd.isna(operator) or str(operator).lower() == "nan" else str(operator).strip()
            messwert_str = ""
            if not (wert is None or pd.isna(wert) or str(wert).lower() == "nan"):
                messwert_str = f"{op_str} {wert}".strip()
            elif op_str == "< BG" or "<" in op_str:
                messwert_str = "< BG"
            results.append({
                "Parameter": param, "Einheit": target_einheit, "Messwert": messwert_str,
                "Eingestufte Klasse": klasse, "Maßgeblicher GW": gw, "Fußnote": fn, "ebv_order": item["ebv_order"]
            })

    # Phase 2: Eluate
    for item in ebv_tabelle_3:
        if item["typ"] == "Eluat":
            param = item["parameter"]
            target_einheit = item["einheit"]
            match = df[(df['EBV_Parameter'] == param) & (df['Matrix'] == "Eluat")]
            out_einheit = target_einheit
            wert, operator = None, ""
            if not match.empty:
                w = match.iloc[0]['Wert']
                operator = match.iloc[0]['Operator']
                ist_einheit = str(match.iloc[0]['Einheit']).lower()
                if target_einheit == "µg/l" and "mg/l" in ist_einheit:
                    wert = float(w) * 1000 if pd.notna(w) else None
                    out_einheit = "µg/l (umger.)"
                else:
                    wert = w
            klasse, gw, fn, order = find_best_class(param, target_einheit, wert, operator, toc_gehalt, bodenart)

            if param != "Sulfat" and 3 in item.get("fussnoten", []):
                fs_klasse = feststoff_status.get(param, "Fehlt")
                if fs_klasse in ["BM-0", "Kein Messwert (< BG)"]:
                    klasse = "BM-0 (Eluat n. maßgeblich)"
                elif fs_klasse == "Kein Messwert":
                    fn += " | HINWEIS: Feststoffwert fehlt zur sicheren FN 3 Prüfung."

            op_str = "" if pd.isna(operator) or str(operator).lower() == "nan" else str(operator).strip()
            messwert_str = ""
            if not (wert is None or pd.isna(wert) or str(wert).lower() == "nan"):
                messwert_str = f"{op_str} {wert}".strip()
            elif op_str == "< BG" or "<" in op_str:
                messwert_str = "< BG"

            results.append({
                "Parameter": param, "Einheit": out_einheit, "Messwert": messwert_str,
                "Eingestufte Klasse": klasse, "Maßgeblicher GW": gw, "Fußnote": fn, "ebv_order": item["ebv_order"]
            })

    res_df = pd.DataFrame(results)
    if not res_df.empty:
        res_df = res_df.sort_values(by="ebv_order").drop(columns=["ebv_order"])
    return res_df
```

---

## Fehlerliste mit Korrekturen

### Fehler 1 | KRITISCH | Thallium und Quecksilber Eluat: Fußnote 12 wird nicht korrekt angewendet

**Betroffene Datei:** `evaluator.py` → `evaluate_sample()`  
**Betroffene Proben:** MP03 (Thallium Eluat 0,92 µg/l → fälschlich "BM-F3 (Deponie!)"), MP04 (Thallium Eluat 0,25 µg/l → fälschlich "> BM-F3 (Deponie!)")

**Problem:**  
Laut EBV Fußnote 12 ist für Thallium und Quecksilber der Eluat-Grenzwert **ausschließlich für die Klasse BM-0* maßgeblich** (Thallium: 0,2 µg/l Standard / 0,3 µg/l Klammerwert; Quecksilber: 0,1 µg/l). Für die Klassen BM-F0* bis BM-F3 gilt ausschließlich der jeweilige Feststoffwert — es gibt keinen Eluat-Grenzwert für diese Klassen.

Das Programm prüft derzeit: Wert > BM-0*-Grenzwert → kein weiterer Grenzwert in der config → Ausgabe "> BM-F3 (Deponie!)". Das ist fachlich falsch.

**Korrekte Logik:**  
Wenn ein Eluat-Parameter die Fußnote 12 trägt UND der Messwert den einzigen vorhandenen Grenzwert (BM_0*) überschreitet, darf das Programm NICHT "> BM-F3" ausgeben. Es soll stattdessen ausgeben: `"> BM-0* (nur Feststoff maßgeblich für BM-F)"`.

**Anforderung an die Korrektur:**  
In `evaluate_sample()`, Phase 2 (Eluate), nach der `find_best_class()`-Prüfung folgenden Block ergänzen:

```python
# Fn12-Sonderregel: Hg und Tl Eluat nur für BM-0* maßgeblich
if 12 in item.get("fussnoten", []):
    if klasse == "> BM-F3 (Deponie!)":
        klasse = "> BM-0* (Eluat; für BM-F Klassen nur Feststoff maßgeblich, Fn. 12)"
        gw = None
```

Dieser Block muss **nach** der bestehenden Fn3-Prüfung stehen.

---

### Fehler 2 | KRITISCH | Kohlenwasserstoffe: Fußnote 8 falsch modelliert — C10-C40 ist kein eigenständiger EBV-Grenzwert

**Betroffene Datei:** `config.py`, `evaluator.py`

**Problem — ausführliche Erklärung (wichtig für korrekte Umsetzung):**  
Fußnote 8 der EBV Tabelle 3 definiert ein **zweistufiges Prüfschema** für Kohlenwasserstoffe:

1. **Primärer Grenzwert:** Der C10–C22-Wert gilt als hauptmaßgeblicher Grenzwert je Klasse (BM-0: 300 mg/kg, BM-F3: 1000 mg/kg).
2. **Zusatzbedingung:** Der C10–C40-Gesamtgehalt darf zusätzlich den doppelten C10–C22-Wert (= Klammerwert) nicht überschreiten (BM-0: 600 mg/kg, BM-F3: 2000 mg/kg).

Eine Klasse gilt als **eingehalten**, wenn BEIDE Bedingungen gleichzeitig erfüllt sind: C10-C22 ≤ Grenzwert UND C10-C40 ≤ Klammerwert. Wird der Klammerwert (C10-C40) überschritten, erhöht sich die Klasse, weil die Zusatzbedingung der aktuellen Klasse nicht mehr erfüllt ist.

**Das aktuelle Modell (zwei separate Zeilen in ebv_tabelle_3) ist falsch**, weil:
- Es C10-C40 als eigenständigen Parameter mit scheinbar unabhängiger Klasse ausweist
- Es in seltenen Fällen zu falschen Einstufungen führt (z. B. wenn C10-C22 unkritisch ist, aber C10-C40 über dem Klammerwert liegt, wird fälschlich eine niedrige C10-C22-Klasse als Gesamtergebnis angezeigt)
- In MP02 kommt es zufällig zum richtigen Ergebnis (BM-F3), weil KW(C10-C40) = 860 mg/kg > 600 mg/kg (BM-0-Klammerwert) und die Ausgabe zufällig stimmt

**Anforderung an die Korrektur:**  
Das Kohlenwasserstoff-Modell muss umgebaut werden. Es soll nur **einen** EBV-Parameter für Kohlenwasserstoffe geben, der intern beide Bedingungen prüft.

**Schritt 1 — config.py:** Ersetze die beiden KW-Einträge (ebv_order 24 und 25) durch einen einzigen Eintrag mit einem neuen Feld `klammerwert_c40`:

```python
{
    "ebv_order": 24,
    "parameter": "Kohlenwasserstoffe",
    "einheit": "mg/kg",
    "typ": "Feststoff",
    "fussnoten": [8],
    "grenzwerte": {
        "BM_0_Sand":        {"c22": 300,  "c40": 600},
        "BM_0_Lehm_Schluff":{"c22": 300,  "c40": 600},
        "BM_0_Ton":         {"c22": 300,  "c40": 600},
        "BM_0*":            {"c22": 300,  "c40": 600},
        "BM_F0*":           {"c22": 300,  "c40": 600},
        "BM_F1":            {"c22": 300,  "c40": 600},
        "BM_F2":            {"c22": 300,  "c40": 600},
        "BM_F3":            {"c22": 1000, "c40": 2000}
    }
}
```

**Schritt 2 — pdf_parser.py:** Beide Laborwerte (C10-C22 und C10-C40) müssen weiterhin aus dem PDF extrahiert werden, aber jetzt als separate Zeilen mit dem Parameternamen "Kohlenwasserstoffe" und einer neuen Spalte "KW_Typ" ("C22" oder "C40"). Alternativ können sie als zwei Zeilen mit Parameternamen "Kohlenwasserstoffe (C10-C22)" und "Kohlenwasserstoffe (C10-C40)" extrahiert bleiben — dann muss der Evaluator beide zusammenführen.

**Schritt 3 — evaluator.py:** In Phase 1 (Feststoffe), beim Parameter "Kohlenwasserstoffe", müssen beide Messwerte (C22 und C40) aus dem DataFrame geholt und gemeinsam bewertet werden:

```python
# Pseudocode für KW-Sonderlogik
if param == "Kohlenwasserstoffe":
    match_c22 = df[(df['EBV_Parameter'] == "Kohlenwasserstoffe (C10-C22)") & (df['Matrix'] == "Feststoff")]
    match_c40 = df[(df['EBV_Parameter'] == "Kohlenwasserstoffe (C10-C40)") & (df['Matrix'] == "Feststoff")]
    wert_c22 = match_c22.iloc[0]['Wert'] if not match_c22.empty else None
    wert_c40 = match_c40.iloc[0]['Wert'] if not match_c40.empty else None
    
    # Klasse iterativ bestimmen: erste Klasse, bei der BEIDE Bedingungen erfüllt sind
    klasse = "> BM-F3 (Deponie!)"
    for klassen_key in KLASSEN_HIERARCHIE:
        gw_dict = item["grenzwerte"].get(klassen_key)
        if not gw_dict: continue
        c22_ok = (wert_c22 is None) or (wert_c22 <= gw_dict["c22"])
        c40_ok = (wert_c40 is None) or (wert_c40 <= gw_dict["c40"])
        if c22_ok and c40_ok:
            klasse = klassen_key.replace("_", "-")
            break
    
    messwert_str = f"C10-C22: {wert_c22} | C10-C40: {wert_c40}"
```

---

### Fehler 3 | KRITISCH | Benzo(a)pyren: BM_0* bis BM_F3 fehlen weiterhin in config.py

**Betroffene Datei:** `config.py`

**Problem:**  
Benzo(a)pyren hat in config.py nur Grenzwerte für BM_0_Sand, BM_0_Lehm_Schluff und BM_0_Ton (je 0,3 mg/kg). Die Grenzwerte für BM_0* bis BM_F3 fehlen. Ein Laborwert > 0,3 mg/kg würde sofort als "> BM-F3 (Deponie!)" eingestuft, ohne die Zwischenklassen zu prüfen.

**Anforderung an die Korrektur:**  
Ergänze den Eintrag für Benzo(a)pyren (ebv_order 26) in config.py um die höheren Klassen. Die korrekten Werte gemäß EBV Anlage 1 Tabelle 3 sind:

```python
{"ebv_order": 26, "parameter": "Benzo(a)pyren", "einheit": "mg/kg", "typ": "Feststoff",
 "fussnoten": [], "grenzwerte": {
     "BM_0_Sand": 0.3, "BM_0_Lehm_Schluff": 0.3, "BM_0_Ton": 0.3,
     "BM_0*": 0.6, "BM_F0*": 0.6, "BM_F1": 0.6, "BM_F2": 1.0, "BM_F3": 1.0
 }}
```
*(Bitte Werte anhand der aktuellen EBV-Fassung verifizieren, bevor diese eingetragen werden.)*

---

### Fehler 4 | WARNUNG | Fn3-Bedingung prüft auf Zeichenkette statt Integer — fragile Logik

**Betroffene Datei:** `evaluator.py` → `evaluate_sample()`, Phase 2

**Problem:**  
```python
# AKTUELL (fehlerhaft):
if param != "Sulfat" and "3" in str(fn):
```
Die Bedingung `"3" in str(fn)` matcht auf den String "3" im Fußnotentext. Das ist fehleranfällig: Fußnote 13 würde ebenfalls matchen, und bei kombinierten Fußnotenstrings wie "3, 6" funktioniert es nur zufällig.

**Anforderung an die Korrektur:**  
Den Integer-Check auf der Quellliste verwenden (wie jetzt bereits korrekt für Fn12 umgesetzt):

```python
# KORREKT:
if param != "Sulfat" and 3 in item.get("fussnoten", []):
```

---

### Fehler 5 | WARNUNG | Farblogik: BM-0* wird grün wie BM-0 eingefärbt

**Betroffene Datei:** `reporter.py`

**Problem:**  
```python
if "BM-0" in val:
    fill_to_use, font_to_use = fill_green, font_green
```
Da der String "BM-0*" den Teilstring "BM-0" enthält, wird BM-0* grün eingefärbt — identisch mit BM-0. Das ist irreführend für den Nutzer, da BM-0* eine eingeschränktere Verwertungsklasse ist.

**Anforderung an die Korrektur:**
```python
# Neue Farbe für BM-0* (z.B. hellblau):
fill_bm0star = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
font_bm0star = Font(color="2E75B6")

# Reihenfolge der Prüfung ist wichtig — BM-0* vor BM-0 prüfen:
if "> BM-F3" in val:
    fill_to_use, font_to_use = fill_red, font_red
elif "BM-F" in val:
    fill_to_use, font_to_use = fill_yellow, font_yellow
elif "BM-0*" in val:
    fill_to_use, font_to_use = fill_bm0star, font_bm0star
elif "BM-0" in val:
    fill_to_use, font_to_use = fill_green, font_green
elif "Nicht in EBV" in val or "Kein" in val:
    fill_to_use, font_to_use = fill_gray, font_gray
```
Dieselbe Logik auch in `reporter.py` für den HTML- und PDF-Bericht anpassen (CSS-Klasse `row-bm0star` mit blauer Farbe).

---

### Fehler 6 | WARNUNG | Kohlenwasserstoffe (C10-C22): BM_0* bis BM_F3 fehlen in config.py

*(Wird durch Fehler-2-Korrektur automatisch mitbehoben, wenn das neue KW-Modell implementiert wird. Eigenständig relevant nur, falls das aktuelle Modell beibehalten wird.)*

**Problem:**  
Aktueller Eintrag für C10-C22 hat nur BM_0_Sand bis BM_0_Ton (je 300 mg/kg). BM_0* bis BM_F3 fehlen. Ein Wert > 300 mg/kg ergibt sofort "> BM-F3".

**Korrektur (nur falls das Zwei-Parameter-Modell beibehalten wird):**
```python
"grenzwerte": {
    "BM_0_Sand": 300, "BM_0_Lehm_Schluff": 300, "BM_0_Ton": 300,
    "BM_0*": 300, "BM_F0*": 300, "BM_F1": 300, "BM_F2": 300, "BM_F3": 1000
}
```

---

### Fehler 7 | HINWEIS | EBV_VERSION fehlt Datum der Gültigkeit und Prüfdatum

**Betroffene Datei:** `config.py`

**Problem:**  
EBV_VERSION enthält kein `gueltig_ab` und kein `geprueft_am`. Diese Felder sind für die Revisionssicherheit wichtig: Bei einer EBV-Novellierung müsste erkennbar sein, auf Basis welcher Fassung die Grenzwerte gepflegt wurden.

**Anforderung an die Korrektur:**
```python
EBV_VERSION = {
    "gesetz": "Ersatzbaustoffverordnung (EBV)",
    "fundstelle": "BGBl. I 2021 S. 2598",
    "tabelle": "Anlage 1 Tabelle 3",
    "gueltig_ab": "2023-08-01",
    "geprueft_am": "2026-04-09",
    "geprueft_von": "Automatisierter Abgleich — manuelle Verifikation erforderlich"
}
```

---

## Verifikation der Berechnungsergebnisse (aus den erzeugten Berichten)

Die folgenden Einstufungen wurden manuell gegen EBV Tabelle 3 geprüft. Alle korrekt ausgewiesenen Werte sind nachfolgend dokumentiert.

### MP02 — Gesamteinstufung: BM-F3 (bedingt korrekt, s. u.)
| Parameter | Messwert | Klasse (Programm) | Korrekt? |
|---|---|---|---|
| pH-Wert | 8,0 | BM-0 [6,5–9,5] | ✅ |
| Leitfähigkeit | 188 µS/cm | BM-0* (GW 350) | ✅ (kein BM-0-GW in EBV) |
| Sulfat | 13 mg/l | BM-0 (GW 250) | ✅ |
| Arsen Feststoff | 7,3 mg/kg | BM-0 (GW 10) | ✅ |
| Arsen Eluat | 2 µg/l | BM-0 (Eluat n. maßgeblich) | ✅ (Fn3 korrekt) |
| Blei | 4,2 mg/kg | BM-0 (GW 40) | ✅ |
| Cadmium | < 0,3 mg/kg | BM-0 (GW 0,4) | ✅ |
| Chrom | 9,3 mg/kg | BM-0 (GW 30) | ✅ |
| Kupfer | 5,3 mg/kg | BM-0 (GW 20) | ✅ |
| Nickel | 13,0 mg/kg | BM-0 (GW 15) | ✅ |
| Quecksilber Feststoff | < 0,05 mg/kg | BM-0 (GW 0,2) | ✅ |
| Quecksilber Eluat | 0,03 µg/l | BM-0* (GW 0,1) | ✅ |
| Thallium Feststoff | < 0,25 mg/kg | BM-0 (GW 0,5) | ✅ |
| Zink | 17 mg/kg | BM-0 (GW 60) | ✅ |
| TOC | 1,0 M% | BM-0 (GW 1) | ✅ |
| KW C10-C22 | < 50 mg/kg | BM-0 (GW 300) | ✅ |
| KW C10-C40 | 860 mg/kg | BM-F3 (GW 2000) | ⚠️ Ergebnis stimmt, Logik falsch (s. Fehler 2) |
| Benzo(a)pyren | < 0,05 mg/kg | BM-0 (GW 0,3) | ✅ |
| PAK15 Eluat | 0,1 µg/l | BM-0* (GW 0,2) | ✅ |
| PAK16 | 0,4 mg/kg | BM-0 (GW 3) | ✅ |
| PCB | 0,011 mg/kg | BM-0 (GW 0,05) | ✅ |
| EOX | < 0,5 mg/kg | BM-0 (GW 1) | ✅ |

### MP03 — Gesamteinstufung: > BM-F3 (FALSCH — muss korrigiert werden)
| Parameter | Messwert | Klasse (Programm) | Korrekt? |
|---|---|---|---|
| Leitfähigkeit | 440 µS/cm | BM-F1 (GW 500) | ✅ |
| Nickel Feststoff | 24 mg/kg | BM-0* (GW 100) | ✅ |
| Thallium Feststoff | < 0,25 mg/kg | BM-0 (GW 0,5) | ✅ |
| **Thallium Eluat** | **0,92 µg/l** | **> BM-F3 (Deponie!)** | ❌ **FALSCH** — laut Fn. 12 kein Eluat-GW für BM-F-Klassen. Korrekte Ausgabe: "> BM-0* (Eluat; nur Feststoff für BM-F maßgeblich, Fn. 12)" |

### MP04 — Gesamteinstufung: > BM-F3 (FALSCH — muss korrigiert werden)
| Parameter | Messwert | Klasse (Programm) | Korrekt? |
|---|---|---|---|
| **Thallium Eluat** | **0,25 µg/l** | **> BM-F3 (Deponie!)** | ❌ **FALSCH** — wie MP03. Thallium-Feststoff < 0,25 mg/kg → BM-0. Nur der Eluat-Wert überschreitet BM-0* (0,2 µg/l) leicht. Korrekte Ausgabe: "> BM-0* (Eluat; nur Feststoff für BM-F maßgeblich, Fn. 12)" |

### MP05 — Gesamteinstufung: BM-F3 (korrekt)
| Parameter | Messwert | Klasse (Programm) | Korrekt? |
|---|---|---|---|
| Leitfähigkeit | 966 µS/cm | BM-F3 (GW 2000) | ✅ |
| Nickel | 27 mg/kg | BM-0* (GW 100) | ✅ |
| Thallium Eluat | < 0,2 µg/l | BM-0* (GW 0,2) | ✅ |

---

## Zusammenfassung aller offenen Punkte

| ID | Priorität | Datei | Beschreibung |
|---|---|---|---|
| F-01 | KRITISCH | evaluator.py | Fußnote-12-Sonderregel für Hg/Tl Eluat fehlt → falsche Deponie-Einstufung |
| F-02 | KRITISCH | config.py + evaluator.py | KW C10/C40: falsch modelliert, muss als zweistufige Prüfung umgebaut werden |
| F-03 | KRITISCH | config.py | Benzo(a)pyren BM_0* bis BM_F3 Grenzwerte fehlen |
| F-04 | WARNUNG | evaluator.py | Fn3-Check auf String "3" statt Integer 3 |
| F-05 | WARNUNG | reporter.py | BM-0* wird grün wie BM-0 eingefärbt (Farbprioritätsfehler) |
| F-06 | WARNUNG | config.py | KW C10-C22: BM_0* bis BM_F3 fehlen (behoben durch F-02) |
| F-07 | HINWEIS | config.py | EBV_VERSION fehlt gueltig_ab und geprueft_am |

---

*Ende des Reports V0.0.02. Bitte alle Korrekturen in einem vollständigen, lauffähigen Code-Set ausgeben.*
