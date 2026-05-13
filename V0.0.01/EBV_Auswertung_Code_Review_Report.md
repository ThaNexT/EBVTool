# Code-Review-Report: EBV_Auswertung V0.0.01
**Erstellt:** 2026-04-08  
**Geprüft von:** Claude Sonnet 4.6 (Anthropic)  
**Zweck:** Vollständige fachliche, technische und rechtliche Analyse zur Korrektur durch ein LLM  
**Gesamturteil:** ❌ Nicht produktionsreif — kritische Fehler in allen drei Kategorien

---

## Aufgabe für den Empfänger dieses Reports

Bitte korrigiere den unten vollständig abgedruckten Python-Code (5 Module) gemäß den aufgeführten Fehlern und Anforderungen. Der Code ist ein automatisiertes Tool zur Klassifizierung von Bodenmaterial nach der deutschen **Ersatzbaustoffverordnung (EBV), Anlage 1 Tabelle 3**. Die Korrekturen müssen fachlich korrekt, rechtssicher und technisch robust sein.

---

## Vollständiger Quellcode (alle 5 Module)

### main.py
```python
import os
from parser import extract_data_from_pdf
from evaluator import evaluate_sample
from reporter import create_excel_report

INPUT_DIR = "input"
OUTPUT_DIR = "output"

def main():
    print("Starte automatisierte EBV-Auswertung...\n" + "="*40)
    
    pdf_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.pdf')]
    
    if not pdf_files:
        print("Keine PDFs im Ordner 'input' gefunden.")
        return

    # HIER: Wir setzen standardmäßig den Worst-Case (Sand) an
    ziel_bodenart = "BM_0_Sand"
    angenommener_toc = 0.1

    for pdf_file in pdf_files:
        pdf_path = os.path.join(INPUT_DIR, pdf_file)
        print(f"\nVerarbeite Datei: {pdf_file}")
        
        print("  - Extrahiere Daten...")
        raw_df = extract_data_from_pdf(pdf_path)
        
        if raw_df.empty:
            print("  - WARNUNG: Keine Daten gefunden. Überspringe PDF.")
            continue
            
        print(f"  - Klassifiziere Werte (Worst-Case Annahme, TOC: {angenommener_toc}%)...")
        evaluated_df = evaluate_sample(raw_df, bodenart=ziel_bodenart, toc_gehalt=angenommener_toc)
        
        print("  - Generiere Excel-Bericht...")
        create_excel_report(evaluated_df, OUTPUT_DIR, pdf_file)
        
    print("\n" + "="*40 + "\nVerarbeitung aller Dateien abgeschlossen.")

if __name__ == "__main__":
    main()
```

### config.py
```python
# config.py

KLASSEN_HIERARCHIE = [
    "BM_0_Sand", "BM_0_Lehm_Schluff", "BM_0_Ton", 
    "BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3"
]

ebv_tabelle_3 = [
    {"parameter": "Mineralische Fremdbestandteile", "einheit": "Vol.-%", "typ": "Feststoff", "fussnoten": [1], "grenzwerte": {"BM_0_Sand": 10, "BM_0_Lehm_Schluff": 10, "BM_0_Ton": 10, "BM_0*": 10, "BM_F0*": 50, "BM_F1": 50, "BM_F2": 50, "BM_F3": 50}},
    {"parameter": "pH-Wert", "einheit": "-", "typ": "Eluat", "fussnoten": [4], "grenzwerte": {"BM_0*": [6.5, 9.5], "BM_F0*": [6.5, 9.5], "BM_F1": [6.5, 9.5], "BM_F2": [6.5, 9.5], "BM_F3": [5.5, 12.0]}},
    {"parameter": "Elektrische Leitfähigkeit", "einheit": "µS/cm", "typ": "Eluat", "fussnoten": [4], "grenzwerte": {"BM_0*": 350, "BM_F0*": 350, "BM_F1": 500, "BM_F2": 500, "BM_F3": 2000}},
    {"parameter": "Sulfat", "einheit": "mg/l", "typ": "Eluat", "fussnoten": [5], "grenzwerte": {"BM_0_Sand": 250, "BM_0_Lehm_Schluff": 250, "BM_0_Ton": 250, "BM_0*": 250, "BM_F0*": 250, "BM_F1": 450, "BM_F2": 450, "BM_F3": 1000}},
    {"parameter": "Arsen", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 10, "BM_0_Lehm_Schluff": 20, "BM_0_Ton": 20, "BM_0*": 20, "BM_F0*": 40, "BM_F1": 40, "BM_F2": 40, "BM_F3": 150}},
    {"parameter": "Arsen", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 8, "klammerwert": 13}, "BM_F0*": 12, "BM_F1": 20, "BM_F2": 85, "BM_F3": 100}},
    {"parameter": "Blei", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 40, "BM_0_Lehm_Schluff": 70, "BM_0_Ton": 100, "BM_0*": 140, "BM_F0*": 140, "BM_F1": 140, "BM_F2": 140, "BM_F3": 700}},
    {"parameter": "Blei", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 23, "klammerwert": 43}, "BM_F0*": 35, "BM_F1": 90, "BM_F2": 250, "BM_F3": 470}},
    {"parameter": "Cadmium", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [6], "grenzwerte": {"BM_0_Sand": 0.4, "BM_0_Lehm_Schluff": 1.0, "BM_0_Ton": 1.5, "BM_0*": 1.0, "BM_F0*": 2.0, "BM_F1": 2.0, "BM_F2": 2.0, "BM_F3": 10}},
    {"parameter": "Cadmium", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 2, "klammerwert": 4}, "BM_F0*": 3.0, "BM_F1": 3.0, "BM_F2": 10, "BM_F3": 15}},
    {"parameter": "Chrom, gesamt", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 30, "BM_0_Lehm_Schluff": 60, "BM_0_Ton": 100, "BM_0*": 120, "BM_F0*": 120, "BM_F1": 120, "BM_F2": 120, "BM_F3": 600}},
    {"parameter": "Chrom, gesamt", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 10, "klammerwert": 19}, "BM_F0*": 15, "BM_F1": 150, "BM_F2": 290, "BM_F3": 530}},
    {"parameter": "Kupfer", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 20, "BM_0_Lehm_Schluff": 40, "BM_0_Ton": 60, "BM_0*": 80, "BM_F0*": 80, "BM_F1": 80, "BM_F2": 80, "BM_F3": 320}},
    {"parameter": "Kupfer", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 20, "klammerwert": 41}, "BM_F0*": 30, "BM_F1": 110, "BM_F2": 170, "BM_F3": 320}},
    {"parameter": "Zink", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 60, "BM_0_Lehm_Schluff": 150, "BM_0_Ton": 200, "BM_0*": 300, "BM_F0*": 300, "BM_F1": 300, "BM_F2": 300, "BM_F3": 1200}},
    {"parameter": "Zink", "einheit": "µg/l", "typ": "Eluat", "fussnoten": [3], "grenzwerte": {"BM_0*": {"standard": 100, "klammerwert": 210}, "BM_F0*": 150, "BM_F1": 160, "BM_F2": 840, "BM_F3": 1600}},
    {"parameter": "Kohlenwasserstoffe", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [8], "grenzwerte": {"BM_0*": {"standard": 300, "klammerwert": 600}, "BM_F0*": {"standard": 300, "klammerwert": 600}, "BM_F1": {"standard": 300, "klammerwert": 600}, "BM_F2": {"standard": 300, "klammerwert": 600}, "BM_F3": {"standard": 1000, "klammerwert": 2000}}},
    {"parameter": "Benzo(a)pyren", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [], "grenzwerte": {"BM_0_Sand": 0.3, "BM_0_Lehm_Schluff": 0.3, "BM_0_Ton": 0.3}},
    {"parameter": "PAK16", "einheit": "mg/kg", "typ": "Feststoff", "fussnoten": [10], "grenzwerte": {"BM_0_Sand": 3, "BM_0_Lehm_Schluff": 3, "BM_0_Ton": 3, "BM_0*": 6, "BM_F0*": 6, "BM_F1": 6, "BM_F2": 9, "BM_F3": 30}},
]

SYNONYM_MAPPING = {
    "Leitfähigkeit": "Elektrische Leitfähigkeit",
    "LF": "Elektrische Leitfähigkeit",
    "PAK nach EPA": "PAK16",
    "PAK (16)": "PAK16",
    "Benzo(a)pyren": "Benzo(a)pyren",
    "Kohlenwasserstoffe C10 - C40": "Kohlenwasserstoffe",
    "KW-Index": "Kohlenwasserstoffe",
    "pH": "pH-Wert",
    "Chrom": "Chrom, gesamt",
    "Chrom (gesamt)": "Chrom, gesamt"
}
```

### parser.py  
*(Achtung: Dateiname muss geändert werden — siehe Fehler F-T-01)*
```python
import pdfplumber
import pandas as pd
import re
from thefuzz import process
from config import ebv_tabelle_3, SYNONYM_MAPPING

VALID_PARAMETERS = list(set([item["parameter"] for item in ebv_tabelle_3]))

# Erweiterte Blacklist um False-Positives bei einzelnen PAKs zu verhindern
BLACKLIST = [
    "temperatur", "datum", "zeit", "färbung", "geruch", "trübung",
    "fluoranthen", "chrysen", "inden", "anthracen", "pyren", 
    "fluoren", "naphthalin", "phenanthren", "acenaphthen", "perylen"
]

def map_parameter_name(raw_name):
    if not raw_name or not isinstance(raw_name, str):
        return None
    
    raw_name_lower = raw_name.strip().lower()
    
    if any(black_word in raw_name_lower for black_word in BLACKLIST):
        return None
        
    if raw_name.strip() in SYNONYM_MAPPING:
        return SYNONYM_MAPPING[raw_name.strip()]
        
    best_match, score = process.extractOne(raw_name, VALID_PARAMETERS)
    if score >= 90:  # Strengerer Schwellenwert!
        return best_match
    return None

def parse_value_and_unit(rohdaten_zeile):
    if len(rohdaten_zeile) < 3: return None, None, None
        
    einheit = rohdaten_zeile[1]
    wert_index = 2
    
    if rohdaten_zeile[2] in ['TS', 'TR', 'OS'] and len(rohdaten_zeile) > 3:
        wert_index = 3
        
    raw_wert = rohdaten_zeile[wert_index]
    operator = ""
    numerischer_wert = None
    
    if "n.n." in str(raw_wert).lower():
        operator = "<"
        if len(rohdaten_zeile) > wert_index + 1:
            raw_wert = rohdaten_zeile[wert_index + 1]
        else:
            raw_wert = "0"

    if "<" in str(raw_wert):
        operator = "<"
        raw_wert = raw_wert.replace("<", "").strip()
        
    raw_wert = str(raw_wert).replace(",", ".")
    
    try:
        treffer = re.findall(r"[-+]?\d*\.\d+|\d+", raw_wert)
        if treffer: numerischer_wert = float(treffer[0])
    except Exception:
        pass 
        
    return einheit, operator, numerischer_wert

def extract_data_from_pdf(pdf_path):
    extracted_data = []
    table_settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables(table_settings)
            for table in tables:
                for row in table:
                    clean_row = [str(cell).strip() if cell else "" for cell in row]
                    if not any(clean_row): continue
                        
                    raw_param_name = clean_row[0]
                    mapped_name = map_parameter_name(raw_param_name)
                    
                    if mapped_name:
                        einheit, operator, wert = parse_value_and_unit(clean_row)
                        extracted_data.append({
                            "EBV_Parameter": mapped_name,
                            "Einheit": einheit,
                            "Operator": operator,
                            "Wert": wert,
                            "Labor_Original_String": clean_row[0] 
                        })
    return pd.DataFrame(extracted_data)

if __name__ == "__main__":
    print("Modul 'parser.py' erfolgreich geladen.")
```

### evaluator.py
```python
import pandas as pd
from config import ebv_tabelle_3, KLASSEN_HIERARCHIE

def clean_unit(unit_str):
    if not unit_str: return ""
    u = str(unit_str).lower().strip()
    if "g/l" in u and "mg/l" not in u and "kg" not in u: return "µg/l"
    if "mg/kg" in u: return "mg/kg"
    if "mg/l" in u: return "mg/l"
    if "--" in u: return "-"
    if "/cm" in u and "µs" not in u: return "µS/cm"
    return str(unit_str).strip()

def find_best_class(param_name, einheit, wert, operator, toc_gehalt, bodenart="BM_0_Sand"):
    if pd.isna(wert) or wert is None:
        return "Kein Messwert", None, ""

    wert_float = float(wert)
    target_ebv_item = None
    
    for item in ebv_tabelle_3:
        if item["parameter"] == param_name:
            if item["einheit"] == einheit:
                target_ebv_item = item
                break
            elif item["einheit"] == "µg/l" and einheit == "mg/l":
                wert_float = wert_float * 1000
                einheit = "µg/l"
                target_ebv_item = item
                break
            
    if not target_ebv_item:
        return "Nicht in EBV Tabelle 3", None, ""
        
    fussnoten_str = ", ".join(map(str, target_ebv_item.get("fussnoten", [])))

    # 1. BM-0 Basisklasse prüfen (wir geben nur noch "BM-0" zurück)
    basis_gw = target_ebv_item["grenzwerte"].get(bodenart)
    if basis_gw is not None:
        if basis_gw >= wert_float or (operator == "<" and basis_gw >= wert_float):
            return "BM-0", basis_gw, fussnoten_str
            
    # 2. Höhere Klassen prüfen
    for klasse in ["BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3"]:
        gw_eintrag = target_ebv_item["grenzwerte"].get(klasse)
        if gw_eintrag is None:
            continue 
            
        grenzwert = None
        if isinstance(gw_eintrag, dict):
            grenzwert = gw_eintrag["klammerwert"] if toc_gehalt >= 0.5 else gw_eintrag["standard"]
        elif isinstance(gw_eintrag, list):
            if gw_eintrag[0] <= wert_float <= gw_eintrag[1]:
                return klasse.replace("_", "-"), f"[{gw_eintrag[0]} - {gw_eintrag[1]}]", fussnoten_str
            continue
        else:
            grenzwert = float(gw_eintrag)

        if grenzwert >= wert_float or (operator == "<" and grenzwert >= wert_float):
            return klasse.replace("_", "-"), grenzwert, fussnoten_str

    return "> BM-F3 (Deponie!)", None, fussnoten_str

def evaluate_sample(df, bodenart="BM_0_Sand", toc_gehalt=0.1):
    results = []
    for index, row in df.iterrows():
        param_name = row['EBV_Parameter']
        einheit = clean_unit(row['Einheit'])
        wert = row['Wert']
        operator = row['Operator']
        
        eingestufte_klasse, massgeblicher_gw, fussnoten = find_best_class(param_name, einheit, wert, operator, toc_gehalt, bodenart)
        
        if einheit == "mg/l" and eingestufte_klasse != "Nicht in EBV Tabelle 3" and param_name != "Sulfat":
             einheit_out = "µg/l (umger.)"
             wert_out = float(wert) * 1000
        else:
             einheit_out = einheit
             wert_out = wert
             
        results.append({
            "Parameter": param_name,
            "Einheit": einheit_out,
            "Messwert": f"{operator} {wert_out}".strip(),
            "Eingestufte Klasse": eingestufte_klasse,
            "Maßgeblicher GW": massgeblicher_gw,
            "Fußnote": fussnoten
        })
    return pd.DataFrame(results)

if __name__ == "__main__":
    print("Modul 'evaluator.py' erfolgreich geladen.")
```

### reporter.py
```python
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

FUSSNOTEN_TEXTE = [
    "1: Die Materialwerte gelten für Bodenmaterial und Baggergut mit bis zu 10 Volumenprozent (BM und BG) oder bis zu 50 Volumenprozent (BM-F und BG-F) mineralischer Fremdbestandteile im Sinne von § 2 Nummer 8 der Bundes-Bodenschutz- und Altlastenverordnung mit nur vernachlässigbaren Anteilen an Störstoffen im Sinne von § 2 Nummer 9 der Bundes-Bodenschutz- und Altlastenverordnung.",
    "2: Bodenarten-Hauptgruppen gemäß Bodenkundlicher Kartieranleitung, 5. Auflage, Hannover 2005 (KA5); stark schluffige Sande, lehmig-schluffige Sande und stark lehmige Sande sowie Materialien, die nicht bodenartspezifisch zugeordnet werden können, sind entsprechend der Bodenart Lehm, Schluff zu bewerten.",
    "3: Die Eluatwerte in Spalte 6 sind mit Ausnahme des Eluatwertes für Sulfat nur maßgeblich, wenn für den betreffenden Stoff der jeweilige Feststoffwert nach Spalte 3 bis 5 überschritten wird.",
    "4: Stoffspezifischer Orientierungswert; bei Abweichungen ist die Ursache zu prüfen.",
    "5: Bei Überschreitung des Wertes ist die Ursache zu prüfen.",
    "6: Der Wert 1 mg/kg gilt für Sand und Lehm/Schluff. Für Ton gilt der Wert 1,5 mg/kg.",
    "8: Die angegebenen Werte gelten für Kohlenwasserstoffverbindungen mit einer Kettenlänge von C10 bis C22.",
    "10: PAK16: stellvertretend für die Gruppe der polyzyklischen aromatischen Kohlenwasserstoffe (PAK) werden 16 ausgewählte PAK nach EPA-Liste untersucht.",
]

def create_excel_report(df, output_dir, original_filename):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    base_name = os.path.splitext(original_filename)[0]
    output_filename = f"Klassifizierung_{base_name}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='EBV_Klassifizierung', startrow=2)
        
    wb = load_workbook(output_path)
    ws = wb['EBV_Klassifizierung']
    
    ws.cell(row=1, column=1).value = "HINWEIS: Für BM-0 wurde standardmäßig der 'Worst-Case' (Sand) angesetzt. Bei Lehm/Schluff oder Ton gelten ggf. höhere Grenzwerte."
    ws.cell(row=1, column=1).font = Font(bold=True, color="FF0000")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    
    fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_green = Font(color="006100")
    fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    font_yellow = Font(color="9C5700")
    fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_red = Font(color="9C0006")
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    font_gray = Font(color="7A7A7A")
    
    for col_idx in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col_idx)
        max_length = 0
        for row_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value:
                try:
                    if len(str(cell_value)) > max_length:
                        max_length = len(str(cell_value))
                except:
                    pass
        ws.column_dimensions[column_letter].width = min(max_length + 3, 50)

    for row in range(4, ws.max_row + 1):
        klasse_cell = ws.cell(row=row, column=4)
        val = str(klasse_cell.value)
        fill_to_use, font_to_use = None, None
        if "BM-0" in val:
            fill_to_use, font_to_use = fill_green, font_green
        elif "BM-F" in val:
            fill_to_use, font_to_use = fill_yellow, font_yellow
        elif "> BM-F3" in val:
            fill_to_use, font_to_use = fill_red, font_red
        elif "Nicht in EBV" in val or "Kein" in val:
            fill_to_use, font_to_use = fill_gray, font_gray
        if fill_to_use:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill_to_use
                ws.cell(row=row, column=col).font = font_to_use

    start_row = ws.max_row + 3
    ws.cell(row=start_row, column=1).value = "Regelbezüge & Fußnoten (Anlage 1 Tabelle 3 EBV):"
    ws.cell(row=start_row, column=1).font = Font(bold=True)
    for i, text in enumerate(FUSSNOTEN_TEXTE):
        cell = ws.cell(row=start_row + 1 + i, column=1)
        cell.value = text
        ws.merge_cells(start_row=start_row + 1 + i, start_column=1, end_row=start_row + 1 + i, end_column=6)
        cell.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[start_row + 1 + i].height = 30

    wb.save(output_path)
    print(f"  --> Excel-Bericht gespeichert: {output_path}")

if __name__ == "__main__":
    print("Modul 'reporter.py' erfolgreich geladen.")
```

---

## Fehlerliste mit Korrekturen

### Kategorie 1: Fachliche Fehler (EBV-Regelwerk)

---

#### F-F-01 | KRITISCH | Eluat-Grenzwertlogik ignoriert Fußnote 3

**Betroffene Datei:** `evaluator.py` → `evaluate_sample()` und `find_best_class()`

**Problem:**  
Laut EBV Anlage 1 Tabelle 3, Fußnote 3 sind Eluat-Grenzwerte (außer Sulfat) nur dann maßgeblich, wenn der jeweilige Feststoffwert überschritten wird. Der aktuelle Code prüft Eluat- und Feststoffwerte vollständig unabhängig voneinander. Das führt dazu, dass ein Material z. B. wegen eines Eluat-Arsen-Wertes schlechter eingestuft wird, obwohl der Feststoff-Arsen-Wert unkritisch ist — was fachlich falsch ist.

**Anforderung an die Korrektur:**
1. In `evaluate_sample()` muss die Auswertung in zwei Phasen erfolgen:
   - Phase 1: Alle Feststoff-Parameter klassifizieren und die schlechteste Klasse je Parameter merken.
   - Phase 2: Eluat-Parameter nur dann prüfen und klassifizieren, wenn für denselben Stoff in Phase 1 der Feststoffwert die BM-0-Klasse überschritten hat. Ausnahme: Sulfat (Eluat) wird immer geprüft.
2. Der Ausgabe-DataFrame muss eine zusätzliche Spalte `"Eluat geprüft"` (True/False) erhalten, die dokumentiert, ob die Eluat-Prüfung aktiviert wurde.
3. Parameter, bei denen keine Feststoffdaten vorhanden sind (z. B. pH-Wert, Leitfähigkeit), werden weiterhin direkt geprüft, da sie keinen Feststoff-Gegenwert haben.

---

#### F-F-02 | KRITISCH | pH-Wert: Fehlende Grenzwerte für BM-0 Basisklassen

**Betroffene Datei:** `config.py`

**Problem:**  
Der pH-Wert ist in `ebv_tabelle_3` nur für `BM_0*` bis `BM_F3` definiert. Der Orientierungswert [6,5–9,5] gilt jedoch auch für BM-0 (Sand, Lehm/Schluff, Ton) als Eluatwert. Ohne diese Einträge gibt das Programm „Nicht in EBV Tabelle 3" zurück, was irreführend ist.

**Anforderung an die Korrektur:**  
Ergänze in `ebv_tabelle_3` für den Eintrag pH-Wert folgende Grenzwerte:
```
"BM_0_Sand": [6.5, 9.5],
"BM_0_Lehm_Schluff": [6.5, 9.5],
"BM_0_Ton": [6.5, 9.5],
```

---

#### F-F-03 | KRITISCH | Benzo(a)pyren: Fehlende Grenzwerte für BM_0* und BM-F-Klassen

**Betroffene Datei:** `config.py`

**Problem:**  
Benzo(a)pyren ist in `ebv_tabelle_3` nur für `BM_0_Sand`, `BM_0_Lehm_Schluff` und `BM_0_Ton` definiert. Gemäß EBV Tabelle 3 existieren jedoch Werte für alle Klassen. Das Fehlen führt zu der Ausgabe „Nicht in EBV Tabelle 3" für Proben, die die BM-0-Werte überschreiten — die tatsächliche Einordnung (z. B. BM-F3: 1 mg/kg) bleibt unsichtbar.

**Anforderung an die Korrektur:**  
Ergänze in `ebv_tabelle_3` für Benzo(a)pyren folgende Grenzwerte (bitte anhand der aktuellen EBV-Tabelle verifizieren und nachtragen):
```
"BM_0*": 0.6,
"BM_F0*": 0.6,
"BM_F1": 0.6,
"BM_F2": 1.0,
"BM_F3": 1.0,
```
*(Werte sind Beispielwerte — der Implementierer muss diese anhand der EBV verifizieren)*

---

#### F-F-04 | KRITISCH | Fehlende EBV-Parameter: Quecksilber, Nickel, Thallium, MKW u. a.

**Betroffene Datei:** `config.py`

**Problem:**  
Die EBV Anlage 1 Tabelle 3 enthält mehr Parameter als die aktuell 19 Einträge in `ebv_tabelle_3`. Fehlende Parameter werden stillschweigend ignoriert. Bei einer Laborprobe mit Quecksilber-Belastung würde das Programm keinen Befund ausgeben.

**Anforderung an die Korrektur:**  
Folgende Parameter müssen mit korrekten Grenzwerten (Feststoff und/oder Eluat) ergänzt werden:
- Quecksilber (Feststoff mg/kg + Eluat µg/l, Fußnote 12 beachten)
- Nickel (Feststoff mg/kg + Eluat µg/l)
- Thallium (Feststoff mg/kg, Fußnote 12 beachten)
- Molybdän (Feststoff mg/kg + Eluat µg/l)
- Antimon (Feststoff mg/kg + Eluat µg/l)
- Vanadium (Feststoff mg/kg + Eluat µg/l)
- MKW / Mineralöl C10–C40 (sofern nicht identisch mit Kohlenwasserstoffe-Eintrag)
- Cyanid, gesamt (Feststoff + Eluat)

Alle Werte anhand der offiziellen EBV (BGBl. 2021 I S. 2598, zuletzt geändert) nachtragen.

---

#### F-F-05 | WARNUNG | Bodenart hardcodiert — keine Nutzereingabe

**Betroffene Datei:** `main.py`

**Problem:**  
`ziel_bodenart = "BM_0_Sand"` ist fest im Code verankert. Für Proben aus Lehm- oder Tonböden gelten z. T. deutlich höhere Grenzwerte. Ein Nutzer kann die Bodenart nicht ohne Code-Änderung anpassen.

**Anforderung an die Korrektur:**  
`main.py` soll die Bodenart als Kommandozeilenargument oder interaktive Abfrage entgegennehmen:
- Gültige Werte: `BM_0_Sand`, `BM_0_Lehm_Schluff`, `BM_0_Ton`
- Bei ungültigem oder fehlendem Wert: Fallback auf `BM_0_Sand` mit expliziter Warnung.
- Beispiel mit `argparse`:
  ```python
  import argparse
  parser_args = argparse.ArgumentParser()
  parser_args.add_argument("--bodenart", default="BM_0_Sand", choices=["BM_0_Sand","BM_0_Lehm_Schluff","BM_0_Ton"])
  parser_args.add_argument("--toc", type=float, default=0.1)
  args = parser_args.parse_args()
  ```

---

#### F-F-06 | WARNUNG | n.n.-Behandlung: Wert 0 statt Bestimmungsgrenze

**Betroffene Datei:** `parser.py` → `parse_value_and_unit()`

**Problem:**  
Bei „n.n." (nicht nachweisbar) wird `raw_wert = "0"` gesetzt, wenn kein Folgewert vorhanden ist. Korrekt wäre es, den Wert als kleiner-als-Bestimmungsgrenze (`< BG`) zu behandeln. Der aktuelle Wert 0 ist zwar konservativ, aber technisch inkorrekt.

**Anforderung an die Korrektur:**  
Wenn „n.n." vorkommt und kein numerischer Folgewert verfügbar ist, soll der Wert `None` bleiben und der Operator auf `"<BG"` gesetzt werden. Im Ausgabe-DataFrame soll dann `"< Bestimmungsgrenze"` erscheinen, nicht `"0"`.

---

### Kategorie 2: Technische Fehler

---

#### F-T-01 | KRITISCH | Namenskonflikt: parser.py mit Python-Builtin

**Betroffene Dateien:** `parser.py`, `main.py`

**Problem:**  
`parser` ist ein eingebautes Python-Standardmodul. In Python 3.9+ kann `from parser import extract_data_from_pdf` zu einem `ImportError` führen, da das Builtin-Modul Vorrang haben kann. Dies ist ein latentem Laufzeitfehler.

**Anforderung an die Korrektur:**
1. Datei `parser.py` umbenennen in `pdf_parser.py`.
2. In `main.py` den Import anpassen:
   ```python
   from pdf_parser import extract_data_from_pdf
   ```

---

#### F-T-02 | KRITISCH | Doppelte Einheitenkonvertierung mg/l → µg/l

**Betroffene Datei:** `evaluator.py` → `find_best_class()` und `evaluate_sample()`

**Problem:**  
In `find_best_class()` wird `wert_float *= 1000` durchgeführt (lokale Variable). In `evaluate_sample()` wird dann erneut `wert_out = float(wert) * 1000` berechnet — aber `wert` ist das originale (nicht konvertierte) `row['Wert']`. Das Ergebnis stimmt zufällig, weil beide Operationen unabhängig auf das Original zugreifen. Jedoch ist die Logik undurchsichtig, fehleranfällig und schwer wartbar. Wenn `find_best_class()` jemals den konvertierten Wert zurückgeben würde, käme es zur Doppelkonvertierung (×1.000.000).

**Anforderung an die Korrektur:**  
Die Einheitenkonvertierung soll an genau einer Stelle erfolgen — in `evaluate_sample()`, bevor `find_best_class()` aufgerufen wird. `find_best_class()` soll keinen internen Unit-Conversion-Sonderfall mehr enthalten. Der bereits konvertierte Wert und die normalisierte Einheit werden dann direkt übergeben.

---

#### F-T-03 | KRITISCH | Keine Fehlerbehandlung bei PDF-Öffnung

**Betroffene Datei:** `parser.py` (bzw. `pdf_parser.py`) → `extract_data_from_pdf()`

**Problem:**  
`pdfplumber.open(pdf_path)` hat kein `try/except`. Ein korruptes, passwortgeschütztes oder leeres PDF bricht das gesamte Programm mit einem unkontrollierten Traceback ab.

**Anforderung an die Korrektur:**
```python
def extract_data_from_pdf(pdf_path):
    extracted_data = []
    table_settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # ... bestehende Logik ...
    except Exception as e:
        logging.error(f"Fehler beim Lesen von {pdf_path}: {e}")
        return pd.DataFrame()
    return pd.DataFrame(extracted_data)
```

---

#### F-T-04 | WARNUNG | Kein requirements.txt vorhanden

**Betroffene Datei:** (fehlt komplett)

**Problem:**  
Die externen Abhängigkeiten `pdfplumber`, `thefuzz`, `openpyxl`, `pandas` sind nicht dokumentiert. `python-levenshtein` (Performance-Beschleunigung für thefuzz) fehlt ebenfalls.

**Anforderung an die Korrektur:**  
Erstelle eine Datei `requirements.txt` mit folgendem Inhalt (Versionen ggf. anpassen):
```
pdfplumber>=0.10.0
thefuzz>=0.20.0
python-levenshtein>=0.21.0
openpyxl>=3.1.0
pandas>=2.0.0
```

---

#### F-T-05 | WARNUNG | Kein Logging — nur print()-Ausgaben

**Betroffene Dateien:** alle Module

**Problem:**  
Alle Statusmeldungen erfolgen über `print()`. Bei automatisiertem Betrieb oder Fehlersuche ist dies unzureichend. Es gibt keine persistente Log-Datei.

**Anforderung an die Korrektur:**  
Füge am Anfang von `main.py` folgende Logging-Konfiguration ein und ersetze alle `print()`-Aufrufe durch `logging.info()` / `logging.warning()` / `logging.error()`:
```python
import logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("ebv_auswertung.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
```

---

#### F-T-06 | HINWEIS | Fuzzy-Matching ohne Konfidenzprotokoll

**Betroffene Datei:** `parser.py` → `map_parameter_name()`

**Problem:**  
Wenn `thefuzz` mit Score ≥ 90 einen Parameter erkennt, wird dies stillschweigend akzeptiert. Bei einer falschen Zuordnung (z. B. „Arsen" → „Arsen" Score 92, aber eigentlich gemeint war ein anderer Parameter) gibt es keine Warnung.

**Anforderung an die Korrektur:**  
Jede Fuzzy-Zuordnung (nicht die exakten SYNONYM_MAPPING-Treffer) soll mit `logging.warning()` protokolliert werden:
```python
logging.warning(f"Fuzzy-Match: '{raw_name}' -> '{best_match}' (Score: {score}). Bitte prüfen!")
```

---

### Kategorie 3: Rechtliche & Compliance-Anforderungen

---

#### F-R-01 | KRITISCH | Kein Haftungsausschluss im Excel-Output

**Betroffene Datei:** `reporter.py` → `create_excel_report()`

**Problem:**  
EBV-Klassifizierungen sind rechtsverbindliche Einstufungen nach § 12 KrWG / EBV. Der aktuelle Excel-Output enthält keinen Hinweis, dass das Ergebnis maschinell erstellt wurde und einer fachgutachterlichen Prüfung bedarf. Im Schadensfall (Fehlklassifizierung → unzulässige Verwertung) besteht erhebliches Haftungsrisiko.

**Anforderung an die Korrektur:**  
Füge in `create_excel_report()` einen zweiten Hinweis-Header (Zeile 2, vor der Datentabelle) ein:
```python
ws.cell(row=2, column=1).value = (
    "RECHTLICHER HINWEIS: Diese Auswertung wurde automatisiert erstellt und ersetzt NICHT "
    "die gutachterliche Prüfung durch eine zugelassene Stelle gemäß EBV. "
    "Maßgeblich ist ausschließlich die Ersatzbaustoffverordnung (EBV) vom 09.07.2021 "
    "(BGBl. I S. 2598) in ihrer jeweils gültigen Fassung. Auswertungsdatum: {datum}"
).format(datum=datetime.date.today().isoformat())
ws.cell(row=2, column=1).font = Font(bold=True, color="9C0006")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
```
Außerdem: `import datetime` am Anfang von `reporter.py` ergänzen. Die Datentabelle muss entsprechend auf `startrow=3` verschoben werden (und alle Row-Indizes um 1 erhöht).

---

#### F-R-02 | KRITISCH | Keine Versionierung der EBV-Tabelle in config.py

**Betroffene Datei:** `config.py`

**Problem:**  
Die Grenzwerte in `ebv_tabelle_3` enthalten keinen Zeitstempel oder Quellenangabe. Bei einer Novellierung der EBV würden veraltete Grenzwerte silent weiterverwendet.

**Anforderung an die Korrektur:**  
Füge am Anfang von `config.py` folgende Metadaten ein:
```python
EBV_VERSION = {
    "gesetz": "Ersatzbaustoffverordnung (EBV)",
    "fundstelle": "BGBl. I 2021 S. 2598",
    "gueltig_ab": "2023-08-01",
    "tabelle": "Anlage 1 Tabelle 3",
    "geprueft_am": "2026-04-08",
    "geprueft_von": "Bitte nachtragen"
}
```
Diese Metadaten sollen auch im Excel-Bericht in einer separaten Zelle erscheinen.

---

#### F-R-03 | WARNUNG | Auswertungsdatum fehlt im Excel-Output

**Betroffene Datei:** `reporter.py`

**Problem:**  
Der Excel-Bericht enthält kein Erstellungsdatum. Für die Nachvollziehbarkeit (z. B. bei späteren Behördenprüfungen) ist das Datum der Auswertung unverzichtbar.

**Anforderung an die Korrektur:**  
Das aktuelle Datum (`datetime.date.today()`) soll im Header des Excel-Bericht erscheinen (siehe auch F-R-01).

---

### Kategorie 4: Fehlende Tests

---

#### F-V-01 | KRITISCH | Keine automatisierten Tests vorhanden

**Problem:**  
Es gibt kein `tests/`-Verzeichnis, keine `pytest`-Konfiguration, keine Unit-Tests.

**Anforderung an die Korrektur:**  
Erstelle eine Datei `tests/test_evaluator.py` mit mindestens folgenden Test-Cases:

```python
import pytest
from evaluator import find_best_class, evaluate_sample
import pandas as pd

# Test 1: Arsen Feststoff unter BM-0-Sand-Grenzwert
def test_arsen_feststoff_bm0():
    klasse, gw, fn = find_best_class("Arsen", "mg/kg", 5.0, "", 0.1, "BM_0_Sand")
    assert klasse == "BM-0"
    assert gw == 10

# Test 2: Arsen Feststoff überschreitet BM-0, fällt in BM_F0*
def test_arsen_feststoff_bmf0():
    klasse, gw, fn = find_best_class("Arsen", "mg/kg", 35.0, "", 0.1, "BM_0_Sand")
    assert klasse == "BM-F0*"
    assert gw == 40

# Test 3: Arsen Feststoff über BM-F3 → Deponie
def test_arsen_feststoff_deponie():
    klasse, gw, fn = find_best_class("Arsen", "mg/kg", 200.0, "", 0.1, "BM_0_Sand")
    assert "> BM-F3" in klasse

# Test 4: Operator < behandlung
def test_arsen_kleiner_operator():
    klasse, gw, fn = find_best_class("Arsen", "mg/kg", 10.0, "<", 0.1, "BM_0_Sand")
    assert klasse == "BM-0"

# Test 5: Kein Messwert
def test_kein_messwert():
    klasse, gw, fn = find_best_class("Arsen", "mg/kg", None, "", 0.1, "BM_0_Sand")
    assert klasse == "Kein Messwert"

# Test 6: pH Bereichswert
def test_ph_im_bereich():
    klasse, gw, fn = find_best_class("pH-Wert", "-", 7.5, "", 0.1, "BM_0_Sand")
    assert "BM" in klasse

# Test 7: TOC-Klammerwert-Logik
def test_toc_klammerwert():
    klasse_niedrig, gw_niedrig, _ = find_best_class("Arsen", "µg/l", 10.0, "", 0.1)
    klasse_hoch, gw_hoch, _ = find_best_class("Arsen", "µg/l", 10.0, "", 0.6)
    # Bei TOC>=0,5 gilt Klammerwert (13 statt 8), daher Klasse gleich oder besser
    assert gw_hoch >= gw_niedrig
```

---

## Zusammenfassung der Änderungen

| ID | Priorität | Datei | Art |
|---|---|---|---|
| F-F-01 | KRITISCH | evaluator.py | Fachlich |
| F-F-02 | KRITISCH | config.py | Fachlich |
| F-F-03 | KRITISCH | config.py | Fachlich |
| F-F-04 | KRITISCH | config.py | Fachlich |
| F-F-05 | WARNUNG | main.py | Fachlich |
| F-F-06 | WARNUNG | parser.py | Fachlich |
| F-T-01 | KRITISCH | parser.py, main.py | Technisch |
| F-T-02 | KRITISCH | evaluator.py | Technisch |
| F-T-03 | KRITISCH | parser.py | Technisch |
| F-T-04 | WARNUNG | (neu) | Technisch |
| F-T-05 | WARNUNG | alle | Technisch |
| F-T-06 | HINWEIS | parser.py | Technisch |
| F-R-01 | KRITISCH | reporter.py | Rechtlich |
| F-R-02 | KRITISCH | config.py | Rechtlich |
| F-R-03 | WARNUNG | reporter.py | Rechtlich |
| F-V-01 | KRITISCH | (neu) | Tests |

---

*Ende des Reports. Bitte alle Korrekturen in einem vollständigen, lauffähigen Code-Set ausgeben.*
