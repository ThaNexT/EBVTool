import pdfplumber
import pandas as pd
import re
import logging
from thefuzz import process
from config import ebv_tabelle_3, SYNONYM_MAPPING

VALID_PARAMETERS = list(set([item["parameter"] for item in ebv_tabelle_3]))
BLACKLIST = [
    "temperatur", "datum", "zeit", "färbung", "geruch", "trübung",
    "fluoranthen", "chrysen", "inden", "anthracen", "pyren", 
    "fluoren", "phenanthren", "acenaphthen", "perylen"
]

def map_parameter_name(raw_name):
    if not raw_name or not isinstance(raw_name, str): return None
    raw_name_lower = raw_name.strip().lower()
    
    if any(black_word in raw_name_lower for black_word in BLACKLIST): return None
    if raw_name.strip() in SYNONYM_MAPPING: return SYNONYM_MAPPING[raw_name.strip()]
        
    best_match, score = process.extractOne(raw_name, VALID_PARAMETERS)
    if score >= 90: return best_match
    return None

def parse_value_and_unit(rohdaten_zeile):
    if len(rohdaten_zeile) < 3: return None, None, None
        
    einheit = rohdaten_zeile[1]
    wert_index = 2
    if rohdaten_zeile[2] in ['TS', 'TR', 'OS'] and len(rohdaten_zeile) > 3: wert_index = 3
        
    raw_wert = rohdaten_zeile[wert_index]
    operator, numerischer_wert = "", None
    
    if "n.n." in str(raw_wert).lower():
        if len(rohdaten_zeile) > wert_index + 1:
            try:
                val = str(rohdaten_zeile[wert_index + 1]).replace(",", ".")
                treffer = re.findall(r"[-+]?\d*\.\d+|\d+", val)
                if treffer: return einheit, "<", float(treffer[0])
            except: pass
        return einheit, "< BG", None

    if "<" in str(raw_wert):
        operator = "<"
        raw_wert = raw_wert.replace("<", "").strip()
        
    raw_wert = str(raw_wert).replace(",", ".")
    try:
        treffer = re.findall(r"[-+]?\d*\.\d+|\d+", raw_wert)
        if treffer: numerischer_wert = float(treffer[0])
    except Exception: pass 
        
    return einheit, operator, numerischer_wert

def extract_all_data_from_pdf(pdf_path):
    extracted_data = []
    table_settings = {
        "vertical_strategy": "text", "horizontal_strategy": "text",
        "intersection_y_tolerance": 5, "intersection_x_tolerance": 5  
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        clean_row = [str(cell).strip() if cell else "" for cell in row]
                        if not any(clean_row): continue
                        
                        raw_str = clean_row[0]
                        mapped_name = map_parameter_name(raw_str)
                        einheit, operator, wert = parse_value_and_unit(clean_row)
                        
                        extracted_data.append({
                            "Labor_Original_String": raw_str,
                            "Ganze_Zeile": " | ".join(clean_row),
                            "Relevant": "X" if mapped_name else "",
                            "EBV_Parameter": mapped_name if mapped_name else "",
                            "Operator": operator,
                            "Wert": wert,
                            "Einheit": einheit
                        })
        return pd.DataFrame(extracted_data)
    except Exception as e:
        logging.error(f"Fehler beim Lesen von {pdf_path}: {e}")
        return pd.DataFrame()