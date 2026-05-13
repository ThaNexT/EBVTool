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
    "fluoren", "phenanthren", "acenaphthen", "perylen",
    "pcb nr", "pcb-", "pcb 28", "pcb 52", "pcb 101", "pcb 118", "pcb 138", "pcb 153", "pcb 180"
]

EXACT_BLACKLIST = ["naphthalin"]

def map_parameter_name(raw_name):
    if not raw_name or not isinstance(raw_name, str): return None
    raw_name_lower = raw_name.strip().lower()
    
    if "ph-wert" in raw_name_lower or "ph wert" in raw_name_lower or raw_name_lower == "ph": return "pH-Wert"
    if "summe" in raw_name_lower and "pcb" in raw_name_lower: return "PCB6 und PCB-118"
    if "pak" in raw_name_lower and "15" in raw_name_lower: return "PAK15"
    if "pak" in raw_name_lower and "16" in raw_name_lower: return "PAK16"
    if "summe" in raw_name_lower and "naphthalin" in raw_name_lower: return "Naphthalin und Methylnaphthaline, gesamt"
    
    if raw_name_lower in SYNONYM_MAPPING: return SYNONYM_MAPPING[raw_name_lower]
    if raw_name_lower in EXACT_BLACKLIST: return None
    if any(black_word in raw_name_lower for black_word in BLACKLIST): return None
        
    best_match, score = process.extractOne(raw_name, VALID_PARAMETERS)
    if score >= 90: return best_match
    return None

def bestimme_matrix(einheit):
    e = str(einheit).lower().strip()
    if "kg" in e or "%" in e or "vol" in e or "ts" in e or "tr" in e: return "Feststoff"
    if "l" in e or "/cm" in e or e in ["-", "--", "—", "–"]: return "Eluat"
    return "Unbekannt"

def parse_value_and_unit(clean_data):
    if len(clean_data) == 1:
        clean_data = clean_data[0].split()
        
    if len(clean_data) < 2: return None, None, None
    
    # FIX 1: Brandmauer gegen Laborverfahren. Sobald "DIN" etc. auftaucht, wird der Rest der Zeile gelöscht.
    pruned_data = []
    for x in clean_data:
        xl = x.lower()
        if "din " in xl or xl.startswith("din") or "en iso" in xl or "verfahren" in xl or xl.startswith("iso"):
            break
        pruned_data.append(x)
    clean_data = pruned_data
    
    if len(clean_data) < 2: return None, None, None
    
    merged_data = [clean_data[0]]
    skip_next = False
    for i in range(1, len(clean_data)):
        if skip_next:
            skip_next = False
            continue
        if clean_data[i] in ["<", ">", "n.n.", "nn", "n.b.", "n.d."] and i+1 < len(clean_data):
            if re.search(r'\d', clean_data[i+1]) or clean_data[i+1] in ["--", "-", "—", "–", "n.a."]:
                merged_data.append(clean_data[i] + " " + clean_data[i+1])
                skip_next = True
                continue
        merged_data.append(clean_data[i])
        
    einheit = "-"
    raw_wert = None
    
    # Erweiterte Platzhalter für PDF-Gedankenstriche
    placeholders = ["n.n.", "nn", "n.b.", "n.d.", "nd", "--", "-", "—", "–", "n.a.", "k.a.", "k.a", "<bg"]
    
    for item in merged_data[1:]:
        lower_item = item.lower()
        
        is_unit = any(u in lower_item for u in ["mg/kg", "µg/l", "mg/l", "vol", "%", "µs/cm", "/cm", "ts", "tr", "os", "m%"]) and not re.search(r'^\d', lower_item)
        if is_unit:
            if einheit == "-": einheit = item
            else: einheit += " " + item
            continue
            
        if lower_item in ["-", "--", "—", "–"]:
            if einheit == "-":
                einheit = item
                continue 
                
        if raw_wert is None:
            if re.search(r'\d', item) or lower_item in placeholders:
                raw_wert = item

    operator, numerischer_wert = "", None
    
    if raw_wert:
        raw_lower = raw_wert.lower()
        # FIX 2: Weist allen Platzhaltern konsequent < BG zu
        if any(p in raw_lower for p in placeholders):
            operator = "< BG"
        elif "<" in raw_wert:
            operator = "<"
            raw_wert = raw_wert.replace("<", "")
        elif ">" in raw_wert:
            operator = ">"
            raw_wert = raw_wert.replace(">", "")
            
        val_str = raw_wert.replace(",", ".")
        treffer = re.findall(r'[-+]?\d*\.\d+|\d+', val_str)
        if treffer:
            numerischer_wert = float(treffer[0])
    else:
        # FIX 3: Wenn die Zeile leer war (z.B. "--" wurde vom PDF verschluckt), 
        # wir aber die DIN-Norm erfolgreich blockiert haben, 
        # gehen wir davon aus, dass der Parameter unauffällig/nicht nachgewiesen war.
        operator = "< BG"
            
    if einheit and einheit != "-":
        if "cm" in einheit.lower() and "µs" not in einheit.lower():
            einheit = "µS/cm"
            
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
                        clean_row = [str(x).strip() for x in row if x and str(x).strip() != ""]
                        if not clean_row: continue
                        
                        raw_str = clean_row[0]
                        mapped_name = map_parameter_name(raw_str)
                        einheit, operator, wert = parse_value_and_unit(clean_row)
                        matrix = bestimme_matrix(einheit)
                        
                        extracted_data.append({
                            "Labor_Original_String": raw_str,
                            "Ganze_Zeile": " | ".join(clean_row),
                            "EBV_Parameter": mapped_name if mapped_name else "",
                            "Matrix": matrix,
                            "Labor_Operator": operator,
                            "Labor_Wert": wert,
                            "Labor_Einheit": einheit
                        })
        return pd.DataFrame(extracted_data)
    except Exception as e:
        logging.error(f"Fehler beim Lesen von {pdf_path}: {e}")
        return pd.DataFrame()