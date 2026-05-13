import re
import logging
import pdfplumber
import pandas as pd
from typing import List, Tuple, Optional
from thefuzz import process
from config import ebv_tabelle_3, SYNONYM_MAPPING

VALID_PARAMETERS: List[str] = list(set([item["parameter"] for item in ebv_tabelle_3]))

BLACKLIST: List[str] = [
    "temperatur", "datum", "zeit", "färbung", "geruch", "trübung",
    "fluoranthen", "chrysen", "inden", "anthracen", "pyren", 
    "fluoren", "phenanthren", "acenaphthen", "perylen",
    "pcb nr", "pcb-", "pcb 28", "pcb 52", "pcb 101", "pcb 118", "pcb 138", "pcb 153", "pcb 180"
]
EXACT_BLACKLIST: List[str] = ["naphthalin"]

def map_parameter_name(raw_name: str) -> Optional[str]:
    """Maps a raw parameter name extracted from the PDF to the official EBV parameter name."""
    if not raw_name or not isinstance(raw_name, str): 
        return None
    raw_name_lower = raw_name.strip().lower()
    
    # Reject lines that talk ABOUT a parameter (e.g. "Bei-Temperatur für pH-Wert")
    # — they contain the parameter name but report a different metric.
    if "temperatur" in raw_name_lower or raw_name_lower.startswith("bei"):
        return None
    if "ph-wert" in raw_name_lower or "ph wert" in raw_name_lower or raw_name_lower == "ph": return "pH-Wert"
    if "summe" in raw_name_lower and "pcb" in raw_name_lower: return "PCB6 und PCB-118"
    if "pak" in raw_name_lower and "15" in raw_name_lower: return "PAK15"
    if "pak" in raw_name_lower and "16" in raw_name_lower: return "PAK16"
    # Lab variant "Summe PAK EPA" / "PAK nach EPA" without "16" — treat as PAK16
    # (RuVA-StB 01 + EBV both use the EPA-PAK16 list). Excludes "PAK 15 EPA" cases
    # by checking the "15" sub-rule first above.
    if "pak" in raw_name_lower and "epa" in raw_name_lower: return "PAK16"
    if "summe" in raw_name_lower and "naphthalin" in raw_name_lower: return "Naphthalin und Methylnaphthaline, gesamt"
    
    if raw_name_lower in SYNONYM_MAPPING: return SYNONYM_MAPPING[raw_name_lower]
    if raw_name_lower in EXACT_BLACKLIST: return None
    if any(black_word in raw_name_lower for black_word in BLACKLIST): return None
        
    best_match, score = process.extractOne(raw_name, VALID_PARAMETERS)
    if score >= 90: return best_match
    return None

def determine_matrix(unit: str) -> str:
    """Determines if the material matrix is solid matter (Feststoff) or eluate (Eluat) based on unit."""
    unit_lower = str(unit).lower().strip()
    if "kg" in unit_lower or "%" in unit_lower or "vol" in unit_lower or "ts" in unit_lower or "tr" in unit_lower: 
        return "Feststoff"
    if "l" in unit_lower or "/cm" in unit_lower or unit_lower in ["-", "--", "—", "–"]: 
        return "Eluat"
    return "Unbekannt"

def parse_value_and_unit(clean_data: List[str]) -> Tuple[str, str, Optional[float]]:
    """Parses a cleaned row of PDF text to extract the unit, operator, and numerical value."""
    if len(clean_data) == 1: clean_data = clean_data[0].split()
    if len(clean_data) < 2: return "-", "", None
    
    pruned_data = []
    for token in clean_data:
        token_lower = token.lower()
        if "din " in token_lower or token_lower.startswith("din") or "en iso" in token_lower or "verfahren" in token_lower or token_lower.startswith("iso"):
            break
        pruned_data.append(token)
    clean_data = pruned_data
    
    if len(clean_data) < 2: return "-", "", None
    
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
        
    unit, raw_value = "-", None
    placeholders = ["n.n.", "nn", "n.b.", "n.d.", "nd", "--", "-", "—", "–", "n.a.", "k.a.", "k.a", "<bg"]
    
    for item in merged_data[1:]:
        lower_item = item.lower()
        is_unit = any(u in lower_item for u in ["mg/kg", "µg/l", "mg/l", "vol", "%", "µs/cm", "/cm", "ts", "tr", "os", "m%"]) and not re.search(r'^\d', lower_item)
        
        if is_unit:
            unit = item if unit == "-" else unit + " " + item
            continue
            
        if lower_item in ["-", "--", "—", "–"]:
            if unit == "-":
                unit = item
                continue 
                
        if raw_value is None:
            if re.search(r'\d', item) or lower_item in placeholders:
                raw_value = item

    operator, numeric_value = "", None
    
    if raw_value:
        raw_lower = raw_value.lower()
        if any(p in raw_lower for p in placeholders): operator = "< BG"
        elif "<" in raw_value:
            operator, raw_value = "<", raw_value.replace("<", "")
        elif ">" in raw_value:
            operator, raw_value = ">", raw_value.replace(">", "")
            
        val_str = raw_value.replace(",", ".")
        number_matches = re.findall(r'[-+]?\d*\.\d+|\d+', val_str)
        if number_matches: numeric_value = float(number_matches[0])
    else:
        operator = "< BG"
            
    if unit and unit != "-" and "cm" in unit.lower() and "µs" not in unit.lower():
        unit = "µS/cm"
            
    return unit, operator, numeric_value

def extract_all_data_from_pdf(pdf_path: str) -> pd.DataFrame:
    """Extracts tabular data from a given PDF file and maps it to EBV parameters."""
    extracted_data = []
    table_settings = {"vertical_strategy": "text", "horizontal_strategy": "text", "intersection_y_tolerance": 5, "intersection_x_tolerance": 5}
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables(table_settings):
                    for row in table:
                        clean_row = [str(x).strip() for x in row if x and str(x).strip() != ""]
                        if not clean_row: continue
                        
                        raw_str = clean_row[0]
                        mapped_name = map_parameter_name(raw_str)
                        unit, operator, value = parse_value_and_unit(clean_row)
                        
                        extracted_data.append({
                            "Lab_Original_String": raw_str,
                            "Full_Row": " | ".join(clean_row),
                            "EBV_Parameter": mapped_name if mapped_name else "",
                            "Matrix": determine_matrix(unit),
                            "Lab_Operator": operator,
                            "Lab_Value": value,
                            "Lab_Unit": unit
                        })
        return pd.DataFrame(extracted_data)
    except Exception as e:
        logging.error(f"Error reading PDF {pdf_path}: {e}")
        return pd.DataFrame()
