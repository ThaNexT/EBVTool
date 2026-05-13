import pandas as pd
import logging
from config import ebv_tabelle_3

def clean_unit(unit_str):
    if not unit_str: return ""
    u = str(unit_str).lower().strip()
    if "g/l" in u and "mg/l" not in u and "kg" not in u: return "µg/l"
    if "mg/kg" in u: return "mg/kg"
    if "mg/l" in u: return "mg/l"
    if "--" in u: return "-"
    if "/cm" in u and "µs" not in u: return "µS/cm"
    if "vol" in u: return "Vol.-%"
    if "%" in u or "massen" in u or "gew" in u or "ts" in u: return "M%"
    return str(unit_str).strip()

def find_best_class(param_name, einheit, wert, operator, toc_gehalt, bodenart):
    target_item = next((item for item in ebv_tabelle_3 if item["parameter"] == param_name and item["einheit"] == einheit), None)
    if not target_item:
        return "Nicht in EBV Tabelle 3", None, "", 999
        
    fussnoten_str = ", ".join(map(str, target_item.get("fussnoten", [])))
    ebv_order = target_item["ebv_order"]

    is_invalid = wert is None or pd.isna(wert) or str(wert).lower() == "nan"

    if is_invalid:
        if operator == "< BG" or "<" in str(operator): return "Kein Messwert (< BG)", None, fussnoten_str, ebv_order
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
    
    if not df.empty:
        df['Einheit_Clean'] = df['Einheit'].apply(clean_unit)
    else:
        df = pd.DataFrame(columns=['EBV_Parameter', 'Einheit_Clean', 'Wert', 'Operator'])
    
    for item in ebv_tabelle_3:
        if item["typ"] == "Feststoff":
            param = item["parameter"]
            target_einheit = item["einheit"]
            match = df[(df['EBV_Parameter'] == param) & (df['Einheit_Clean'] == target_einheit)]
            
            wert, operator = None, ""
            if not match.empty:
                wert = match.iloc[0]['Wert']
                operator = match.iloc[0]['Operator']
                
            klasse, gw, fn, order = find_best_class(param, target_einheit, wert, operator, toc_gehalt, bodenart)
            feststoff_status[param] = klasse
            
            messwert_str = ""
            if not (wert is None or pd.isna(wert) or str(wert).lower() == "nan"):
                messwert_str = f"{operator} {wert}".replace("None", "").strip()
            elif operator == "< BG" or "<" in str(operator):
                messwert_str = "< BG"

            results.append({
                "Parameter": param, "Einheit": target_einheit, "Messwert": messwert_str,
                "Eingestufte Klasse": klasse, "Maßgeblicher GW": gw, "Fußnote": fn, "ebv_order": order
            })

    for item in ebv_tabelle_3:
        if item["typ"] == "Eluat":
            param = item["parameter"]
            target_einheit = item["einheit"]
            match = df[(df['EBV_Parameter'] == param) & (df['Einheit_Clean'] == target_einheit)]
            
            out_einheit = target_einheit
            wert, operator = None, ""

            if match.empty and target_einheit == "µg/l":
                match_mg = df[(df['EBV_Parameter'] == param) & (df['Einheit_Clean'] == "mg/l")]
                if not match_mg.empty:
                    w = match_mg.iloc[0]['Wert']
                    wert = float(w) * 1000 if pd.notna(w) else None
                    operator = match_mg.iloc[0]['Operator']
                    out_einheit = "µg/l (umger.)"
            elif not match.empty:
                wert = match.iloc[0]['Wert']
                operator = match.iloc[0]['Operator']

            klasse, gw, fn, order = find_best_class(param, target_einheit, wert, operator, toc_gehalt, bodenart)

            if param != "Sulfat" and "3" in str(fn):
                fs_klasse = feststoff_status.get(param, "Fehlt")
                if fs_klasse in ["BM-0", "Kein Messwert (< BG)"]:
                    klasse = "BM-0 (Eluat n. maßgeblich)"
                elif fs_klasse == "Kein Messwert":
                    fn += " | HINWEIS: Feststoffwert fehlt zur sicheren FN 3 Prüfung."

            messwert_str = ""
            if not (wert is None or pd.isna(wert) or str(wert).lower() == "nan"):
                messwert_str = f"{operator} {wert}".replace("None", "").strip()
            elif operator == "< BG" or "<" in str(operator):
                messwert_str = "< BG"

            results.append({
                "Parameter": param, "Einheit": out_einheit, "Messwert": messwert_str,
                "Eingestufte Klasse": klasse, "Maßgeblicher GW": gw, "Fußnote": fn, "ebv_order": order
            })

    res_df = pd.DataFrame(results)
    if not res_df.empty:
        res_df = res_df.sort_values(by="ebv_order").drop(columns=["ebv_order"])
    return res_df

if __name__ == "__main__":
    print("Modul 'evaluator.py' erfolgreich geladen.")