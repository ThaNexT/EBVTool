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
        # FIX: Fängt Operator "< BG" ab, wenn es keine Zahl gibt (z.B. bei "--" im Laborbericht)
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
    
    # Fallback, falls die Matrix-Spalte aus irgendeinem Grund fehlt
    if 'Matrix' not in df.columns:
        df['Matrix'] = "Unbekannt"
    
    # 1. FESTSTOFFE BEWERTEN
    for item in ebv_tabelle_3:
        if item["typ"] == "Feststoff":
            param = item["parameter"]
            target_einheit = item["einheit"]
            
            # FIX: Zuweisung über Parameter + Matrix, NICHT mehr über die Labor-Einheiten
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

    # 2. ELUATE BEWERTEN
    for item in ebv_tabelle_3:
        if item["typ"] == "Eluat":
            param = item["parameter"]
            target_einheit = item["einheit"]
            
            # FIX: Zuweisung über Parameter + Matrix
            match = df[(df['EBV_Parameter'] == param) & (df['Matrix'] == "Eluat")]
            
            out_einheit = target_einheit
            wert, operator = None, ""

            if not match.empty:
                w = match.iloc[0]['Wert']
                operator = match.iloc[0]['Operator']
                ist_einheit = str(match.iloc[0]['Einheit']).lower()
                
                # Wenn das Labor mg/l geliefert hat, EBV aber µg/l fordert -> umrechnen!
                if target_einheit == "µg/l" and "mg/l" in ist_einheit:
                    wert = float(w) * 1000 if pd.notna(w) else None
                    out_einheit = "µg/l (umger.)"
                else:
                    wert = w

            klasse, gw, fn, order = find_best_class(param, target_einheit, wert, operator, toc_gehalt, bodenart)

            if param != "Sulfat" and "3" in str(fn):
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

if __name__ == "__main__":
    print("Modul 'evaluator.py' erfolgreich geladen.")