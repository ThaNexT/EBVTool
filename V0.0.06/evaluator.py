import pandas as pd
import re
from typing import Tuple, Dict, Any, List
from config import ebv_tabelle_3

#: Eluat parameters whose Fn3 cross-reference points to a DIFFERENT Feststoff
#: parameter (per EBV Anlage 1 Tab. 3 footnote 3, sentence 2: "Der Eluatwert
#: für PAK15 und Naphthalin und Methylnaphthaline, gesamt, ist maßgeblich,
#: wenn der Feststoffwert für PAK16 nach Spalte 3 bis 5 überschritten wird.").
ELUAT_FESTSTOFF_XREF: Dict[str, str] = {
    "PAK15": "PAK16",
    "Naphthalin und Methylnaphthaline, gesamt": "PAK16",
}

#: BM-rank used when an Eluat parameter exceeds its only-defined limit
#: (BM_0*) but EBV does not define higher classes for it. The sample's
#: actual higher class is determined by the parent/aggregate parameter
#: (e.g. PAK15 for Naphthalin), so we cap the rank at BM-F0* rather than
#: cascading the sample to "Landfill!".
_BM_F0_STAR_LABEL: str = "> BM-0* (höhere Klassen nicht definiert; übergeordneter Parameter maßgeblich)"

def get_all_gw_string(target_item: Dict[str, Any], toc_gehalt: float, bodenart: str, active_gw: Any) -> str:
    """Creates a string of all class limits, highlighting the applicable one in bold."""
    gws = []
    basis_val = target_item["grenzwerte"].get(bodenart)
    if basis_val is not None:
        val = basis_val if not isinstance(basis_val, list) else f"[{basis_val[0]}-{basis_val[1]}]"
        gws.append(f"<b>{val}</b>" if active_gw == basis_val else str(val))
    
    for klasse in ["BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3"]:
        gw_eintrag = target_item["grenzwerte"].get(klasse)
        if gw_eintrag is None: continue
        
        current_gw = gw_eintrag["klammerwert"] if isinstance(gw_eintrag, dict) and toc_gehalt >= 0.5 else (gw_eintrag["standard"] if isinstance(gw_eintrag, dict) else gw_eintrag)
        val = current_gw if not isinstance(current_gw, list) else f"[{current_gw[0]}-{current_gw[1]}]"
        gws.append(f"<b>{val}</b>" if active_gw == current_gw else str(val))
            
    return " / ".join(gws)

def find_best_class(param_name: str, einheit: str, wert: Any, operator: str, toc_gehalt: float, bodenart: str) -> Tuple[str, Any, List[int], int]:
    """Determines the worst-case class for a single parameter based on EBV limits."""
    target_item = next((item for item in ebv_tabelle_3 if item["parameter"] == param_name and item["einheit"] == einheit), None)
    if not target_item: return "Not in EBV", None, [], 999
        
    ebv_order = target_item["ebv_order"]
    fussnoten = target_item.get("fussnoten", [])
    
    if wert is None or pd.isna(wert) or str(wert).lower() == "nan":
        if str(operator).strip() == "< BG" or "<" in str(operator):
            return "BM-0", None, fussnoten, ebv_order
        return "No Value", None, fussnoten, ebv_order

    wert_float = float(wert)
    
    basis_gw = target_item["grenzwerte"].get(bodenart)
    if basis_gw is not None:
        if (isinstance(basis_gw, list) and basis_gw[0] <= wert_float <= basis_gw[1]) or \
           (not isinstance(basis_gw, list) and (basis_gw >= wert_float or ("<" in str(operator) and basis_gw >= wert_float))):
            return "BM-0", get_all_gw_string(target_item, toc_gehalt, bodenart, basis_gw), fussnoten, ebv_order

    # Sub-threshold convention for Eluat-only parameters (Hg, Tl, PCB6,
    # Naphthalin, …): the BM-0* limit is a COMPLIANCE bound, not a class
    # assignment (per FN12 sentence "Der Eluatwert der Materialklasse BM-0*
    # ist einzuhalten"). So a value ≤ BM-0* limit means BM-0, not BM-0*.
    bm0_sand_lehm_ton_defined = any(
        target_item["grenzwerte"].get(k) is not None
        for k in ("BM_0_Sand", "BM_0_Lehm_Schluff", "BM_0_Ton")
    )
    if not bm0_sand_lehm_ton_defined:
        bm0_star_entry = target_item["grenzwerte"].get("BM_0*")
        if bm0_star_entry is not None:
            bm0_star_limit = (
                bm0_star_entry["klammerwert"]
                if isinstance(bm0_star_entry, dict) and toc_gehalt >= 0.5
                else (bm0_star_entry["standard"] if isinstance(bm0_star_entry, dict) else bm0_star_entry)
            )
            if not isinstance(bm0_star_limit, list):
                if float(bm0_star_limit) >= wert_float or ("<" in str(operator) and float(bm0_star_limit) >= wert_float):
                    return "BM-0", get_all_gw_string(target_item, toc_gehalt, bodenart, bm0_star_limit), fussnoten, ebv_order

    for klasse_key in ["BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3"]:
        gw_eintrag = target_item["grenzwerte"].get(klasse_key)
        if gw_eintrag is None: continue
        
        limit = gw_eintrag["klammerwert"] if isinstance(gw_eintrag, dict) and toc_gehalt >= 0.5 else (gw_eintrag["standard"] if isinstance(gw_eintrag, dict) else gw_eintrag)
        
        if (isinstance(limit, list) and limit[0] <= wert_float <= limit[1]) or \
           (not isinstance(limit, list) and (float(limit) >= wert_float or ("<" in str(operator) and float(limit) >= wert_float))):
            return klasse_key.replace("_", "-"), get_all_gw_string(target_item, toc_gehalt, bodenart, limit), fussnoten, ebv_order

    # Special case: parameter has ONLY a BM_0* limit defined (e.g. Naphthalin,
    # PCB6/118 Eluat). Per EBV the higher-class determination is governed by
    # the aggregate parameter (PAK15 for Naphthalin etc.), so cap at BM-F0*
    # level rather than cascading to "Landfill".
    gw_dict = target_item["grenzwerte"]
    defined_klassen = [k for k in ("BM_0*", "BM_F0*", "BM_F1", "BM_F2", "BM_F3") if gw_dict.get(k) is not None]
    if defined_klassen == ["BM_0*"]:
        return _BM_F0_STAR_LABEL, get_all_gw_string(target_item, toc_gehalt, bodenart, None), fussnoten, ebv_order

    return "> BM-F3 (Landfill!)", get_all_gw_string(target_item, toc_gehalt, bodenart, None), fussnoten, ebv_order

def evaluate_sample(df: pd.DataFrame, bodenart: str = "BM_0_Sand", toc_override: float = -1.0) -> pd.DataFrame:
    """Evaluates all parameters of a sample against EBV limits and applies legal footnotes."""
    if 'Matrix' not in df.columns: df['Matrix'] = "Unbekannt"

    toc_gehalt = 0.1
    if toc_override != -1.0:
        toc_gehalt = toc_override
    else:
        match_toc = df[(df['EBV_Parameter'] == 'TOC') & (df['Matrix'] == 'Feststoff')]
        if not match_toc.empty:
            try: toc_gehalt = float(match_toc.iloc[0]['Wert'])
            except: pass

    results = []
    feststoff_status = {}
    
    for typ in ["Feststoff", "Eluat"]:
        for item in [i for i in ebv_tabelle_3 if i["typ"] == typ]:
            param = item["parameter"]
            match = df[(df['EBV_Parameter'] == param) & (df['Matrix'] == typ)]
            
            wert, operator, ist_einheit = None, "", item["einheit"]
            if not match.empty:
                wert, operator = match.iloc[0]['Wert'], match.iloc[0]['Operator']
                if typ == "Eluat" and item["einheit"] == "µg/l" and "mg/l" in str(match.iloc[0]['Einheit']).lower():
                    wert = float(wert) * 1000 if pd.notna(wert) else None
                    ist_einheit = "µg/l (umger.)"

            klasse, gw_str, fn_list, order = find_best_class(param, item["einheit"], wert, operator, toc_gehalt, bodenart)
            if typ == "Feststoff": feststoff_status[param] = klasse


            if typ == "Eluat":
                # NOTE: per-param FN3 downgrade is intentionally OMITTED at
                # this level — the Eluat CLASS in the per-sample sheet shows
                # the strict threshold result (e.g. Leitfähigkeit 966 →
                # BM-F3 even though it is orientation-only). The "Eluat
                # nicht maßgeblich" decision applies at the Zusammenfassung
                # Gesamt-class stage (reporter._classify_split), where the
                # Feststoff status of each parameter governs whether its
                # Eluat result is allowed to drive the sample's final class.
                pass

                # FN9 cap: PAK15 Eluat alone cannot push the sample beyond
                # BM-F0* ONLY when the parent Feststoff (PAK16) is still in
                # BM-0. Per EBV Anlage 1 Tab. 3 FN3 sentence 2: "Der
                # Eluatwert für PAK15 ... ist maßgeblich, wenn der Feststoff-
                # wert für PAK16 nach Spalte 3 bis 5 überschritten wird."
                # i.e. when Feststoff PAK16 is already in BM-F1/F2/F3, the
                # Eluat PAK15 IS decisive and must cascade to its natural
                # F-class (no cap). When PAK16 Feststoff is BM-0, the cap
                # applies — PAK15 Eluat alone cannot escalate the sample.
                if 9 in fn_list and klasse in ("BM-F0*", "BM-F1", "BM-F2", "BM-F3"):
                    parent_param = ELUAT_FESTSTOFF_XREF.get(param, "PAK16")
                    parent_klasse = feststoff_status.get(parent_param, "")
                    # Cap only when parent Feststoff is BM-0 (or unknown — no
                    # signal that Feststoff escalates). If parent is BM-F class
                    # the Eluat cascades naturally.
                    if parent_klasse == "BM-0" or parent_klasse in ("", "Not in EBV", "No Value"):
                        klasse = "BM-F0* (Eluat cap; PAK16 Feststoff für höhere Klassen maßgeblich)"
                        bm0_gw = item["grenzwerte"].get("BM_0*")
                        dyn_gw = bm0_gw["klammerwert"] if isinstance(bm0_gw, dict) and toc_gehalt >= 0.5 else (bm0_gw["standard"] if isinstance(bm0_gw, dict) else bm0_gw)
                        gw_str = f"<b>{dyn_gw}</b>"

                if 12 in fn_list and klasse == "> BM-F3 (Landfill!)":
                    klasse = "> BM-0* (Eluat; für BM-F nur Feststoff maßgeblich)"
                    bm0_gw = item["grenzwerte"].get("BM_0*")
                    dyn_gw = bm0_gw["klammerwert"] if isinstance(bm0_gw, dict) and toc_gehalt >= 0.5 else (bm0_gw["standard"] if isinstance(bm0_gw, dict) else bm0_gw)
                    gw_str = f"<b>{dyn_gw}</b>"

            is_toc_active = toc_gehalt >= 0.5
            has_klammer = any(isinstance(v, dict) for v in item["grenzwerte"].values())
            italic = (is_toc_active and has_klammer) or (param == "TOC" and is_toc_active)
            fn_str = ", ".join([f"<b>{f}</b>" if (f == 7 and is_toc_active) else str(f) for f in fn_list])

            op_str = "" if pd.isna(operator) or str(operator).lower() == "nan" else str(operator).strip()
            m_str = f"{op_str} {wert}".strip() if pd.notna(wert) else ("< BG" if "<" in op_str else "")

            results.append({
                "Parameter": param, "Einheit": ist_einheit, "Messwert": m_str,
                "Eingestufte Klasse": klasse, "Maßgeblicher GW": gw_str, "Fußnote": fn_str,
                "ebv_order": item["ebv_order"], "Format_Italic": italic
            })

    return pd.DataFrame(results).sort_values("ebv_order").drop(columns="ebv_order")
