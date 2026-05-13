import os
import argparse
import pandas as pd
from datetime import datetime
from evaluator import evaluate_sample
from reporter import create_excel_report

VALIDATION_DIR = "1_validation"

def main():
    print("SCHRITT 2: Auswertung der validierten Daten gestartet...")
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--bodenart", default="BM_0_Sand", choices=["BM_0_Sand", "BM_0_Lehm_Schluff", "BM_0_Ton"])
    parser.add_argument("--toc", type=float, default=0.1)
    args = parser.parse_args()

    subdirs = [os.path.join(VALIDATION_DIR, d) for d in os.listdir(VALIDATION_DIR) if os.path.isdir(os.path.join(VALIDATION_DIR, d))]
    if not subdirs:
        print("Keine Validierungs-Ordner gefunden. Bitte erst Schritt 1 ausfuehren.")
        return
        
    latest_dir = max(subdirs, key=os.path.getmtime)
    val_path = os.path.join(latest_dir, "Validierung_Alle_Proben.xlsx")
    
    if not os.path.exists(val_path):
        print(f"Datei {val_path} nicht gefunden.")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    out_dir = os.path.join("2_output", f"{timestamp}_Auswertung")
    os.makedirs(out_dir, exist_ok=True)

    print(f"Lese Daten aus: {val_path}")
    
    try:
        excel_tabs = pd.read_excel(val_path, sheet_name=None)
    except PermissionError:
        print("FEHLER: Die Validierungs-Excel ist noch geoeffnet. Bitte schliessen Sie Excel.")
        return

    for sheet_name, df in excel_tabs.items():
        print(f"  - Bewerte Probe: {sheet_name}")
        
        # FIX: Wir löschen nur noch Zeilen, denen der EBV-Parameter fehlt. 
        # So retten wir "< BG" Werte, die keine Zahl (Labor_Wert = nan) haben!
        df_clean = df.dropna(subset=['EBV_Parameter']).copy()
        df_clean = df_clean[df_clean['EBV_Parameter'] != ""]
        
        df_clean = df_clean.rename(columns={
            "Labor_Operator": "Operator",
            "Labor_Wert": "Wert",
            "Labor_Einheit": "Einheit"
        })
        
        evaluated_df = evaluate_sample(df_clean, bodenart=args.bodenart, toc_gehalt=args.toc)
        create_excel_report(evaluated_df, out_dir, f"{sheet_name}.xlsx", bodenart=args.bodenart)
        
    print(f"\nVerarbeitung abgeschlossen. Berichte liegen in: {out_dir}")

if __name__ == "__main__":
    main()