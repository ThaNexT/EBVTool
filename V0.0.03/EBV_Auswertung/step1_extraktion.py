import os
import pandas as pd
from datetime import datetime
from pdf_parser import extract_all_data_from_pdf
from config import ebv_tabelle_3

INPUT_DIR = "input"

def generate_validation_html(df, output_path, title):
    """Erzeugt eine visuelle HTML-Version der Validierungsdatei."""
    html = f"""
    <!DOCTYPE html>
    <html><head><meta charset="utf-8">
    <style>
        body {{ font-family: Arial, sans-serif; font-size: 13px; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ccc; padding: 6px; text-align: left; }}
        th {{ background-color: #e0e0e0; }}
        .missing {{ background-color: #FFC7CE; color: #9C0006; font-weight: bold; text-align: center; }}
        .found {{ background-color: #C6EFCE; color: #006100; }}
        .unmapped {{ background-color: #f9f9f9; color: #777; }}
    </style></head>
    <body>
        <h2>Validierungs-Übersicht: {title}</h2>
        <p>Rote Felder (X) müssen in der Excel-Datei nachgetragen werden.</p>
        <table>
            <tr><th>X</th><th>EBV Parameter</th><th>Matrix</th><th>Einheit (Soll)</th><th>Ausgelesener String</th><th>Operator</th><th>Wert</th><th>Einheit (Ist)</th></tr>
    """
    for _, row in df.iterrows():
        x_mark = row.get("Fehlend_Bitte_Eintragen", "")
        param = row.get("EBV_Parameter", "")
        
        row_cls = ""
        if param == "": row_cls = "unmapped"
        elif x_mark == "X": row_cls = ""
        else: row_cls = "found"
            
        x_cls = "missing" if x_mark == "X" else ""
        
        html += f"<tr class='{row_cls}'>"
        html += f"<td class='{x_cls}'>{x_mark}</td>"
        html += f"<td>{param}</td>"
        html += f"<td>{row.get('Matrix', '')}</td>"
        html += f"<td>{row.get('EBV_Einheit', '')}</td>"
        html += f"<td>{row.get('Labor_Original_String', '')}</td>"
        html += f"<td>{row.get('Labor_Operator', '')}</td>"
        html += f"<td>{row.get('Labor_Wert', '')}</td>"
        html += f"<td>{row.get('Labor_Einheit', '')}</td>"
        html += "</tr>"
        
    html += "</table></body></html>"
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

def main():
    print("SCHRITT 1: Extraktion & Vorbereitung zur Validierung gestartet...")
    
    pdf_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.pdf')]
    if not pdf_files:
        print("Keine PDFs im Input-Ordner gefunden.")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    val_dir = os.path.join("1_validation", f"{timestamp}_Extraktion")
    os.makedirs(val_dir, exist_ok=True)
    
    out_path_excel = os.path.join(val_dir, "Validierung_Alle_Proben.xlsx")

    ebv_master = pd.DataFrame([{
        "EBV_Sortierung": item["ebv_order"],
        "EBV_Parameter": item["parameter"],
        "Matrix": item["typ"],
        "EBV_Einheit": item["einheit"]
    } for item in ebv_tabelle_3])

    try:
        with pd.ExcelWriter(out_path_excel, engine='openpyxl') as writer:
            for pdf_file in pdf_files:
                print(f"Verarbeite: {pdf_file}")
                pdf_path = os.path.join(INPUT_DIR, pdf_file)
                raw_df = extract_all_data_from_pdf(pdf_path)
                
                if raw_df.empty: continue
                
                raw_unique = raw_df.drop_duplicates(subset=["EBV_Parameter", "Matrix"], keep="first")
                mapped_df = pd.merge(ebv_master, raw_unique, on=["EBV_Parameter", "Matrix"], how="left")
                
                mapped_df["Fehlend_Bitte_Eintragen"] = mapped_df["Labor_Wert"].apply(
                    lambda x: "X" if pd.isna(x) or str(x).strip() == "" else ""
                )
                
                unmapped_df = raw_df[raw_df["EBV_Parameter"] == ""].copy()
                unmapped_df["EBV_Sortierung"] = 999
                unmapped_df["EBV_Einheit"] = "---"
                unmapped_df["Fehlend_Bitte_Eintragen"] = ""
                
                final_view = pd.concat([mapped_df, unmapped_df], ignore_index=True)
                
                cols = ["EBV_Sortierung", "Fehlend_Bitte_Eintragen", "EBV_Parameter", "Matrix", "EBV_Einheit", "Labor_Original_String", "Labor_Operator", "Labor_Wert", "Labor_Einheit"]
                final_view = final_view[cols].sort_values("EBV_Sortierung")
                
                # Excel schreiben
                sheet_name = pdf_file[:30] 
                final_view.to_excel(writer, index=False, sheet_name=sheet_name)
                
                # HTML Visualisierung schreiben
                html_path = os.path.join(val_dir, f"Ansicht_{sheet_name}.html")
                generate_validation_html(final_view, html_path, pdf_file)
                
        print(f"\n--> Validierungsdatei erfolgreich erstellt: {out_path_excel}")
        print("HTML-Ansichten zur schnellen Ueberpruefung wurden im selben Ordner abgelegt.")
        
    except PermissionError:
        print(f"FEHLER: Die Datei {out_path_excel} ist in Excel geoeffnet. Bitte schliessen Sie sie.")

if __name__ == "__main__":
    main()