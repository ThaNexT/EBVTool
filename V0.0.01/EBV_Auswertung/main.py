import os
import argparse
import logging
from pdf_parser import extract_data_from_pdf
from evaluator import evaluate_sample
from reporter import create_excel_report, create_raw_data_log

INPUT_DIR = "input"
OUTPUT_DIR = "output"

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("ebv_auswertung.log", encoding="utf-8"),
            logging.StreamHandler()
        ]
    )

def main():
    setup_logging()
    
    parser = argparse.ArgumentParser(description="Automatisierte EBV-Auswertung für Bodenmaterial")
    parser.add_argument("--bodenart", default="BM_0_Sand", choices=["BM_0_Sand", "BM_0_Lehm_Schluff", "BM_0_Ton"])
    parser.add_argument("--toc", type=float, default=0.1, help="TOC Gehalt in Prozent")
    args = parser.parse_args()

    logging.info("Starte automatisierte EBV-Auswertung...")
    logging.info(f"Referenz-Bodenart: {args.bodenart}, TOC: {args.toc}%")
    
    pdf_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.pdf')]
    if not pdf_files:
        logging.warning("Keine PDFs im Ordner 'input' gefunden.")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(INPUT_DIR, pdf_file)
        logging.info(f"Verarbeite Datei: {pdf_file}")
        
        # 1. Rohdaten extrahieren
        raw_df = extract_data_from_pdf(pdf_path)
        if raw_df.empty:
            logging.warning("Keine extrahierbaren Daten gefunden. Überspringe PDF.")
            continue
            
        # 2. Zur Kontrolle: Simple Textdatei mit Rohdaten-Auszug generieren
        create_raw_data_log(raw_df, OUTPUT_DIR, pdf_file)
            
        # 3. Klassifizierung und Excel-Export
        evaluated_df = evaluate_sample(raw_df, bodenart=args.bodenart, toc_gehalt=args.toc)
        create_excel_report(evaluated_df, OUTPUT_DIR, pdf_file, bodenart=args.bodenart)
        
    logging.info("Verarbeitung aller Dateien abgeschlossen.")

if __name__ == "__main__":
    main()