@echo off
title EBV Auswertungstool
color 0A

echo ==========================================
echo UEBERPRUEFE UND INSTALLIERE BIBLIOTHEKEN...
echo ==========================================
python -m pip install -q pandas openpyxl pdfplumber thefuzz pytest reportlab
echo Alle Bibliotheken sind aktuell!
echo.

:menu
echo ==========================================
echo HAUPTMENUE
echo ==========================================
echo 1 - Schritt 1: PDFs einlesen und Validierungs-Datei erstellen
echo 2 - Schritt 2: Validierte Excel auswerten und Bericht generieren
echo 3 - Beenden
echo.

set /p choice="Waehlen Sie eine Option (1-3): "

if "%choice%"=="1" goto step1
if "%choice%"=="2" goto step2
if "%choice%"=="3" goto end

echo Ungueltige Eingabe.
goto menu

:step1
echo.
python step1_extraktion.py
echo.
pause
goto menu

:step2
echo.
python step2_auswertung.py
echo.
pause
goto menu

:end
exit