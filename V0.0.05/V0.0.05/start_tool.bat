@echo off
title EBV Classification Tool
color 0A

echo ==========================================
echo INITIALIZING WORKSPACE...
echo ==========================================
if not exist "0_input"                      mkdir "0_input"
if not exist "0_input\EBV"                  mkdir "0_input\EBV"
if not exist "0_input\Aggressivitaet"       mkdir "0_input\Aggressivitaet"
if not exist "0_input\PAK"                  mkdir "0_input\PAK"
if not exist "1_validation"                 mkdir "1_validation"
if not exist "1_validation\EBV"             mkdir "1_validation\EBV"
if not exist "1_validation\Aggressivitaet"  mkdir "1_validation\Aggressivitaet"
if not exist "1_validation\PAK"             mkdir "1_validation\PAK"
if not exist "2_output"                     mkdir "2_output"
if not exist "2_output\EBV"                 mkdir "2_output\EBV"
if not exist "2_output\Aggressivitaet"      mkdir "2_output\Aggressivitaet"
if not exist "2_output\PAK"                 mkdir "2_output\PAK"
if not exist "templates"                    mkdir "templates"
echo Directories verified.

echo.
echo Checking and installing dependencies...
python -m pip install -q pandas openpyxl pdfplumber thefuzz pytest reportlab
echo All dependencies are up to date!
echo.

:menu
echo ==========================================
echo MAIN MENU - Three flows: EBV / Aggressivitaet / PAK
echo ==========================================
echo 1 - Step 1 + Step 2 (all three flows)
echo 2 - Step 1 only (all three flows)
echo 3 - Step 2 only (all three flows)
echo 4 - Step 1 EBV only
echo 5 - Step 2 EBV only
echo 6 - Step 1 PAK only
echo 7 - Step 2 PAK only
echo 8 - Step 1 Aggressivitaet only
echo 9 - Step 2 Aggressivitaet only
echo 0 - Exit
echo.

set /p choice="Select an option (0-9): "

if "%choice%"=="1" goto both_all
if "%choice%"=="2" goto step1_all
if "%choice%"=="3" goto step2_all
if "%choice%"=="4" goto step1_ebv
if "%choice%"=="5" goto step2_ebv
if "%choice%"=="6" goto step1_pak
if "%choice%"=="7" goto step2_pak
if "%choice%"=="8" goto step1_aggr
if "%choice%"=="9" goto step2_aggr
if "%choice%"=="0" goto end

echo Invalid input.
goto menu

:both_all
python step1_extraktion.py --flow all
python step2_auswertung.py --flow all
pause
goto menu

:step1_all
python step1_extraktion.py --flow all
pause
goto menu

:step2_all
python step2_auswertung.py --flow all
pause
goto menu

:step1_ebv
python step1_extraktion.py --flow ebv
pause
goto menu

:step2_ebv
python step2_auswertung.py --flow ebv
pause
goto menu

:step1_pak
python step1_extraktion.py --flow pak
pause
goto menu

:step2_pak
python step2_auswertung.py --flow pak
pause
goto menu

:step1_aggr
python step1_extraktion.py --flow aggressivität
pause
goto menu

:step2_aggr
python step2_auswertung.py --flow aggressivität
pause
goto menu

:end
exit
