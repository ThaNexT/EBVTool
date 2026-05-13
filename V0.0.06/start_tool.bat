@echo off
REM EBV Tool V0.0.06 — unified launcher.

setlocal
cd /d "%~dp0"

echo.
echo === EBV Tool V0.0.06 — extraction + classification in one run ===
echo.

python run.py --flow all
if errorlevel 1 (
    echo.
    echo !! Something went wrong. Scroll up for the error message.
    pause
    exit /b 1
)

echo.
echo Output is in 2_output\^<timestamp^>_Evaluation\
echo.
pause
