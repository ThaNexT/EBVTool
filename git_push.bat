@echo off
REM ====================================================================
REM EBVTool — one-click git commit + push to GitHub.
REM
REM First-time setup:
REM   - Install Git for Windows: https://gitforwindows.org/
REM   - (Optional, recommended) Install GitHub CLI for friction-free auth:
REM       winget install --id GitHub.cli
REM       gh auth login
REM     After that this script runs without any auth prompts.
REM
REM On every run:
REM   - inits the repo if EBVTool isn't already a git repo
REM   - links origin = https://github.com/ThaNexT/EBVTool.git
REM   - writes a .gitignore (input PDFs, run dirs, large templates excluded;
REM     folders kept via .gitkeep so the structure is still pulled)
REM   - stages, commits, and pushes
REM ====================================================================

setlocal
cd /d "%~dp0"

set "REMOTE_URL=https://github.com/ThaNexT/EBVTool.git"

echo === EBVTool git push ===
echo Repo: %cd%
echo Remote: %REMOTE_URL%
echo.

REM Init repo if it's not already one
if not exist ".git" (
    echo Initializing fresh git repository ...
    git init
    git branch -M main
    git remote add origin "%REMOTE_URL%"
) else (
    git remote get-url origin >nul 2>nul
    if errorlevel 1 (
        git remote add origin "%REMOTE_URL%"
    )
)

REM Always re-write .gitignore so the policy stays current.
(
    echo # ---- Generated outputs ^(recreated on every run^) ----
    echo V0.0.06/1_validation/*
    echo V0.0.06/2_output/*
    echo V0.0.05/V0.0.05/1_validation/*
    echo V0.0.05/V0.0.05/2_output/*
    echo.
    echo # ---- Confidential / heavy lab input — never push to GitHub ----
    echo V0.0.06/0_input/EBV/*
    echo V0.0.06/0_input/PAK/*
    echo V0.0.06/0_input/Aggressivität/*
    echo V0.0.05/V0.0.05/0_input/EBV/*
    echo V0.0.05/V0.0.05/0_input/PAK/*
    echo V0.0.05/V0.0.05/0_input/Aggressivität/*
    echo.
    echo # ---- Heavy / proprietary templates ^(distributed out-of-band^) ----
    echo V0.0.06/templates/2604XX_Mantelverordnung.xlsx
    echo V0.0.06/templates/2604XX_Rohdaten ^& Aggressivität.xlsx
    echo V0.0.06/templates/A_4_3_1_Auswertung_Labor.pdf
    echo V0.0.05/V0.0.05/2604XX_Mantelverordnung.xlsx
    echo V0.0.05/V0.0.05/2604XX_Rohdaten ^& Aggressivität.xlsx
    echo V0.0.05/V0.0.05/A_4_3_1_Auswertung_Labor.pdf
    echo.
    echo # ---- Keep .gitkeep files anywhere so empty folders persist ----
    echo !.gitkeep
    echo.
    echo # ---- Python / Office / editor noise ----
    echo __pycache__/
    echo *.pyc
    echo *.tmp
    echo ~$*
    echo .~lock.*
    echo .DS_Store
    echo Thumbs.db
    echo .idea/
    echo .vscode/
) > .gitignore

REM Drop .gitkeep placeholders so the folder skeleton is preserved on clone.
for %%D in (
    "V0.0.06\0_input\EBV"
    "V0.0.06\0_input\PAK"
    "V0.0.06\0_input\Aggressivität"
    "V0.0.06\1_validation"
    "V0.0.06\2_output"
) do (
    if not exist "%%~D\.gitkeep" (
        if exist "%%~D" (
            echo. > "%%~D\.gitkeep"
        )
    )
)

echo.
echo Staging changes ...
git add -A

set /p COMMIT_MSG="Commit message (Enter for default 'v0.0.06 release'): "
if "%COMMIT_MSG%"=="" set "COMMIT_MSG=v0.0.06 release"

git commit -m "%COMMIT_MSG%"

echo.
echo Pushing to %REMOTE_URL% ...
git push -u origin main

if errorlevel 1 (
    echo.
    echo Push was rejected. Most common cause: the GitHub repo already has
    echo history from V0.0.04 that your fresh local repo doesn't share.
    echo.
    echo Quick fix ^(treats local V0.0.06 as the new canonical state^):
    echo    git push --force-with-lease origin main
    echo.
    echo Or merge histories instead:
    echo    git pull origin main --allow-unrelated-histories --no-rebase
    echo    git push origin main
    pause
    exit /b 1
)

echo.
echo Done. Open https://github.com/ThaNexT/EBVTool to verify.
pause
