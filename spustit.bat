@echo off
REM Dvojklikem spustis aplikaci — otevře se v prohlizeci automaticky

cd /d "%~dp0"

echo === Fight Card Generator ===
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo Python 3 neni nainstalovan.
    echo Stahni ho z https://www.python.org a nainstaluj.
    pause
    exit /b 1
)

echo Instaluji zavislosti...
python -m pip install -r requirements.txt -q

echo.
echo Spoustim aplikaci...
echo Pro ukonceni zavri toto okno nebo stiskni Ctrl+C
echo.

python -m streamlit run app_karta.py --browser.gatherUsageStats false

pause
