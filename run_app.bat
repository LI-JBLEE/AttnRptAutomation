@echo off
REM GSC Attainment Report Automator - Auto-update and Run Script
REM This script pulls the latest code from GitHub and starts the Streamlit app

echo ========================================
echo GSC Attainment Report Automator
echo ========================================
echo.

REM Change to the script's directory
cd /d "%~dp0"

echo [1/3] Checking for updates from GitHub...
git pull origin main

echo.
echo [2/3] Checking Python dependencies...
pip install -q -r requirements.txt

echo.
echo [3/3] Starting Streamlit app...
echo.
echo The app will open in your browser automatically.
echo Press Ctrl+C to stop the app.
echo.

python -m streamlit run app.py

pause
