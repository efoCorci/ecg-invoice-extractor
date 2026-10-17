@echo off
title ECG Invoice Extractor App
echo ========================================================
echo         ECG Invoice Extractor to Excel App
echo ========================================================
echo.
echo Starting Web Server...
cd /d "%~dp0"

start "" http://localhost:5000
python app.py
pause
