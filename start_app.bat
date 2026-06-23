@echo off
title Coconut Coir BI System - Starting...
echo ================================================
echo   Coconut Coir Sales Prediction System
echo   Starting Flask Server...
echo ================================================
echo.

cd /d "%~dp0"

echo Starting server on http://localhost:5000
echo Please wait...
echo.

start "" "http://localhost:5000"

python app.py

pause
