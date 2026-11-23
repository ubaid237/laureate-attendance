@echo off
title Laureate Group - Staff Attendance & Payroll Analyzer
color 0A
cls
echo ==============================================================
echo       LAUREATE GROUP - STAFF ATTENDANCE & PAYROLL ANALYZER
echo ==============================================================
echo.
echo Starting application... Please wait.
echo.

cd /d "%~dp0"
streamlit run app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ==============================================================
    echo Streamlit command failed. Trying python -m streamlit...
    echo ==============================================================
    python -m streamlit run app.py
)

pause
