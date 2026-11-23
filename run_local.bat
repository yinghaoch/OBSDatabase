@echo off
title HLK OBS Database
cd /d "%~dp0"
python -m streamlit run HLK_OBS_Database.py
if errorlevel 1 (
    echo.
    echo Failed to start the app. Check Python and installed dependencies.
    pause
)
