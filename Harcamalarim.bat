@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Harcamalarim
rem Your private data folder is configured in home.txt or HARCAMA_HOME.
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py
) else (
    python app.py
)
if errorlevel 1 pause
