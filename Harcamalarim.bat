@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Harcamalarim
rem The setup window uses only Python's built-in libraries.
py -3 -c "import tkinter, venv; import sys; assert sys.version_info >= (3, 12)" >nul 2>&1
if not errorlevel 1 (
    py -3 launcher.py
    goto finished
)
python -c "import tkinter, venv; import sys; assert sys.version_info >= (3, 12)" >nul 2>&1
if not errorlevel 1 (
    python launcher.py
    goto finished
)
echo Python 3.12 veya sonrasi gerekli.
echo Python kurulumunda "Add python.exe to PATH" ve "tcl/tk" seceneklerini etkinlestirin.
echo Indirme adresi: https://www.python.org/downloads/windows/
pause
exit /b 1
:finished
if errorlevel 1 pause
