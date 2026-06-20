@echo off
chcp 65001 > nul
title Small Group Guide Generator

:: Run EXE if available
if exist "%~dp0dist\smallgroup-guide-generator.exe" (
    start "" "%~dp0dist\smallgroup-guide-generator.exe"
    exit /b
)
if exist "%~dp0smallgroup-guide-generator.exe" (
    start "" "%~dp0smallgroup-guide-generator.exe"
    exit /b
)

:: Fallback: run via Python
echo EXE not found, running via Python...
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [Error] Python is not installed.
    echo Please install Python 3.10+ or run build.bat to create the EXE.
    pause
    exit /b 1
)

python -m pip install -r "%~dp0requirements.txt" -q
python "%~dp0app.py"
