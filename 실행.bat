@echo off
chcp 65001 > nul
title 소그룹교안 생성기

:: EXE가 있으면 바로 실행
if exist "%~dp0소그룹교안생성기.exe" (
    start "" "%~dp0소그룹교안생성기.exe"
    exit /b
)

:: EXE가 없으면 Python으로 실행 (개발/백업용)
echo EXE를 찾을 수 없어 Python으로 실행합니다...
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python 3.10 이상을 설치하거나 빌드EXE.bat을 실행해 EXE를 생성하세요.
    pause
    exit /b 1
)

python -m pip install -r "%~dp0requirements.txt" -q
python "%~dp0교안생성기.py"
