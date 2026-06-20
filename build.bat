@echo off
chcp 65001 > nul
title 소그룹교안 생성기 — EXE 빌드

echo.
echo ================================================
echo   소그룹교안 생성기 EXE 빌드 스크립트
echo ================================================
echo.

:: Python 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    pause & exit /b 1
)

:: pip 패키지 설치
echo [1/3] 필요 패키지 설치 중...
python -m pip install pyinstaller pillow --quiet
python -m pip install -r "%~dp0requirements.txt" --quiet
echo       완료.
echo.

:: PyInstaller 빌드
echo [2/3] EXE 빌드 중 (1~3분 소요)...
python -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --name "smallgroup-guide-generator" ^
    --hidden-import pdfplumber ^
    --hidden-import docx ^
    --hidden-import anthropic ^
    --hidden-import openai ^
    --hidden-import google.generativeai ^
    --collect-all pdfplumber ^
    --collect-all docx ^
    "%~dp0app.py"

if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
    pause & exit /b 1
)

echo.
echo [3/3] 완료!
echo.
echo   생성 파일: %~dp0dist\smallgroup-guide-generator.exe
echo.
echo   이 EXE 파일 하나만 복사하면 어디서든 실행할 수 있습니다.
echo   (Python 없이도 실행 가능)
echo.
pause
