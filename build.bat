@echo off
chcp 65001 > nul
title 내 주식 대시보드 v1.1.0 — 빌드

echo.
echo ================================================
echo   내 주식 대시보드 v1.1.0  빌드 스크립트
echo ================================================
echo.

:: ── 1. Python 확인 ──────────────────────────────────────
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo   https://www.python.org 에서 Python 3.10+ 를 설치해주세요.
    pause
    exit /b 1
)

:: ── 2. PyInstaller 설치/확인 ─────────────────────────────
echo [1/4] PyInstaller 확인 중...
python -c "import PyInstaller" > nul 2>&1
if errorlevel 1 (
    echo        PyInstaller 설치 중...
    pip install pyinstaller --quiet
)
echo        완료

:: ── 3. 의존성 확인 ──────────────────────────────────────
echo [2/4] 필수 패키지 확인 중...
pip install -r requirements.txt --quiet
echo        완료

:: ── 4. 이전 빌드 정리 ────────────────────────────────────
echo [3/4] 이전 빌드 파일 정리 중...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo        완료

:: ── 5. PyInstaller 빌드 ──────────────────────────────────
echo [4/4] PyInstaller 빌드 시작... (수 분 소요)
echo.
pyinstaller dashboard.spec --noconfirm

if errorlevel 1 (
    echo.
    echo [오류] PyInstaller 빌드 실패!
    pause
    exit /b 1
)

echo.
echo ================================================
echo   빌드 완료!  dist\StockDashboard\ 폴더 확인
echo ================================================
echo.

:: ── 6. Inno Setup 설치기 생성 (선택) ────────────────────
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if exist %ISCC% (
    echo Inno Setup 설치기 생성 중...
    %ISCC% installer.iss
    if not errorlevel 1 (
        echo.
        echo 설치기 생성 완료: StockDashboard_v1.1.0_Setup.exe
    )
) else (
    echo.
    echo [안내] Inno Setup 미설치 — EXE 설치기 생성 건너뜀
    echo        https://jrsoftware.org/isinfo.php 에서 설치 후
    echo        다시 실행하면 설치기(.exe)까지 자동 생성됩니다.
)

echo.
echo 작업 폴더:  dist\StockDashboard\StockDashboard.exe
echo 설치기:     StockDashboard_v1.1.0_Setup.exe  (Inno Setup 필요)
echo.
pause
