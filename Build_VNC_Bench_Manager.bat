@echo off
title VNC Bench Manager - EXE Builder
cd /d "%~dp0"

echo.
echo ==========================================
echo   VNC Bench Manager - EXE Build
echo ==========================================
echo.

python -m pip install pyinstaller
if errorlevel 1 (
    echo.
    echo Failed to install/check PyInstaller.
    pause
    exit /b 1
)

if not exist "VNC_Bench_Manager.py" (
    echo ERROR: VNC_Bench_Manager.py not found.
    pause
    exit /b 1
)

if not exist "VNC_Bench_Manager.ico" (
    echo ERROR: VNC_Bench_Manager.ico not found.
    pause
    exit /b 1
)

echo.
echo Building EXE with VNC icon...
python -m PyInstaller --onefile --windowed --clean --icon="VNC_Bench_Manager.ico" --name="VNC_Bench_Manager" "VNC_Bench_Manager.py"

if errorlevel 1 (
    echo.
    echo EXE BUILD FAILED.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   BUILD COMPLETE
echo ==========================================
echo EXE: %~dp0dist\VNC_Bench_Manager.exe
echo.
pause
