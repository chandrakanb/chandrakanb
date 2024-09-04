@echo off
set /p "D:\KITE\KITE"

:: Check if the path exists
if exist "D:\KITE\KITE" (
    cd /d "D:\KITE\KITE"
    start cmd
) else (
    echo The path does not exist.
    pause
)
