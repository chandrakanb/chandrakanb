@echo off
set /p "D:\KITE\KITE\ExternalLib\adb"

:: Check if the path exists
if exist "D:\KITE\KITE\ExternalLib\adb" (
    cd /d "D:\KITE\KITE\ExternalLib\adb"
    start cmd
) else (
    echo The path does not exist.
    pause
)
