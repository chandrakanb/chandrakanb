@echo off
set /p "D:\KITE_DATA\GIT_Resources\OBU_A14_Repo\DRT_Jenkins\Test_data\GoldenData"

:: Check if the path exists
if exist "D:\KITE_DATA\GIT_Resources\OBU_A14_Repo\DRT_Jenkins\Test_data\GoldenData" (
    cd /d "D:\KITE_DATA\GIT_Resources\OBU_A14_Repo\DRT_Jenkins\Test_data\GoldenData"
    start cmd
) else (
    echo The path does not exist.
    pause
)
