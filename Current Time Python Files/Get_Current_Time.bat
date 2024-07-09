@echo off
:: To set batch file directory to the current directory
cd /d "%~dp0"

:: To run the Python script to get the current time
python Get_Current_Time.py

:: To prevent the script from pausing
REM No pause