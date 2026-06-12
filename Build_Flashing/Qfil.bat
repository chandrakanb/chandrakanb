@echo off
cd /d "D:\Build_Flashing\"
python Qfil.py
ping -n 61 localhost >nul
EXIT /B %ERRORLEVEL%
REM No pause