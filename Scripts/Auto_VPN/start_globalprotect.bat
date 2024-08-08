@echo off
setlocal

REM Start GlobalProtect
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Auto_VPN\GlobalProtect.lnk"

REM Wait for the program to start (5 seconds)
timeout /t 5

REM Run VBScript to send keystrokes
cscript //nologo "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Auto_VPN\SendKeys.vbs"

endlocal
