@echo off

rem Start TightVNC Viewer and connect to the VNC server
start "" "C:\Program Files\TightVNC\tvnviewer.exe" -host=10.52.52.32 -port=5900

rem Change directory to the location where the VBScript is located
cd /d "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\TightVNC Files"

rem Run the VBScript to input the password
wscript.exe Password_DRT@123.vbs
