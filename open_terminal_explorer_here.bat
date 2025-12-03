@echo off
start explorer "%~dp0Reports"
timeout /t 10 /nobreak > nul
start cmd /k "%~dp0"
exit