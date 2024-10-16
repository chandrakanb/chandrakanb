@echo off
cd /d "%~dp0reports"
start explorer "%cd%"
timeout /t 3 /nobreak
cd /d "%~dp0"
start cmd