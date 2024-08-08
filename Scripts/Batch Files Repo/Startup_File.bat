@echo off

:: WiFi Connection and Auto VPN
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Auto_VPN\wifi_vpn.bat"

:: Wait for batch files to finish
timeout /t 30 /nobreak

:: Excel_Files.bat
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Batch Files Repo\Excel_Files.bat"

:: Wait for batch files to finish
timeout /t 5 /nobreak

:: Startup_Apps.bat
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Batch Files Repo\Startup_Apps.bat"

:: Wait for batch files to finish
timeout /t 5 /nobreak

:: Chrome_Tabs.bat
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Batch Files Repo\Chrome_Tabs.bat"

:: Wait for batch files to finish
timeout /t 30 /nobreak

:: Bench 06.bat
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\RDC Files\Bench 06.bat"

:: Wait for batch files to finish (adjust the timeout value as needed)
timeout /t 30 /nobreak

:: Close all CMD windows
taskkill /f /im cmd.exe
