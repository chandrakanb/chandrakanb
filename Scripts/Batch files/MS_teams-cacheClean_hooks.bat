@ECHO OFF

 [+] Description: a simple script to clean up Microsoft Team's Cache (release a bit memory)
 
:main

taskkill /f /t /fi "IMAGENAME eq teams.exe"
echo "[x] Microsoft Team Processes were closed!!!!"
del /f /q "%appdata%\Microsoft\teams\application cache\cache\*.*" > nul 2>&1
del /f /q "%appdata%\Microsoft\teams\blob_storage\*.*" > nul 2>&1
del /f /q "%appdata%\Microsoft\teams\databases\*.*" > nul 2>&1
del /f /q "%appdata%\Microsoft\teams\GPUcache\*.*" > nul 2>&1
del /f /q "%appdata%\Microsoft\teams\IndexdDB\*.db" > nul 2>&1
del /f /q "%appdata%\Microsoft\teams\Local Storage\*.*" > nul 2>&1
del /f /q "%appdata%\Microsoft\teams\tmp\*.*" > nul 2>&1
echo {"enableSso":false,"forceEnableNativeWamAnonFlow":true,"enableSlimCoreLogging":true,"enableSsoLog":true} > %appdata%\Microsoft\Teams\hooks.json
echo "[+] Microsoft Team's Cache was cleaned!!!"
C:\Users\%USERNAME%\AppData\Local\Microsoft\Teams\Update.exe --processStart Teams.exe
echo "[+] Microsoft Team Started!!!"
echo "DONE"