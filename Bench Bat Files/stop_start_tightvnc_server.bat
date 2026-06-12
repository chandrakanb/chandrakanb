@echo off
:: Stop the TightVNC Server service
net stop "TightVNC Server"

:: Start the TightVNC Server service
net start "TightVNC Server"

:: Notify the user
echo TightVNC Server service has been restarted.
pause
