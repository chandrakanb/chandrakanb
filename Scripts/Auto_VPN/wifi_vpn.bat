@echo off

:: Define Wi-Fi network names
set "wifi1=C3_HondaUser"
set "wifi2=KPIT-AD-USER"
set "wifi3=TP-Link_0220_5G"
set "wifi4=Omkar"
set "wifi5=Jay"

:: Connect to Wi-Fi networks in priority order
netsh wlan connect name="%wifi1%" >nul 2>&1
if not errorlevel 1 (
    echo Connected to %wifi1%
    goto :eof
)

netsh wlan connect name="%wifi2%" >nul 2>&1
if not errorlevel 1 (
    echo Connected to %wifi2%
    goto :eof
)

netsh wlan connect name="%wifi3%" >nul 2>&1
if not errorlevel 1 (
    echo Connected to %wifi3%
    goto :check_connection
)

netsh wlan connect name="%wifi4%" >nul 2>&1
if not errorlevel 1 (
    echo Connected to %wifi4%
    goto :check_connection
)

netsh wlan connect name="%wifi5%" >nul 2>&1
if not errorlevel 1 (
    echo Connected to %wifi5%
    goto :check_connection
)

:check_connection
:: Execute start_globalprotect.bat here
start "" "D:\OneDrive - KPIT Technologies Ltd\Documents\Scripts\Auto_VPN\start_globalprotect.bat"

:eof