@echo off
for /F "tokens=2 delims= " %%A in ('adb shell ip -f inet addr show wlan0 ^| findstr /r /c:"inet "') do (
    for /F "delims=/" %%B in ("%%A") do set device_ip=%%B
)
adb tcpip 5555
adb connect %device_ip%:5555
pause
