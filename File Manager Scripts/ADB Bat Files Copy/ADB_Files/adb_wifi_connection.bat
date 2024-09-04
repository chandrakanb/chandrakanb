cd "D:\KITE\KITE\ExternalLib\adb"
ECHO OFF
adb -s %1 tcpip 5557
adb connect 10.52.147.92:5557


