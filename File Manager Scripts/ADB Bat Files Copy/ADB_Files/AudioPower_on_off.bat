cd "D:\KITE\KITE\ExternalLib\adb"
ECHO OFF
adb -s %1 shell dumpsys car_service inject-key 181