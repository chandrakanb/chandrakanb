@echo off
cd /d C:\windows\system32
net stop w32time
w32tm /unregister
w32tm /register
net start w32time
w32tm /resync
