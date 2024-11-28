@echo off
:: Replace <username> with the current user's name or leave %username% to automatically take the current user
net localgroup administrators chandrakanb /delete

::echo Admin rights removed for %username%.
pause
