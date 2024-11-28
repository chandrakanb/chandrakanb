@echo off
:: Replace <username> with the current user's name or leave %username% to automatically take the current user
net localgroup administrators chandrakanb /add

::echo Admin rights granted to %username%.
pause
