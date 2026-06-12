import os

# Reconnecting the USB drive using PowerShell
os.system('''
powershell -Command "Get-PnpDevice -PresentOnly | Where-Object { $_.InstanceId -like '*USBSTOR*' } | Enable-PnpDevice -Confirm:$false"
''')