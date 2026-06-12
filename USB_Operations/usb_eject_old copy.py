import os

# Ejecting the USB drive (Drive E:)
# os.system('powershell $driveEject = New-Object -comObject Shell.Application; $driveEject.Namespace(17).ParseName("""E:""").InvokeVerb("""Eject""")')
os.system(f'powershell "Get-WmiObject -Class Win32_Volume | Where-Object {{ $_.DriveLetter -eq \'E:\' }} | ForEach-Object {{ $_.Dismount(\'false\',\'true\') }}"')
print(f"Ejected drive E:")