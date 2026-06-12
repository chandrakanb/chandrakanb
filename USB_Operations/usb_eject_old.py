import os

# Ejecting the USB drive (Drive E:)
os.system('powershell $driveEject = New-Object -comObject Shell.Application; $driveEject.Namespace(17).ParseName("""E:""").InvokeVerb("""Eject""")')
