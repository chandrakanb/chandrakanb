import os

def remove_usb_drive(drive_letter):
    try:
        # PowerShell command to safely remove the USB drive without fully ejecting it
        remove_cmd = f'powershell "Get-WmiObject -Query \\"SELECT * FROM Win32_Volume WHERE DriveLetter = \'{drive_letter}:\'\\" | ForEach-Object {{ $_.Dismount($false, $true) }}"'
        os.system(remove_cmd)
        print(f"Drive {drive_letter} removed (but not ejected).")

    except Exception as e:
        print(f"Error removing drive {drive_letter}: {e}")

# Example usage for drive E:
remove_usb_drive('E')
