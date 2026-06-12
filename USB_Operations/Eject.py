import os

def eject_usb_drive():
    drives = [d for d in os.popen("wmic logicaldisk get caption") if ":" in d]
    if not drives:
        print("No USB drives connected.")
        return

    for drive in drives:
        drive_letter = drive.strip()
        print(drive)
        try:
            os.system(f'powershell "Get-WmiObject -Class Win32_Volume | Where-Object {{ $_.DriveLetter -eq \'{drive_letter}:\' }} | ForEach-Object {{ $_.Dismount(\'false\',\'true\') }}"')
            print(f"Ejected drive {drive_letter}")
        except Exception as e:
            print(f"Error ejecting drive {drive_letter}: {e}")

eject_usb_drive()
