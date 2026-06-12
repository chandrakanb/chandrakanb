import os

def eject_usb_drive():
    # PowerShell script to get only removable (USB) drives
    powershell_script = """
    Get-Volume | Where-Object { $_.DriveType -eq 'Removable' } | Select-Object -ExpandProperty DriveLetter
    """
    
    # Execute PowerShell script and get the list of removable drives
    drive_letters = os.popen(f'powershell -Command "{powershell_script}"').read().splitlines()

    if not drive_letters:
        print("No USB drives connected.")
        return

    # Loop through each drive and eject it
    for drive_letter in drive_letters:
        try:
            eject_command = f'powershell "Get-Volume -DriveLetter {drive_letter} | ForEach-Object {{ $_.Dismount(\'false\',\'true\') }}"'
            os.system(eject_command)
            print(f"Ejected USB drive {drive_letter}:")
        except Exception as e:
            print(f"Error ejecting USB drive {drive_letter}: {e}")

eject_usb_drive()
