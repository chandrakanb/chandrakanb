import subprocess

def eject_usb_drive(drive_letter):
    # PowerShell script to eject the USB drive
    powershell_command = f'''
    $driveEject = New-Object -comObject Shell.Application;
    $driveEject.Namespace(17).ParseName("{drive_letter}:").InvokeVerb("Eject")
    '''
    
    try:
        # Execute the PowerShell script and capture the output
        result = subprocess.run(
            ["powershell", "-Command", powershell_command],
            capture_output=True, text=True, check=True
        )
        
        # Check if there's any output from the command
        if result.stdout:
            print(f"Success: {result.stdout}")
        else:
            print(f"USB drive {drive_letter}: successfully ejected.")
    
    except subprocess.CalledProcessError as e:
        # Capture any error messages and print them
        print(f"Error ejecting USB drive {drive_letter}: {e.stderr}")

# Eject the USB drive (Drive E)
eject_usb_drive("E")
