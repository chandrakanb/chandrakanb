import os
import string

# Function to check if the drive exists and then eject it
def eject_drive(drive_letter):
    if os.path.exists(f"{drive_letter}:\\"):
        print(f"Ejecting drive {drive_letter}:")
        # Eject the drive using PowerShell
        os.system(f'powershell -Command "$driveEject = New-Object -comObject Shell.Application; $driveEject.Namespace(17).ParseName(\'{drive_letter}:\').InvokeVerb(\'Eject\')"')
    else:
        print(f"Drive {drive_letter}: not found.")

# Loop through drives from E: to Z: and eject if connected
for drive_letter in string.ascii_uppercase[4:]:
    eject_drive(drive_letter)
