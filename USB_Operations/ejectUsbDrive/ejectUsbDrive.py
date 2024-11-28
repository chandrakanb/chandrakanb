import os

# Define USB Name
usbName = "Flash_Drive"

# Find drive letter of USB from USB Name
findDriveLetter = (
    f'powershell -ExecutionPolicy Bypass -Command "(Get-Volume | '
    f'Where-Object {{ $_.FileSystemLabel -eq \'{usbName}\' }}).DriveLetter"'
)

# Get drive letter
driveLetter = os.popen(findDriveLetter).read().strip()

# Print drive letter
print(f"Drive letter found: {driveLetter}")

# Check for drive letter
if driveLetter:
    # Eject USB drive
    ejectDrive = (
        f'powershell -ExecutionPolicy Bypass -Command "$driveEject = New-Object '
        f'-comObject Shell.Application; $driveEject.Namespace(17).ParseName(\'{driveLetter}:\').InvokeVerb(\'Eject\')"'
    )
    os.system(ejectDrive)
    print(f"USB drive '{usbName}' with drive letter '{driveLetter}' has been ejected.")
else:
    print(f"USB drive with name '{usbName}' not found.")
