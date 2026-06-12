import os

def remount_usb(drive_letter, volume_number):
    # PowerShell script to rescan and remount the USB using diskpart
    diskpart_script = f"""
    rescan
    list volume
    select volume {volume_number}
    assign letter={drive_letter}
    """

    # Save the diskpart script to a temporary file
    with open("C:\\temp\\diskpart_script.txt", "w") as script_file:
        script_file.write(diskpart_script)

    # Run the diskpart command with the script
    os.system('diskpart /s C:\\temp\\diskpart_script.txt')
    print(f"USB drive {drive_letter}: remounted.")

# Example usage for drive E: and volume 2
remount_usb("E", 2)
