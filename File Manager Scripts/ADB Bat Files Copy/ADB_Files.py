import os
import shutil

# List of files to copy
files_to_copy = [
    r"D:\KITE\KITE\ExternalLib\adb\install.bat",
    r"D:\KITE\KITE\ExternalLib\adb\volume_dec_by_1.bat",
    r"D:\KITE\KITE\ExternalLib\adb\volume_dec_by_20.bat",
    r"D:\KITE\KITE\ExternalLib\adb\volume_inc_by_1.bat",
    r"D:\KITE\KITE\ExternalLib\adb\volume_inc_by_20.bat",
    r"D:\KITE\KITE\ExternalLib\adb\AudioPower_on_off.bat",
    r"D:\KITE\KITE\ExternalLib\adb\adb_wifi_connection.bat"
]

# Create a folder called "ADB_Files" in the same directory as the script
destination_folder = os.path.join(os.getcwd(), "ADB_Files")

# Create the directory if it doesn't exist
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# Copy the files to the "ADB_Files" folder
for file in files_to_copy:
    try:
        shutil.copy(file, destination_folder)
        print(f"Copied {file} to {destination_folder}")
    except Exception as e:
        print(f"Failed to copy {file}: {e}")
