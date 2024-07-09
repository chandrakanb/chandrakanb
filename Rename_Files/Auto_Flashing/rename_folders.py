import os
import re

# Define the directory containing the folders as the current working directory
directory = os.getcwd()

# Define the new fixed parts of the folder names
fixed_prefix = "Bench09"
fixed_suffix = "Auto_Flashing_Reports"

# Function to rename folders
def rename_folders(directory):
    for folder_name in os.listdir(directory):
        # Match folders with the naming convention for Bench09 and Bench06
        match = re.match(r"Bench0[69]_AutoFlashing_Reports_(\d{1,2}[a-zA-Z]{2}_\w+)", folder_name)
        if match:
            date_part = match.group(1)
            new_folder_name = f"{fixed_prefix}_{fixed_suffix}_{date_part}"
            old_folder_path = os.path.join(directory, folder_name)
            new_folder_path = os.path.join(directory, new_folder_name)
            os.rename(old_folder_path, new_folder_path)
            print(f"Renamed '{folder_name}' to '{new_folder_name}'")

# Call the function
rename_folders(directory)
