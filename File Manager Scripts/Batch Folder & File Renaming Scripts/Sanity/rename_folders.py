import os
import re

#Sanity_Bench09_9th_July_ChandrakantBhalekar
#Bench09_Sanity_Reports_9th_July_Chandrakant

# Define the directory containing the folders as the current working directory
directory = os.getcwd()

# Define the new fixed parts of the folder names
fixed_prefix = "Bench09"
fixed_suffix = "Sanity_Reports"
test_engineer = "Chandrakant"

# Function to rename folders
def rename_folders(directory):
    for folder_name in os.listdir(directory):
        # Match folders with the naming convention for Bench09 and Bench06
        match = re.match(r"Sanity_Bench09_(\d{1,2}[a-zA-Z]{2}_\w+)_ChandrakantBhalekar", folder_name)
        if match:
            date_part = match.group(1)
            new_folder_name = f"{fixed_prefix}_{fixed_suffix}_{date_part}_{test_engineer}"
            old_folder_path = os.path.join(directory, folder_name)
            new_folder_path = os.path.join(directory, new_folder_name)
            os.rename(old_folder_path, new_folder_path)
            print(f"Renamed '{folder_name}' to '{new_folder_name}'")

# Call the function
rename_folders(directory)
