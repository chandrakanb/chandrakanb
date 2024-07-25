import os
import shutil

# List of folder names to move
folders_to_move = [
	"Bench09_Auto_Flashing_Reports_10th_May",
	"Bench09_Auto_Flashing_Reports_14th_May",
	"Bench09_Auto_Flashing_Reports_15th_May",
	"Bench09_Auto_Flashing_Reports_16th_May",
	"Bench09_Auto_Flashing_Reports_17th_May",
	"Bench09_Auto_Flashing_Reports_21st_May",
	"Bench09_Auto_Flashing_Reports_23rd_May",
	"Bench09_Auto_Flashing_Reports_29th_May",
	"Bench09_Auto_Flashing_Reports_7th_May",
	"Bench09_Auto_Flashing_Reports_8th_May",
	"Bench09_Auto_Flashing_Reports_9th_May"
]

# Define the current path (where the script is located)
current_path = os.path.dirname(os.path.abspath(__file__))

# Define the path for the new folder
new_folder_path = os.path.join(current_path, "new_folder")

# Create the new folder if it doesn't exist
os.makedirs(new_folder_path, exist_ok=True)

# Iterate through the list of folder names and move them to the new folder
for folder in folders_to_move:
    source_path = os.path.join(current_path, folder)
    destination_path = os.path.join(new_folder_path, folder)
    
    try:
        shutil.move(source_path, destination_path)
        print(f"Moved {folder} to new_folder successfully.")
    except Exception as e:
        print(f"Error moving {folder}: {e}")
