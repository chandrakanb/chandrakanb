import os
from datetime import datetime

def get_date_suffix(day):
    # Determine the suffix for the day
    if 4 <= day <= 20 or 24 <= day <= 30:
        suffix = 'th'
    else:
        suffix = ['st', 'nd', 'rd'][day % 10 - 1]
    return suffix

def create_date_folder(base_path):
    # Get current date
    today = datetime.today()

    # Format the date
    day = today.day
    suffix = get_date_suffix(day)
    formatted_date = f"{day}{suffix}_{today.strftime('%B')}_{today.year}"

    # Create the folder path
    global folder_path
    folder_path = os.path.join(base_path, formatted_date)

    # Create the folder
    os.makedirs(folder_path, exist_ok=True)

    print(f"Folder created at: {folder_path}")

# Example Usage
base_path = r"D:\GIT_DATA\Overnight_Sanity_Checklist\Bench09"  # Specify the base path where you want to create the folder
create_date_folder(base_path)

# Open the path in File Explorer
os.startfile(folder_path)

