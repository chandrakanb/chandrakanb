# Date-Based Folder Creation Script

This Python script automatically creates a folder named based on the current date, including a proper day suffix (e.g., `1st`, `2nd`, `3rd`). The folder is created in a specified base directory and is formatted in the structure `DaySuffix_Month_Year`.

## Features

- **Automatic Folder Naming**: The folder name is based on the current date, using the format `DaySuffix_Month_Year`. For example, `1st_September_2024`.
- **Day Suffix**: The script correctly appends the appropriate suffix to the day of the month (e.g., `1st`, `2nd`, `3rd`, `4th`, etc.).
- **Customizable Base Path**: You can specify the base directory where the folder should be created.
- **Auto Folder Creation**: The script automatically creates the folder, handling any existing directories gracefully using `exist_ok=True`.

## Requirements

- **Python 3.x** must be installed on your system.

## Installation

1. Ensure Python 3.x is installed on your machine.
2. Clone or download the script (`create_date_folder.py`) to your desired location.

## Usage

1. **Configure the Base Path**:
   Modify the `base_path` variable in the script to point to the directory where you want the date-named folder to be created. For example:
   
   ```python
   base_path = r"D:\GIT_DATA\GIT_Sanity\build\Bench09"
