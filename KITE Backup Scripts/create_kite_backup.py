import os
import shutil
from datetime import datetime

def check_and_create_backup_folder(Backup_Folder):
    """
    Check if the backup folder exists, create it if not.
    """
    # Get the current directory
    current_directory = os.getcwd()
    
    # Construct the full path of the folder
    folder_path = os.path.join(current_directory, Backup_Folder)
    
    # Check if the folder exists
    if not os.path.exists(folder_path):
        # If not, create the folder
        os.makedirs(folder_path)
        print(f"Folder '{Backup_Folder}' created.")
    else:
        print(f"Folder '{Backup_Folder}' already exists.")

def get_ordinal_suffix(day):
    """
    Return the ordinal suffix for a given day.
    """
    if 10 <= day % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return suffix

def get_unique_date_folder_name(date_folder_name):
    """
    Generate a unique folder name by appending a number if the folder already exists.
    """
    counter = 1
    folder_name = date_folder_name
    while os.path.exists(folder_name):
        folder_name = f"{date_folder_name}_{counter}"
        counter += 1
    return folder_name

def check_and_create_folder_with_date(Backup_Folder):
    """
    Create a backup folder with the current date in the format "ddth_Month_yyyy".
    """
    # Get the current date
    current_date = datetime.now()
    
    # Get the day and its ordinal suffix
    day = current_date.day
    int_day = int(day)
    suffix = get_ordinal_suffix(day)
    
    # Format the date as "ddth_Month_yyyy"
    date_folder_name = current_date.strftime(f"{int_day}{suffix}_%B_%Y")
    
    # Construct the full path of the folder
    folder_path = os.path.join(Backup_Folder, date_folder_name)
    
    # Get a unique folder name if the folder already exists
    unique_date_folder_path = get_unique_date_folder_name(folder_path)
    
    # Create the folder
    os.makedirs(unique_date_folder_path)
    print(f"Folder '{unique_date_folder_path}' created.")
    
    return unique_date_folder_path

def copy_files_with_extensions(source_directory, backup_directory, extensions):
    """
    Copy files with specific extensions from the source_directory and its 'KITE_DATA' subdirectory
    to the backup_directory.

    :param source_directory: Path to the source directory.
    :param backup_directory: Path to the backup directory.
    :param extensions: Tuple of file extensions to be copied.
    """
    # Function to copy files from a specific directory
    def copy_from_directory(directory):
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            if os.path.isfile(file_path) and file.endswith(tuple(extensions)):
                # Calculate the destination path
                relative_path = os.path.relpath(directory, source_directory)
                destination_dir = os.path.join(backup_directory, relative_path)
                
                # Create the destination directory if it doesn't exist
                if not os.path.exists(destination_dir):
                    os.makedirs(destination_dir)
                
                # Copy the file to the destination directory
                shutil.copy2(file_path, destination_dir)
                print(f"'{file_path}' Copied.")

    # Copy files from the source directory
    copy_from_directory(source_directory)
    
    # Copy files from the 'KITE_DATA' subdirectory
    kite_data_directory = os.path.join(source_directory, "KITE_DATA")
    if os.path.exists(kite_data_directory):
        copy_from_directory(kite_data_directory)
    else:
        print(f"Directory '{kite_data_directory}' does not exist.")

def copy_license_file(source_directory, backup_directory):
    """
    Copy License file from the 'KITE_DATA\\License' subdirectory to the backup directory.
    
    :param source_directory: Path to the source directory containing the License file.
    """
    Source_License_File_Dir = os.path.join(source_directory, "KITE_DATA", "License")
    Destination_License_File_Dir = os.path.join(backup_directory, "KITE_DATA", "License")
    
    # Copy the LiveStatus folder to the backup directory
    shutil.copytree(Source_License_File_Dir, Destination_License_File_Dir)
    print(f"'{Source_License_File_Dir}' Copied.")

def copy_adb_folder(source_directory, backup_directory):
    """
    Copy License file from the 'KITE_DATA\\License' subdirectory to the backup directory.
    
    :param source_directory: Path to the source directory containing the License file.
    """
    Source_ADB_Folder_Dir = os.path.join(source_directory, "ExternalLib", "adb")
    Destination_ADB_Folder_Dir = os.path.join(backup_directory, "ExternalLib", "adb")
    
    # Copy the LiveStatus folder to the backup directory
    shutil.copytree(Source_ADB_Folder_Dir, Destination_ADB_Folder_Dir)
    print(f"'{Source_ADB_Folder_Dir}' Copied.")

def copy_Live_Status_folder(source_directory, backup_directory):
    """
    Copy the LiveStatus folder from the source path to the backup directory.
    
    :param source_directory: Path to the source directory containing the LiveStatus folder.
    :param backup_directory: Path to the destination directory.
    """
    Source_LiveStatus_Folder_Dir = os.path.join(source_directory, "LiveStatus")
    Destination_LiveStatus_Folder_Dir = os.path.join(backup_directory, "LiveStatus")
    
    # Copy the LiveStatus folder to the backup directory
    shutil.copytree(Source_LiveStatus_Folder_Dir, Destination_LiveStatus_Folder_Dir)
    print(f"'{Source_LiveStatus_Folder_Dir}' Copied.")

def copy_Settings_folder(source_directory, backup_directory):
    """
    Copy License file from the 'KITE_DATA\\License' subdirectory to the backup directory.
    
    :param source_directory: Path to the source directory containing the License file.
    """
    Source_Settings_Folder_Dir = os.path.join(source_directory, "KITE_DATA", "Settings")
    Destination_Settings_Folder_Dir = os.path.join(backup_directory, "KITE_DATA", "Settings")
    
    # Copy the LiveStatus folder to the backup directory
    shutil.copytree(Source_Settings_Folder_Dir, Destination_Settings_Folder_Dir)
    print(f"'{Source_Settings_Folder_Dir}' Copied.")

if __name__ == "__main__":
    # Specify the source directory and file extensions to search for
    source_directory = os.path.join(os.getcwd(), "KITE")
    file_extensions = ('.xlsx', '.bat', '.cfg', '.cfg.bak', '.db', '.py', '.ttf')

    # Define backup folder name and create it if needed
    Backup_Folder = "KITE_BACKUP"
    check_and_create_backup_folder(Backup_Folder)
    
    # Create the backup folder with the current date and get its path
    backup_directory = check_and_create_folder_with_date(Backup_Folder)
    
    # Copy the files to the backup folder
    copy_files_with_extensions(source_directory, backup_directory, file_extensions)

    # Copy the Live Status folder to the backup folder
    copy_Live_Status_folder(source_directory, backup_directory)
    
    # Copy Kite License file to the backup folder
    copy_license_file(source_directory, backup_directory)
    
    # Copy adb folder to the backup folder
    copy_adb_folder(source_directory, backup_directory)
    
    # Copy Settings folder to the backup folder
    copy_Settings_folder(source_directory, backup_directory)
    
    print("KITE Backup Created Successfully.")
    input("Press Enter to exit...")