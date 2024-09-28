import os

def list_folders_in_subfolders():
    current_dir = os.getcwd()  # Get the current directory
    folders_in_subfolders = {}  # Dictionary to store subfolder names for each folder

    # Iterate through all items in the current directory
    for root, dirs, files in os.walk(current_dir):
        if root == current_dir:
            # For each folder in the current directory
            for folder in dirs:
                folder_path = os.path.join(current_dir, folder)
                # List all subfolders within each folder
                subfolders = [subfolder for subfolder in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, subfolder))]
                folders_in_subfolders[folder] = subfolders

    # Print the results
    for folder, subfolders in folders_in_subfolders.items():
        print(f"Folder: {folder}")
        if subfolders:
            for subfolder in subfolders:
                print(f"  - Subfolder: {subfolder}")
        else:
            print("  - No subfolders")
        print()

# Run the function
list_folders_in_subfolders()
