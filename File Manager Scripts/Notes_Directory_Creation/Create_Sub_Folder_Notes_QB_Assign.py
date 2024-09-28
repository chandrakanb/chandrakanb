import os

def create_common_subfolders():
    current_dir = os.getcwd()  # Get the current directory

    # Define the common subfolders to create in each folder
    common_subfolders = ["Notes", "Assignment", "Question_Bank"]

    # Iterate through all folders in the current directory
    for folder in os.listdir(current_dir):
        folder_path = os.path.join(current_dir, folder)

        # Check if the item is a directory
        if os.path.isdir(folder_path):
            # For each folder, create the common subfolders if they don't exist
            for subfolder in common_subfolders:
                subfolder_path = os.path.join(folder_path, subfolder)
                if not os.path.exists(subfolder_path):
                    os.makedirs(subfolder_path)
                    print(f"Created {subfolder} in {folder}")
                else:
                    print(f"{subfolder} already exists in {folder}")

# Run the function
create_common_subfolders()
