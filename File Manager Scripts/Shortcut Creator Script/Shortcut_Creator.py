import os
import shutil
import win32com.client

def create_shortcut(src, dest):
    """
    Creates a shortcut to the source file at the destination path.
    
    Parameters:
    - src: Path to the source file.
    - dest: Path where the shortcut will be created.
    """
    try:
        # Create a shell object for shortcut operations
        shell = win32com.client.Dispatch("WScript.Shell")
        
        # Create a shortcut object and set its target path
        shortcut = shell.CreateShortcut(dest)
        shortcut.TargetPath = src
        
        # Save the shortcut
        shortcut.save()
        print(f"    Created shortcut for \"{os.path.basename(src)}\" at \"{dest}\"")
    except Exception as e:
        print(f"    Error creating shortcut for \"{os.path.basename(src)}\": {e}")

def main():
    """
    Main function to manage the process of creating shortcuts, copying directory contents,
    and cleaning up the temporary directory.
    """
    # Get the current directory where the script is running
    current_directory = os.getcwd()

    # Set the path for the destination directory to store shortcuts
    dest_directory = input("    Enter the destination directory: ").strip() #"D:\\temp"
    
    # Ensure the destination directory exists; create it if not
    if not os.path.exists(dest_directory):
        try:
            os.makedirs(dest_directory)
            print(f"    Created destination directory: \"{dest_directory}\"")
        except Exception as e:
            print(f"    Error creating destination directory: {e}")
            return

    # Loop through each file in the current directory
    for filename in os.listdir(current_directory):
        file_path = os.path.join(current_directory, filename)
        # Ensure it's a file and not a directory, and skip specific files
        if os.path.isfile(file_path):
            if filename.lower() in ['readme.txt', 'shortcut_creator.py']:
                continue
            else:
                # Define the path for the shortcut
                shortcut_path = os.path.join(temp_directory, f"{os.path.splitext(filename)[0]}.lnk")
                # Create the shortcut
                create_shortcut(file_path, shortcut_path)

if __name__ == "__main__":
    main()
