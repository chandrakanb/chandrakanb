import os
import shutil
import subprocess

def delete_pendrive_contents(destination_path):
    try:
        # List all files and directories in the folder
        for root, dirs, files in os.walk(destination_path, topdown=False):
            for name in files:
                file_path = os.path.join(root, name)
                try:
                    os.remove(file_path)  # Remove the file
                except FileNotFoundError:
                    print(f"File {file_path} not found. It may have been deleted.")
                except Exception as e:
                    print(f"An error occurred while deleting file {file_path}: {e}")

            for name in dirs:
                dir_path = os.path.join(root, name)
                try:
                    os.rmdir(dir_path)  # Remove the directory
                except FileNotFoundError:
                    print(f"Directory {dir_path} not found. It may have been deleted.")
                except Exception as e:
                    print(f"An error occurred while deleting directory {dir_path}: {e}")

        print(f"All contents of {destination_path} have been deleted successfully.")
    except FileNotFoundError:
        print(f"Folder {destination_path} not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def copy_sysucom_canwhitelist(source_path, destination_path):
    # Copy SysUcom and CANWhitelist
    try:
        shutil.copytree(source_path, destination_path, dirs_exist_ok=True)
        print(f"Successfully copied contents from {source_path} to {destination_path}.")
    except FileExistsError:
        print(f"The destination directory {destination_path} already exists. Files may have been overwritten.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    source_path = "D:/BaseBuildFlashing/BaseBuild/SysUcomCANWhitelistFiles/ACC100X_0029"
    destination_path = "D:/BENCH_USB/"

    # Delete the contents of the Pen Drive folder
    delete_pendrive_contents(destination_path)

    # Copy the SysUcom and CANWhitelist
    copy_sysucom_canwhitelist(source_path, destination_path)
