# restore_backup.py

import os
import shutil

def list_backups(backup_folder):
    """
    List all existing backups in the backup folder.
    """
    backups = []
    for item in os.listdir(backup_folder):
        item_path = os.path.join(backup_folder, item)
        if os.path.isdir(item_path):
            backups.append(item)
    return backups

def restore_backup(source_directory, backup_folder, backup_name):
    """
    Restore the specified backup folder to the source directory.
    """
    backup_path = os.path.join(backup_folder, backup_name)
    if os.path.exists(backup_path):
        for item in os.listdir(backup_path):
            s = os.path.join(backup_path, item)
            d = os.path.join(source_directory, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, dirs_exist_ok=True)
            else:
                shutil.copy2(s, d)
        print(f"Backup '{backup_name}' restored to '{source_directory}'.")
    else:
        print(f"Backup '{backup_name}' does not exist.")

def main():
    source_directory = os.path.join(os.getcwd(), "KITE")
    backup_folder = "KITE_BACKUP"

    # Check if the backup folder exists
    if not os.path.exists(backup_folder):
        print(f"The backup folder '{backup_folder}' does not exist.")
        return

    # List all backups
    backups = list_backups(backup_folder)
    if not backups:
        print("No backups found.")
        return

    print("Available backups:")
    for idx, backup in enumerate(backups, start=1):
        print(f"{idx}. {backup}")

    # Ask the user which backup to restore
    try:
        choice = int(input("Enter the number of the backup to restore (0 to cancel): "))
        if choice == 0:
            print("No backup was restored.")
            return
        elif 1 <= choice <= len(backups):
            restore_backup(source_directory, backup_folder, backups[choice - 1])
        else:
            print("Invalid choice.")
    except ValueError:
        print("Invalid input. Please enter a number.")

if __name__ == "__main__":
    main()
