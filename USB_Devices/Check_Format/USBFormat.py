import tkinter as tk
from tkinter import ttk
import psutil
import subprocess

def get_connected_drives():
    """Get a list of connected drive letters."""
    drives = []
    for disk in psutil.disk_partitions():
        drive_letter = disk.mountpoint.replace('\\', '')
        if drive_letter not in ['C:', 'D:']:
            drives.append(drive_letter)
    return drives

def check_drive_format(drive_letter):
    """Check the format of the selected drive."""
    try:
        output = subprocess.check_output(f'fsutil fsinfo volumeinfo {drive_letter} | find "File System Name"', shell=True).decode('utf-8')
        return output.strip()
    except Exception as e:
        return str(e)

def refresh_drives():
    """Refresh the drives list every 10 seconds."""
    drives = get_connected_drives()
    drive_menu['values'] = drives
    if drive_var.get() not in drives:
        drive_var.set(drives[0] if drives else '')
    root.after(10000, refresh_drives)  # Call this function again after 10 seconds

def update_format_label(event=None):
    """Update the format label when the drive letter is selected."""
    drive_letter = drive_var.get()
    if drive_letter:
        format_label.config(text=f"Format: {check_drive_format(drive_letter)}")

root = tk.Tk()
root.title("USB Drive Format Checker")

# Get connected drive letters
drives = get_connected_drives()

# Create a dropdown menu with connected drive letters
drive_var = tk.StringVar()
drive_menu = ttk.Combobox(root, textvariable=drive_var)
drive_menu['values'] = drives
drive_menu.current(0)
drive_menu.pack(padx=10, pady=10)
drive_menu.bind("<<ComboboxSelected>>", update_format_label)

# Create a label to display the drive format
format_label = tk.Label(root, text="")
format_label.pack(padx=10, pady=10)

# Start refreshing drives list
refresh_drives()

root.mainloop()