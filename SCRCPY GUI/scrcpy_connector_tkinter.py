import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess
import threading
import os

def get_adb_path():
    return os.environ.get('ADB_PATH', 'adb')  # Default to 'adb' if not set

def get_scrcpy_path():
    return os.environ.get('SCRCPY_PATH', 'scrcpy')  # Default to 'scrcpy' if not set

def get_adb_devices():
    adb_path = get_adb_path()
    try:
        output = subprocess.check_output([adb_path, 'devices'], universal_newlines=True)
        lines = output.splitlines()
        devices = []
        for line in lines:
            if 'device' in line and not 'List of devices attached' in line:
                device_id = line.split('\t')[0]
                devices.append(device_id)
        return devices
    except FileNotFoundError:
        return ["ADB not found. Please check the ADB_PATH environment variable."]
    except subprocess.CalledProcessError:
        return ["Error executing ADB command. Check your ADB installation."]

def refresh_devices():
    devices = get_adb_devices()
    device_listbox.delete(0, tk.END)  # Clear the current list
    for device in devices:
        device_listbox.insert(tk.END, device)

def connect_to_device():
    scrcpy_path = get_scrcpy_path()
    try:
        selected_indices = device_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Info", "Please select a device from the list.")
            return

        selected_index = selected_indices[0]
        selected_device = device_listbox.get(selected_index)

        # Run scrcpy in a separate thread
        threading.Thread(target=run_scrcpy, args=(selected_device, scrcpy_path)).start()
    except FileNotFoundError:
        messagebox.showerror("Error", "Scrcpy not found. Please check the SCRCPY_PATH environment variable.")
    except Exception as e:
         messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def run_scrcpy(device_id, scrcpy_path):
    subprocess.Popen([scrcpy_path, '-s', device_id])

# Create the main window
root = tk.Tk()
root.title("Scrcpy - Chandrakant")

# Apply a theme
style = ttk.Style()
style.theme_use('clam')  # Experiment with different themes: 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'

# UI elements
refresh_button = ttk.Button(root, text="Refresh", command=refresh_devices)
refresh_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew") # Use grid instead of pack

scrollbar = tk.Scrollbar(root)
scrollbar.grid(row=1, column=1, sticky="ns")

device_listbox = tk.Listbox(root, selectmode=tk.SINGLE, width=40, height=10, yscrollcommand=scrollbar.set)
device_listbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

connect_button = ttk.Button(root, text="Connect", command=connect_to_device)
connect_button.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

# Configure grid column and row weights to allow resizing
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)

# Initial device refresh
refresh_devices()

# Start the main event loop
root.mainloop()