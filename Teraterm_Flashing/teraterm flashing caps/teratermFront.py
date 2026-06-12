import subprocess
import time
import ctypes
import win32gui
import win32process

def bring_window_to_foreground(pid):
    # Define the callback function to bring the window to the front
    def enum_window_callback(hwnd, pid):
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
        if found_pid == pid:
            win32gui.SetForegroundWindow(hwnd)
            return False  # Stop enumerating windows
        return True

    # Enumerate through all windows and find the one that matches the PID
    win32gui.EnumWindows(enum_window_callback, pid)

def launch_teraterm(com_port):
    try:
        # Launch Tera Term with the specified COM port and baud rate
        process = subprocess.Popen(
            [r"C:\Program Files (x86)\teraterm\ttermpro.exe", f'/C={com_port[3:]}', '/BAUD=115200'],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        print(f"Tera Term launched on {com_port} with baud rate 115200.")

        # Wait for the application to start
        time.sleep(2)

        # Get the process ID (PID)
        pid = process.pid
        
        # Bring Tera Term window to the foreground
        bring_window_to_foreground(pid)

        return True
    except FileNotFoundError:
        print("Tera Term executable not found. Ensure the path is correct.")
        return False

# Usage Example
com_port = "COM5"  # Example COM port, replace as needed
launch_teraterm(com_port)
