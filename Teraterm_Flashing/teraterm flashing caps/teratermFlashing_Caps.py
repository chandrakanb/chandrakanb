import subprocess
import pyautogui
import time
import serial.tools.list_ports
import os
import sys
import ctypes

def find_silicon_labs_port():
    for port in serial.tools.list_ports.comports():
        if "Silicon Labs CP210x USB to UART Bridge" in port.description:
            return port.device  # Return COM port name (e.g., COM9)
    return None

def launch_teraterm(com_port):
    try:
        # Launch Tera Term with the specified COM port and baud rate
        subprocess.Popen(
            [r"C:\Program Files (x86)\teraterm\ttermpro.exe", f'/C={com_port[3:]}', '/BAUD=115200'],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        print(f"Tera Term launched on {com_port} with baud rate 115200.")
        return True
    except FileNotFoundError:
        print("Tera Term executable not found. Ensure the path is correct.")
        return False

def close_teraterm():
    subprocess.run(["taskkill", "/f", "/im", "ttermpro.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print("Tera Term closed.")

def manage_caps_lock(state):
    current_state = ctypes.windll.user32.GetKeyState(0x14) & 0xFFFF != 0
    if state != current_state:
        pyautogui.press('capslock')

def send_commands(commands):
    for command in commands:
        time.sleep(1)  # Wait to ensure Tera Term is ready
        pyautogui.typewrite(command)
        pyautogui.press("enter")

if __name__ == "__main__":
    if len(sys.argv) != 3:  # Check for exactly two additional arguments
        sys.exit("Usage: python script.py <hard_variant_code> <fcp_switch_command>")

    # Extract arguments
    hard_variant_code = sys.argv[1].upper()
    fcp_switch_command = sys.argv[2].upper()

    # Find Silicon Labs COM port
    com_port = find_silicon_labs_port()
    if not com_port:
        sys.exit("No Silicon Labs COM port found. Exiting.")

    # Launch Tera Term
    if not launch_teraterm(com_port):
        sys.exit("Failed to launch Tera Term. Exiting.")

    # Manage Caps Lock state
    original_caps_state = ctypes.windll.user32.GetKeyState(0x14) & 0xFFFF != 0
    manage_caps_lock(False)

    # Send commands
    commands = [command.upper() for command in ["MCU_DEBUG00", "MCU_DEBUG01"]]
    send_commands(commands)
    send_commands(["switch command".lower()])
    
    # Send custom commands
    send_commands([hard_variant_code] * 4)
    send_commands([fcp_switch_command] * 4)

    # Restore Caps Lock state if necessary
    manage_caps_lock(original_caps_state)

    # Close Tera Term
    time.sleep(2)
    close_teraterm()
