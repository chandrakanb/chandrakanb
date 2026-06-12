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
    if len(sys.argv) == 2: 
        if len(sys.argv[1]) == 17 or len(sys.argv[1]) == 23:
            argument1 = sys.argv[1]
            argument_commands = [command.upper() for command in [argument1]]
        else:
            sys.exit("Usage: python script.py <hard_variant_code(17 chars)> or/and <fcp_switch_command(23 chars)>.")
    elif len(sys.argv) == 3:
        if len(sys.argv[1]) == 17 and len(sys.argv[2]) == 23:
            argument1 = sys.argv[1]
            argument2 = sys.argv[2]
            argument_commands = [command.upper() for command in [argument1, argument2]]
        elif len(sys.argv[1]) == 23 and len(sys.argv[2]) == 17:
            argument1 = sys.argv[2]
            argument2 = sys.argv[1]
            argument_commands = [command.upper() for command in [argument1, argument2]]
        else:
            sys.exit("Usage: python script.py <hard_variant_code(17 chars)> or/and <fcp_switch_command(23 chars)>.")
    else:
        sys.exit("Usage: python script.py <hard_variant_code(17 chars)> or/and <fcp_switch_command(23 chars)>.")

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
    
    # Send argument commands
    for command in argument_commands:
        send_commands([command] * 4)

    # Restore Caps Lock state if necessary
    manage_caps_lock(original_caps_state)

    # Close Tera Term
    time.sleep(2)
    close_teraterm()
