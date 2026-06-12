import subprocess
import pyautogui
import time
import serial.tools.list_ports
import os
import sys
import win32com.client

def find_silicon_labs_port():
    for port in serial.tools.list_ports.comports():
        if "Silicon Labs CP210x USB to UART Bridge" in port.description:
            return port.device  # Return COM port name (e.g., COM9)
    return None

def launch_teraterm(com_port):
    try:
        # Launch Tera Term with the specified COM port and baud rate
        subprocess.Popen([r"C:\Program Files (x86)\teraterm\ttermpro.exe", f'/C={com_port[3:]}', '/BAUD=115200'])
        print(f"Tera Term launched on {com_port} with baud rate 115200")
        return True
    except FileNotFoundError:
        print("Tera Term executable not found. Ensure the path is correct.")
        return False

def send_command(command):
    time.sleep(3)  # Wait to ensure Tera Term is ready
    pyautogui.typewrite(command)
    pyautogui.press("enter")

def close_teraterm():
    subprocess.run(["taskkill", "/f", "/im", "ttermpro.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print("Tera Term closed.")

if __name__ == "__main__":
    if len(sys.argv) != 3:  # Check for exactly one additional argument
        sys.exit("Usage: python script.py <hard_variant_code> <fcp_switch_command>")
        
    # Find Silicon Labs COM port
    com_port = find_silicon_labs_port()
    if not com_port:
        sys.exit("No Silicon Labs COM port found. Exiting.")

    # Launch Tera Term
    if not launch_teraterm(com_port):
        sys.exit("Failed to launch Tera Term. Exiting.")

    # Send commands to Tera Term
    commands = [command.upper() for command in ["MCU_DEBUG00", "MCU_DEBUG01"]]
    for command in commands:
        send_command(command)

    commands = [command.lower() for command in ["switch command"]]
    for command in commands:
        send_command(command)
    
    hw_variant = sys.argv[1].upper()
    for i in range(4):  # Send the hard variant code 4 times
        send_command(hw_variant)
        time.sleep(1)
    
    fcp_switch_command = sys.argv[2].upper()
    for i in range(4):  # Send the FCP switch command 4 times
        send_command(fcp_switch_command)
        time.sleep(1)
    
    # Close Tera Term
    time.sleep(2)
    close_teraterm()
