import os
import pandas as pd

# Path to the Excel file
excel_path = "input.xlsx"  # Replace with your file path

# Read Excel file
df = pd.read_excel(excel_path)

# Iterate through each row in the dataframe
for index, row in df.iterrows():
    main_folder = row['Main_folder']
    subfolder = row.get('VehicleModel', '')  # Default to an empty string if VehicleModel is missing or NaN
    script_name = row['VariantCode']
    command = row['command']
    com_port = 'com_port'
    com_port_1 = 'com_port[3:]'

    # Check if subfolder is a valid string (not NaN or float)
    if isinstance(subfolder, str) and subfolder.strip():
        subfolder_path = os.path.join(main_folder, subfolder)
    else:
        subfolder_path = main_folder  # Use main folder if subfolder is not valid

    # Create the folder if it doesn't exist
    os.makedirs(subfolder_path, exist_ok=True)

    # Prepare Python script content
    script_content = f"""
import subprocess
import pyautogui
import time
import serial.tools.list_ports
import os
import sys

def find_silicon_labs_port():
    for port in serial.tools.list_ports.comports():
        if "Silicon Labs CP210x USB to UART Bridge" in port.description:
            return port.device  # Return COM port name (e.g., COM9)
    return None

def launch_teraterm(com_port):
    try:
        # Launch Tera Term with the specified COM port and baud rate
        subprocess.Popen([r"C:\Program Files (x86)\teraterm\ttermpro.exe", f'/C={com_port_1}', '/BAUD=115200'])
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
	# Find Silicon Labs COM port
    com_port = find_silicon_labs_port()
    if not com_port:
        sys.exit("No Silicon Labs COM port found. Exiting.")

    # Launch Tera Term
    if not launch_teraterm(com_port):
        sys.exit("Failed to launch Tera Term. Exiting.")

    # Send commands to Tera Term
    commands = ["MCU_DEBUG00", "MCU_DEBUG01"]
    for command in commands:
        send_command(command)
		
	for i in range(4):  # Send the variant code 4 times
		send_command({command})
        time.sleep(1)

    # Close Tera Term
    time.sleep(2)
    close_teraterm()
"""

    # Write the script to a file
    script_path = os.path.join(subfolder_path, script_name)
    try:
        with open(script_path, 'w') as script_file:
            script_file.write(script_content)
        print(f"Script created: {script_path}")
    except Exception as e:
        print(f"Error creating script {script_name}: {e}")

print("All folders, subfolders, and scripts created successfully!")
