import os
import subprocess
import serial.tools.list_ports
import sys
import time

def find_silicon_labs_port():
    for port in serial.tools.list_ports.comports():
        if "Silicon Labs CP210x USB to UART Bridge" in port.description:
            return port.device  # Return COM port name (e.g., COM9)
    return None

def create_ttl_file(ttl_file_path, commands, variant_code):
    """
    Create a .ttl macro file with the specified commands and variant code.
    """
    with open(ttl_file_path, "w") as ttl_file:
        ttl_file.write(f"; TTL Macro File\n")
        ttl_file.write(f"setbaud 115200\n")  # Set baud rate
        ttl_file.write(f"connect '{find_silicon_labs_port()}'\n")  # Connect to the COM port
        for command in commands:
            ttl_file.write(f"sendln '{command}'\n")  # Send each command
        for _ in range(4):  # Send variant code 4 times
            ttl_file.write(f"sendln '{variant_code}'\n")
            ttl_file.write(f"wait 1\n")  # Wait 1 second between each send
        ttl_file.write(f"disconnect\n")  # Disconnect from the session
        ttl_file.write(f"end\n")  # End the macro

def launch_teraterm_with_ttl(ttl_file_path):
    """
    Launch Tera Term with the created .ttl macro file.
    """
    try:
        # Get the path to the Tera Term executable
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "."))
        tterm_path = os.path.join(base_dir, "teraterm", "ttermpro.exe")
        
        # Launch Tera Term with the macro file
        subprocess.Popen([tterm_path, f'/M={ttl_file_path}'])
        print(f"Tera Term launched with macro: {ttl_file_path}")
        return True
    except FileNotFoundError:
        print("Tera Term executable not found. Ensure the path is correct.")
        return False

def cleanup_ttl_file(ttl_file_path):
    """
    Delete the TTL macro file after execution.
    """
    if os.path.exists(ttl_file_path):
        os.remove(ttl_file_path)
        print(f"Deleted TTL macro file: {ttl_file_path}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: python script.py <variant_code>")

    variant_code = sys.argv[1]

    # Find the Silicon Labs COM port
    com_port = find_silicon_labs_port()
    if not com_port:
        sys.exit("No Silicon Labs COM port found. Exiting.")

    # Define TTL file path
    ttl_file_path = os.path.join(os.path.dirname(__file__), "temp_macro.ttl")

    # Define commands
    commands = ["MCU_DEBUG00", "MCU_DEBUG01"]

    # Create TTL macro file
    create_ttl_file(ttl_file_path, commands, variant_code)

    # Launch Tera Term with the macro file
    if not launch_teraterm_with_ttl(ttl_file_path):
        cleanup_ttl_file(ttl_file_path)  # Clean up even if Tera Term fails
        sys.exit("Failed to launch Tera Term. Exiting.")

    # Wait for the macro execution to complete
    time.sleep(10)  # Adjust based on expected macro execution time

    # Clean up TTL macro file
    cleanup_ttl_file(ttl_file_path)
