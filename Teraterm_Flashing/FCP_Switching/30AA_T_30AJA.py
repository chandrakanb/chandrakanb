import subprocess
import pyautogui
import time

def launch_teraterm():
    try:
        # Launch Tera Term with the specified COM port and baud rate
        subprocess.Popen(["./teraterm/ttermpro.exe", "/C=21", "/BAUD=115200"])
        print("Tera Term launched on COM port 21 with baud rate 115200")
    except FileNotFoundError:
        print("Tera Term executable not found. Ensure it's in the same directory or provide the correct path.")

def send_command(command):
    time.sleep(3)  # Wait to ensure Tera Term is ready
    pyautogui.typewrite(command)
    pyautogui.press("enter")

if __name__ == "__main__":
    # Launch Tera Term
    launch_teraterm()

    # Send commands to Tera Term
    send_command("MCU_DEBUG00")
    send_command("MCU_DEBUG01")
    time.sleep(2)

    variant_code = "MCU_DEBUG143330414A4101"
    for i in range(1, 5):  # Send the variant code 4 times
        send_command(variant_code)
        time.sleep(1)  # Small delay between commands

    # Close Tera Term
    time.sleep(2)
    subprocess.run(["taskkill", "/f", "/im", "ttermpro.exe"])
    print("Tera Term closed")
