import subprocess
import pyautogui
import time
import serial

def launch_teraterm():
    try:
        # Replace 'ttermpro.exe' with the path to your Tera Term executable if it's different
        subprocess.Popen(['D:/Softwares/teraterm-4.105/ttermpro.exe', '/C=24', '/BAUD=115200']) #COM Port = "COM24", Baude_Rate = "115200" 
        print(f"Tera Term launched with COM port 24 and baud rate 115200")
    except FileNotFoundError:
        print("Tera Term executable not found. Make sure Tera Term is installed and its directory is in the system PATH.")

def send_command(command):
    time.sleep(2)  # Wait for Tera Term window to open
    pyautogui.typewrite(command)
    pyautogui.press('enter')


if __name__ == "__main__":
    launch_teraterm()
    send_command("MCU_DEBUG00")
    send_command("MCU_DEBUG01")
    send_command("MCU_DEBUG160000003C")
    time.sleep(5)
    pyautogui.hotkey('alt','F4')
    
    
