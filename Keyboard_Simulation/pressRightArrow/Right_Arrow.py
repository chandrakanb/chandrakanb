import pyautogui
import time
import keyboard

while True:
    if keyboard.is_pressed('q'):  # Exit loop when 'q' is pressed
        break
    pyautogui.press('right')  # Simulate Left Arrow key press
    time.sleep(1.5)          # Wait for 1.5 seconds
