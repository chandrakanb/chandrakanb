"""
import pyautogui

command = "hello world"
pyautogui.typewrite(command.upper())
"""

"""
import pyautogui
import ctypes

def is_capslock_on():
    # Check the state of the Caps Lock key
    return ctypes.windll.user32.GetKeyState(0x14) & 0xFFFF != 0

def set_capslock(state):
    # Toggle Caps Lock state if necessary
    if is_capslock_on() != state:
        pyautogui.press('capslock')

def type_in_caps(command):
    ogCapsState = is_capslock_on()
    set_capslock(False)  # Turn off Caps Lock
    pyautogui.typewrite(command.upper())
    set_capslock(ogCapsState)  # Restore original Caps Lock state

command = "hello world"
type_in_caps(command)
"""

"""
import pyautogui
import ctypes

def toggle_capslock():
    # ctypes.windll.user32.keybd_event(0x14, 0, 0, 0)  # Press Caps Lock
    ctypes.windll.user32.keybd_event(0x14, 0, 2, 0)  # Release Caps Lock

command = "hello world"
toggle_capslock()  # Toggle Caps Lock
pyautogui.typewrite(command.upper())  # Type in uppercase
toggle_capslock()  # Restore original Caps Lock state
"""


import pyautogui
import ctypes

def type_in_caps(command):
    ogCapsState = ctypes.windll.user32.GetKeyState(0x14) & 0xFFFF != 0
    if ogCapsState:
        pyautogui.press('capslock')
    pyautogui.typewrite(command.upper())
    if ogCapsState:
        pyautogui.press('capslock')
command = "hello world"
type_in_caps(command)
