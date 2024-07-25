import os
import time
import webbrowser
import pyautogui

def take_screenshot(html_file_path, screenshot_path, crop_coordinates):
    # Get the currently active window title before opening the HTML file
    active_window_title_before = pyautogui.getActiveWindow().title
    
    # Construct the file URL using 'file://' protocol
    file_url = 'file:///' + os.path.abspath(html_file_path).replace('\\', '/')

    # Open the file URL in a web browser (opens in a new tab or window)
    webbrowser.open_new_tab(file_url)
    time.sleep(5)  # Adjust as necessary to ensure the pyautoguie loads completely

    # Take a screenshot of the entire screen
    screenshot = pyautogui.screenshot()

    # Crop the screenshot based on the provided coordinates
    cropped_screenshot = screenshot.crop(crop_coordinates)

    # Save the cropped screenshot to the specified path
    cropped_screenshot.save(screenshot_path)

    # Close the browser tab
    pyautogui.hotkey('ctrl', 'w')  # Close the tab (Ctrl + W)
    
    # Switch back to the previously active window
    pyautogui.getWindowsWithTitle(active_window_title_before)[0].activate()

if __name__ == "__main__":

    html_file_path = 'MainDetailedReport.html'
    screenshot_path = 'Execution_Summary.png'
    crop_coordinates = (90, 160, 1350, 450)  # crop coordinates: (left, upper, right, lower)

    take_screenshot(html_file_path, screenshot_path, crop_coordinates)
