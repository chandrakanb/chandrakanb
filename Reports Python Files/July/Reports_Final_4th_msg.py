import os
import re
import bs4
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import xml.etree.ElementTree as ET
import openpyxl
import shutil
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import webbrowser
import pyautogui
import pyperclip

if __name__ == "__main__":
    print('\n|' + '‾' * 45)
    
    # Get the currently active window title before opening the HTML file
    active_window_title_before = pyautogui.getActiveWindow().title
    
    # Parameters for creating folder and copying files
    result = "Result"
    
    # Take DryRun2.x from User
    cycle = 'DryRun2.1'
    print('|' + '_' * 45 + '\n|\t')
    
    # Take Build Type from User
    build = "Dev_Build"
    print('|' + '_' * 45 + '\n|\t')
    
    # Take Date from User and add correct suffix
    date = "4th"
    print('|' + '_' * 45 + '\n|\t')

    # Take Month from User
    month = "July"
    print('|' + '_' * 45 + '\n|\t')
    
    # Take Execution Type from User
    overnight = "Overnight_Execution"
    flashing = "Auto_Flashing"
    flag = input("|\tExecution:\n|\t\t1. Overnight Execution\n|\t\t2. Auto Flashing\n|\tEnter your choice (1 or 2): ")
    execution = overnight if flag == "1" else flashing if flag == "2" else None
    print('|' + '_' * 45 + '\n|\t')
    
    # Take Bench Number from User
    bench = "Bench09_TYAW"
    print('|' + '_' * 45 + '\n|\t')
    
    # Construct input_string
    parameters = [execution, date, month, bench]
    input_string = '_'.join(parameters) + '_'
    
    # Folder name
    folder_name = input_string.strip("_").replace(" ", "_")
    
    # Print folder Name
    print(input_string)
    
    # Construct message_string
    if execution == overnight:
        date_month = f"{date} {month}, "
        execution_result = f"{execution.replace('_', ' ')} {result}:"
        flashing_result = f"{flashing.replace('_', ' ')} {result}:"
    
        message_string = f"{date_month}{execution.replace('_', ' ')} {bench.replace('_', ' ')}:\nDryRun2.3 {execution_result}\nDryRun2.4 {execution_result}"
        
        # Copy the output string to the clipboard
        pyperclip.copy(message_string)
        output_message = pyperclip.paste()
        if output_message == message_string:
            print(f"|\t{output_message}")
            print("|\tMessage copied successfully.")
            print('|' + '_' * 45 + '\n|\t')

    # Switch back to the previously active window
    pyautogui.getWindowsWithTitle(active_window_title_before)[0].activate()
    
print("|\tAll tasks completed successfully.")
print('|' + '_' * 45 + '\n')
input("Press Enter to exit...")