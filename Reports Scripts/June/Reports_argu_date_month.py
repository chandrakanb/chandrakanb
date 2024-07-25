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

def get_date(input_date):
    # Convert number to string
    input_date_str = str(input_date)
    
    # Determine suffix based on last digit
    if input_date_str.endswith('1') and not input_date_str.endswith('11'):
        suffix = 'st'
    elif input_date_str.endswith('2') and not input_date_str.endswith('12'):
        suffix = 'nd'
    elif input_date_str.endswith('3') and not input_date_str.endswith('13'):
        suffix = 'rd'
    else:
        suffix = 'th'
    
    # Return ordinal string
    return input_date_str + suffix

if __name__ == "__main__":
    print('\n' + '-' * 45)
    
    # Parameters for creating folder and copying files
    cycle = "DryRun2.3"
    build = "Build"
    bench = "Bench9_TYAW"
    result = "Result"

    # Take Build Type from User
    input_build = int(input("|\tBuild Type:\n|\t\t1.Dev\n|\t\t2.Rel\n|\tSelect Build Type:"))
    if input_build == 1:
        dev_or_rel = "Dev"
    elif input_build == 2:
        dev_or_rel = "Rel"
    print('-' * 45 + '\n')
'''    
    # Take Date from User and add correct suffix
    input_date = int(input("Enter Date: "))
    date = get_date(input_date)
    print('-' * 45 + '\n')

    # Take Month from User
    input_month = input("Enter month: ")
    month = input_month.capitalize()
    print('-' * 45 + '\n')
    
    # Take Execution Type from User
    overnight = "Overnight_Execution"
    flashing = "Auto_Flashing"
    flag = input("Execution:\n\t1. Overnight Execution\n\t2. Auto Flashing\nEnter your choice (1 or 2): ")
    if flag == "1":
        execution = overnight
    elif flag == "2":
        execution = flashing
    print('-' * 45 + '\n')
    
    # Construct message_string
    if execution == overnight:
        date_month = [date, month]
        cycle_dev_or_rel_build_execution_bench = [cycle, dev_or_rel, build, execution, bench]
        execution_result = [execution, result]
        flashing_result = [flashing, result]
        output_string_1 = ' '.join(date_month)+', '+' '.join(cycle_dev_or_rel_build_execution_bench).replace("_", " ")+':'
        output_string_2 = ' '.join(execution_result).replace("_", " ")+':'
        output_string_3 = ' '.join(flashing_result).replace("_", " ")+':'
        message_string = f"{output_string_1}\n{output_string_2}\n{output_string_3}"
    
    
        # Copy the output string to the clipboard
        pyperclip.copy(message_string)
        output_message = pyperclip.paste()
        if output_message == message_string:
            print("Message copied successfully.")
            print('-' * 45 + '\n')
'''