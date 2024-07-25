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
    # Parameters for creating folder and copying files
    cycle = "DryRun2.3"
    dev_or_rel = "Dev"
    build = "Build"
    month = "June"
    bench = "Bench9_TYAW"
    result = "Result"
    
    input_date = int(input("Enter Date: "))
    date = get_date(input_date)
    print(date)
    
    overnight = "Overnight_Execution"
    flashing = "Auto_Flashing"
    flag = input("Execution:\n1. Overnight Execution\n2. Auto Flashing\nEnter your choice (1 or 2): ")
    if flag == "1":
        execution = overnight
    elif flag == "2":
        execution = flashing
        
    print(execution)