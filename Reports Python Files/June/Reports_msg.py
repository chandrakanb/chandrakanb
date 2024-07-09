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
    # Parameters for creating folder and copying files
    cycle = "DryRun2.3"
    build = "Dev_Build"
    execution = "Overnight_Execution"
    flashing = "Auto_Flashing"
    date = "23rd"
    month = "June"
    bench = "Bench9_TYAW"
    result = "Result"

    # Construct input_string
    parameters = [cycle, build, execution, date, month, bench]
    input_string = '_'.join(parameters) + '_'

    # Construct message_string
    date_month = [date, month]
    cycle_build_execution_bench = [cycle, build, execution, bench]
    execution_result = [execution, result]
    flashing_result = [flashing, result]
    output_string_1 = ' '.join(date_month)+', '+' '.join(cycle_build_execution_bench).replace("_", " ")+':'
    output_string_2 = ' '.join(execution_result).replace("_", " ")+':'
    output_string_3 = ' '.join(flashing_result).replace("_", " ")+':'
    message_string = f"{output_string_1}\n{output_string_2}\n{output_string_3}"
    
    # Copy the output string to the clipboard
    pyperclip.copy(message_string)
    