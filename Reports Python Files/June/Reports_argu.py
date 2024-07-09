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
    overnight = "Overnight_Execution"
    flashing = "Auto_Flashing"
    
    # Prompt user for input
    flag = input("Execution:\n1. Overnight\n2. Auto\nEnter your choice (1 or 2): ")
    
    # Check user input and assign execution accordingly
    if flag == "1":
        execution = overnight
    elif flag == "2":
        execution = flashing
    else:
        print("Invalid input. Please enter 1 or 2.")
        exit()  # Exit the program if input is invalid
    
    # Print the chosen execution
    print("Selected execution:", execution)
