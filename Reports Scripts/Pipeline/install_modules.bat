@echo off
REM Batch file to install necessary Python modules

echo Installing required Python modules...

REM Ensure pip is installed and upgraded
python -m ensurepip --upgrade
python -m pip install --upgrade pip

REM Install required packages
pip install beautifulsoup4
pip install pandas
pip install openpyxl
pip install pyautogui
pip install pyperclip
pip install requests
pip install html2image
pip install selenium
pip install Pillow
pip install webdriver-manager

echo All modules installed successfully!

pause
