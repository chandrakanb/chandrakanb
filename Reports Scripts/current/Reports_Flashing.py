import os
import shutil
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from bs4 import BeautifulSoup
from PIL import Image
import pyautogui

def get_execution_date_time():
    """
    This function reads an HTML report file to extract the start time of the execution.
    It then parses the start time to obtain the date, month, year, hours, and minutes.
    The date is adjusted if the hour is between 12 midnight and 12 noon of the next day.
    Finally, it converts the abbreviated month name to the full month name and returns the date and month.
    """
    # Read the content of the HTML file
    with open('MainDetailedReport.html', 'r', encoding='utf-8') as file:
        html_content = file.read()

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')

    # Find the start time
    start_time_tag = soup.find('th', string='Start Time')
    start_time_value = start_time_tag.find_next_sibling('td').text if start_time_tag else None

    # Split the start time into date, month, year, hours, and minutes
    date_parts = start_time_value.split(' ')
    date_digit = int(date_parts[0])  # "date" as an integer
    month_abbr = date_parts[1]  # "month"
    year = date_parts[2]  # "year"
    time_parts = date_parts[3].split(':')
    time_hh = time_parts[0]  # "hours"
    time_mm = time_parts[1]  # "minutes"

    # Adjust the date based on the hour
    # If hours is between 12 midnight and next day's 12 noon, then date minus one
    if int(time_hh) <= 12:
        date_digit -= 1
    
    # Add correct suffix to the date based on last digit
    date = get_date(date_digit)
    
    # Convert abbreviated month to full month name
    month = get_month(month_abbr)

    return date, month

def create_folder(folder_name):
    """Create a directory if it does not exist."""
    os.makedirs(folder_name, exist_ok=True)

def copy_pdf_files(folder_name, reports_prefix):
    """Copy PDF files to the specified folder."""
    pdf_reports_dir = os.path.join("XMLReports", "PdfReports")
    shutil.copy(os.path.join(pdf_reports_dir, "DashboardReport.pdf"), os.path.join(folder_name, f"{reports_prefix}_DashboardReport.pdf"))
    shutil.copy(os.path.join(pdf_reports_dir, "DetailedReport.pdf"), os.path.join(folder_name, f"{reports_prefix}_DetailedReport.pdf"))

def take_screenshot(html_file_path, screenshot_path, crop_coordinates):
    """Take a screenshot of the specified HTML file and save it."""
    options = Options()
    options.use_chromium = True
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    
    file_url = 'file:///' + os.path.abspath(html_file_path).replace('\\', '/')
    
    driver = webdriver.Edge(options=options)
    try:
        driver.get(file_url)
        screenshot_data = driver.get_screenshot_as_png()
        temp_screenshot_path = 'temp_screenshot.png'
        with open(temp_screenshot_path, 'wb') as f:
            f.write(screenshot_data)
        
        screenshot = Image.open(temp_screenshot_path)
        cropped_screenshot = screenshot.crop(crop_coordinates)
        cropped_screenshot.save(screenshot_path)
    finally:
        driver.quit()
        os.remove(temp_screenshot_path)

def get_date(date_digit):
    """Return the date with the appropriate ordinal suffix."""
    date_str = str(date_digit)
    suffix = 'st' if date_str.endswith('1') and not date_str.endswith('11') else \
             'nd' if date_str.endswith('2') and not date_str.endswith('12') else \
             'rd' if date_str.endswith('3') and not date_str.endswith('13') else \
             'th'
    return date_str + suffix

def get_month(month_abbr):
    """Return the full month name based on abbreviated month names."""
    months = {
        "Jan": "January", "Feb": "February", "Mar": "March",
        "Apr": "April", "May": "May", "Jun": "June",
        "Jul": "July", "Aug": "August", "Sep": "September",
        "Oct": "October", "Nov": "November", "Dec": "December"
    }
    return months.get(month_abbr, month_abbr)

def get_bench(input_bench):
    """Return the bench name with the appropriate variant name."""
    bench = (
        f"Bench0{input_bench}_T_Variant" if input_bench == 7 else
        f"Bench0{input_bench}_R_Variant" if input_bench == 6 else
        f"Bench0{input_bench}_TYAW" if input_bench == 9 else
        f"Bench{input_bench}_3BMA" if input_bench == 10 else
        f"Bench0{input_bench}_Q_Variant"
    )
    return bench

if __name__ == "__main__":
    # Get the execution date and month from the HTML report
    date, month = get_execution_date_time()
    
    # Set the execution type
    execution = "Auto_Flashing"
    
    # Prompt the user to enter the bench number
    input_bench = int(input("Enter Bench Number: "))
    bench = get_bench(input_bench)
    
    # Print the date, month, execution type, and bench information
    print(f"{date} {month}")
    print(execution.replace("_", " "))
    print(bench.replace("_", " "))
    
    # Create a string with the parameters separated by underscores
    parameters = [execution, date, month, bench]
    input_string = '_'.join(parameters) + '_'
    
    # Define the folder name and report prefix
    folder_name = f"{execution}_Reports"
    reports_prefix = input_string.strip("_").replace(" ", "_")
    
    # Create the folder and copy the PDF files into it
    create_folder(folder_name)
    copy_pdf_files(folder_name, reports_prefix)
    
    # Define the HTML file path, screenshot path and crop coordinates for screenshot
    html_file_path = 'MainDetailedReport.html'
    screenshot_path = os.path.join(folder_name, f"{execution}_Summary.png")
    crop_coordinates = (90, 63, 1365, 341)
    
    # Take a screenshot of the HTML report
    take_screenshot(html_file_path, screenshot_path, crop_coordinates)
    
    # Prompt the user indicating that all tasks are completed successfully
    input("All tasks completed successfully...")
