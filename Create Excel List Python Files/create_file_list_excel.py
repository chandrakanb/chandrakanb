import os
from openpyxl import Workbook

# Function to get list of files in the current directory
def get_file_list(directory):
    file_list = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_list.append(file)
    return file_list

# Function to create an Excel sheet with file names
def create_excel(file_list, excel_name='file_list.xlsx'):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'File List'

    # Adding column title
    sheet['A1'] = 'Sr. No.'
    sheet['B1'] = 'File Name'

    # Adding Sr. No. and file names to the sheet
    for index, file_name in enumerate(file_list, start=2):
        sheet[f'A{index}'] = index-1
        sheet[f'B{index}'] = file_name

    workbook.save(excel_name)
    print(f"Excel file '{excel_name}' created successfully!")

# Main function
if __name__ == '__main__':
    # Get the current directory
    current_directory = os.getcwd()
    
    # Get the list of files in the current directory
    file_list = get_file_list(current_directory)
    
    # Create the Excel file
    create_excel(file_list)
