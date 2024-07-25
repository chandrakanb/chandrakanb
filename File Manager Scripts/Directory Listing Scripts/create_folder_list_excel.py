import os
from openpyxl import Workbook

# Function to get list of folders in the current directory
def get_folder_list(directory):
    folder_list = []
    for root, dirs, files in os.walk(directory):
        for dir in dirs:
            folder_list.append(dir)
        break  # Only get top-level directories
    return folder_list

# Function to create an Excel sheet with folder names
def create_excel(folder_list, excel_name='folder_list.xlsx'):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Folder List'

    # Adding column title
    sheet['A1'] = 'Sr. No.'
    sheet['B1'] = 'Folder Name'

    # Adding Sr. No. and folder names to the sheet
    for index, folder_name in enumerate(folder_list, start=2):
        sheet[f'A{index}'] = index-1
        sheet[f'B{index}'] = folder_name

    workbook.save(excel_name)
    print(f"Excel file '{excel_name}' created successfully!")

# Main function
if __name__ == '__main__':
    # Get the current directory
    current_directory = os.getcwd()
    
    # Get the list of folders in the current directory
    folder_list = get_folder_list(current_directory)
    
    # Create the Excel file
    create_excel(folder_list)
