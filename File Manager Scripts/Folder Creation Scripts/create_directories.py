import os

directories = [
    r"C:\Users\chandrakanb\Downloads\Bench09_TYAW\Bench09_Auto_Flashing_Reports_10th_July",
    r"C:\Users\chandrakanb\Downloads\Bench09_TYAW\Bench09_Auto_Flashing_Reports_16th_July",
    r"C:\Users\chandrakanb\Downloads\Bench09_TYAW\Bench09_Auto_Flashing_Reports_17th_July",
    r"C:\Users\chandrakanb\Downloads\Bench09_TYAW\Bench09_Auto_Flashing_Reports_18th_July",
    r"C:\Users\chandrakanb\Downloads\Bench09_TYAW\Bench09_Auto_Flashing_Reports_22nd_July",
    r"C:\Users\chandrakanb\Downloads\Bench09_TYAW\Bench09_Auto_Flashing_Reports_23rd_July"
]

for directory in directories:
    try:
        os.makedirs(directory, exist_ok=True)
        print(f"Directory {directory} created successfully.")
    except Exception as e:
        print(f"Error creating directory {directory}: {e}")
