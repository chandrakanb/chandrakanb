import pandas as pd

# Define the path for the new Excel file
file_path = 'Daily_Activity_Tracker_August.xlsx'

# Define the data for the new sheet
data = {
    'Column1': [1, 2, 3],
    'Column2': ['A', 'B', 'C']
}

# Create a DataFrame from the data
df = pd.DataFrame(data)

# Write the DataFrame to a new Excel file with a new sheet named "new_sheet"
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='new_sheet', index=False)

print(f"Excel file created with a new sheet named 'new_sheet' at {file_path}")
