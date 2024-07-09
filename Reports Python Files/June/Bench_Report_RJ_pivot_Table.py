import os
import pandas as pd

def create_pivot_table(df):
    # Filter and add 'Script Type' column
    filtered_dfs = []

    filtered_df_post = df[df['Test ID'].str.contains('Post', na=False)]
    filtered_df_post.insert(0, 'Script Type', 'Postcondition')
    filtered_dfs.append(filtered_df_post)

    filtered_df_kpit = df[df['Test ID'].str.contains('KPIT_TC', na=False)]
    filtered_df_kpit.insert(0, 'Script Type', 'Test Script')
    filtered_dfs.append(filtered_df_kpit)

    filtered_df_precon = df[df['Test ID'].str.contains('Precon', na=False)]
    filtered_df_precon.insert(0, 'Script Type', 'Precondition')
    filtered_dfs.append(filtered_df_precon)

    # Concatenate the filtered DataFrames
    filtered_df = pd.concat(filtered_dfs)

    # Create the pivot table
    pivot_table_df = pd.pivot_table(filtered_df, index='Script Type', columns='Status', aggfunc='size', fill_value=0)
    
    return pivot_table_df

def save_to_same_excel(df, pivot_table_df, excel_file_path):
    # Load the existing workbook
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        pivot_table_df.to_excel(writer, sheet_name='Pivot Table')
    print(f"Pivot table saved to new sheet in the same workbook: {excel_file_path}")

if __name__ == "__main__":
    # Define the file paths
    excel_file_path = "Execution_Report.xlsx"
    new_excel_file_path = "pivot_table.xlsx"
    
    # Check if the file exists
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"Excel file '{excel_file_path}' not found.")
    
    # Read the Excel file
    df = pd.read_excel(excel_file_path)
    
    # Create the pivot table DataFrame
    pivot_table_df = create_pivot_table(df)
    
    # Save to the same Excel file in a new sheet
    save_to_same_excel(df, pivot_table_df, excel_file_path)
    
    print(pivot_table_df)
