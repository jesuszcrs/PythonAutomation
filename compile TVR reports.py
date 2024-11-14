import pandas as pd
import os

def merge_excel_sheets(folder_path, output_file):
    # Get a list of all Excel files in the specified folder
    excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

    # Check if there are any Excel files in the folder
    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    # Initialize an empty DataFrame to store the merged data
    merged_data = pd.DataFrame()

    # Iterate through each Excel file and merge its sheet into the DataFrame
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)
        merged_data = merged_data.append(df, ignore_index=True)

    # Write the merged data to a new Excel file
    merged_data.to_excel(output_file, index=False)
    print(f"Merged data written to {output_file}.")

# Specify the folder containing the Excel files and the output file name
folder_path = f'H:\MSC\Public\Finance\Reports\Rebates\{Year}\True Value Rewards'
output_file = f'merged_{Year}_TVRData.xlsx'

# Call the function to merge Excel sheets
merge_excel_sheets(folder_path, output_file)
