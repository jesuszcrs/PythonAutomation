import pandas as pd
import os

Period = input("What Period are you combining data for: (1-12): ")
Eff_Year = int(input("Effective Year: "))

# Ask if previous year's data is needed
previous_year = input("Do you need data from the previous year? (yes/no): ").lower()

if previous_year == "yes":
    Year = Eff_Year - 1
else:
    Year = None  # Set Year to None if previous year's data is not needed

# List of Excel files to merge
rebate_files = [
    f'Conversion Funds\\Conversion Funds - DW\\Conversion Funds P{Period} {Eff_Year}.xlsx',
    f'Defective Program\\Receivable\\Defective Program Receivable P{Period} {Eff_Year}.xlsx',
    f'Retail Marketing Fund\\Retail Marketing Fund Receivable P{Period} {Eff_Year}.xlsx',
    f'Vendor Funding\\Vendor Funding Receivable P{Period} {Eff_Year}.xlsx',
    f'Rebate Earnings\\Period {Period} {Eff_Year} Rebate - Earned Report.xlsx'
]


if Year is not None:
    rebate_files.extend([
        f'Conversion Funds\\Conversion Funds - DW\\{Year} Conversion Funds P{Period} {Eff_Year}.xlsx',
        f'Defective Program\\Receivable\\{Year} Defective Program Receivable P{Period} {Eff_Year}.xlsx',
        f'Retail Marketing Fund\\{Year} Retail Marketing Fund Receivable P{Period} {Eff_Year}.xlsx',
        f'Vendor Funding\\{Year} Vendor Funding Receivable P{Period} {Eff_Year}.xlsx',
        f'Rebate Earnings\\{Year} Rebate - Earned Report thru P{Period} {Eff_Year}.xlsx'
    ])


incentive_files = rebate_files  # Use the same file paths as rebate files

# Directory where the Excel files are located
Eff_Year_Directory = f'H:\\MSC\\Public\\Finance\\Reports\\Rebates\\{Eff_Year}'
PrevYear_Directory = f'H:\\MSC\\Public\\Finance\\Reports\\Rebates\\{Year}'

# Output Excel file where the merged data will be saved
output_combined_file = 'VS_Receivable.xlsx'

# Create lists to store DataFrames
rebate_data_frames = []
incentive_data_frames = []

# Function to construct file paths based on year
def get_file_path(file, year):
    if year is not None:
        return os.path.join(Eff_Year_Directory if os.path.exists(os.path.join(Eff_Year_Directory, file)) else PrevYear_Directory, file)
    return os.path.join(Eff_Year_Directory, file)

# Loop through each Excel file for rebates
for file in rebate_files:
    file_path = get_file_path(file, Year)
    if not os.path.exists(file_path):
        continue
#    print("File Path:", file_path)
    xls = pd.ExcelFile(file_path)

    # List of sheet names to merge from each file
    sheets_to_merge = [
        f'CONV P{Period} {Eff_Year}',
        f'Defective Rpt_{Period} {Eff_Year}',
        f'RMKTF P{Period}',
        f'VENDFUND - P{Period}'
    ]

    # Add sheets for the previous year if needed
    if Year is not None:
        sheets_to_merge.extend([
            f'CONV P12 {Year}',
            f'Defective Rpt_{Year} P{Period} {Eff_Year}',
            f'RMKTF P12',
            f'VENDFUND - P12',
            'ADV', 'FUNC', 'RIF'
        ])

    for sheet_name in xls.sheet_names:
        if sheet_name in sheets_to_merge:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Filter out rows where the 'Year' column is not null
            df = df[df['Year'].notnull()]
            
            rebate_data_frames.append(df)

# Loop through each Excel file for incentives
for file in incentive_files:
    file_path = get_file_path(file, Year)
    if not os.path.exists(file_path):
        continue
    xls = pd.ExcelFile(file_path)

    # List of sheet names to merge from each file for incentives
    sheets_to_merge = [
        f'RMTKFINC P{Period}',  # Add for the effective year
    ]
    if Year is not None:
        sheets_to_merge.append(f'RMTKFINC P12')  # Add for the previous year
    common_INC_sheet = ['INCENT']
    sheets_to_merge.extend(common_INC_sheet)

    for sheet_name in xls.sheet_names:
        if sheet_name in sheets_to_merge:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Filter out rows where the 'Year' column is not null
            df = df[df['Year'].notnull()]
            
            incentive_data_frames.append(df)

# Concatenate all DataFrames for rebates and incentives
merged_rebate_data = pd.concat(rebate_data_frames)
merged_incentive_data = pd.concat(incentive_data_frames)

# Specify the desired sheet names
output_rebate_sheet = 'Rebates_Combined'
output_incentive_sheet = 'INC_TIER_Combined'

# Define the output directory based on whether previous year's data is needed
output_directory = Eff_Year_Directory

# Write merged data to Excel file
output_file_path = os.path.join(output_directory, output_combined_file)
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    merged_rebate_data.to_excel(writer, sheet_name=output_rebate_sheet, index=False)
    merged_incentive_data.to_excel(writer, sheet_name=output_incentive_sheet, index=False)

print(f'Merged data saved to {output_combined_file} with sheets {output_rebate_sheet} and {output_incentive_sheet}')

# Open the file
os.system(f'start excel "{os.path.join(output_directory, output_combined_file)}"')
