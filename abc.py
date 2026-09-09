import pandas as pd
from sqlalchemy import create_engine

# 1. Database Connection
DATABASE_URL = "postgresql://postgres.mxqtzquofvbcewtaowxo:Dhineshmelroy@aws-0-ap-southeast-2.pooler.supabase.com:6543/postgres"
engine = create_engine(DATABASE_URL)

file_path = 'PeoplesBank_CRM_UAT_Filtered_Report_20260814_0339.xlsx'

# 2. Search through all sheets to find the real data sheet containing 'TC ID'
excel_file = pd.ExcelFile(file_path)
target_df = None
found_sheet = None

for sheet in excel_file.sheet_names:
    # Read the sheet raw to inspect its content
    temp_raw = excel_file.parse(sheet, header=None)
    
    # Check if 'TC ID' exists anywhere in this sheet's cells
    for idx, row in temp_raw.iterrows():
        row_vals = [str(val).strip().upper() for val in row.values]
        if 'TC ID' in row_vals or 'TC_ID' in row_vals:
            # Found the correct sheet and header row!
            target_df = excel_file.parse(sheet, header=idx)
            found_sheet = sheet
            break
    if target_df is not None:
        break

if target_df is None:
    raise ValueError("Could not find any sheet containing 'TC ID' in this workbook. Please check the Excel file.")

print(f"Successfully found data on sheet: '{found_sheet}'")

# 3. Clean up column headers to match the DB columns exactly (Caps + Underscores)
target_df.columns = target_df.columns.astype(str).str.strip().str.upper().str.replace(' ', '_')

# 4. Remove completely empty rows/columns and unnamed artifacts
target_df = target_df.dropna(how='all')
target_df = target_df.dropna(axis=1, how='all')
target_df = target_df.loc[:, ~target_df.columns.str.contains('^UNNAMED', case=False)]

print("Cleaned Columns Found:", target_df.columns.tolist())

# 5. Dynamically locate the test case identifier column
id_col = next((col for col in target_df.columns if 'TC' in col), None)

if id_col:
    print(f"Identified primary ID column as: {id_col}")
    target_df = target_df.dropna(subset=[id_col])
    target_df = target_df[target_df[id_col].astype(str).str.strip() != '']
    target_df = target_df[~target_df[id_col].astype(str).str.contains('TC_ID', case=False)]
    
    if id_col != 'TC_ID':
        target_df = target_df.rename(columns={id_col: 'TC_ID'})
else:
    raise ValueError("CRITICAL ERROR: Could not locate the test case identifier column.")

# 6. UPSERT / Append to Database
try:
    target_df.to_sql('uat_test_cases_v2', engine, if_exists='append', index=False, method='multi')
    print("Friday progress and new modules merged successfully!")
except Exception as e:
    print(f"An error occurred during insertion: {e}")