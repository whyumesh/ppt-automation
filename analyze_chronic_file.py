"""
Analyze Chronic Missing Report file to understand structure
"""

import pandas as pd
import os

raw_file = "Data/Apr 2025/Reports/Chronic Missing Report AIL - Jan to Mar.xlsx"
working_file = "Data/Apr 2025/AIL LT Working file.xlsx"
working_sheet = "Chronic & Overcalling"

print("=" * 80)
print("ANALYZING CHRONIC MISSING REPORT FILE")
print("=" * 80)

# Load raw file
print("\n📄 RAW FILE STRUCTURE:")
print("-" * 80)

if os.path.exists(raw_file):
    excel_file = pd.ExcelFile(raw_file)
    print(f"Sheets: {excel_file.sheet_names}")
    
    if excel_file.sheet_names:
        sheet_name = excel_file.sheet_names[0]
        print(f"\nAnalyzing sheet: {sheet_name}")
        
        # Try different header rows
        for header_row in [0, 1, 2, 3, 4]:
            try:
                df = pd.read_excel(raw_file, sheet_name=sheet_name, header=header_row, nrows=20)
                print(f"\nWith header={header_row}:")
                print(f"  Columns: {list(df.columns)[:15]}")
                print(f"  Shape: {df.shape}")
                if len(df) > 0:
                    print(f"  First row: {df.iloc[0].tolist()[:10]}")
            except Exception as e:
                print(f"  Error with header={header_row}: {e}")
else:
    print(f"File not found: {raw_file}")

# Load Working File sheet
print("\n\n📊 WORKING FILE SHEET STRUCTURE:")
print("-" * 80)

if os.path.exists(working_file):
    try:
        # Try different header rows
        for header_row in [0, 1, 2, 3, 4]:
            try:
                df_working = pd.read_excel(working_file, sheet_name=working_sheet, header=header_row, nrows=20)
                print(f"\nWith header={header_row}:")
                print(f"  Columns: {list(df_working.columns)[:10]}")
                print(f"  Shape: {df_working.shape}")
                if len(df_working) > 0:
                    print(f"  First row: {df_working.iloc[0].tolist()[:10]}")
            except Exception as e:
                print(f"  Error with header={header_row}: {e}")
    except Exception as e:
        print(f"Error loading Working File sheet: {e}")
else:
    print(f"File not found: {working_file}")

print("\n" + "=" * 80)

