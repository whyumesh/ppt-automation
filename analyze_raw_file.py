"""
Analyze Raw File vs Working File Sheet
Compare raw file structure with corresponding Working File sheet to identify operations
"""

import pandas as pd
import os
from pathlib import Path

def analyze_raw_file(raw_file_path: str, working_file_path: str, working_sheet_name: str):
    """Compare raw file with Working File sheet."""
    print("=" * 80)
    print(f"ANALYZING: {os.path.basename(raw_file_path)}")
    print(f"COMPARING WITH: Working File -> {working_sheet_name} sheet")
    print("=" * 80)
    
    # Load raw file
    print("\n📄 RAW FILE STRUCTURE:")
    print("-" * 80)
    
    file_ext = os.path.splitext(raw_file_path)[1].lower()
    
    if file_ext == '.xlsb':
        from pyxlsb import open_workbook
        with open_workbook(raw_file_path) as wb:
            sheets = list(wb.sheets)
            print(f"Sheets: {sheets}")
            
            # Analyze first sheet
            if sheets:
                sheet_name = sheets[0]
                print(f"\nAnalyzing sheet: {sheet_name}")
                
                rows = []
                for i, row in enumerate(wb.get_sheet(sheet_name)):
                    if i < 20:  # First 20 rows
                        row_data = [cell.v if hasattr(cell, 'v') else None for cell in row]
                        rows.append(row_data)
                        if i < 10:
                            print(f"Row {i}: {row_data[:10]}")  # First 10 columns
                
                # Try to identify header row
                print(f"\nTotal rows analyzed: {len(rows)}")
    else:
        # .xlsx file
        excel_file = pd.ExcelFile(raw_file_path)
        print(f"Sheets: {excel_file.sheet_names}")
        
        if excel_file.sheet_names:
            sheet_name = excel_file.sheet_names[0]
            print(f"\nAnalyzing sheet: {sheet_name}")
            
            # Try different header rows
            for header_row in [0, 1, 2]:
                try:
                    df = pd.read_excel(raw_file_path, sheet_name=sheet_name, header=header_row, nrows=20)
                    print(f"\nWith header={header_row}:")
                    print(f"  Columns: {list(df.columns)[:10]}")
                    print(f"  Shape: {df.shape}")
                    print(f"  First row: {df.iloc[0].tolist()[:10] if len(df) > 0 else 'Empty'}")
                except Exception as e:
                    print(f"  Error with header={header_row}: {e}")
    
    # Load Working File sheet
    print("\n\n📊 WORKING FILE SHEET STRUCTURE:")
    print("-" * 80)
    
    try:
        # Try different header rows for Working File
        for header_row in [0, 1, 2, 3, 4]:
            try:
                df_working = pd.read_excel(working_file_path, sheet_name=working_sheet_name, header=header_row, nrows=20)
                print(f"\nWith header={header_row}:")
                print(f"  Columns: {list(df_working.columns)}")
                print(f"  Shape: {df_working.shape}")
                if len(df_working) > 0:
                    print(f"  First row: {df_working.iloc[0].tolist()}")
                    print(f"  Sample data types: {df_working.dtypes.to_dict()}")
            except Exception as e:
                print(f"  Error with header={header_row}: {e}")
    except Exception as e:
        print(f"Error loading Working File sheet: {e}")
    
    print("\n" + "=" * 80)
    print("COMPARISON COMPLETE")
    print("=" * 80)

if __name__ == "__main__":
    # Analyze consent file
    raw_file = "Data/Apr 2025/Reports/AIL Consented Status HCP's_02.04.2025.xlsb"
    working_file = "Data/Apr 2025/AIL LT Working file.xlsx"
    working_sheet = "consent"
    
    if os.path.exists(raw_file) and os.path.exists(working_file):
        analyze_raw_file(raw_file, working_file, working_sheet)
    else:
        print(f"Files not found:")
        print(f"  Raw: {raw_file} - {os.path.exists(raw_file)}")
        print(f"  Working: {working_file} - {os.path.exists(working_file)}")

