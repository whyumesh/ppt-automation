"""
Extract precise mappings from Excel files by reading actual data
"""
import pandas as pd
import json
from pathlib import Path

def analyze_excel_sheet_precise(excel_path, sheet_name):
    """Analyze a specific sheet with precise column identification"""
    try:
        # Read without header first to see raw structure
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=20)
        
        # Try to identify header row
        header_row = None
        for idx in range(min(5, len(df_raw))):
            row = df_raw.iloc[idx]
            # Check if row looks like headers (mostly strings, few numbers)
            string_count = sum(1 for val in row if isinstance(val, str) and str(val).strip())
            if string_count > len(row) * 0.5:
                header_row = idx
                break
        
        # Read with identified header
        if header_row is not None:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row, nrows=50)
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=50)
        
        analysis = {
            "sheet_name": sheet_name,
            "header_row": header_row if header_row is not None else 0,
            "columns": list(df.columns),
            "row_count": len(df),
            "sample_data": df.head(10).to_dict('records'),
            "data_types": {col: str(dtype) for col, dtype in df.dtypes.items()}
        }
        
        return analysis
    except Exception as e:
        return {"sheet_name": sheet_name, "error": str(e)}

def main():
    excel_path = "Data/Apr 2025/AIL LT Working file.xlsx"
    
    # Get all sheet names
    excel_file = pd.ExcelFile(excel_path)
    sheets = excel_file.sheet_names
    
    print("="*80)
    print("PRECISE EXCEL SHEET ANALYSIS")
    print("="*80)
    
    results = {}
    
    for sheet_name in sheets:
        print(f"\n{'='*80}")
        print(f"Sheet: {sheet_name}")
        print(f"{'='*80}")
        
        analysis = analyze_excel_sheet_precise(excel_path, sheet_name)
        results[sheet_name] = analysis
        
        if "error" in analysis:
            print(f"  ERROR: {analysis['error']}")
            continue
        
        print(f"  Header Row: {analysis['header_row']}")
        print(f"  Total Rows: {analysis['row_count']}")
        print(f"  Columns ({len(analysis['columns'])}):")
        
        for i, col in enumerate(analysis['columns']):
            dtype = analysis['data_types'].get(col, 'unknown')
            print(f"    {i}: '{col}' ({dtype})")
        
        print(f"\n  Sample Data (first 3 rows):")
        for i, row in enumerate(analysis['sample_data'][:3]):
            print(f"    Row {i}:")
            for col, val in list(row.items())[:5]:  # Show first 5 columns
                val_str = str(val)[:50] if val is not None else "NaN"
                print(f"      {col}: {val_str}")
            if len(row) > 5:
                print(f"      ... and {len(row) - 5} more columns")
    
    # Save results
    output_file = "analysis/precise_excel_analysis.json"
    Path("analysis").mkdir(exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str, ensure_ascii=False)
    
    print(f"\n{'='*80}")
    print(f"Analysis saved to: {output_file}")
    print(f"{'='*80}")

if __name__ == "__main__":
    main()

