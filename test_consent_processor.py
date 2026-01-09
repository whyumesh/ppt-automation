"""
Test Consent Processor
Test the consent file processor and compare output with Working File
"""

import pandas as pd
from src.raw_file_processors import ConsentedStatusProcessor

# Test the processor
raw_file = "Data/Apr 2025/Reports/AIL Consented Status HCP's_02.04.2025.xlsb"
working_file = "Data/Apr 2025/AIL LT Working file.xlsx"

print("=" * 80)
print("TESTING CONSENT PROCESSOR")
print("=" * 80)

# Process raw file
processor = ConsentedStatusProcessor()
processed_df = processor.process(raw_file)

print("\n📊 PROCESSED DATA (from raw file):")
print("-" * 80)
print(f"Shape: {processed_df.shape}")
print(f"Columns: {list(processed_df.columns)}")
print("\nFirst 5 rows:")
print(processed_df.head().to_string())

# Load Working File for comparison
print("\n\n📊 WORKING FILE DATA (for comparison):")
print("-" * 80)
df_working = pd.read_excel(working_file, sheet_name="consent", header=0)

# Select same columns for comparison
comparison_cols = ['Division Name', 'DVL', '# HCP Consent', '% Consent Require']
available_cols = [col for col in comparison_cols if col in df_working.columns]

if available_cols:
    df_working_subset = df_working[available_cols].head(10)
    print(f"Shape: {df_working_subset.shape}")
    print(f"Columns: {list(df_working_subset.columns)}")
    print("\nFirst 5 rows:")
    print(df_working_subset.head().to_string())

# Compare
print("\n\n🔍 COMPARISON:")
print("-" * 80)
if 'Division Name' in processed_df.columns and 'Division Name' in df_working.columns:
    processed_divisions = set(processed_df['Division Name'].str.strip().str.lower())
    working_divisions = set(df_working['Division Name'].str.strip().str.lower())
    
    print(f"Processed divisions: {sorted(processed_divisions)}")
    print(f"Working file divisions: {sorted(working_divisions)}")
    print(f"\nMatch: {processed_divisions == working_divisions}")

