import pandas as pd
import os
from datetime import datetime

def combine_monthly_visits(input_files, output_file):
    """
    Combines multiple monthly DCR files and creates a single dataset
    that includes all original columns + a new column 'Visited Date'
    listing all unique visit days (dd,dd,dd...) for each unique
    (Territory Code, Account: Customer Code) pair across all months.

    Parameters
    ----------
    input_files : list[str]
        List of CSV file paths for different months
    output_file : str
        Path where the combined output CSV will be saved
    """

    # === Step 1: Validate input files ===
    for file in input_files:
        if not os.path.exists(file):
            raise FileNotFoundError(f"Input file not found: {file}")

    # === Step 2: Read and combine all monthly files ===
    all_data = []
    for file in input_files:
        try:
            df = pd.read_csv(file)
            df["Source File"] = os.path.basename(file)  # Optional: keep track of origin
            all_data.append(df)
        except Exception as e:
            raise ValueError(f"Error reading file {file}: {e}")

    combined_df = pd.concat(all_data, ignore_index=True)

    # === Step 3: Validate required columns ===
    required_cols = ["Territory Code", "Account: Customer Code", "Date"]
    missing_cols = [col for col in required_cols if col not in combined_df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    # === Step 4: Clean and normalize ===
    combined_df.columns = combined_df.columns.str.strip()
    combined_df["Territory Code"] = (
        combined_df["Territory Code"].astype(str).str.strip().str.replace(";", "", regex=False)
    )
    combined_df["Account: Customer Code"] = combined_df["Account: Customer Code"].astype(str).str.strip()

    # === Step 5: Convert 'Date' column to datetime and extract day ===
    combined_df["Date"] = pd.to_datetime(combined_df["Date"], errors="coerce")
    combined_df = combined_df.dropna(subset=["Date"])
    combined_df["Day"] = combined_df["Date"].dt.day.astype(str).str.zfill(2)

    # === Step 6: Create mapping for (Territory, Account) → all visit days across months ===
    visit_map = (
        combined_df.groupby(["Territory Code", "Account: Customer Code"])["Day"]
        .apply(lambda x: ",".join(sorted(x.unique())))
        .to_dict()
    )

    # === Step 7: Map 'Visited Date' back to each record ===
    combined_df["Visited Date"] = combined_df.apply(
        lambda row: visit_map.get((row["Territory Code"], row["Account: Customer Code"]), ""),
        axis=1
    )

    # === Step 8: Save final combined file ===
    combined_df.to_csv(output_file, index=False, encoding="utf-8-sig")

    print(f"✅ Combined file successfully generated: {output_file}")
    print(f"📊 Total records: {len(combined_df)}")
    print("🆕 Added column: 'Visited Date' (aggregated across all months)")


# === Example Usage ===
if __name__ == "__main__":
    input_files = [
        "DCR Report APC July.csv",
        "DCR Report APC August.csv",
        "DCR Report APC September.csv"
    ]
    output_file = "DCR_Report_July_Aug_Sep_Combined.csv"

    combine_monthly_visits(input_files, output_file)
