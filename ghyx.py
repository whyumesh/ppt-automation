import pandas as pd
import os
from datetime import datetime

def add_visited_dates(input_file: str, output_file: str) -> None:
    """
    Reads a DCR report, identifies how many times each unique
    'Account: Customer Code' has visited a specific 'Territory Code',
    and appends a new column 'Visited Date' with all visit days
    (in dd,dd,dd... format). Keeps all original columns intact.

    Parameters
    ----------
    input_file : str
        Path to input CSV file (DCR report)
    output_file : str
        Path where output CSV will be saved
    """

    # === Step 1: Validate file existence ===
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    # === Step 2: Read CSV ===
    try:
        df = pd.read_csv(input_file)
    except Exception as e:
        raise ValueError(f"Error reading CSV file: {e}")

    # === Step 3: Validate required columns ===
    required_cols = ["Territory Code", "Account: Customer Code", "Date"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    # === Step 4: Clean and normalize columns ===
    df.columns = df.columns.str.strip()
    df["Territory Code"] = df["Territory Code"].astype(str).str.strip().str.replace(";", "", regex=False)
    df["Account: Customer Code"] = df["Account: Customer Code"].astype(str).str.strip()

    # === Step 5: Convert 'Date' to datetime safely ===
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])  # remove invalid dates
    df["Day"] = df["Date"].dt.day.astype(str).str.zfill(2)

    # === Step 6: Create mapping of (Territory, Account) → all visit days ===
    visit_map = (
        df.groupby(["Territory Code", "Account: Customer Code"])["Day"]
        .apply(lambda x: ",".join(sorted(x.unique())))
        .to_dict()
    )

    # === Step 7: Map 'Visited Date' back to each row ===
    df["Visited Date"] = df.apply(
        lambda row: visit_map.get((row["Territory Code"], row["Account: Customer Code"]), ""),
        axis=1
    )

    # === Step 8: Save the updated DataFrame ===
    df.to_csv(output_file, index=False, encoding="utf-8-sig")

    print(f"✅ File successfully processed: {output_file}")
    print(f"📊 Total records: {len(df)}")
    print("🆕 Added column: 'Visited Date'")


# === Example Usage ===
if __name__ == "__main__":
    input_path = "DCR Report APC Sep.csv"           # Input file name
    output_path = "DCR_Report_with_Visited_Date.csv"  # Output file name
    add_visited_dates(input_path, output_path)
