import pandas as pd
import os
from datetime import datetime

def summarize_visits(input_file: str, output_file: str) -> None:
    """
    Summarizes how many times each unique 'Account: Customer Code'
    visited a particular 'Territory Code', and lists all visit dates.

    Parameters
    ----------
    input_file : str
        Path to the input CSV file (DCR report)
    output_file : str
        Path where the summary CSV will be saved
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
    df["Territory Code"] = df["Territory Code"].astype(str).str.strip()
    df["Account: Customer Code"] = df["Account: Customer Code"].astype(str).str.strip()

    # Remove any trailing semicolons in territory codes if present
    df["Territory Code"] = df["Territory Code"].str.replace(";", "", regex=False)

    # === Step 5: Convert 'Date' to datetime and back to uniform format (dd-mm-yyyy) ===
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])  # drop rows with invalid dates
    df["Date"] = df["Date"].dt.strftime("%d-%m-%Y")

    # === Step 6: Group data ===
    grouped = (
        df.groupby(["Territory Code", "Account: Customer Code"], as_index=False)
        .agg(
            Visit_Count=("Date", "count"),
            Date=("Date", lambda x: ",".join(sorted(x.unique())))
        )
    )

    # === Step 7: Save the output ===
    grouped.to_csv(output_file, index=False, encoding="utf-8-sig")

    print(f"✅ Summary successfully generated: {output_file}")
    print(f"📊 Total records: {len(grouped)}")

# === Example Usage ===
if __name__ == "__main__":
    input_path = "DCR Report APC Sep.csv"  # your input file
    output_path = "territory_customer_visit_summary.csv"  # output file
    summarize_visits(input_path, output_path)
