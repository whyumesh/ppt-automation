import pandas as pd
import os
from datetime import datetime

def add_monthly_visit_columns(july_file, august_file, september_file, output_file):
    """
    Combines DCR data from July, August, and September and adds
    3 new columns ('Jul', 'Aug', 'Sep') containing all unique visit
    days (dd,dd,dd...) for each (Territory Code, Account: Customer Code)
    while keeping all other original columns intact.
    """

    # === Step 1: Validate files ===
    for file in [july_file, august_file, september_file]:
        if not os.path.exists(file):
            raise FileNotFoundError(f"Input file not found: {file}")

    # === Step 2: Read and clean monthly files ===
    def read_and_prepare(file, month_label):
        df = pd.read_csv(file)
        df.columns = df.columns.str.strip()
        if "Territory Code" not in df.columns or "Account: Customer Code" not in df.columns or "Date" not in df.columns:
            raise ValueError(f"Missing required columns in {file}")
        df["Territory Code"] = df["Territory Code"].astype(str).str.strip().str.replace(";", "", regex=False)
        df["Account: Customer Code"] = df["Account: Customer Code"].astype(str).str.strip()
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
        df["Day"] = df["Date"].dt.day.astype(str).str.zfill(2)
        return df

    df_jul = read_and_prepare(july_file, "Jul")
    df_aug = read_and_prepare(august_file, "Aug")
    df_sep = read_and_prepare(september_file, "Sep")

    # === Step 3: Create month-wise visit mappings ===
    def create_visit_map(df):
        return (
            df.groupby(["Territory Code", "Account: Customer Code"])["Day"]
            .apply(lambda x: ",".join(sorted(x.unique())))
            .to_dict()
        )

    visit_map_jul = create_visit_map(df_jul)
    visit_map_aug = create_visit_map(df_aug)
    visit_map_sep = create_visit_map(df_sep)

    # === Step 4: Use September file as base (or any month) to keep all original columns ===
    base_df = pd.concat([df_jul, df_aug, df_sep], ignore_index=True)
    base_df = base_df.drop_duplicates(subset=["Territory Code", "Account: Customer Code", "Date"])

    # === Step 5: Add month columns ===
    base_df["Jul"] = base_df.apply(
        lambda r: visit_map_jul.get((r["Territory Code"], r["Account: Customer Code"]), ""), axis=1
    )
    base_df["Aug"] = base_df.apply(
        lambda r: visit_map_aug.get((r["Territory Code"], r["Account: Customer Code"]), ""), axis=1
    )
    base_df["Sep"] = base_df.apply(
        lambda r: visit_map_sep.get((r["Territory Code"], r["Account: Customer Code"]), ""), axis=1
    )

    # === Step 6: Save final file ===
    base_df.to_csv(output_file, index=False, encoding="utf-8-sig")

    print(f"✅ File successfully created: {output_file}")
    print(f"📊 Total records: {len(base_df)}")
    print("🆕 Added columns: 'Jul', 'Aug', 'Sep'")


# === Example Usage ===
if __name__ == "__main__":
    july_path = "DCR Report APC Jul 0810.csv"
    august_path = "DCR Report APC Aug 0810.csv"
    september_path = "DCR Report APC Sep 0810.csv"
    output_path = "DCR_Report_with_Jul_Aug_Sep_Visits.csv"

    add_monthly_visit_columns(july_path, august_path, september_path, output_path)
