import pandas as pd
import os

def add_monthly_visited_dates(month_files: dict, output_file: str) -> None:
    """
    Reads multiple monthly DCR CSVs and combines them into one,
    adding new columns (e.g. Jul, Aug, Sep) that contain the visited dates
    for each (Territory Code, Account: Customer Code).

    Parameters
    ----------
    month_files : dict
        Dictionary in format {'Jul': 'file_july.csv', 'Aug': 'file_august.csv', ...}
    output_file : str
        Path for saving the combined CSV file
    """

    all_month_data = []

    for month_name, file_path in month_files.items():
        if not os.path.exists(file_path):
            print(f"⚠️ File not found: {file_path}")
            continue

        # Read file
        df = pd.read_csv(file_path)
        required_cols = ["Territory Code", "Account: Customer Code", "Date"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"{file_path} missing columns: {missing_cols}")

        # Clean data
        df.columns = df.columns.str.strip()
        df["Territory Code"] = df["Territory Code"].astype(str).str.strip().str.replace(";", "", regex=False)
        df["Account: Customer Code"] = df["Account: Customer Code"].astype(str).str.strip()
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
        df["Day"] = df["Date"].dt.day.astype(str).str.zfill(2)

        # Aggregate visit days
        visit_days = (
            df.groupby(["Territory Code", "Account: Customer Code"])["Day"]
            .apply(lambda x: ",".join(sorted(x.unique())))
            .reset_index()
            .rename(columns={"Day": month_name})
        )

        all_month_data.append(visit_days)

    if not all_month_data:
        raise ValueError("❌ No valid input files processed.")

    # Merge all months on Territory + Account
    combined_df = all_month_data[0]
    for df_month in all_month_data[1:]:
        combined_df = pd.merge(
            combined_df,
            df_month,
            on=["Territory Code", "Account: Customer Code"],
            how="outer"
        )

    # Save final combined file
    combined_df.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"✅ Combined file created successfully: {output_file}")
    print(f"📊 Total unique (Territory, Account) pairs: {len(combined_df)}")
    print(f"🆕 Added columns: {', '.join(month_files.keys())}")



# === Example Usage ===
if __name__ == "__main__":
    month_files = {
        "Jul": "DCR Report ASC Jul.csv",
        "Aug": "DCR Report ASC Aug.csv",
        "Sep": "DCR Report ASC Sep.csv"
    }
    output_path = "DCR_Report_ASC_Jul_Aug_Sep_Combined.csv"

    add_monthly_visited_dates(month_files, output_path)
