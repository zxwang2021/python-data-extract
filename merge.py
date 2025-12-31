import os
import pandas as pd

# Folder containing your .xlsx files
INPUT_FOLDER = r"D:\Workspace\python\dataExtractor"
OUTPUT_FILE = "merged_output.xlsx"   # or .xls if you prefer

def merge_xlsx_files(folder):
    dataframes = []

    for file in os.listdir(folder):
        if file.lower().endswith(".xlsx"):
            path = os.path.join(folder, file)
            print(f"Reading: {path}")
            df = pd.read_excel(path, engine='openpyxl')
            dataframes.append(df)

    if not dataframes:
        print("No .xlsx files found.")
        return

    # Concatenate — union of all columns
    merged_df = pd.concat(dataframes, ignore_index=True, sort=True)

    # Save result
    merged_df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
    print(f"Saved merged file as: {OUTPUT_FILE}")

if __name__ == "__main__":
    merge_xlsx_files(INPUT_FOLDER)