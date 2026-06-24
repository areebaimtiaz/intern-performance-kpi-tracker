"""
 loads the CLEANED CSV (interns_cleaned.csv) into a SQLite database.
Run this every time you have a new month's cleaned export.

Usage:
    python load_data.py interns_cleaned.csv
"""
import sys
import pandas as pd
import sqlite3

DB_PATH = "interns.db"
TABLE_NAME = "interns"

def load_csv(csv_path: str):
    df = pd.read_csv(csv_path)
    conn = sqlite3.connect(DB_PATH)
    df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)
    conn.close()
    print(f"Loaded {len(df)} rows from {csv_path} into {DB_PATH} (table: {TABLE_NAME})")

if __name__ == "__main__":
    csv_file = sys.argv[1] if len(sys.argv) > 1 else "interns_cleaned.csv"
    load_csv(csv_file)
