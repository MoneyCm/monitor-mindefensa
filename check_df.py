import pandas as pd
import sys
try:
    df = pd.read_excel(sys.argv[1], engine='openpyxl', nrows=5)
    print("COLUMNS:")
    print(df.columns.tolist())
    print("FIRST ROWS:")
    print(df.head())
except Exception as e:
    print(f"ERROR: {e}")
