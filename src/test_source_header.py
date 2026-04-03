import pandas as pd

SOURCE_FILE = r"C:\Users\140086\Documents\okz-data-validation\data\GrantsTracking.xlsx"

for i in range(0, 8):
    df = pd.read_excel(SOURCE_FILE, dtype=str, header=i, nrows=2)
    print(f"\nHEADER ROW = {i}")
    print(df.columns.tolist())