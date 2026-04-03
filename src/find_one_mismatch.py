import pandas as pd

source_file = r"C:\Users\140086\Documents\okz-data-validation\data\GrantsTracking.xlsx"
target_file = r"C:\Users\140086\Documents\okz-data-validation\data\OKZ_data.csv"


def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


# Load files
source = pd.read_excel(source_file, dtype=str)
target = pd.read_csv(target_file, dtype=str)

# Clean column names
source.columns = [col.strip() for col in source.columns]
target.columns = [col.strip() for col in target.columns]

# Find common columns
common_cols = [col for col in source.columns if col in target.columns]

print("Comparing columns:", common_cols)

# Clean values
for col in common_cols:
    source[col] = source[col].apply(clean)
    target[col] = target[col].apply(clean)

# Create row keys
source["key"] = source[common_cols].agg("||".join, axis=1)
target["key"] = target[common_cols].agg("||".join, axis=1)

# Find mismatch
mismatch = source[~source["key"].isin(target["key"])]

if mismatch.empty:
    print("✅ All rows match perfectly")
else:
    print("\n❌ Found mismatch row:\n")
    print(mismatch.iloc[0])   # 👈 ONLY ONE ROW