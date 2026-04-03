import os
import pandas as pd

BASE_DIR = r"C:\Users\140086\Documents\okz-data-validation"
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

SOURCE_FILE = os.path.join(DATA_DIR, "GrantsTracking.xlsx")
TARGET_FILE = os.path.join(DATA_DIR, "OKZ_data.csv")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "full_row_column_match_validation.xlsx")


def clean_column_name(col):
    return str(col).strip().replace("\n", " ").replace("\r", " ")


def clean_value(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    source_df = pd.read_excel(SOURCE_FILE, dtype=str)
    target_df = pd.read_csv(TARGET_FILE, dtype=str)

    source_df.columns = [clean_column_name(col) for col in source_df.columns]
    target_df.columns = [clean_column_name(col) for col in target_df.columns]

    common_columns = [col for col in source_df.columns if col in target_df.columns]

    if not common_columns:
        print("No common columns found.")
        return

    print("Common columns used for comparison:")
    for col in common_columns:
        print(col)

    source_compare = source_df[common_columns].copy()
    target_compare = target_df[common_columns].copy()

    for col in common_columns:
        source_compare[col] = source_compare[col].apply(clean_value)
        target_compare[col] = target_compare[col].apply(clean_value)

    source_compare = source_compare.reset_index(drop=True)
    target_compare = target_compare.reset_index(drop=True)

    max_rows = max(len(source_compare), len(target_compare))
    source_compare = source_compare.reindex(range(max_rows)).fillna("")
    target_compare = target_compare.reindex(range(max_rows)).fillna("")

    result_df = pd.DataFrame()
    result_df["Row_Number"] = range(1, max_rows + 1)

    mismatch_flags = []

    for col in common_columns:
        result_df[f"Source_{col}"] = source_compare[col]
        result_df[f"Target_{col}"] = target_compare[col]
        result_df[f"{col}_Match"] = source_compare[col] == target_compare[col]

    for i in range(max_rows):
        row_match = True
        for col in common_columns:
            if source_compare.at[i, col] != target_compare.at[i, col]:
                row_match = False
                break
        mismatch_flags.append("Match" if row_match else "Mismatch")

    result_df["Row_Status"] = mismatch_flags

    mismatch_only_df = result_df[result_df["Row_Status"] == "Mismatch"].copy()

    summary_df = pd.DataFrame({
        "Metric": [
            "Source Row Count",
            "Target Row Count",
            "Compared Row Count",
            "Common Column Count",
            "Matched Rows",
            "Mismatched Rows"
        ],
        "Value": [
            len(source_df),
            len(target_df),
            max_rows,
            len(common_columns),
            (result_df["Row_Status"] == "Match").sum(),
            (result_df["Row_Status"] == "Mismatch").sum()
        ]
    })

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="Row_Level_Comparison", index=False)
        mismatch_only_df.to_excel(writer, sheet_name="Mismatch_Only", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    print("\nValidation completed.")
    print(f"Source row count: {len(source_df)}")
    print(f"Target row count: {len(target_df)}")
    print(f"Compared row count: {max_rows}")
    print(f"Matched rows: {(result_df['Row_Status'] == 'Match').sum()}")
    print(f"Mismatched rows: {(result_df['Row_Status'] == 'Mismatch').sum()}")
    print(f"Output created: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()