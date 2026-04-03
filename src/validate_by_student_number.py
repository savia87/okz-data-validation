import os
import pandas as pd

BASE_DIR = r"C:\Users\140086\Documents\okz-data-validation"
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

SOURCE_FILE = os.path.join(DATA_DIR, "GrantsTracking.xlsx")
TARGET_FILE = os.path.join(DATA_DIR, "OKZ_data.csv")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "student_number_based_validation.xlsx")

SOURCE_HEADER_ROW = 0


def clean_value(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def clean_column_name(col):
    return " ".join(str(col).replace("\n", " ").replace("\r", " ").split()).strip()


def build_compare_name(col):
    text = clean_column_name(col).lower()

    replacements = {
        "grant name": "grantname",
        "student number": "studentnumber",
        "total iss/ pass": "totalisspass",
        "total iss/pass": "totalisspass",
        "social studies": "socialstudies",
        "conduct code q1 ela": "conductcodeq1ela",
        "conduct code q2 ela": "conductcodeq2ela",
        "conduct code q3 ela": "conductcodeq3ela",
        "conduct code q4 ela": "conductcodeq4ela",
    }

    if text in replacements:
        text = replacements[text]

    text = (
        text.replace("social studies", "socialstudies")
            .replace(" / ", "/")
            .replace("/ ", "/")
            .replace(" /", "/")
            .replace(" ", "")
    )

    return text


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    source = pd.read_excel(SOURCE_FILE, dtype=str, header=SOURCE_HEADER_ROW)
    target = pd.read_csv(TARGET_FILE, dtype=str)

    source.columns = [clean_column_name(col) for col in source.columns]
    target.columns = [clean_column_name(col) for col in target.columns]

    print("Source columns:")
    print(source.columns.tolist())
    print("\nTarget columns:")
    print(target.columns.tolist())

    source_map = {build_compare_name(col): col for col in source.columns}
    target_map = {build_compare_name(col): col for col in target.columns}

    source_key_name = build_compare_name("Student Number")
    target_key_name = build_compare_name("StudentNumber")

    if source_key_name not in source_map:
        print("❌ Source student number column not found")
        return

    if target_key_name not in target_map:
        print("❌ Target student number column not found")
        return

    source_key_col = source_map[source_key_name]
    target_key_col = target_map[target_key_name]

    print(f"\nSource key column: {source_key_col}")
    print(f"Target key column: {target_key_col}")

    source[source_key_col] = source[source_key_col].apply(clean_value)
    target[target_key_col] = target[target_key_col].apply(clean_value)

    source = source[source[source_key_col] != ""].copy()
    target = target[target[target_key_col] != ""].copy()

    source = source.drop_duplicates(subset=[source_key_col]).copy()
    target = target.drop_duplicates(subset=[target_key_col]).copy()

    source = source.set_index(source_key_col)
    target = target.set_index(target_key_col)

    common_keys = source.index.intersection(target.index)
    source_only_keys = source.index.difference(target.index)
    target_only_keys = target.index.difference(source.index)

    compare_pairs = []
    for source_norm, source_col in source_map.items():
        if source_norm == source_key_name:
            continue
        if source_norm in target_map:
            compare_pairs.append((source_col, target_map[source_norm], source_norm))

    print("\nMatched columns for comparison:")
    for source_col, target_col, _ in compare_pairs:
        print(f"{source_col}  <-->  {target_col}")

    result_rows = []
    mismatch_detail_rows = []

    for key in common_keys:
        row_result = {
            "StudentNumber": key,
            "Row_Status": "Match",
            "Mismatch_Columns": ""
        }

        mismatch_cols = []

        for source_col, target_col, normalized_name in compare_pairs:
            source_val = clean_value(source.at[key, source_col])
            target_val = clean_value(target.at[key, target_col])

            row_result[f"{source_col}_Source"] = source_val
            row_result[f"{target_col}_Target"] = target_val
            row_result[f"{source_col}_Match"] = "Y" if source_val == target_val else "N"

            if source_val != target_val:
                row_result["Row_Status"] = "Mismatch"
                mismatch_cols.append(f"{source_col} <> {target_col}")

                mismatch_detail_rows.append({
                    "StudentNumber": key,
                    "Source_Column": source_col,
                    "Target_Column": target_col,
                    "Source_Value": source_val,
                    "Target_Value": target_val
                })

        row_result["Mismatch_Columns"] = ", ".join(mismatch_cols)
        result_rows.append(row_result)

    result_df = pd.DataFrame(result_rows)
    mismatch_only_df = result_df[result_df["Row_Status"] == "Mismatch"].copy()
    mismatch_detail_df = pd.DataFrame(mismatch_detail_rows)

    source_only_df = source.loc[source_only_keys].reset_index()
    target_only_df = target.loc[target_only_keys].reset_index()

    summary_df = pd.DataFrame({
        "Metric": [
            "Source Row Count",
            "Target Row Count",
            "Common Student Numbers",
            "Missing in Target",
            "Extra in Target",
            "Matched Rows",
            "Mismatched Rows"
        ],
        "Value": [
            len(source),
            len(target),
            len(common_keys),
            len(source_only_keys),
            len(target_only_keys),
            (result_df["Row_Status"] == "Match").sum(),
            (result_df["Row_Status"] == "Mismatch").sum()
        ]
    })

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="Full_Comparison", index=False)
        mismatch_only_df.to_excel(writer, sheet_name="Mismatch_Only", index=False)
        mismatch_detail_df.to_excel(writer, sheet_name="Mismatch_Details", index=False)
        source_only_df.to_excel(writer, sheet_name="Missing_in_Target", index=False)
        target_only_df.to_excel(writer, sheet_name="Extra_in_Target", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    print("\n✅ Validation complete")
    print(f"Common student numbers: {len(common_keys)}")
    print(f"Matched rows: {(result_df['Row_Status'] == 'Match').sum()}")
    print(f"Mismatched rows: {(result_df['Row_Status'] == 'Mismatch').sum()}")
    print(f"Missing in target: {len(source_only_keys)}")
    print(f"Extra in target: {len(target_only_keys)}")
    print(f"Output file: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()