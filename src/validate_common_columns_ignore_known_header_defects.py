import pandas as pd
import os
import re

# ===== FILE LOCATION =====
BASE_PATH = r"C:\Users\140086\Documents\okz-data-validation\data"

GRANT_FILE = "GrantsTracking (1).xlsx"
OKZ_FILE = "OKZ_data.csv"

GRANT_PATH = os.path.join(BASE_PATH, GRANT_FILE)
OKZ_PATH = os.path.join(BASE_PATH, OKZ_FILE)

OUTPUT_MATCHED = os.path.join(BASE_PATH, "Common_Columns_Row_Level_Validation.xlsx")
OUTPUT_MISMATCH = os.path.join(BASE_PATH, "Common_Columns_Mismatch_Only.xlsx")


def clean_value(series):
    return (
        series.astype(str)
        .str.replace("\n", " ", regex=True)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )


def clean_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def normalize_column_name(col_name):
    """
    Normalize header names so known formatting/spelling differences do not block comparison.
    """
    col = str(col_name).strip().lower()
    col = col.replace("\n", " ")
    col = re.sub(r"\s+", " ", col)

    # remove spaces, hyphens, slashes, periods for easier comparison
    col = re.sub(r"[ \-\/\.]", "", col)

    # known naming/spelling normalization
    col = col.replace("grantname", "grantname")
    col = col.replace("firstname", "firstname")
    col = col.replace("lastname", "lastname")
    col = col.replace("middlename", "middlename")
    col = col.replace("studentnumber", "studentnumber")

    # known spelling issue
    col = col.replace("unexecused", "unexcused")

    # social studies formatting
    col = col.replace("socialstudies", "socialstudies")

    # ese / lep
    col = col.replace("eseother", "eseother")
    col = col.replace("lepcode", "lepcode")

    # dates/comments/year formatting
    col = col.replace("entrydate", "entrydate")
    col = col.replace("exitdate", "exitdate")
    col = col.replace("startyear", "startyear")
    col = col.replace("endyear", "endyear")
    col = col.replace("entrancecomments", "entrancecomments")
    col = col.replace("exitcomments", "exitcomments")

    return col


def main():
    # Grant Tracking headers are in row 2
    grant_df = pd.read_excel(GRANT_PATH, header=1)
    okz_df = pd.read_csv(OKZ_PATH)

    grant_df = clean_columns(grant_df)
    okz_df = clean_columns(okz_df)

    print("Grant Tracking columns:")
    print(grant_df.columns.tolist())
    print("\nOKZ Data columns:")
    print(okz_df.columns.tolist())

    # business key
    grant_key = "Student Number"
    okz_key = "StudentNumber"

    grant_df[grant_key] = clean_value(grant_df[grant_key])
    okz_df[okz_key] = clean_value(okz_df[okz_key])

    # remove blank keys
    grant_df = grant_df[~grant_df[grant_key].isin(["", "nan", "None"])].copy()
    okz_df = okz_df[~okz_df[okz_key].isin(["", "nan", "None"])].copy()

    # add normalized column maps
    grant_map = {normalize_column_name(c): c for c in grant_df.columns}
    okz_map = {normalize_column_name(c): c for c in okz_df.columns}

    # known extra columns in OKZ to ignore
    ignore_okz_only = {
        "entitycode",
        "dimstudentid",
        "rundate",
    }

    # find common normalized columns
    common_normalized_cols = sorted(
        (set(grant_map.keys()) & set(okz_map.keys())) - ignore_okz_only
    )

    # remove key from compare list, but keep it for join
    compare_normalized_cols = [c for c in common_normalized_cols if c != "studentnumber"]

    print("\nCommon columns used for comparison:")
    for c in compare_normalized_cols:
        print(f"{grant_map[c]}  <-->  {okz_map[c]}")

    # rename common columns to normalized names for comparison
    grant_compare = grant_df[[grant_map[c] for c in common_normalized_cols]].copy()
    okz_compare = okz_df[[okz_map[c] for c in common_normalized_cols]].copy()

    grant_compare = grant_compare.rename(columns={grant_map[c]: c for c in common_normalized_cols})
    okz_compare = okz_compare.rename(columns={okz_map[c]: c for c in common_normalized_cols})

    # clean all comparable values
    for col in grant_compare.columns:
        grant_compare[col] = clean_value(grant_compare[col])
    for col in okz_compare.columns:
        okz_compare[col] = clean_value(okz_compare[col])

    # if duplicates exist, row-level merge on studentnumber can multiply rows
    # still useful for review, but flag it
    grant_dup_count = grant_compare.duplicated(subset=["studentnumber"], keep=False).sum()
    okz_dup_count = okz_compare.duplicated(subset=["studentnumber"], keep=False).sum()

    print("\nDuplicate key rows:")
    print("Grant Tracking duplicate Student Number rows:", grant_dup_count)
    print("OKZ duplicate StudentNumber rows:", okz_dup_count)

    merged = grant_compare.merge(
        okz_compare,
        on="studentnumber",
        how="left",
        suffixes=("_grant", "_okz"),
        indicator=True
    )

    # compare field by field
    mismatch_summary = []
    for col in compare_normalized_cols:
        grant_col = f"{col}_grant"
        okz_col = f"{col}_okz"
        status_col = f"{col}_status"

        merged[status_col] = merged.apply(
            lambda row: "Match" if row[grant_col] == row[okz_col] else "Mismatch",
            axis=1
        )

        mismatch_count = (merged[status_col] == "Mismatch").sum()
        mismatch_summary.append({
            "Column": col,
            "Grant Header": grant_map[col],
            "OKZ Header": okz_map[col],
            "Mismatch Count": int(mismatch_count)
        })

    summary_df = pd.DataFrame(mismatch_summary)

    # rows where at least one compared column mismatched
    status_cols = [f"{c}_status" for c in compare_normalized_cols]
    mismatch_only = merged[
        (merged["_merge"] != "both") |
        (merged[status_cols].eq("Mismatch").any(axis=1))
    ].copy()

    # write output
    with pd.ExcelWriter(OUTPUT_MATCHED, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Row_Level_Validation", index=False)
        summary_df.to_excel(writer, sheet_name="Mismatch_Summary", index=False)

    mismatch_only.to_excel(OUTPUT_MISMATCH, index=False)

    print("\n===== VALIDATION SUMMARY =====")
    print("Rows in Grant Tracking:", len(grant_compare))
    print("Rows in OKZ Data:", len(okz_compare))
    print("Rows in merged output:", len(merged))
    print("Rows with mismatch or missing match:", len(mismatch_only))
    print(f"\nDetailed file created: {OUTPUT_MATCHED}")
    print(f"Mismatch-only file created: {OUTPUT_MISMATCH}")


if __name__ == "__main__":
    main()