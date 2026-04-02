import pandas as pd
import os

# ===== FILE LOCATION =====
BASE_PATH = r"C:\Users\140086\Documents\okz-data-validation\data"

GRANT_FILE = "GrantsTracking (1).xlsx"
OKZ_FILE = "OKZ_data.csv"

GRANT_PATH = os.path.join(BASE_PATH, GRANT_FILE)
OKZ_PATH = os.path.join(BASE_PATH, OKZ_FILE)
OUTPUT_FILE = os.path.join(BASE_PATH, "Missing_StudentNumber.xlsx")


def clean_student_number(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )


def clean_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ", regex=True)   # 🔴 FIX line breaks
        .str.replace("  ", " ", regex=True)   # remove double spaces
        .str.strip()
    )
    return df


def main():
    grant_df = pd.read_excel(GRANT_PATH, header=1)
    okz_df = pd.read_csv(OKZ_PATH)

    # 🔴 CLEAN COLUMN NAMES (CRITICAL FIX)
    grant_df = clean_columns(grant_df)
    okz_df = clean_columns(okz_df)

    print("Grant Tracking columns:")
    print(grant_df.columns.tolist())

    print("\nOKZ Data columns:")
    print(okz_df.columns.tolist())

    # Now this will work
    grant_col = "Student Number"
    okz_col = "StudentNumber"

    # Clean values
    grant_df[grant_col] = clean_student_number(grant_df[grant_col])
    okz_df[okz_col] = clean_student_number(okz_df[okz_col])

    # Remove blanks
    grant_students = set(grant_df[grant_col].dropna()) - {"", "nan", "None"}
    okz_students = set(okz_df[okz_col].dropna()) - {"", "nan", "None"}

    # Compare
    missing_in_okz = sorted(grant_students - okz_students)

    print("\n===== VALIDATION RESULT =====")
    print("Total StudentNumber in Grant Tracking:", len(grant_students))
    print("Total StudentNumber in OKZ Data:", len(okz_students))
    print("Missing in OKZ Data:", len(missing_in_okz))

    if missing_in_okz:
        missing_df = pd.DataFrame(missing_in_okz, columns=["StudentNumber"])
        missing_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n❌ Mismatch found. File created: {OUTPUT_FILE}")
    else:
        print("\n✅ Validation Passed: All StudentNumbers are present.")


if __name__ == "__main__":
    main()