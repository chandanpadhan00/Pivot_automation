import pandas as pd
from pathlib import Path

# ---------- CONFIG ----------
CSV_PATH = r"/path/to/source.csv"              # <-- your csv
OUTPUT_XLSX = r"/path/to/output.xlsx"          # <-- your excel file
PIVOT_SHEET = "Pivot_Summary"
SOURCE_SHEET = "Source"
# ---------------------------

csv_path = Path(CSV_PATH)
out_path = Path(OUTPUT_XLSX)

# Read CSV
df = pd.read_csv(csv_path)

# Basic cleanup (optional but helpful)
df["drug"] = df["drug"].astype(str).str.strip()
df["case_sub_status_reason_code"] = df["case_sub_status_reason_code"].astype(str).str.strip()

# Ensure case_count is numeric
df["case_count"] = pd.to_numeric(df["case_count"], errors="coerce").fillna(0)

# Create pivot (drug -> reason_code, sum of case_count)
pivot = pd.pivot_table(
    df,
    index=["drug", "case_sub_status_reason_code"],
    values="case_count",
    aggfunc="sum",
    fill_value=0
).reset_index()

# Also create a "collapsed" view like Excel pivot subtotal per drug (optional)
drug_total = df.groupby("drug", as_index=False)["case_count"].sum()
drug_total["case_sub_status_reason_code"] = ""  # blank reason row for totals
drug_total = drug_total[["drug", "case_sub_status_reason_code", "case_count"]]
drug_total = drug_total.rename(columns={"case_count": "case_count"})

# Write into Excel (replace pivot sheet if exists)
mode = "a" if out_path.exists() else "w"

with pd.ExcelWriter(out_path, engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
    # Optional: write source
    df.to_excel(writer, sheet_name=SOURCE_SHEET, index=False)

    # Write pivot summary
    pivot.to_excel(writer, sheet_name=PIVOT_SHEET, index=False)

print(f"Done. Pivot written to: {out_path} (sheet: {PIVOT_SHEET})")
