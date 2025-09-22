import os
import pandas as pd

# Get current script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# File paths (Excel file already on Desktop)
file_path = r"C:\Users\PARAPATLA SAI KUMAR\OneDrive\Desktop\internship DAY-1.xlsx"

# Load Excel file
df = pd.read_excel(file_path)
print("✅ File loaded successfully from:", file_path)
print(df.head())

# Initialize audit info
audit = {
    "original_rows": int(df.shape[0]),
    "original_columns": int(df.shape[1]),
    "duplicates_removed": 0,
    "nan_counts_before": int(df.isna().sum().sum()),
    "nan_counts_after": None,
    "notes": []
}

# Cleaning steps

# 1) Drop exact duplicate rows
dups = df.duplicated()
num_dups = int(dups.sum())
if num_dups > 0:
    df = df.drop_duplicates().reset_index(drop=True)
audit["duplicates_removed"] = num_dups
audit["notes"].append(f"Removed {num_dups} duplicate rows.")

# 2) Handle NaNs
# - Example policies (adjust as needed):
#   a) For numeric columns, fill with median
#   b) For object/string columns, fill with empty string or propagate a value
#   c) For critical columns, drop rows with NaN in those columns (optional)
# Here, we do a simple approach: fill numeric NaNs with median, object NaNs with empty string.
for col in df.columns:
    if pd.api.types.is_numeric_dtype(df[col]):
        median = df[col].median()
        if pd.notna(median):
            df[col] = df[col].fillna(median)
    else:
        df[col] = df[col].fillna("")

nan_after = int(df.isna().sum().sum())
audit["nan_counts_after"] = nan_after
audit["notes"].append(f"Filled NaNs. Total NaNs after cleaning: {nan_after}.")

# 3) Standardize text columns
text_cols = [c for c in df.columns if df[c].dtype == "object"]
for col in text_cols:
    df[col] = df[col].astype(str).str.strip()  # remove leading/trailing spaces
    df[col] = df[col].str.lower()  # optional: normalize casing
audit["notes"].append(f"Standardized text columns: stripped whitespace and lowered case for {len(text_cols)} columns.")

# 4) Convert date-like columns to datetime (best-effort)
for col in df.columns:
    # Try to parse if dtype is object and looks like a date
    if df[col].dtype == "object":
        try:
            parsed = pd.to_datetime(df[col], errors="ignore")
            if not parsed.equals(df[col]):
                df[col] = parsed
                audit["notes"].append(f"Converted column '{col}' to datetime where possible.")
        except Exception:
            pass

# 5) (Optional) Reorder columns or further processing can go here

# Ensure output path exists
output_file = os.path.join(script_dir, "internship_DAY-1_cleaned.xlsx")

# Save cleaned file
df.to_excel(output_file, index=False)
print("✅ Cleaned file saved to:", output_file)

# Generate README report
readme_path = os.path.join(script_dir, "README.md")
readme_lines = []
readme_lines.append("# Cleaning Report")
readme_lines.append("")
readme_lines.append("This report describes the cleaning performed on the input Excel file.")
readme_lines.append("")
readme_lines.append("## Summary")
readme_lines.append("")
readme_lines.append(f"- Original rows: {audit['original_rows']}")
readme_lines.append(f"- Original columns: {audit['original_columns']}")
readme_lines.append(f"- Duplicates removed: {audit['duplicates_removed']}")
readme_lines.append(f"- NaNs before: {audit['nan_counts_before']}")
readme_lines.append(f"- NaNs after: {audit['nan_counts_after']}")
readme_lines.append("")
readme_lines.append("## Details")
readme_lines.append("")
for note in audit["notes"]:
    readme_lines.append(f"- {note}")
readme_lines.append("")


# Write README
with open(readme_path, "w", encoding="utf-8") as f:
    f.write("\n".join(readme_lines))

print("✅ README report generated at:", readme_path)
