import pandas as pd

# ================= CONFIG =================
INPUT_FILE = r"D:\appa\Students Graduation 1021719.xlsx"
OUTPUT_FILE = r"D:\appa\outputcontact.xlsx"

SEARCH_TEXT = "tamil nadu"   # or "coimbatore"
# =========================================

print("Reading Excel file (no column assumptions)...")

# Read Excel WITHOUT trusting headers
df = pd.read_excel(INPUT_FILE, engine="openpyxl", header=0)

# Convert everything to string (safe for scanning)
df_str = df.astype(str).apply(lambda x: x.str.lower())

# Find matching rows ANYWHERE in the row
mask = df_str.apply(
    lambda row: row.str.contains(SEARCH_TEXT.lower(), na=False).any(),
    axis=1
)

filtered_df = df[mask]

# Save full row details
filtered_df.to_excel(OUTPUT_FILE, index=False)

print(f"Done! Output saved as: {OUTPUT_FILE}")
print(f"Matched rows: {len(filtered_df)}")
