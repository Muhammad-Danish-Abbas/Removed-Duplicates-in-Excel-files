import io
import re
import pandas as pd
from google.colab import files

# ---------- 1. File upload ----------
print("Apni Excel file select karein...")
uploaded = files.upload()
file_name = list(uploaded.keys())[0]
print(f"\nUpload hui: {file_name}")

# ---------- 2. File padhein ----------
if file_name.lower().endswith(".csv"):
    df = pd.read_csv(io.BytesIO(uploaded[file_name]), dtype=str)
else:
    df = pd.read_excel(io.BytesIO(uploaded[file_name]), dtype=str)

print(f"Total rows: {len(df)}")
print("\nColumns:")
for i, c in enumerate(df.columns):
    letter = chr(ord("A") + i) if i < 26 else str(i)
    print(f"  {letter}: {c}")

# ---------- 3. ID column chunein ----------
id_input = input("\nID column ka letter ya naam likhein (jaise B): ").strip()
keep = input("Konsi entry rakhni hai? first (pehli) / last (aakhri): ").strip().lower()
if keep not in ("first", "last"):
    keep = "first"


def find_id_column(df, value):
    for col in df.columns:
        if str(col).strip().lower() == value.lower():
            return col
    if re.fullmatch(r"[A-Za-z]{1,2}", value):
        n = 0
        for ch in value.upper():
            n = n * 26 + (ord(ch) - ord("A") + 1)
        if n - 1 < len(df.columns):
            return df.columns[n - 1]
    raise ValueError(f"Column '{value}' nahi mila.")


id_col = find_id_column(df, id_input)

# ---------- 4. Duplicates hatayein ----------
df["_id_clean"] = (
    df[id_col].fillna("").astype(str).str.strip()
    .str.replace(r"\s+", "", regex=True).str.upper()
)

empty_rows = df[df["_id_clean"] == ""]
with_id = df[df["_id_clean"] != ""]

dup_mask = with_id.duplicated(subset="_id_clean", keep=keep)
removed = with_id[dup_mask]
cleaned = pd.concat([with_id[~dup_mask], empty_rows]).sort_index()

cleaned = cleaned.drop(columns="_id_clean")
removed = removed.drop(columns="_id_clean")
empty_rows = empty_rows.drop(columns="_id_clean")

# ---------- 5. Nayi file banayein aur download karein ----------
out_name = re.sub(r"\.(xlsx|xls|csv)$", "", file_name, flags=re.I) + "_clean.xlsx"

with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
    cleaned.to_excel(writer, sheet_name="Clean Data", index=False)
    removed.to_excel(writer, sheet_name="Removed Duplicates", index=False)
    if len(empty_rows):
        empty_rows.to_excel(writer, sheet_name="Empty ID Rows", index=False)

print("\n" + "=" * 40)
print(f"Total rows        : {len(df)}")
print(f"Duplicates hataye : {len(removed)}")
print(f"Khali ID rows     : {len(empty_rows)} (rakhi gayi)")
print(f"Clean rows        : {len(cleaned)}")
print("=" * 40)

files.download(out_name)
