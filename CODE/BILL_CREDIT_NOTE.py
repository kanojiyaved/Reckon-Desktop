import pandas as pd
import os
import random

# Load data
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /BILL CREDIT NOTE/Reckon Desktop Bill Credit - Sheet1.csv")
df_COA_mapping = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /BILL CREDIT NOTE/COA - MAPPING NEW.xlsx - Sheet1.csv")

# Strip column names
df.columns = df.columns.str.strip()
df_COA_mapping.columns = df_COA_mapping.columns.str.strip()

# Filter Bill rows
df = df[(df['Type'] == 'Bill') & (df['Account Type'] != 'Accounts Payable')]
df = df[~df['Account'].str.contains("Tax Payable", na=False)]

# Column renaming
field_mapping = {
    "Trans #": "Bill number",
    "Source Name": "Supplier",
    "Date": "Transaction date",
    "Due Date": "Due date",
    "Item": "Item",
    "Description": "Description",
    "Account": "Account No.",
    "Amount": "Amount ($)",
    "Tax Code": "Tax code",
    "Tax Amount": "Tax amount ($)",
    "Class": "Job name"
}
df = df.rename(columns=field_mapping)

# ✅ Create Supplier invoice number = Bill number
df["Supplier invoice number"] = df["Bill number"]

# ✅ Limit Description column to 1000 characters
if "Description" in df.columns:
    df["Description"] = df["Description"].astype(str).str.slice(0, 1000)

# ✅ Replace NaN/blank in Description with NULL (empty cell in Excel)
df["Description"] = df["Description"].replace("nan", pd.NA)
df["Description"] = df["Description"].replace(r'^\s*$', pd.NA, regex=True)

# Extract digits from Account No.
df["Account No."] = df["Account No."].astype(str).str.extract(r"(\d+)")
old_code_col = next((col for col in df_COA_mapping.columns if "old" in col.lower() and "code" in col.lower()), None)
new_code_col = next((col for col in df_COA_mapping.columns if "new" in col.lower() and "code" in col.lower()), None)
df_COA_mapping[old_code_col] = df_COA_mapping[old_code_col].astype(str).str.extract(r"(\d+)")
df_COA_mapping[new_code_col] = df_COA_mapping[new_code_col].astype(str)
account_mapping = dict(zip(df_COA_mapping[old_code_col], df_COA_mapping[new_code_col]))
df["Account No."] = df["Account No."].map(account_mapping)

# Clean amount fields
df["Amount ($)"] = df["Amount ($)"].astype(str).str.replace(",", "").astype(float)
df["Tax amount ($)"] = df["Tax amount ($)"].astype(str).str.replace(",", "").astype(float).abs()

# Ensure Unit Price is positive
original_sign = df["Amount ($)"].apply(lambda x: -1 if x < 0 else 1)
df["Amount ($)"] = df["Amount ($)"].abs()
df["Unit Price"] = (df["Amount ($)"] * original_sign.map({1: -1, -1: 1})).abs()

# Remove rows where Unit Price is NaN
df = df.dropna(subset=["Unit Price"])

# Default values
df["Amounts are"] = "Tax Exclusive"
df["No. of Unit"] = "-1"
df["Discount %"] = pd.NA

# ✅ Job Mapping using Class column
Job_mapping = {
    "Blinds, Awnings & Shutters": "11",
    "Builders": "12",
    "Flooring Hervey Bay": "13",
    "Patios": "14",
    "Sheds": "15",
    "Tiles Hervey Bay": "16"
}
df["Job name"] = df["Job name"]
df["Job No."] = df["Job name"].map(Job_mapping)

# Clean tax codes
Tax_Code_Mapping = {
    "CAG": "GST",
    "GST": "GST",
    "NCF": "FRE",
    "NCG": "GST",
    "NRG": "N-T",
    "": "N-T"
}
df["Tax code"] = df["Tax code"].map(Tax_Code_Mapping)
df["Tax code"] = df["Tax code"].replace(r'^\s*$', pd.NA, regex=True).fillna("N-T")

# Remove commas from Bill number and Supplier invoice number
df["Bill number"] = df["Bill number"].astype(str).str.replace(",", "")
df["Supplier invoice number"] = df["Supplier invoice number"].astype(str).str.replace(",", "")

# Final column order
columns_order = [
    "Bill number", "Supplier", "Transaction date", "Due date", "Supplier invoice number", "Amounts are", "Item", "Description", "Account No.", "No. of Unit", "Unit Price", "Discount %",
    "Amount ($)", "Tax code", "Tax amount ($)", "Job No.", "Job name"
]
final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export to Excel
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /BILL CREDIT NOTE/MYOB - BILL_CREDIT.xlsx"
df.to_excel(output_path, sheet_name="Bill_CREDIT", index=False)

print(df.head())
