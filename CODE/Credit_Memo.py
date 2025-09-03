import pandas as pd
import os
import random
import numpy as np

# Load data
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Credit Memo/Reckon Desktop Adjustment Note - Sheet1.csv", low_memory=False)
df_COA_mapping = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/CODE/COA/Coa Mapping - Sheet1.csv")
df_Job = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /JOB File /Reckon Desktop Job - Sheet1.csv")

# Strip column names
df.columns = df.columns.str.strip()
df_COA_mapping.columns = df_COA_mapping.columns.str.strip()

# ✅ Print columns to confirm
print("COA Mapping Columns:", df_COA_mapping.columns.tolist())

# ✅ Case-insensitive match for Old/New Code
old_code_col = next((col for col in df_COA_mapping.columns if "old" in col.lower() and "code" in col.lower()), None)
new_code_col = next((col for col in df_COA_mapping.columns if "new" in col.lower() and "code" in col.lower()), None)

if old_code_col is None or new_code_col is None:
    raise ValueError("❌ Could not find 'Old Code' or 'New Code' column in COA mapping file.")

# Clean df
df = df[df['Account Type'] != 'Accounts Receivable']
df = df[~df['Account'].str.contains("Tax Payable", na=False)]

field_mapping = {
    "Num": "Invoice number",
    "Name": "Customer",
    "Date": "Transaction date",
    "Due Date": "Due date",
    "Item": "Item",
    "Description": "Description",
    "Account": "Account No.",
    "Debit": "Unit Price",
    "Amount": "Amount ($)",
    "Tax Code": "Tax code",
    "Tax Amount": "Tax amount ($)",
    "Class": "Job name"
}

df = df.rename(columns=field_mapping)

df["Amounts are"] = "Tax exclusive"
df["No. of Unit"] = "-1"
df["Customer PO No."] = df["Invoice number"]

# Extract digits from Account No.
df["Account No."] = df["Account No."].astype(str).str.extract(r"(\d+)")

# Extract digits from COA mapping
df_COA_mapping[old_code_col] = df_COA_mapping[old_code_col].astype(str).str.extract(r"(\d+)")
df_COA_mapping[new_code_col] = df_COA_mapping[new_code_col].astype(str)

# Perform mapping
account_mapping = dict(zip(df_COA_mapping[old_code_col], df_COA_mapping[new_code_col]))
df["Account No."] = df["Account No."].map(account_mapping)

# Clean amount
df["Amount ($)"] = df["Amount ($)"].astype(str).str.replace(",", "").astype(float)

df["Tax amount ($)"] = df["Tax amount ($)"].astype(str).str.replace(",", "").astype(float).abs()

original_sign = df["Amount ($)"].apply(lambda x: -1 if x < 0 else 1)
df["Amount ($)"] = df["Amount ($)"].abs()

df["Discount %"] = pd.NA
df["Job No."] = pd.NA




Job_mapping = {
    "Blinds, Awnings & Shutters": "11",
    "Builders": "12",
    "Flooring Hervey Bay": "13",
    "Patios": "14",
    "Sheds": "15",
    "Tiles Hervey Bay": "16"
}
df["Job No."] = df["Job name"].map(Job_mapping)

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


columns_order = [
    "Invoice number", "Customer", "Transaction date", "Due date", "Customer PO No.", "Amounts are", "Item", "Description",
    "Account No.", "No. of Unit", "Unit Price", "Discount %", "Amount ($)", "Tax code", "Tax amount ($)", "Job No.", "Job name"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Credit Memo/MYOB - CREDIT_MEMO.xlsx"
df.to_excel(output_path, sheet_name="Invoice", index=False)

print(df.head())
print("✅ Conversion Successful!")
