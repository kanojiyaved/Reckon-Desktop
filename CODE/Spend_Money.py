import pandas as pd
import os
import numpy as np
import random

# Load data
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /ALL DATA/All Data - All Data.csv.csv")
df_COA_mapping = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Spend Money/COA - MAPPING NEW.xlsx - Sheet1-2.csv")

# Strip column names
df.columns = df.columns.str.strip()
df_COA_mapping.columns = df_COA_mapping.columns.str.strip()

# Filter Cheque rows
df = df[(df['Type'] == 'Cheque') & (df['Account Type'] != 'Accounts Receivable')]
df = df[~df['Account'].str.contains("Cheque", na=False)]

# Rename columns
field_mapping = {
    "Source Name": "Contact",
    "Trans #": "Reference number",
    "Date": "Date",
    "Amount": "Amount",
    "Description": "Description",
    "Tax Code": "Tax Code"
}
df = df.rename(columns=field_mapping)

# Add required columns
df["Description of transaction"] = pd.NA
df["Amounts are"] = "Tax Exclusive"
df["Quantity"] = pd.NA
df["Job"] = pd.NA

# ✅ Limit Description column to 1000 characters
if "Description" in df.columns:
    df["Description"] = df["Description"].astype(str).str.slice(0, 1000)

# ✅ Replace NaN/blank in Description with NULL
df["Description"] = df["Description"].replace("nan", pd.NA)
df["Description"] = df["Description"].replace(r'^\s*$', pd.NA, regex=True)

# ✅ Clean Account column (remove everything after '.' and keep only digits)
df["Account"] = df["Account"].astype(str).str.split(".").str[0]
df["Account"] = df["Account"].str.extract(r"(\d+)")

# ✅ Clean mapping file
old_code_col = next((col for col in df_COA_mapping.columns if "old" in col.lower() and "code" in col.lower()), None)
new_code_col = next((col for col in df_COA_mapping.columns if "new" in col.lower() and "code" in col.lower()), None)

df_COA_mapping[old_code_col] = df_COA_mapping[old_code_col].astype(str).str.extract(r"(\d+)")
df_COA_mapping[new_code_col] = df_COA_mapping[new_code_col].astype(str)

# ✅ Create mapping dictionary
account_mapping = dict(zip(df_COA_mapping[old_code_col], df_COA_mapping[new_code_col]))

# ✅ Map Account → new code in two places
# df["Bank Account"] = df["Account"].map(account_mapping)   # for Bank Account column
df["Account"] = df["Account"].map(account_mapping)        # keep mapped Account also

df["Bank Account"] = df["Split"].str.extract(r"^(\d+)")
df["Bank Account"] = df["Bank Account"].map(account_mapping)

# ✅ Clean tax codes
Tax_Code_Mapping = {
    "CAG": "GST",
    "GST": "GST",
    "NCF": "FRE",
    "NCG": "GST",
    "NRG": "N-T",
    "": "N-T"
}
df["Tax Code"] = df["Tax Code"].map(Tax_Code_Mapping)
df["Tax Code"] = df["Tax Code"].replace(r'^\s*$', pd.NA, regex=True).fillna("N-T")
# ✅ Remove rows where Tax Code is "N-T"
df = df[df["Tax Code"] != "N-T"]

# ✅ Remove rows where Amount is negative
df = df[pd.to_numeric(df["Amount"], errors="coerce") >= 0]

# ✅ Final column order
columns_order = [
    "Bank Account", "Contact", "Description of transaction", "Reference number", 
    "Date", "Amounts are", "Account", "Amount",
    "Quantity", "Description", "Job", "Tax Code"
]
final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export to Excel
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Spend Money/MYOB - SPEND MONEY.xlsx"
df.to_excel(output_path, sheet_name="Spend Money", index=False)

print(df.head())
