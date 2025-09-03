import pandas as pd
import os
import random
import numpy as np

# Load data
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Journal/Reckon Desktop - Sheet1.csv", low_memory=False)
df_coa = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Journal/COA - MAPPING NEW.xlsx - Sheet1-2.csv")

# Strip column names
df.columns = df.columns.str.strip()
df_coa.columns = df_coa.columns.str.strip()

field_mapping = {
    "Date": "Date",
    "Trans #": "Reference number",
    "Source Name": "Description of transaction",
    "Account": "Account",
    "Debit": "Debit Amount",
    "Credit": "Credit Amount",
    "Description": "Description",
    "Tax Code": "Tax Code"
}

df = df.rename(columns=field_mapping)

# Replace NaN in Debit and Credit with 0
df["Debit Amount"] = df["Debit Amount"].fillna(0)
df["Credit Amount"] = df["Credit Amount"].fillna(0)

df["Quantity"] = pd.NA
df["Job"] = pd.NA
df["Amounts are"] = "Tax Exclusive"

# ✅ Extract only numeric part from Account column
df["Account"] = df["Account"].astype(str).str.extract(r'(\d+)')[0]

# ✅ Map Account with df_coa (Old Code → New Code)
coa_map = dict(zip(df_coa["Old Code"].astype(str), df_coa["New Code"].astype(str)))
df["Account"] = df["Account"].map(coa_map).fillna(df["Account"])  # keep original if no match

# ✅ Fix Description column (make NaN → empty/null in Excel)
df["Description"] = df["Description"].astype(str).str[:1000]
df["Description"] = df["Description"].replace({"nan": None, "NaN": None})

# ✅ Remove rows where Date is null
df = df.dropna(subset=["Date"])

# Tax code mapping
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

columns_order = [
    "Date", "Reference number", "Description of transaction", "Account", "Debit Amount", "Credit Amount", "Quantity",
    "Description", "Job", "Tax Code", "Amounts are"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Journal/MYOB - JOURNAL.xlsx"
df.to_excel(output_path, sheet_name="Journal", index=False)

print(df.head())
print("✅ Conversion Successful! (Null Date rows removed)")
