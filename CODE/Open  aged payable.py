import pandas as pd
import os
import random
import numpy as np

# Load data
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Open Aged Payable/Reckon Desktop - Open AP.csv")
df_coa = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Open Aged Payable/COA - MAPPING NEW.xlsx - Sheet1.csv")

# Strip column names
df.columns = df.columns.str.strip()
df_coa.columns = df_coa.columns.str.strip()

field_mapping = {
    "Trans #": "Bill number",
    "Source Name": "Supplier",
    "Date": "Transaction date",
    "Due Date": "Due date",
    "Description": "Description",
    "Account": "Account No.",
    "Open Balance": "Unit Price"
    # "Tax Code": "Tax Code",
    # "Tax Amount": "Tax amount ($)"
    # "Class": "Job name"
}

df = df.rename(columns=field_mapping)

df["Amounts are"] = "Tax Exclusive"
df["Item"] = pd.NA
df["Description"] = pd.NA
df["Account No."] = "3-8000"
df["Discount %"] = pd.NA
df["Job No."] = pd.NA
df["Supplier invoice number"] = df["Bill number"]
df["Job name"] = pd.NA
df["Tax amount ($)"] = pd.NA   # ✅ fixed column name

# ✅ Adjust No. of Unit and make Unit Price positive
df["No. of Unit"] = np.where(df["Unit Price"] < 0, -1, 1)
df["Unit Price"] = df["Unit Price"].abs()

# Amount = Unit Price (since No. of Unit is used separately)
df["Amount ($)"] = df["Unit Price"]

# Default tax code
df["Tax code"] = "N-T"

# ✅ Remove rows where Bill number is null
df = df[df["Bill number"].notna()]

columns_order = [
    "Bill number", "Supplier", "Transaction date", "Due date", "Supplier invoice number", "Amounts are", "Item", "Description", "Account No.",
    "No. of Unit", "Unit Price", "Discount %", "Amount ($)", "Tax code", "Tax amount ($)", "Job No.", "Job name"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# --- Export to Excel ---
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Open Aged Payable/MYOB - PAYABLE.xlsx"
df.to_excel(output_path, sheet_name="Opening aged", index=False)
