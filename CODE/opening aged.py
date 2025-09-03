import pandas as pd
import numpy as np
import os
import random

# --- Load data ---
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Opening age recievable(invoice)/Reckon Desktop - Open AR.csv")
df_coa = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Journal/COA - MAPPING NEW.xlsx - Sheet1-2.csv")

# --- Clean column names ---
df.columns = df.columns.str.strip()
df_coa.columns = df_coa.columns.str.strip()

# --- Map Reckon fields → MYOB fields ---
field_mapping = {
    "Num": "Invoice number",
    "Source Name": "Customer",
    "Date": "Transaction date",
    "Due Date": "Due date",
    "Trans #": "Customer PO No.",
    "Open Balance": "Unit Price",   # keep Amount ($)
    # "Tax Code": "Tax code",
    # "Tax Amount": "Tax amount ($)",
    "Class": "Job name"
}
df = df.rename(columns=field_mapping)

# --- Add new MYOB fields ---
df["Amounts are"] = "Tax Exclusive"
df["Item"] = pd.NA
df["Description"] = pd.NA
df["Account No."] = "3-8000"
df["Discount %"] = pd.NA
df["Job No."] = pd.NA
df["Tax code"] = "N-T"
df["Tax amount ($)"] = pd.NA

# ✅ Adjust No. of Unit and make Unit Price positive
df["No. of Unit"] = np.where(df["Unit Price"] < 0, -1, 1)
df["Unit Price"] = df["Unit Price"].abs()

# Amount = Unit Price (since No. of Unit is used separately)
df["Amount ($)"] = df["Unit Price"]

# Default tax code
df["Tax code"] = "N-T"

columns_order = [
    "Invoice number", "Customer", "Transaction date", "Due date", "Customer PO No.", "Amounts are", "Item", "Description", "Account No.", "No. of Unit",
    "Unit Price", "Discount %", "Amount ($)", "Tax code", "Tax amount ($)", "Job No.", "Job name"
]

# ✅ Remove rows where Invoice number is null
df = df[df["Invoice number"].notna()]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# --- Export to Excel ---
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Opening age recievable(invoice)/output.xlsx"
df.to_excel(output_path, sheet_name="Opening aged", index=False)

