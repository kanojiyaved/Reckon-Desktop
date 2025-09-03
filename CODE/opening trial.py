import pandas as pd
import numpy as np

# Load data
df = pd.read_csv(
    "/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /OPEN TRIAL/Reckon Desktop - Open TB new.csv"
)
df_coa = pd.read_csv(
    "/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Open Aged Payable/COA - MAPPING NEW.xlsx - Sheet1.csv"
)

# Strip column names
df.columns = df.columns.str.strip()
df_coa.columns = df_coa.columns.str.strip()

# Rename fields if needed
field_mapping = {
    "Account": "Account",
    "Debit": "Debit Amount",
    "Credit": "Credit Amount"
}
df = df.rename(columns=field_mapping)

# Add required columns
df["Date"] = "30/06/2023"
df["Reference number"] = "Closing Bal."
df["Description of transaction"] = pd.NA
df["Quantity"] = pd.NA
df["Description"] = pd.NA
df["Job"] = pd.NA
df["Tax Code"] = "N-T"
df["Amounts are"] = "Tax Exclusive"

# Final column order
columns_order = [
    "Date", "Reference number", "Description of transaction", "Account",
    "Debit Amount", "Credit Amount", "Quantity", "Description", "Job",
    "Tax Code", "Amounts are"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export to Excel
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /OPEN TRIAL/output.xlsx"
df.to_excel(output_path, sheet_name="Opening Trial", index=False)
