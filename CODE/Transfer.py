import pandas as pd
import os

# Load main data
df = pd.read_csv(
    r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Transfer/All Data - All Data.csv.csv",
    thousands=",",        
    low_memory=False
)

# ✅ Corrected COA mapping path
df_coa = pd.read_csv(
    r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/CODE/COA/Coa Mapping - Sheet1.csv"
)

# Strip column names
df.columns = df.columns.str.strip()
df_coa.columns = df_coa.columns.str.strip()

# Extract digits from 'Account' and 'Split'
df['Account'] = df['Account'].str.extract(r'(\d+)')
df['Split'] = df['Split'].str.extract(r'(\d+)')

# Convert Amount to numeric
df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')

# Filter only: Type = 'Transfer', Account Type = 'Bank', and positive Amount
df = df[(df['Type'] == 'Transfer') & (df['Account Type'] == 'Bank') & (df['Amount'] > 0)]

# Rename columns
field_mapping = {
    "Date": "Date",
    "Trans #": "Reference number",
    "Amount": "Amount",
    "Description": "Description of transaction",
    "Account": "Bank account from",
    "Split": "Bank account to"
}
df = df.rename(columns=field_mapping)

# Set uniform description
df["Description of transaction"] = "Funds Transfer"

# ---- Map Bank Accounts using df_coa ----
df["Bank account from"] = df["Bank account from"].astype(str)
df["Bank account to"] = df["Bank account to"].astype(str)

def get_code(code):
    row = df_coa[df_coa["Old code"].astype(str) == str(code)]
    if not row.empty:
        return row["New Code"].values[0]
    return code  # fallback if no match

df["Bank account from"] = df["Bank account from"].apply(get_code)
df["Bank account to"] = df["Bank account to"].apply(get_code)





# Reorder and select final columns
columns_order = [
    "Date", "Reference number", "Amount", "Description of transaction", "Bank account from", "Bank account to"
]
final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export to Excel
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Transfer/MYOB _ Transfer.xlsx"
df.to_excel(output_path, sheet_name="Transfer", index=False)
print("✅ Successfully exported with COA mapping applied")
