import pandas as pd
import random
import os

# Load CSV file
df = pd.read_csv("/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/Coa Mapping - Sheet1.csv")
df_columns = list(df.columns)

# Field mapping from source to target
field_mapping = {
    "Accnt. #": "Account Number",
    "Account": "Account Name",
    "Type": "Account Type",
    "Header": "Header",
    "Parent Code": "Parent Account Number",
    "Class": "Classification",
}

# Rename columns
df = df.rename(columns=field_mapping)
df["Header"] = pd.NA
df["Tax Code"] = pd.NA
df["Account Type"] = df["Account Type"].str.title()

# Type mapping (Reckon → MYOB)
type_mapping = {
    "Accounts Payable": "AccountsPayable",
    "Accounts Receivable": "AccountReceivable",
    "Bank": "Bank",
    "Cost Of Goods Sold": "CostofSales",
    "Credit Card": "Creditcard",
    "Equity": "Equity",
    "Income": "Income",
    "Other Current Asset": "OtherCurrentAsset",
    "Other Asset": "OtherCurrentAsset",
    "Other Current Liability": "OtherCurrentLiability",
    "Fixed Asset": "FixedAsset",
    "Expense": "Expense",
    "Long Term Liability": "LongTermLiability",
    "Suspense": "LongTermLiability",
    "Non-Posting": "OtherCurrentLiability",
    "Other Income": "OtherIncome",
    "Other Expense": "OtherExpense"
}

# Apply Account Type mapping
df["Account Type"] = df["Account Type"].map(type_mapping).fillna(df["Account Type"])

# Tax Code Mapping
Tax_Code_Mapping = {
    "CAG": "GST",
    "GST": "GST",
    "NCF": "FRE",
    "NCG": "GST",
    "NRG": "N-T",
    "": "N-T"
}
# df["Tax Code"] = df["Tax Code"].replace(pd.NA, "N-T")
# df["Tax Code"] = df["Tax Code"].map(Tax_Code_Mapping)

df['Tax Code'] = df['Tax Code'].replace(r'^\s*$', pd.NA, regex=True)

# Replace NA (including None) with 'N-T'
df['Tax Code'] = df['Tax Code'].fillna('N-T')

# Parent Account Number based on Account Type
Parent_Account_Number = {
    "AccountsPayable": None,
    "AccountReceivable": None,
    "Bank": "1-0000",
    "CostofSales": "5-0000",
    "Creditcard": "2-0000",
    "Equity": "3-0000",
    "Income": "4-0000",
    "OtherCurrentAsset": "1-0000",
    "OtherCurrentLiability": "2-0000",
    "FixedAsset": "1-0000",
    "Expense": "6-0000",
    "LongTermLiability": "2-0000",
    "OtherIncome": "8-0000",
    "OtherExpense": "9-0000"
}

df["Parent Account Number"] = df["Account Type"].map(Parent_Account_Number)
df["Classification"] = df["Parent Account Number"].str.extract(r'^(\d+)')

# Clean Account Number column
df["Account Number"] = df["Account Number"].replace(['', '-'], pd.NA)
col = "Account Number"

# Clean column (treat blank strings as missing)
df[col] = df[col].replace('', pd.NA)
# Replace '-' with NaN first
df[col] = df[col].replace('-', pd.NA)

# For existing account numbers, ensure they match the format Classification-XXXX
def standardize_account_number(row):
    account_num = row[col]
    classification = row['Classification']
    
    if pd.notna(account_num):
        # If account number exists, check if it follows the correct format
        if '-' in str(account_num):
            prefix, suffix = str(account_num).split('-', 1)
            # If prefix doesn't match classification, update it
            if prefix != str(classification):
                return f"{classification}-{suffix}"
            else:
                return account_num
        else:
            # If no dash, assume it's just the numeric part
            return f"{classification}-{account_num}"
    return account_num

# Apply standardization to existing account numbers
df[col] = df.apply(standardize_account_number, axis=1)

# Get existing 4-digit codes (extract numeric part after dash)
existing_codes = set()
for acc_num in df[col].dropna():
    if '-' in str(acc_num):
        numeric_part = str(acc_num).split('-')[1]
        try:
            existing_codes.add(int(numeric_part))
        except ValueError:
            continue

# Generate random unique 4-digit numbers for missing account numbers
needed = df[col].isna().sum()
available = list(set(range(1000, 10000)) - existing_codes)

if needed > len(available):
    raise ValueError("Not enough unique 4-digit codes available.")

new_codes = random.sample(available, needed)

# Fill missing account numbers with formatted classification + random code
na_indices = df[df[col].isna()].index
for i, idx in enumerate(na_indices):
    class_prefix = df.at[idx, 'Classification']
    df.at[idx, col] = f"{class_prefix}-{new_codes[i]}"

# column order 
columns_order=[
    "Account Number","Account Name","Account Type","Header","Parent Account Number","Tax Code","Classification"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns] 

# Final column order
columns_order = [
    "Account Number", "Account Name", "Account Type", "Header",
    "Parent Account Number", "Tax Code", "Classification"
]
final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export final file
output_path = r"/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/RECKONS FILES\Final COA to MYOB9.xlsx"
df.to_excel(output_path,sheet_name="COA", index=False)

print(df.head())
print("✅ Conversion Successful!")