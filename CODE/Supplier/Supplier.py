import pandas as pd
import random
import os 
import numpy as np 

df = pd.read_csv("/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/Supplier/File/Supplier Reckon Desktop - Sheet1.csv")
df_columns = list(df.columns)

field_mapping = {
    "Contact ID": "Contact ID",
    "Supplier": "Name",
    "Phone": "Phone",
    "Email": "Email",
    "Street1": "Street",
    "City": "City",
    "State": "State",
    "Post Code": "Postcode",
    "Country": "Country",
    "Tax ID": "ABN",
    "Fax": "Fax",
    "Contact": "Contact name"
}
df=df.rename(columns=field_mapping)

df["Type"] = "Supplier"
df["Balance ($)"] = pd.NA
df["Status"] = "Active"
df["Ship Street"] = pd.NA
df["Ship City"] = pd.NA
df["Ship State"] = pd.NA
df["Ship Postcode"] = pd.NA
df["Ship Country"] = pd.NA
df["Payment method"] = pd.NA
df["WWW"] = pd.NA
df["BSB number"] = pd.NA
df["Bank acct No."] = pd.NA
df["Bank acct name"] = pd.NA
df["Bank value"] = pd.NA
df["Credit card No."] = pd.NA
df["Card name"] = pd.NA
df["Memo"] = pd.NA
df["Notes"] = pd.NA
df["ABN"] = pd.NA


df['Contact ID'] = ['S-' + str(i) for i in range(1, len(df) + 1)]
df['Name'] = df['Name'].astype(str).str[-50:]

def clean_postcode(x):
        try:
            x_str = str(int(float(x)))  # Handles float or int postcodes like 1793127.0
        except:
            return x  # Leave non-numeric as-is

        # If it's all digits and longer than 4, treat as invalid
        if x_str.isdigit() and len(x_str) > 4:
            return np.nan
        else:
            return int(x_str)

 # Apply cleaner
df['Postcode'] = df['Postcode'].apply(clean_postcode)

# Convert to nullable integer (allows <NA>)
df['Postcode'] = pd.to_numeric(df['Postcode'], errors='coerce').astype('Int64')



columns_order = [
    "Contact ID", "Name", "Phone", "Type", "Email", "Balance ($)", "Status",
    "Street", "City", "State", "Postcode", "Country", "Ship Street", "Ship City",
    "Ship State", "Ship Postcode", "Ship Country", "ABN", "Payment method", "Fax",
    "WWW", "Contact name", "BSB number", "Bank acct No.", "Bank acct name",
    "Bank value", "Credit card No.", "Card name", "Memo", "Notes"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]


# Export final file
output_path = r"/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/RECKONS FILES\Supplier-MYOB.xlsx"
df.to_excel(output_path, index=False)

print(df.head())
print("✅ Conversion Successful!")