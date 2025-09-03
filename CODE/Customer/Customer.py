import pandas as pd
import random
import os 

# Loading CSV FILE

df = pd.read_csv("/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/Customer/FILES/Reckon Customer - Sheet1.csv")
df_columns = list(df.columns)

# Field mapping from the source to target 
field_mapping = {
    "Contact ID": "Contact ID",
    "Customer": "Name",
    "Phone": "Phone",
    # Type column bhi lagana hai usme default value Customer hai 
    "Email": "Email",
    # Balance ($)
    # Status column bhi add karna hai default value is Active
    "Ship To Street1": "Street",
    "City": "City",
    "State": "State",
    "Post Code": "Postcode",
    "Country": "Country",
    "Ship To Street2": "Ship Street",
    "Ship To City": "Ship City",
    "Ship To State": "Ship State",
    "Ship To Post Code": "Ship Postcode",
    "Ship To Country": "Ship Country",
    "Tax Code": "ABN",
    # payment method add karna hai
    "Fax": "Fax",
    # WWW add karna hai
    "Contact": "Contact name"
}
df=df.rename(columns=field_mapping)

df["Type"] = "Customer"
df["Balance ($)"] = pd.NA
df["Status"] = "Active"
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

df['Contact ID'] = ['C-' + str(i) for i in range(1, len(df) + 1)]
df['Name'] = df['Name'].astype(str).str[-50:]


columns_order = [
    "Contact ID", "Name", "Phone", "Type", "Email", "Balance ($)", "Status", "Street", "City", "State", "Postcode", "Country",
    "Ship Street", "Ship City", "Ship State", "Ship Postcode", "Ship Country", "ABN", "Payment method", "Fax", "WWW",
    "Contact name", "BSB number", "Bank acct No.", "Bank acct name", "Bank value", "Credit card No.", "Card name",
    "Memo", "Notes"
]


final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export final file
output_path = r"/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/RECKONS FILES\Customer-MYOB.xlsx"
df.to_excel(output_path, index=False)

print(df.head())
print("✅ Conversion Successful!")