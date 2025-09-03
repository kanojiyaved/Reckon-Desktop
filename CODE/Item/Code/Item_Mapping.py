import pandas as pd
import random
import os
df = pd.read_csv("/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/Item/Files/Item Reckon Desktop - Sheet1.csv")

field_mapping = {
    "Item": "Item",
    "Description": "Description",

}

df = df.rename(columns=field_mapping)

df['New Item'] = df['Item'].str[-30:]

df["Item ID"] = df["Item"].str[:20]

df["Type"] = "Service"

df["Expense Code"] = pd.NA

df["Income"] = pd.NA

df["Tax Code"] = "GST"

# Function to generate code in format D-DDDD
def generate_code():
    return f"{random.randint(0, 9)}-{random.randint(1000, 9999)}"

# Create the new column 'Expense' with the random codes
df['Expense'] = [generate_code() for _ in range(len(df))]



# Function to generate code in format D-DDDD
def generate_code():
    return f"{random.randint(0, 9)}-{random.randint(1000, 9999)}"

# Create the new column 'Expense' with the random codes
df['Income'] = [generate_code() for _ in range(len(df))]

columns_order = [
    "Item", "New Item", "Item ID", "Description", "Type", "Expense",
    "Income", "Tax Code"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]



# Export final file
output_path = r"/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/RECKONS FILES\Item-Mapping-MYOB.xlsx"
df.to_excel(output_path,sheet_name="item mapping", index=False)

print(df.head())
print("✅ Conversion Successful!")