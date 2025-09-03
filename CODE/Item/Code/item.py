import pandas as pd
import random
import os
df = pd.read_csv("/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/Item/Files/Item Reckon Desktop - Sheet1.csv")

field_mapping = {
    "Item": "Name",
    "Description": "Description",
    # Item ID add
    "Price": "Selling Price",
    "Account": "Income account for tracking sales",
    "Tax Code": "Tax code",
    "Cost": "Buying price",
    "COGS Account": "Expense account for tracking purchases"
    
}
df = df.rename(columns=field_mapping)



# Ensure 'Name' column is treated as string
df['Name'] = df['Name'].astype(str)

# Create 'Item ID' by slicing first 20 characters (spaces allowed)
df['Item ID'] = df['Name'].str[:20]

df["Primary supplier for reorders"] = pd.NA
df["Default reorder quantity (per buying unit)"]  = pd.NA



# Ensure 'Tax code' is treated as string
df['Tax code'] = df['Tax code'].astype(str).str.strip()

# Replace 'null', 'nan', or empty values with NaN first
df['Tax code'].replace(['', 'null', 'nan', 'None', "NCF"], pd.NA, inplace=True)

# Now use fillna for 'N-T', others set to 'GST'
df['Tax code'] = df['Tax code'].fillna('N-T').apply(lambda x: 'GST' if x != 'N-T' else x)



# Ensure the column is string type
df['Income account for tracking sales'] = df['Income account for tracking sales'].astype(str)

# Extract first 4 characters
first_four = df['Income account for tracking sales'].str[:4]

# Generate random digit (1–9) for each row
random_digit = [str(random.randint(1, 9)) for _ in range(len(df))]

# Create the custom code in the format X-XXXX
df['Income account for tracking sales'] = [f"{r}-{f}" for r, f in zip(random_digit, first_four)]

# Optional: Preview
print(df['Income account for tracking sales'].head())



# Ensure the column is string type
df['Expense account for tracking purchases'] = df['Expense account for tracking purchases'].astype(str)

def generate_code(value):
    if pd.isna(value) or value.strip().lower() in ['nan', 'none', '']:
        # If value is null or blank → random X-XXXX
        return f"{random.randint(1,9)}-{random.randint(1000, 9999)}"
    else:
        # If value is not null → X-<first 4 chars>
        return f"{random.randint(1,9)}-{value[:4]}"

# Apply function to column
df['Expense account for tracking purchases'] = df['Expense account for tracking purchases'].apply(generate_code)



columns_order = [
    "Name", "Description", "Item ID", "Selling Price", "Income account for tracking sales", "Tax code",
    "Buying price", "Expense account for tracking purchases", "Tax code", 
    "Primary supplier for reorders", "Default reorder quantity (per buying unit)"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]



# Export final file
output_path = r"/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/CODE/RECKONS FILES\Item-MYOB.xlsx"
df.to_excel(output_path, index=False)

print(df.head())
print("✅ Conversion Successful!")

