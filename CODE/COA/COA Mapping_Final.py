import pandas as pd
import os

# Load CSV file
df = pd.read_csv("/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /Journal/Deepak sir COA-MAPPING SHEET.csv")

print("CSV Columns before renaming:", df.columns.tolist())  # 👈 Debug step

# Correct field mapping (adjust based on print output above)
field_mapping = {
    "Account": "Old Account Name",   # <-- update this to match actual column
    "Accnt. #": "Old Code",
    "Type": "Type",
    "Parent Code": "Parent Code",
    "New Code": "New Code"
}

# Rename columns
df = df.rename(columns=field_mapping)

# Now safe to access
df['Account Name'] = df['Old Account Name'].astype(str).str[-50:]
df['Parent Code'] = df['Parent Code'].str.replace(r'-\d+$', '-0000', regex=True)

columns_order = [
    "Old Account Name", "Account Name", "Type", "Parent Code", "Old Code", "New Code"
]
final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export to Excel
output_path = r"/Users/vedantkanojiya/Desktop/Vedant kanojiya/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /COA FILES/COA - MAPPING NEW.xlsx"
df.to_excel(output_path, index=False)

print(df.head())
print("✅ Conversion Successful!")
