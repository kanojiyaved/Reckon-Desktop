import pandas as pd 
import random
import os 

df = pd.read_csv("RECKONS FILES /JOB File /Reckon Desktop Job - Sheet1.csv")

field_mapping = {
    "Class": "Job name"
}

df = df.rename(columns=field_mapping)

df["Job number"] = pd.NA

# Generate random 2-digit codes (00 to 99)
df['Job number'] = [f"{random.randint(0, 99):02}" for _ in range(len(df))]

columns_order = [
    "Job number", "Job name"
]

final_columns = [col for col in columns_order if col in df.columns]
df = df[final_columns]

# Export final file
output_path = r"/Users/vedantkanojiya/Desktop/MMC INTERNSHIP/RECKON DESKTOP TO MYOB/RECKONS FILES /JOB File /DESKTOP - MYOB - JOB.xlsx"
df.to_excel(output_path, sheet_name="JOB", index=False)

print(df.head())
print("✅ Conversion Successful!")
