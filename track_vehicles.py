import pandas as pd

# File paths
excel_file = "Regular Track of Intellicar Vehicles.xlsx"
csv_file = "ViewReport (74).csv"
output_file = "output.xlsx"

# Read Excel file (header starts from row 3)
df_excel = pd.read_excel(excel_file, header=2)

# Read CSV file
df_csv = pd.read_csv(csv_file)

# Standardize column names (strip spaces & convert to lowercase)
df_excel.columns = df_excel.columns.str.strip().str.lower()
df_csv.columns = df_csv.columns.str.strip().str.lower()

# Keep only the required columns from Excel
required_excel_columns = ["vehicle group", "vehicle vin", "device id", "vehicle model", "business type", "owner"]
df_excel = df_excel[required_excel_columns]

# Keep only required columns from CSV and rename them
required_csv_columns = ["device id", "status", "gps time", "can time"]
df_csv = df_csv[required_csv_columns].rename(columns={
    "gps time": "latest gps time",
    "can time": "latest can time"
})

# Trim spaces and lowercase 'device id' for proper merging
df_excel["device id"] = df_excel["device id"].astype(str).str.strip().str.lower()
df_csv["device id"] = df_csv["device id"].astype(str).str.strip().str.lower()

# Print unique 'device id' values to check if they match
print("Excel Device IDs:", df_excel["device id"].unique()[:5])  # Print first 5 unique values
print("CSV Device IDs:", df_csv["device id"].unique()[:5])  # Print first 5 unique values

# Merge dataframes on 'device id'
merged_df = pd.merge(df_excel, df_csv, on="device id", how="inner")

# Check if merge returned data
if merged_df.empty:
    print("❌ No matching 'device id' found between Excel and CSV. Check for typos or extra spaces!")

# Save to new Excel file
merged_df.to_excel(output_file, index=False)

print(f"✅ Processed file saved as {output_file} with {len(merged_df)} rows.")

#This is a test version still incomplete