import pandas as pd
import os

# Define file path
excel_path = r"\\srvnetapp02.phsabc.ehcnet.ca\bcca\docs\CCSI\PlanningModule\Physics Patient QA\VitesseBackup\Vitesse DICOM Archive log.xlsx"

# Define the output CSV path
output_csv = os.path.join(os.path.dirname(excel_path), "patients_misnamed_BCCIDs.csv")

# Define specifier list
specifier_list = ["No", "No "]  # Include both variations

# Read the Excel file, ensuring Column C is treated as a string
df = pd.read_excel(excel_path, engine='openpyxl', dtype={"ID": str})  # Read column C as a string

# Filter rows where column F contains exactly "No"
filtered_df = df[df.iloc[:, 5].isin(specifier_list)]  # Column F is index 5 (0-based)

# Extract values from column C (index 2)
result_df = filtered_df.iloc[:, [2]].copy()  # Selecting only column C

# Rename the column for clarity
result_df.columns = ["BCCID"]

# Ensure leading zeros are retained
result_df["BCCID"] = result_df["BCCID"].astype(str)

# Save to CSV, ensuring leading zeros are not lost
result_df.to_csv(output_csv, index=False, quoting=1)  # quoting=1 ensures proper string handling

print(f"Filtered patient IDs saved to: {output_csv}")
