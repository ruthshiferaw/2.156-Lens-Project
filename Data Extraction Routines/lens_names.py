'''
create csv of all lens names from Vignetting folder
saves in CSVExports as Vignetting_Lenses.csv
'''

import os
import pandas as pd

# Define paths
base_folder = r"Prime Lenses + Data\CSVExports"
input_folder = os.path.join(base_folder, "Vignetting")
output_lenses = os.path.join(base_folder, "Lenses_Vignetting.csv")

# Get all CSV files in the folder
csv_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]

# Create list of lens names (remove trailing part)
lens_names = [os.path.splitext(f)[0].replace("Vignetting", "").strip("_ -") for f in csv_files]

# Save lens names as a CSV
pd.DataFrame(lens_names, columns=["LensName"]).to_csv(output_lenses, index=False)

print(f"✅ Lenses list saved to: {output_lenses}")
