'''
create csv of all lens names from Longitudinal folder
saves in CSVExports as Lenses.csv
'''

import os
import pandas as pd

# Define paths
base_folder = r"C:\Users\Ruth\Downloads\Lens Project\Prime Lenses + Data\CSVExports"
input_folder = os.path.join(base_folder, "Longitudinal")
output_lenses = os.path.join(base_folder, "Lenses.csv")

# Create list of lens names (remove trailing "Longitudinal_LA" part)
lens_names = [os.path.splitext(f)[0].replace("Longitudinal_LA", "").strip("_ -") for f in csv_files]

# Save lens names as a CSV
pd.DataFrame(lens_names, columns=["LensName"]).to_csv(output_lenses, index=False)

print(f"✅ Lenses list saved to: {output_lenses}")
