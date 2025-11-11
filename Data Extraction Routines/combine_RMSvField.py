import os

# Define paths
base_folder = r"C:\Users\Ruth\Downloads\Lens Project\Prime Lenses + Data\CSVExports"
input_folder = os.path.join(base_folder, "RMSvField")
output_combined = os.path.join(base_folder, "RMSvField.csv")

# Get all CSV files in the folder
csv_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]

# Combine them cleanly by text
with open(output_combined, "w", encoding="utf-8", newline="") as outfile:
    for i, fname in enumerate(csv_files):
        fpath = os.path.join(input_folder, fname)
        with open(fpath, "r", encoding="utf-8") as infile:
            lines = infile.readlines()
            if i == 0:
                outfile.writelines(lines)  # keep header for first file
            else:
                outfile.writelines(lines[1:])  # skip header for others

print(f"✅ Combined CSV saved to: {output_combined}")
