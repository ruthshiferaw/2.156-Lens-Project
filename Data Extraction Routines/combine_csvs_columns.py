import os
import csv

# Define paths
base_folder = r"C:\Users\Ruth\Documents\GitHub\2.156-Lens-Project\Prime Lenses + Data\CSVExports"

# 0 = FieldCurvature, 1 = Longitudinal, 2 = RMSvField, 3 = Vignetting
name = 2

if name == 0:
    input_folder = os.path.join(base_folder, "FieldCurvature")
    output_combined = os.path.join(base_folder, "FieldCurvature.csv")
elif name == 1:
    input_folder = os.path.join(base_folder, "Longitudinal")
    output_combined = os.path.join(base_folder, "Longitudinal.csv")
elif name == 2:
    input_folder = os.path.join(base_folder, "RMSvField")
    output_combined = os.path.join(base_folder, "RMSvField.csv")
elif name == 3:
    input_folder = os.path.join(base_folder, "Vignetting")
    output_combined = os.path.join(base_folder, "Vignetting.csv")
else:
    raise ValueError("Invalid value for 'name' (must be 0–3).")

# Get all CSV files in the folder
csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".csv")]
csv_files.sort()  # optional: sort for consistent order

if not csv_files:
    raise FileNotFoundError(f"No CSV files found in {input_folder}")

# ---------------- Build global header set + collect rows ----------------
global_headers = []       # union of all column names (in first-seen order)
rows = []                 # list of dicts, one per data row

# # Optional: include which file each row came from
# add_source_col = True
# source_col_name = "SourceFile"
# if add_source_col:
#     global_headers.append(source_col_name)

for fname in csv_files:
    fpath = os.path.join(input_folder, fname)
    # print(f"Reading {fpath}")

    with open(fpath, "r", encoding="utf-8", newline="") as infile:
        reader = csv.reader(infile)

        try:
            # Read header row
            headers = next(reader)
        except StopIteration:
            # Empty file, skip
            print(f"  ⚠️ Skipping {fname}: empty file")
            continue

        # Strip whitespace from header names
        headers = [h.strip() for h in headers]

        # Update global header list with any new columns
        for h in headers:
            if h and h not in global_headers:
                global_headers.append(h)

        # Read data rows
        for row in reader:
            # Skip completely empty rows
            if not any(cell.strip() for cell in row):
                continue

            row_dict = {}

            # Map this file's columns into the row dict
            for h, val in zip(headers, row):
                h = h.strip()
                if h:  # ignore empty header names just in case
                    row_dict[h] = val

            # # Track source file if requested
            # if add_source_col:
            #     row_dict[source_col_name] = fname

            rows.append(row_dict)

# ---------------- Write combined CSV ----------------
print(f"\nWriting combined CSV to: {output_combined}")
with open(output_combined, "w", encoding="utf-8", newline="") as outfile:
    writer = csv.writer(outfile)

    # Write unified header row
    writer.writerow(global_headers)

    # Write each data row, aligning columns by header name
    for row_dict in rows:
        row_out = [row_dict.get(col, "") for col in global_headers]
        writer.writerow(row_out)

print(f"✅ Combined CSV saved to: {output_combined}")
print(f"   Total input files: {len(csv_files)}")
print(f"   Total rows: {len(rows)}")
print(f"   Total columns: {len(global_headers)}")
