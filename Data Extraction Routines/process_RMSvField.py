# import os
# import re
# import pandas as pd

# # Root folder where all AnalysisExports subfolders are stored
# root_dir = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses\AnalysisExports"

# # Output directory for CSVs (optional: can be same as input)
# output_dir = os.path.join(root_dir, "CSV_Exports")
# os.makedirs(output_dir, exist_ok=True)

# # Pattern to find header line (the line that has column names)
# header_pattern = re.compile(r'Field\s+Poly', re.IGNORECASE)

# # Loop through all subfolders and files
# for subdir, _, files in os.walk(root_dir):
#     for filename in files:
#         if "RMSvField" in filename and filename.endswith(".txt"):
#             filepath = os.path.join(subdir, filename)
#             print(f"Processing {filepath}")

#             # Read lines
#             with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
#                 lines = f.readlines()

#             # Find where the table starts
#             header_idx = None
#             for i, line in enumerate(lines):
#                 if header_pattern.search(line):
#                     header_idx = i
#                     break

#             if header_idx is None:
#                 print(f"⚠️ No header found in {filename}")
#                 continue

#             # Next line after header line is dashed separator
#             data_start = header_idx + 2

#             # Extract headers
#             headers = re.split(r'\s{2,}', lines[header_idx].strip())
#             headers = [h for h in headers if h]  # remove empties

#             # Extract data until empty line or non-numeric start
#             data = []
#             for line in lines[data_start:]:
#                 if not line.strip():
#                     break
#                 parts = re.split(r'\s{2,}', line.strip())
#                 # If first item isn't numeric, stop
#                 if not re.match(r'^[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?$', parts[0]):
#                     break
#                 data.append(parts)

#             # Convert to DataFrame
#             df = pd.DataFrame(data, columns=headers)

#             # Convert numeric columns
#             df = df.apply(pd.to_numeric, errors='ignore')

#             # Save as CSV
#             outname = os.path.splitext(filename)[0] + ".csv"
#             outpath = os.path.join(output_dir, outname)
#             df.to_csv(outpath, index=False)
#             print(f"✅ Saved CSV: {outpath}")

# print("All done!")


import os
import re
import pandas as pd
import codecs

def read_text_auto(path):
    """Read a text file, trying UTF-8 and UTF-16 automatically."""
    # Try UTF-8 first
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.readlines()
    except UnicodeDecodeError:
        # Fall back to UTF-16
        with open(path, "r", encoding="utf-16") as f:
            return f.readlines()

# Root folder containing the AnalysisExports folders
root_dir = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses\AnalysisExports"

# Where to save CSVs
output_dir = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses\CSVExports\RMSvField"
os.makedirs(output_dir, exist_ok=True)

# File matching rule
file_pattern = "RMSvField"

# A few regex helpers
header_regex = re.compile(r'\bField\b.*\bPoly\b', re.IGNORECASE)
numeric_line = re.compile(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?')

for subdir, _, files in os.walk(root_dir):
    for filename in files:
        if file_pattern.lower() in filename.lower() and filename.endswith(".txt"):
            filepath = os.path.join(subdir, filename)
            print(f"Processing: {filepath}")

            # Read all lines (auto-detect UTF-8 vs UTF-16)
            try:
                lines = [line.rstrip("\n") for line in read_text_auto(filepath)]
            except Exception as e:
                print(f"⚠️ Skipping {filename}: error reading file ({e})")
                continue

            # --- Find header line ---
            header_idx = None
            for i, line in enumerate(lines):
                if header_regex.search(line):
                    header_idx = i
                    break

            if header_idx is None:
                print(f"⚠️ Skipping {filename}: couldn't find header line.")
                continue

            # Parse header (split by tabs OR multiple spaces)
            header_line = lines[header_idx].strip()
            headers = re.split(r'\t+|\s{2,}', header_line)
            headers = [h.strip() for h in headers if h.strip()]

            # --- Find where numeric data starts ---
            data = []
            for line in lines[header_idx + 1:]:
                if not line.strip():
                    continue
                # Split and clean
                parts = [p for p in re.split(r'\t+|\s{2,}', line.strip()) if p.strip()]

                # Only accept rows that start with a number
                if parts and numeric_line.match(parts[0]):
                    # Ensure same number of columns as header
                    if len(parts) < len(headers):
                        # Pad with empty strings if missing
                        parts += [''] * (len(headers) - len(parts))
                    elif len(parts) > len(headers):
                        # Trim extra columns if any
                        parts = parts[:len(headers)]
                    data.append(parts)
                elif data:
                    break  # stop once data section ends

            if not data:
                print(f"⚠️ No data found in {filename}")
                continue

            # --- Build DataFrame ---
            df = pd.DataFrame(data, columns=headers)
            # Convert to numeric where possible
            df = df.apply(pd.to_numeric, errors='ignore')

            # --- Save CSV ---
            csv_name = os.path.splitext(filename)[0] + ".csv"
            outpath = os.path.join(output_dir, csv_name)
            df.to_csv(outpath, index=False)
            print(f"✅ Saved: {outpath}")

print("All done!")
