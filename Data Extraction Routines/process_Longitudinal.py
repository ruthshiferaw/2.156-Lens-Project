import os
import re
import csv

# ---------------- UTF Auto-detect Helper ----------------
def read_text_auto(filepath):
    """Read a text file, trying UTF-8 and UTF-16 automatically."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.readlines()
    except UnicodeDecodeError:
        with open(filepath, 'r', encoding='utf-16') as f:
            return f.readlines()

# ---------------- Paths ----------------
root_dir = r"Prime Lenses + Data\AnalysisExports"
output_dir = r"Prime Lenses + Data\CSVExports\Longitudinal"
os.makedirs(output_dir, exist_ok=True)

# ---------------- File Matching ----------------
file_pattern = "_Longitudinal"  # Only process files ending with _Longitudinal

# ---------------- Regex Patterns ----------------
header_pattern = re.compile(r'^\s*Rel\.?\s*Pupil', re.IGNORECASE)  # Header line
numeric_line = re.compile(r'^[+-]?\d')  # Lines starting with a number

# ---------------- Main Processing Loop ----------------
for subdir, _, files in os.walk(root_dir):
    for filename in files:
        if file_pattern.lower() in filename.lower() and filename.lower().endswith('.txt'):
            filepath = os.path.join(subdir, filename)
            print(f"Processing: {filepath}")

            # Read lines with auto-detect encoding
            try:
                lines = [line.rstrip("\n") for line in read_text_auto(filepath)]
            except Exception as e:
                print(f"⚠️ Skipping {filename}: could not read ({e})")
                continue

            # --- Find header line ---
            header_idx = None
            for i, line in enumerate(lines):
                if header_pattern.search(line):
                    header_idx = i
                    break

            if header_idx is None:
                print(f"⚠️ Skipping {filename}: no header line found.")
                continue

            headers = [h.strip() for h in re.split(r'\t+|\s{2,}', lines[header_idx].strip()) if h.strip()]
            data = []

            # --- Extract data lines ---
            number_pattern = re.compile(r'[+-]?\d*\.?\d+(?:[Ee][+-]?\d+)?')

            for line in lines[header_idx + 1:]:
                if not line.strip():
                    continue

                # Extract numbers (handles exponents correctly)
                parts = number_pattern.findall(line)

                if not parts or not numeric_line.match(parts[0]):
                    if data:
                        break  # stop after data ends
                    else:
                        continue

                # Normalize to header length
                if len(parts) < len(headers):
                    parts += [''] * (len(headers) - len(parts))
                elif len(parts) > len(headers):
                    parts = parts[:len(headers)]

                data.append(parts)

            if not data:
                print(f"⚠️ Skipping {filename}: no data rows found.")
                continue

            # --- Save CSV ---
            output_filename = os.path.splitext(filename)[0] + ".csv"
            output_path = os.path.join(output_dir, output_filename)

            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(headers)
                writer.writerows(data)

            print(f"✅ Saved: {output_filename} ({len(data)} rows)")

print("\n Done Processing")
