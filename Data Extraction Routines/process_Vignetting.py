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
root_dir = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses\AnalysisExports"
output_dir = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses\CSVExports\Vignetting"
os.makedirs(output_dir, exist_ok=True)

# ---------------- File Matching ----------------
file_pattern = "_Vignetting"  # Only process files ending with _Vignetting

# ---------------- Regex Patterns ----------------
header_pattern = re.compile(r'^\s*Y\s*Field', re.IGNORECASE)
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
            for line in lines[header_idx + 1:]:
                if not line.strip():
                    continue
                parts = [p for p in re.split(r'\t+|\s{2,}', line.strip()) if p.strip()]

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
            output_filename = os.path.splitext(filename)[0] + "_RI.csv"
            output_path = os.path.join(output_dir, output_filename)

            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(headers)
                writer.writerows(data)

            print(f"✅ Saved: {output_filename} ({len(data)} rows)")

print("\n Done Processing")
