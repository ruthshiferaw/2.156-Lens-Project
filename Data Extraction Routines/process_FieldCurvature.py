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
output_dir = r"Prime Lenses + Data\CSVExports\FieldCurvature"
os.makedirs(output_dir, exist_ok=True)

# ---------------- File Matching ----------------
file_pattern = "_FieldCurvature"  # File suffix to match

# ---------------- Regex Patterns ----------------
table_start_pattern = re.compile(
    r"^(?:Table series:\s*)?Data for wavelength\s*:\s*(.+)$",
    re.IGNORECASE
)
header_pattern = re.compile(r'^\s*Y\s*(Angle|Height)', re.IGNORECASE)
numeric_line = re.compile(r'^[+-]?\d')  # Lines starting with a number

# ---------------- Main Processing Loop ----------------
for subdir, _, files in os.walk(root_dir):
    for filename in files:
        if file_pattern.lower() in filename.lower() and filename.lower().endswith('.txt'):
            filepath = os.path.join(subdir, filename)
            print(f"\nProcessing: {filepath}")

            try:
                lines = [line.rstrip("\n") for line in read_text_auto(filepath)]
            except Exception as e:
                print(f"⚠️ Skipping {filename}: could not read ({e})")
                continue

            current_wavelength = None
            headers = []
            all_data = []

            for i, line in enumerate(lines):
                # Detect new table start
                table_match = table_start_pattern.search(line)
                if table_match:
                    current_wavelength = table_match.group(1).strip()
                    continue

                # Detect header line
                if header_pattern.search(line):
                    headers = [h.strip() for h in re.split(r'\t+|\s{2,}', line.strip()) if h.strip()]
                    # Add wavelength column if not present
                    if 'Wavelength' not in headers:
                        headers.append('Wavelength')
                    continue

                # Parse numeric data lines
                if headers and numeric_line.match(line.strip()):
                    # Regex to capture all numeric values (handles exponentials)
                    number_pattern = re.compile(r'[+-]?\d*\.?\d+(?:[Ee][+-]?\d+)?')

                    # Parse numeric data lines
                    if headers and numeric_line.match(line.strip()):
                        parts = number_pattern.findall(line)
                        
                        # Pad or trim to match headers (excluding wavelength)
                        if len(parts) < len(headers) - 1:
                            parts += [''] * (len(headers) - 1 - len(parts))
                        elif len(parts) > len(headers) - 1:
                            parts = parts[:len(headers) - 1]
                        
                        parts.append(current_wavelength or '')
                        all_data.append(parts)
                    # Pad or trim
                    if len(parts) < len(headers) - 1:
                        parts += [''] * (len(headers) - 1 - len(parts))
                    elif len(parts) > len(headers) - 1:
                        parts = parts[:len(headers) - 1]
                    parts.append(current_wavelength or '')
                    all_data.append(parts)

            # Write all combined data for this file
            if all_data:
                output_filename = os.path.splitext(filename)[0] + ".csv"
                output_path = os.path.join(output_dir, output_filename)
                with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(headers)
                    writer.writerows(all_data)
                print(f"✅ Saved: {output_filename} ({len(all_data)} rows)")
            else:
                print(f"⚠️ No data found in {filename}")

print("\n✅ Done Processing")
