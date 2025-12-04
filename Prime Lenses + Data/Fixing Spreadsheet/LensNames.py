
import os

folder_path = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses"

# Get all .zmx files and clean them
zmx_files = [
    f.lower().replace(".zmx", "").strip() 
    for f in os.listdir(folder_path)
    if f.lower().endswith(".zmx") and os.path.isfile(os.path.join(folder_path, f))
]
zmx_files = [f.replace("_", " ") for f in zmx_files]

# print(zmx_files)


for i in range(len(zmx_files)):
    zmx_files[i] = os.path.splitext(zmx_files[i])[0]

file_path = r"C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Data Processing\SpreadsheetRows.txt"

with open(file_path, "r", encoding="utf-8") as f:
    lines = [line.strip().lower() for line in f]

lines = [f.replace("_", " ") for f in lines]
# print(lines)

# Items in zmx_files but not in cleaned_lines
only_in_zmx_files = [x for x in zmx_files if x not in lines]

# Items in cleaned_lines but not in zmx_files
only_in_lines = [x for x in lines if x not in zmx_files]

# All differences (preserving order from both lists)
all_differences = only_in_zmx_files + only_in_lines
all_differences.sort()

print("\n Only in zmx_files:", only_in_zmx_files)
print("\n Only in lines:", only_in_lines)
# print("All differences:", all_differences)
