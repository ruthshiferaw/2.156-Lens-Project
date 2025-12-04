import os
import shutil
BASE = r"Prime Lenses + Data"
ZMX_DIR = BASE
RMS_DIR = os.path.join(BASE, "CSVExports", "RMSvField")
FAILED_DIR = os.path.join(BASE, "failed", "RMSvField")
# Ensure failed folder exists
os.makedirs(FAILED_DIR, exist_ok=True)
# -------------------------------------------------------
# Collect RMSvField basenames (strip _RMSvField.csv)
# -------------------------------------------------------
rms_basenames = set()
for f in os.listdir(RMS_DIR):
    if f.lower().endswith(".csv") and "_RMSvField" in f:
        name = f.replace("_RMSvField", "").replace(".csv", "")
        rms_basenames.add(name)
print(f"Found {len(rms_basenames)} RMSvField lens names.")
# -------------------------------------------------------
# Check ZMX files
# -------------------------------------------------------
moved = []
for f in os.listdir(ZMX_DIR):
    if f.lower().endswith(".zmx"):
        base = f.replace(".zmx", "")
        # If base name is not found in RMS field CSVs → move the .zmx file
        if base not in rms_basenames:
            src = os.path.join(ZMX_DIR, f)
            dst = os.path.join(FAILED_DIR, f)
            shutil.move(src, dst)
            moved.append(f)
# -------------------------------------------------------
# Summary
# -------------------------------------------------------
print(f"\nMoved {len(moved)} missing RMSvField designs:")
for f in moved:
    print("  -", f)