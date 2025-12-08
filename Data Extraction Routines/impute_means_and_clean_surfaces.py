#!/usr/bin/env python3
"""
impute_means_and_clean_surfaces.py

Creates:
 - Prime Lenses + Data/CSVExports/file_lens_summary_normalized_imputed_mean.csv
 - Prime Lenses + Data/CSVExports/Materials_cleaned/  (all per-surface CSVs with numeric NaN/Inf imputed)

Behavior:
 - TARGET_COLS in the summary CSV: convert to numeric, replace +/-inf with NaN, impute column mean (or 0 if column empty).
 - Per-surface numeric columns ('Radius','Thickness','SemiDiameter'): compute global mean across all files (ignoring NaN/Inf),
   then replace NaN/Inf per-file with that mean. 'Material' missing -> "None".
 - Preserves other columns in each surface CSV (only overwrites the cleaned columns).
"""

import os
import io
import glob
import numpy as np
import pandas as pd

# ---------------------------
# User-editable paths
# ---------------------------
BASE = r"Prime Lenses + Data"
CSV_EXPORTS = os.path.join(BASE, "CSVExports")
ROOT_SUMMARY_CSV = os.path.join(CSV_EXPORTS, "file_lens_summary_normalized.csv")
CLEANED_SUMMARY_CSV = os.path.join(CSV_EXPORTS, "file_lens_summary_normalized_imputed_mean.csv")
MATERIALS_DIR = os.path.join(BASE, "LensDataExportsRenamedMaterials")
MATERIALS_CLEANED_DIR = os.path.join(BASE, "LensDataExportsRenamedMaterialsCleaned")
os.makedirs(MATERIALS_CLEANED_DIR, exist_ok=True)

# ---------------------------
TARGET_COLS = [
    "Tan Shift", "Sag Shift",
    "Long_0.4861", "Long_0.5876", "Long_0.6563",
    "Poly",
    "RMS_0.4861", "RMS_0.5876", "RMS_0.6563",
    "Effective F/#" #"Rel. Ill", 
]

SURFACE_NUM_COLS = ["Radius", "Thickness", "SemiDiameter"]
MAT_COL = "Material"

# ---------------------------
# Robust CSV reader (tries several encodings)
def read_csv_tolerant(path):
    encs = ["utf-8", "utf-8-sig", "utf-16", "latin1", "cp1252"]
    for enc in encs:
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            continue
    # final fallback: read bytes and decode with replacement
    with open(path, "rb") as fh:
        text = fh.read().decode("utf-8", errors="replace")
    return pd.read_csv(io.StringIO(text))

# ---------------------------
# 1) Load summary CSV and impute TARGET_COLS by column mean
if not os.path.isfile(ROOT_SUMMARY_CSV):
    raise SystemExit(f"Summary CSV not found: {ROOT_SUMMARY_CSV}")

print("Loading summary:", ROOT_SUMMARY_CSV)
summary = read_csv_tolerant(ROOT_SUMMARY_CSV)
print("Summary shape:", summary.shape)

# Ensure File Name present
if "File Name" not in summary.columns:
    raise SystemExit("Summary CSV missing 'File Name' column. Aborting.")

# Convert target columns to numeric, replace inf with NaN
for c in TARGET_COLS:
    if c in summary.columns:
        summary[c] = pd.to_numeric(summary[c], errors="coerce")
    else:
        # create missing column with NaN -> will be imputed to 0
        summary[c] = np.nan

summary[TARGET_COLS] = summary[TARGET_COLS].replace([np.inf, -np.inf], np.nan)

# Compute means & impute (use 0 fallback when no finite values)
col_means = {}
for c in TARGET_COLS:
    finite_vals = summary[c].dropna().values
    if finite_vals.size == 0:
        col_means[c] = 0.0
    else:
        col_means[c] = float(np.mean(finite_vals))
    summary[c] = summary[c].fillna(col_means[c])

print("Imputed summary TARGET_COLS with column means (examples):")
for c in TARGET_COLS[:6]:
    print(f"  {c}: mean={col_means[c]:.6g}")

# Save cleaned summary
summary.to_csv(CLEANED_SUMMARY_CSV, index=False)
print("Wrote cleaned summary to:", CLEANED_SUMMARY_CSV)

# ---------------------------
# 2) Collect surface files and compute global means for numeric surface columns
surface_glob = os.path.join(MATERIALS_DIR, "*.csv")
surface_files = sorted(glob.glob(surface_glob))
if len(surface_files) == 0:
    raise SystemExit(f"No surface CSVs found in {MATERIALS_DIR}")

print(f"Found {len(surface_files)} surface CSVs in {MATERIALS_DIR} — scanning for numeric values...")

num_pool = {col: [] for col in SURFACE_NUM_COLS}
files_with_issues = 0

for p in surface_files:
    try:
        df = read_csv_tolerant(p)
    except Exception as e:
        print("Warning: failed to read", p, ":", e)
        continue

    # ensure columns exist; coerce numeric and collect finite values
    for col in SURFACE_NUM_COLS:
        if col in df.columns:
            vals = pd.to_numeric(df[col], errors="coerce")
            vals = vals.replace([np.inf, -np.inf], np.nan).dropna().astype(float).tolist()
            num_pool[col].extend(vals)
        else:
            # column missing -> nothing to add
            pass

# compute means (fallback to 0.0)
surface_means = {}
for col in SURFACE_NUM_COLS:
    vals = num_pool[col]
    if len(vals) == 0:
        surface_means[col] = 0.0
    else:
        surface_means[col] = float(np.mean(vals))

print("Computed global surface means:")
for col in SURFACE_NUM_COLS:
    print(f"  {col}: mean = {surface_means[col]:.6g}  (collected {len(num_pool[col])} finite values)")

# ---------------------------
# 3) Second pass: write cleaned surface CSVs to MATERIALS_CLEANED_DIR (preserve other columns)
written = 0
for p in surface_files:
    fname = os.path.basename(p)
    base = os.path.splitext(fname)[0]
    try:
        df = read_csv_tolerant(p)
    except Exception as e:
        print("Skipping unreadable file:", p, "->", e)
        continue

    # Ensure columns exist and coerce/clean numeric columns
    for col in SURFACE_NUM_COLS:
        if col not in df.columns:
            # create column filled with mean
            df[col] = surface_means[col]
        else:
            # coerce numeric and replace inf with NaN, then fill with global mean
            df[col] = pd.to_numeric(df[col], errors="coerce").replace([np.inf, -np.inf], np.nan).astype(float)
            df[col] = df[col].fillna(surface_means[col])

    # Material column: fill missing with "None"
    if MAT_COL in df.columns:
        df[MAT_COL] = df[MAT_COL].astype(str).fillna("None")
        # df[MAT_COL].fillna("None").astype(str)
    else:
        df[MAT_COL] = "None"

    out_path = os.path.join(MATERIALS_CLEANED_DIR, fname)
    # write preserving all columns (we modified the numeric ones in-place)
    df.to_csv(out_path, index=False)
    written += 1

print(f"Wrote {written} cleaned surface CSVs to: {MATERIALS_CLEANED_DIR}")

# ---------------------------
# 4) Summary report
print("\nDone. Summary:")
print(" - Cleaned summary saved:", CLEANED_SUMMARY_CSV)
print(" - Cleaned materials folder:", MATERIALS_CLEANED_DIR)
print(" - Example cleaned file (first):", os.path.join(MATERIALS_CLEANED_DIR, os.listdir(MATERIALS_CLEANED_DIR)[0]) if os.listdir(MATERIALS_CLEANED_DIR) else "(none)")
