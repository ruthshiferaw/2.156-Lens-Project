#!/usr/bin/env python3
"""
robust_normalize_summary.py

Robust min-max normalization:
 - Attempts to convert every column to numeric (coerce errors -> NaN).
 - Normalizes columns that have at least one finite numeric value to [0,1].
 - Leaves non-numeric columns unchanged.
 - Preserves NaNs (does not impute).
 - Writes diagnostics to stdout.
"""
import os
import pandas as pd
import numpy as np

INPUT = r"Prime Lenses + Data/CSVExports/file_lens_summary.csv"
OUTPUT = r"Prime Lenses + Data/CSVExports/file_lens_summary_normalized.csv"

df = pd.read_csv(INPUT)
print("Loaded:", INPUT, "shape:", df.shape)

df_out = df.copy()   # we'll replace normalized numeric columns here

cols = df.columns.tolist()
normalized_cols = []
skipped_cols = []
constant_cols = []
problem_cols = []

for col in cols:
    # Try to convert column to numeric (coerce errors -> NaN)
    converted = pd.to_numeric(df[col], errors='coerce')

    finite_mask = np.isfinite(converted.values)
    n_finite = int(finite_mask.sum())

    if n_finite == 0:
        # No numeric values at all (or all NaN/inf after conversion) -> skip
        skipped_cols.append(col)
        continue

    # compute min/max on finite values
    col_min = float(np.nanmin(converted.values))
    col_max = float(np.nanmax(converted.values))

    if col_min == col_max:
        # constant column: set finite entries to 0 (preserve NaNs)
        tmp = converted.copy()
        tmp[finite_mask] = 0.0
        df_out[col] = tmp
        constant_cols.append((col, col_min))
        normalized_cols.append(col)
    else:
        # normal min-max scale; only map finite values
        denom = (col_max - col_min)
        tmp = converted.copy()
        tmp_vals = (tmp[finite_mask] - col_min) / denom
        # numerical safety: clip to [0,1] (tiny rounding can push outside)
        tmp_vals = np.clip(tmp_vals, 0.0, 1.0)
        tmp.loc[finite_mask] = tmp_vals
        df_out[col] = tmp
        normalized_cols.append(col)

# diagnostics
print(f"\nNormalization complete. Columns normalized: {len(normalized_cols)}")
if normalized_cols:
    print("  sample normalized cols:", normalized_cols[:8])
if constant_cols:
    print("\nConstant columns (min==max) set to 0 for finite entries:")
    for c, v in constant_cols[:10]:
        print(f"  {c}: constant_value={v}")
if skipped_cols:
    print("\nSkipped (no numeric values after coercion):", skipped_cols[:10])
if len(skipped_cols) > 10:
    print(f"  ... ({len(skipped_cols)} skipped total)")

# final check: any numeric column still outside [0,1]?
out_of_bounds = []
for col in normalized_cols:
    # only check finite entries
    series = pd.to_numeric(df_out[col], errors='coerce')
    finite = series[np.isfinite(series)]
    if finite.size == 0:
        continue
    if finite.min() < -1e-9 or finite.max() > 1.0 + 1e-9:
        out_of_bounds.append((col, float(finite.min()), float(finite.max())))
if out_of_bounds:
    print("\nWarning: these normalized columns have values outside [0,1] (due to rounding or unexpected values):")
    for col, mn, mx in out_of_bounds:
        print(f"  {col}: min={mn}, max={mx}")
else:
    print("\nAll normalized columns are within [0,1] for finite entries.")

# Save
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
df_out.to_csv(OUTPUT, index=False)
print("\nSaved normalized file to:", OUTPUT)
