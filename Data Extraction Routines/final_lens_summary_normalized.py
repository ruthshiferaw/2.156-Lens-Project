import pandas as pd
import numpy as np
import os

# ----------------------------------------------------
# Paths
# ----------------------------------------------------
INPUT = r"Prime Lenses + Data/CSVExports/file_lens_summary.csv"
OUTPUT = r"Prime Lenses + Data/CSVExports/file_lens_summary_normalized.csv"

# ----------------------------------------------------
# Load file
# ----------------------------------------------------
df = pd.read_csv(INPUT)

# Identify numeric columns only
num_cols = df.select_dtypes(include=[np.number]).columns.tolist()

print("Numeric columns found:", num_cols)

# ----------------------------------------------------
# Clean data (handle NaN, inf, -inf)
# ----------------------------------------------------
df[num_cols] = df[num_cols].replace([np.inf, -np.inf], np.nan)

# ----------------------------------------------------
# Min–max normalization per column
# ----------------------------------------------------
df_norm = df.copy()

for col in num_cols:
    col_min = df_norm[col].min(skipna=True)
    col_max = df_norm[col].max(skipna=True)

    # If all values are identical or all NaN → make column 0
    if pd.isna(col_min) or pd.isna(col_max) or col_min == col_max:
        df_norm[col] = 0
    else:
        df_norm[col] = (df_norm[col] - col_min) / (col_max - col_min)

# ----------------------------------------------------
# Save output
# ----------------------------------------------------
df_norm.to_csv(OUTPUT, index=False)
print(f"Saved normalized file to: {OUTPUT}")
