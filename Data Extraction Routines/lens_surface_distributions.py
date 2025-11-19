import os
import math
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ============================================================
# CONFIGURATION
# ============================================================

ROOT_DIR = r"C:\Users\Ruth\Documents\GitHub\2.156-Lens-Project\Prime Lenses + Data\LensDataExports"

# Folder where summary CSVs and plots will be written
OUTPUT_DIR = os.path.join(
    r"C:\Users\Ruth\Documents\GitHub\2.156-Lens-Project\Prime Lenses + Data\CSVExports",
    "LensDataAnalysis"
)
os.makedirs(OUTPUT_DIR, exist_ok=True)

PLOTS_DIR = os.path.join(OUTPUT_DIR, "plots")
os.makedirs(PLOTS_DIR, exist_ok=True)

# Columns we expect to be numeric (based on your example)
NUMERIC_COLS = {
    "Surface",
    "Radius",
    "Thickness",
    "SemiDiameter",
    "Conic",
    "A2", "A4", "A6", "A8", "A10", "A12", "A14", "A16"
}

# ============================================================
# HELPER: TRY TO READ CSV WITH DIFFERENT ENCODINGS
# ============================================================

def read_csv_robust(path):
    """
    Try reading CSV as UTF-8, then UTF-16 if needed.
    Returns a pandas DataFrame or raises the last exception.
    """
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-16")


# ============================================================
# STEP 1: WALK ALL CSV FILES & COLLECT DATAFRAMES
# ============================================================

all_dfs = []
row_counts = []  # list of dicts: {"file": ..., "n_rows": ...}

print(f"Scanning CSV files under:\n  {ROOT_DIR}\n")

for root, dirs, files in os.walk(ROOT_DIR):
    for fname in files:
        if not fname.lower().endswith(".csv"):
            continue

        fpath = os.path.join(root, fname)
        # print(f"Reading: {fpath}")

        try:
            df = read_csv_robust(fpath)
        except Exception as e:
            print(f"  ⚠️ Skipping (could not read): {e}")
            continue

        # Number of data rows (excluding header) = len(df)
        n_rows = len(df)
        row_counts.append({
            "file": fpath,
            "n_rows": n_rows
        })

        all_dfs.append(df)

if not all_dfs:
    print("No CSV files found or readable. Exiting.")
    raise SystemExit

print(f"\nLoaded {len(all_dfs)} CSV files.")

# ============================================================
# STEP 2: COMBINE ALL DATA
# ============================================================

combined = pd.concat(all_dfs, ignore_index=True, sort=False)
print(f"Combined DataFrame shape: {combined.shape}")

# ============================================================
# STEP 3: DISTRIBUTION OF ROW COUNTS PER FILE (+ PLOT)
# ============================================================

row_counts_df = pd.DataFrame(row_counts)

# Save distribution of number of rows per file
row_counts_path = os.path.join(OUTPUT_DIR, "row_counts_per_file.csv")
row_counts_df.to_csv(row_counts_path, index=False)

print(f"\nSaved row-count distribution to:\n  {row_counts_path}")

print("\nRow count stats (per file):")
print(row_counts_df["n_rows"].describe())

# Plot histogram of row counts per file
plt.figure()
row_counts_df["n_rows"].hist(bins=20)
plt.xlabel("Number of data rows per file")
plt.ylabel("Count of files")
plt.title("Distribution of number of rows per file")
plt.tight_layout()
row_hist_path = os.path.join(PLOTS_DIR, "row_counts_histogram.png")
plt.savefig(row_hist_path, dpi=200)
plt.close()
print(f"Saved row-count histogram to:\n  {row_hist_path}")

# ============================================================
# STEP 4: NUMERIC COLUMN DISTRIBUTIONS (+ PLOTS)
# ============================================================

numeric_summary = []

# Only consider numeric columns that actually exist
numeric_cols_present = [c for c in combined.columns if c in NUMERIC_COLS]

for col in numeric_cols_present:
    s_raw = combined[col]

    # Convert to numeric, coercing non-numeric to NaN
    s = pd.to_numeric(s_raw, errors="coerce")

    total = len(s)
    nan_count = s.isna().sum()
    posinf_count = np.isposinf(s).sum()
    neginf_count = np.isneginf(s).sum()
    finite_mask = np.isfinite(s)
    finite_count = finite_mask.sum()

    # Stats over finite vals only
    if finite_count > 0:
        finite_vals = s[finite_mask]
        mean = finite_vals.mean()
        std = finite_vals.std()
        vmin = finite_vals.min()
        vmax = finite_vals.max()
    else:
        finite_vals = pd.Series([], dtype=float)
        mean = std = vmin = vmax = np.nan

    # Count how many entries in the original column
    # were non-empty but couldn't be parsed as numbers
    non_empty_original = s_raw.notna().sum()
    non_numeric_original = non_empty_original - (total - nan_count)

    numeric_summary.append({
        "column": col,
        "total_entries": int(total),
        "finite_count": int(finite_count),
        "nan_count": int(nan_count),
        "posinf_count": int(posinf_count),
        "neginf_count": int(neginf_count),
        "non_numeric_original_count": int(max(non_numeric_original, 0)),
        "finite_fraction": finite_count / total if total else math.nan,
        "nan_fraction": nan_count / total if total else math.nan,
        "posinf_fraction": posinf_count / total if total else math.nan,
        "neginf_fraction": neginf_count / total if total else math.nan,
        "mean_over_finite": mean,
        "std_over_finite": std,
        "min_over_finite": vmin,
        "max_over_finite": vmax,
    })

    # -------- Plot histogram for this numeric column --------
    if finite_count > 0:
        plt.figure()
        # Finite values histogram
        finite_vals.hist(bins=40)
        plt.xlabel(col)
        plt.ylabel("Count")
        plt.title(
            f"{col} (finite values only)\n"
            f"N={finite_count}, NaN={nan_count}, +Inf={posinf_count}, -Inf={neginf_count}"
        )
        plt.tight_layout()
        col_hist_path = os.path.join(PLOTS_DIR, f"{col}_histogram.png")
        plt.savefig(col_hist_path, dpi=200)
        plt.close()
        print(f"Saved numeric histogram for '{col}' to:\n  {col_hist_path}")
    else:
        print(f"No finite values for numeric column '{col}', skipping histogram.")

numeric_summary_df = pd.DataFrame(numeric_summary)
numeric_summary_path = os.path.join(OUTPUT_DIR, "numeric_column_summary.csv")
numeric_summary_df.to_csv(numeric_summary_path, index=False)

print(f"\nSaved numeric column summary to:\n  {numeric_summary_path}")


# ============================================================
# STEP 5: CATEGORICAL / TEXT COLUMN DISTRIBUTIONS (+ PLOTS)
# ============================================================

# Define categorical columns as "everything that's not in NUMERIC_COLS"
categorical_cols = [c for c in combined.columns if c not in NUMERIC_COLS]

if categorical_cols:
    cat_dir = os.path.join(OUTPUT_DIR, "categorical_distributions")
    os.makedirs(cat_dir, exist_ok=True)

    for col in categorical_cols:
        vc = combined[col].fillna("<NaN>").value_counts(dropna=False)

        # Save full value counts table
        out_path = os.path.join(cat_dir, f"{col}_value_counts.csv")
        vc.to_csv(out_path, header=["count"])
        print(f"Saved categorical distribution for '{col}' to:\n  {out_path}")

        # Plot top N categories (to keep plots readable)
        top_n = 20
        top_vals = vc.head(top_n)

        plt.figure(figsize=(max(6, 0.4 * len(top_vals)), 4))
        top_vals.plot(kind="bar")
        plt.xlabel(col)
        plt.ylabel("Count")
        plt.title(f"Top {len(top_vals)} values for '{col}'")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()

        cat_plot_path = os.path.join(PLOTS_DIR, f"{col}_top_values.png")
        plt.savefig(cat_plot_path, dpi=200)
        plt.close()
        print(f"Saved categorical bar plot for '{col}' to:\n  {cat_plot_path}")
else:
    print("\nNo categorical columns detected (all columns treated as numeric?).")

print("\n✅ Done. All summaries and plots written to:")
print(f"  {OUTPUT_DIR}")
print(f"  {PLOTS_DIR}")
