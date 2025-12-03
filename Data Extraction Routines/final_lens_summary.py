# to do
# use EXCEL_XLSX = os.path.join(BASE, "Summary - primes.xlsx")
# OR LensDataExportsRenamed file names (combine with Materials folder? reroute path below)
# to get 1/2 diag, f/#, efl
# add giga

import os
import glob
import pandas as pd

# ---------------------------
# User-editable paths
# ---------------------------
BASE = r"Prime Lenses + Data"
CSV_EXPORTS = os.path.join(BASE, "CSVExports")

MATERIALS_DIR = os.path.join(CSV_EXPORTS, "Materials")
FIELD_CURV_DIR = os.path.join(CSV_EXPORTS, "FieldCurvature")
LONG_DIR = os.path.join(CSV_EXPORTS, "Longitudinal")
RMS_DIR = os.path.join(CSV_EXPORTS, "RMSvField")
VIGN_DIR = os.path.join(CSV_EXPORTS, "Vignetting")

# Where to write outputs
OUTPUT_CSV = os.path.join(CSV_EXPORTS, "file_lens_summary.csv")

# ---------------------------
# Helper: robust CSV reader
# ---------------------------
def read_csv_robust(path):
    """Try to read CSV with utf-8, then utf-16. Returns DataFrame or raises."""
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-16")

# ---------------------------
# Helper: find corresponding file
# ---------------------------
def find_lens_file(directory, base_name):
    """
    Attempt to find the CSV in directory matching base_name (which is e.g. CH321571_Example01P).
    Prefer exact match with '_LensData.csv', else try glob patterns.
    Returns path or None.
    """
    if not os.path.isdir(directory):
        return None

    # exact expected filename
    expected = os.path.join(directory, base_name + "_LensData.csv")
    if os.path.isfile(expected):
        return expected

    # sometimes files may be named slightly differently or without .csv extension in examples:
    patterns = [
        os.path.join(directory, base_name + "_LensData*.*"),   # any extension
        os.path.join(directory, base_name + "*LensData*.csv"),
        os.path.join(directory, "*" + base_name + "*LensData*.csv"),
        os.path.join(directory, base_name + "*.csv"),
        os.path.join(directory, "*" + base_name + "*.csv"),
    ]
    for p in patterns:
        matches = glob.glob(p)
        if matches:
            # prefer shortest (most exact) then first
            matches.sort(key=lambda x: (len(x), x))
            return matches[0]
    return None

# ---------------------------
# Columns to extract (source -> list of column names)
# ---------------------------
# Materials: Thickness, Material, SemiDiameter (first row)
MATERIALS_COLS = ["Thickness", "Material", "SemiDiameter"]

# FieldCurvature: Tan Shift, Sag Shift
FIELD_CURV_COLS = ["Tan Shift", "Sag Shift"]

# Longitudinal: wavelengths 0.4861, 0.5876, 0.6563
LONG_COLS = ["0.4861", "0.5876", "0.6563"]

# RMSvField: Poly, 0.4861, 0.5876, 0.6563
RMS_COLS = ["Poly", "0.4861", "0.5876", "0.6563"]

# Vignetting: Rel. Ill, Effective F/#
VIGN_COLS = ["Rel. Ill", "Effective F/#"]

# We'll prefix repeated numeric-named columns so headers are unique in the final CSV:
LONG_PREFIX = "Long_"
RMS_PREFIX = "RMS_"

# ---------------------------
# Build list of lens base names from Materials folder
# ---------------------------
materials_glob = os.path.join(MATERIALS_DIR, "*_LensData*.csv")
materials_files = glob.glob(materials_glob)
materials_files.sort()

if not materials_files:
    print("No material LensData files found in:", MATERIALS_DIR)
    raise SystemExit(1)

# derive base names (strip directory and trailing _LensData and extension)
bases = []
for p in materials_files:
    fname = os.path.basename(p)
    # remove extension
    name_no_ext = os.path.splitext(fname)[0]
    # remove trailing _LensData if present
    if name_no_ext.endswith("_LensData"):
        base = name_no_ext[:-len("_LensData")]
    else:
        # if file name is like CH..._LensData_extra or CH..._LensData-v1, find first occurrence
        base = name_no_ext.replace("_LensData", "")
    base = base.strip()
    if base:
        bases.append(base)

# Deduplicate bases while preserving order
seen = set()
bases_unique = []
for b in bases:
    if b not in seen:
        bases_unique.append(b)
        seen.add(b)

# ---------------------------
# Iterate and extract first-row values
# ---------------------------
rows = []
for base in bases_unique:
    row = {"File Name": base}

    # --- Materials ---
    mat_file = find_lens_file(MATERIALS_DIR, base)
    if mat_file:
        try:
            df_mat = read_csv_robust(mat_file)
            if not df_mat.empty:
                first = df_mat.iloc[0]
                for col in MATERIALS_COLS:
                    # tolerate slight column name whitespace differences
                    if col in df_mat.columns:
                        val = first[col]
                    else:
                        # try case-insensitive match/strip
                        matches = [c for c in df_mat.columns if c.strip().lower() == col.strip().lower()]
                        val = first[matches[0]] if matches else pd.NA
                    row[col] = val
            else:
                for col in MATERIALS_COLS:
                    row[col] = pd.NA
        except Exception as e:
            for col in MATERIALS_COLS:
                row[col] = pd.NA
    else:
        for col in MATERIALS_COLS:
            row[col] = pd.NA

    # --- Field Curvature ---
    fc_file = find_lens_file(FIELD_CURV_DIR, base)
    if fc_file:
        try:
            df_fc = read_csv_robust(fc_file)
            if not df_fc.empty:
                first = df_fc.iloc[0]
                for col in FIELD_CURV_COLS:
                    if col in df_fc.columns:
                        val = first[col]
                    else:
                        matches = [c for c in df_fc.columns if c.strip().lower() == col.strip().lower()]
                        val = first[matches[0]] if matches else pd.NA
                    row[col] = val
            else:
                for col in FIELD_CURV_COLS:
                    row[col] = pd.NA
        except Exception:
            for col in FIELD_CURV_COLS:
                row[col] = pd.NA
    else:
        for col in FIELD_CURV_COLS:
            row[col] = pd.NA

    # --- Longitudinal (prefix names to avoid header collisions) ---
    long_file = find_lens_file(LONG_DIR, base)
    if long_file:
        try:
            df_long = read_csv_robust(long_file)
            if not df_long.empty:
                first = df_long.iloc[0]
                for col in LONG_COLS:
                    out_name = LONG_PREFIX + col
                    if col in df_long.columns:
                        val = first[col]
                    else:
                        matches = [c for c in df_long.columns if c.strip().lower() == col.strip().lower()]
                        val = first[matches[0]] if matches else pd.NA
                    row[out_name] = val
            else:
                for col in LONG_COLS:
                    row[LONG_PREFIX + col] = pd.NA
        except Exception:
            for col in LONG_COLS:
                row[LONG_PREFIX + col] = pd.NA
    else:
        for col in LONG_COLS:
            row[LONG_PREFIX + col] = pd.NA

    # --- RMSvField (prefix for wavelength columns) ---
    rms_file = find_lens_file(RMS_DIR, base)
    if rms_file:
        try:
            df_rms = read_csv_robust(rms_file)
            if not df_rms.empty:
                first = df_rms.iloc[0]
                # Poly first
                if "Poly" in df_rms.columns:
                    row["Poly"] = first["Poly"]
                else:
                    matches = [c for c in df_rms.columns if c.strip().lower() == "poly"]
                    row["Poly"] = first[matches[0]] if matches else pd.NA
                # wavelengths
                for col in ["0.4861", "0.5876", "0.6563"]:
                    out_name = RMS_PREFIX + col
                    if col in df_rms.columns:
                        val = first[col]
                    else:
                        matches = [c for c in df_rms.columns if c.strip().lower() == col.strip().lower()]
                        val = first[matches[0]] if matches else pd.NA
                    row[out_name] = val
            else:
                row["Poly"] = pd.NA
                for col in ["0.4861", "0.5876", "0.6563"]:
                    row[RMS_PREFIX + col] = pd.NA
        except Exception:
            row["Poly"] = pd.NA
            for col in ["0.4861", "0.5876", "0.6563"]:
                row[RMS_PREFIX + col] = pd.NA
    else:
        row["Poly"] = pd.NA
        for col in ["0.4861", "0.5876", "0.6563"]:
            row[RMS_PREFIX + col] = pd.NA

    # --- Vignetting ---
    vig_file = find_lens_file(VIGN_DIR, base)
    if vig_file:
        try:
            df_vig = read_csv_robust(vig_file)
            if not df_vig.empty:
                first = df_vig.iloc[0]
                for col in VIGN_COLS:
                    if col in df_vig.columns:
                        val = first[col]
                    else:
                        matches = [c for c in df_vig.columns if c.strip().lower() == col.strip().lower()]
                        val = first[matches[0]] if matches else pd.NA
                    row[col] = val
            else:
                for col in VIGN_COLS:
                    row[col] = pd.NA
        except Exception:
            for col in VIGN_COLS:
                row[col] = pd.NA
    else:
        for col in VIGN_COLS:
            row[col] = pd.NA

    rows.append(row)

# ---------------------------
# Assemble DataFrame and write output
# ---------------------------
df_out = pd.DataFrame(rows)

# Re-order columns to match your requested order (with disambiguation for duplicates)
ordered_cols = ["File Name",
                "Thickness", "Material", "SemiDiameter",
                "Tan Shift", "Sag Shift",
                LONG_PREFIX + "0.4861", LONG_PREFIX + "0.5876", LONG_PREFIX + "0.6563",
                "Poly", RMS_PREFIX + "0.4861", RMS_PREFIX + "0.5876", RMS_PREFIX + "0.6563",
                "Rel. Ill", "Effective F/#"]

# keep only columns that exist in df_out
ordered_cols = [c for c in ordered_cols if c in df_out.columns]
df_out = df_out[ordered_cols]

# Save CSV
df_out.to_csv(OUTPUT_CSV, index=False)
print("Wrote summary CSV to:", OUTPUT_CSV)
