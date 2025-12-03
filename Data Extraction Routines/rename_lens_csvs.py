import os
import shutil
import pandas as pd
import re

# --- CONFIG ---
BASE = r"Prime Lenses + Data"
EXCEL_XLSX = os.path.join(BASE, "Summary - primes.xlsx")
EXCEL_CSV_FALLBACK = os.path.join(BASE, "Summary - primes.csv")
INPUT = os.path.join(BASE, "LensDataExports")
OUTPUT = os.path.join(BASE, "LensDataExportsRenamed")
os.makedirs(OUTPUT, exist_ok=True)

NOTE_PATTERN = re.compile(r'^\s*Note\s+(\d+)\s+Fig\s+(\d+)\s*$', flags=re.IGNORECASE)
ILLEGAL_RE = re.compile(r'[<>:"/\\|?*]')

# --- HELPERS ---
def read_summary():
    """Try Excel first, fall back to CSV if reading xlsx fails."""
    if os.path.exists(EXCEL_XLSX):
        try:
            df = pd.read_excel(EXCEL_XLSX, sheet_name=0, header=0, dtype=object)
            print(f"Read Excel: {EXCEL_XLSX} (rows={len(df)})")
            return df
        except Exception as e:
            print(f"Warning reading XLSX: {e}. Trying CSV fallback...")
    if os.path.exists(EXCEL_CSV_FALLBACK):
        df = pd.read_csv(EXCEL_CSV_FALLBACK, header=0, dtype=object)
        print(f"Read CSV fallback: {EXCEL_CSV_FALLBACK} (rows={len(df)})")
        return df
    raise FileNotFoundError(
        f"Could not read '{EXCEL_XLSX}'. Either install openpyxl in your venv "
        f"or export the spreadsheet to CSV at '{EXCEL_CSV_FALLBACK}'."
    )

def transform_colA_only_two_rules(raw):
    """
    Apply exactly the two rules requested:
      - If string matches "Note <n> Fig <m>" -> return "Note{n:02d}Fig{m}" (no underscores)
      - Else: replace spaces with underscores and strip
    Do NOT perform any other normalization.
    """
    if raw is None:
        return "NA"
    s = str(raw).strip()
    if s == "" or s.lower() in ["nan", "na", "none"]:
        return "NA"
    m = NOTE_PATTERN.match(s)
    if m:
        n = int(m.group(1))
        m2 = int(m.group(2))
        return f"Note{n:02d}Fig{m2}"
    # default: replace spaces with underscores only
    return s.replace(" ", "_")

def numeric_or_NA(cell):
    """Return a numeric formatted string if cell is numeric (accepts comma decimal), otherwise 'NA'."""
    if cell is None:
        return "NA"
    s = str(cell).strip()
    if s == "" or s.lower() in ["nan", "na", "none"]:
        return "NA"
    # replace comma decimal -> dot
    s_num = s.replace(",", ".")
    try:
        f = float(s_num)
        if pd.isna(f):
            return "NA"
        # integer without decimal point
        if float(f).is_integer():
            return str(int(round(f)))
        # trim trailing zeros
        return f"{f:.6f}".rstrip('0').rstrip('.')
    except Exception:
        return "NA"

def sanitize_for_filename(s):
    """Make s safe for Windows filenames (also preserve 'NA')."""
    if s is None:
        return "NA"
    s = str(s).strip()
    if s == "" or s.lower() in ["nan", "na", "none"]:
        return "NA"
    s = ILLEGAL_RE.sub("_", s)
    s = s.rstrip(" .")
    s = re.sub(r'__+', '_', s)
    return s

def safe_copy(src, dest_base):
    """Copy src to dest folder using dest_base (no extension); avoid collisions by adding _1/_2..."""
    folder = os.path.dirname(dest_base)
    base = os.path.splitext(os.path.basename(dest_base))[0]
    ext = os.path.splitext(dest_base)[1] or ".csv"
    candidate = os.path.join(folder, base + ext)
    i = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f"{base}_{i}{ext}")
        i += 1
    shutil.copy2(src, candidate)
    return candidate

# --- MAIN ---
def main():
    df = read_summary()
    ncols = df.shape[1]
    # Build lookup: transformed_colA -> (d_str, e_str, f_str)
    lookup = {}
    for _, row in df.iterrows():
        rawA = row.iloc[0] if ncols > 0 else None
        key = transform_colA_only_two_rules(rawA)

        # D/E/F are columns indices 3,4,5 (zero-based)
        def get_col(i):
            if i >= ncols:
                return "NA"
            return numeric_or_NA(row.iloc[i])
        d_val = get_col(3)
        e_val = get_col(4)
        f_val = get_col(5)

        lookup[key] = (d_val, e_val, f_val)

    total = 0
    renamed = 0
    copied_unchanged = 0
    errors = []

    for fname in sorted(os.listdir(INPUT)):
        if not fname.lower().endswith(".csv"):
            continue
        total += 1
        src = os.path.join(INPUT, fname)
        base_no_ext = fname[:-4] if fname.lower().endswith(".csv") else fname

        # Per your instruction: drop trailing "_LensData" before comparing
        if base_no_ext.endswith("_LensData"):
            compare_key = base_no_ext[:-9]
        else:
            compare_key = base_no_ext

        # Compare using only the two-rule transformation on spreadsheet side
        if compare_key in lookup:
            d_raw, e_raw, f_raw = lookup[compare_key]
            # sanitize the three values for filenames
            d_safe = sanitize_for_filename(d_raw)
            e_safe = sanitize_for_filename(e_raw)
            f_safe = sanitize_for_filename(f_raw)
            new_base = f"{base_no_ext}_{d_safe}_{e_safe}_{f_safe}"
            dest_base = os.path.join(OUTPUT, new_base + ".csv")
            try:
                safe_copy(src, dest_base)
                renamed += 1
            except Exception as exc:
                errors.append((fname, str(exc)))
        else:
            # not found: copy unchanged
            dest_base = os.path.join(OUTPUT, fname)
            try:
                safe_copy(src, dest_base)
                copied_unchanged += 1
            except Exception as exc:
                errors.append((fname, str(exc)))

    print(f"Total processed: {total}")
    print(f"Renamed/copied with metadata: {renamed}")
    print(f"Copied unchanged (no match): {copied_unchanged}")
    if errors:
        print("\nErrors copying the following files:")
        for fn, msg in errors:
            print(" -", fn, ":", msg)
    print("\nOutput folder:", OUTPUT)

if __name__ == "__main__":
    main()
