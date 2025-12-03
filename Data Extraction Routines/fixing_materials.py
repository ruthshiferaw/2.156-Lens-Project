import os
import pandas as pd

# Folders (relative paths based on your repo structure)
input_folder = r"Prime Lenses + Data\LensDataExports"
output_folder = r"Prime Lenses + Data\CSVExports\Materials"

# Create output directory if missing
os.makedirs(output_folder, exist_ok=True)

def fix_material_column(df):
    """
    Fix material propagation, including touching-lens cases.
    - Treat real NaNs correctly.
    - If the first surface (i==0) is blank but the next surface has a material,
      copy the next surface's material into surface 0.
    - If two consecutive rows both list materials (back-to-back), treat as touching
      lenses by repeating the previous material on the first of the pair, then update
      the current material after.
    - Otherwise blanks inherit the most-recent 'current' material.
    """
    # Normalize materials to a list of cleaned strings, treating real NaNs as ""
    raw = df["Material"].tolist()
    materials = []
    for x in raw:
        if pd.isna(x):
            materials.append("")
        else:
            materials.append(str(x).strip())

    fixed = []
    current = ""

    for i in range(len(materials)):
        mat = materials[i]

        # Case: empty cell
        if mat == "":
            # Lookahead special-case: first surface blank but next has material -> copy next
            if i == 0 and len(materials) > 1 and materials[1] != "":
                mat = materials[1]
                fixed.append(mat)
                current = mat
                continue

            # Normal blank: inherit the current material (may be "" if none seen yet)
            fixed.append(current)
            continue

        # Non-empty material listed
        if i > 0:
            prev_mat = materials[i - 1]
            if prev_mat != "":
                # Back-to-back material listing → touching lenses
                # RULE: repeat previous material instead of switching immediately
                fixed.append(fixed[-1])  # repeat previous material
                current = mat            # update current after assigning
                continue

        # Normal case: materials appear every other surface
        current = mat
        fixed.append(mat)

    df["Material"] = fixed
    return df

def process_all_csvs():
    csv_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]

    if not csv_files:
        print("⚠️ No CSV files found.")
        return

    for csv_name in csv_files:
        in_path = os.path.join(input_folder, csv_name)
        out_path = os.path.join(output_folder, csv_name)

        df = pd.read_csv(in_path)
        df_fixed = fix_material_column(df)

        df_fixed.to_csv(out_path, index=False)
        print(f"✅ Fixed: {csv_name} → saved to Materials")

    print("\n🎉 All materials corrected and saved!")


# Run processing
process_all_csvs()
