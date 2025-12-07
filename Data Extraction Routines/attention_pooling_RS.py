#!/usr/bin/env python3
"""
attention_pooling.py

Deterministic attention pooling over per-surface tokens (no NN).
Writes artifacts into artifacts_attention/ (does NOT overwrite transformer artifacts).
"""
import os
import glob
import io
import random

import numpy as np
import pandas as pd
import joblib

from sklearn.preprocessing import MinMaxScaler, LabelEncoder
from sklearn.model_selection import train_test_split

# ---------------------------
# CONFIG
# ---------------------------
CLEANED_SUMMARY_CSV = r"Prime Lenses + Data/CSVExports/file_lens_summary_normalized_imputed_mean.csv"
MATERIALS_CLEANED_FOLDER = r"Prime Lenses + Data/LensDataExportsRenamedMaterialsCleaned"

OUT_DIR = "artifacts"
os.makedirs(OUT_DIR, exist_ok=True)

RANDOM_SEED = 42
TEST_SIZE = 0.15

TARGET_COLS = [
    "Tan Shift", "Sag Shift",
    "Long_0.4861", "Long_0.5876", "Long_0.6563",
    "Poly",
    "RMS_0.4861", "RMS_0.5876", "RMS_0.6563",
    "Rel. Ill", "Effective F/#"
]

SURFACE_NUM_COLS = ["Radius", "Thickness", "SemiDiameter"]
MAT_COL = "Material"

# ---------------------------
# Helpers
# ---------------------------
def read_csv_tolerant(path):
    encs = ["utf-8", "utf-8-sig", "utf-16", "latin1", "cp1252"]
    for enc in encs:
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            continue
    with open(path, "rb") as fh:
        text = fh.read().decode("utf-8", errors="replace")
    return pd.read_csv(io.StringIO(text))

def softmax_1d(x):
    # numerically stable softmax for 1D array
    xm = x - np.max(x)
    e = np.exp(xm)
    s = e / (np.sum(e) + 1e-12)
    return s

# ---------------------------
# 1) load summary
# ---------------------------
if not os.path.isfile(CLEANED_SUMMARY_CSV):
    raise SystemExit(f"Summary CSV not found: {CLEANED_SUMMARY_CSV}")

summary = read_csv_tolerant(CLEANED_SUMMARY_CSV)
print("Loaded summary:", summary.shape)

missing_targets = [c for c in TARGET_COLS if c not in summary.columns]
if missing_targets:
    raise SystemExit(f"Missing target columns: {missing_targets}")

# Ensure File Name present and extract list of allowed bases
if "File Name" not in summary.columns:
    raise SystemExit("Summary missing 'File Name' column")

allowed_bases = [str(fn).strip() for fn in summary["File Name"].astype(str).values]
allowed_bases_set = set(allowed_bases)

# Build lens_info dict with targets and parsed globals
lens_info = {}
for _, row in summary.iterrows():
    fname = str(row["File Name"]).strip()
    parts = fname.split("_")
    try:
        diag = float(parts[-3]); fnum = float(parts[-2]); efl = float(parts[-1])
    except Exception:
        diag, fnum, efl = 0.0, 0.0, 0.0
    targets = row[TARGET_COLS].values.astype(float)
    lens_info[fname] = {"targets": targets, "globals": np.array([diag, fnum, efl], dtype=float)}

# ---------------------------
# 2) collect per-surface files that match the summary
# ---------------------------
surface_files = sorted(glob.glob(os.path.join(MATERIALS_CLEANED_FOLDER, "*.csv")))
if len(surface_files) == 0:
    raise SystemExit(f"No surface CSVs found in {MATERIALS_CLEANED_FOLDER}")

all_surface_files_by_base = {os.path.splitext(os.path.basename(p))[0]: p for p in surface_files}
# only keep those bases that are in summary
surface_files_by_name = {b: p for b, p in all_surface_files_by_base.items() if b in allowed_bases_set}
print(f"Found {len(surface_files)} files; {len(surface_files_by_name)} match summary entries")

missing = sorted(set(allowed_bases_set) - set(surface_files_by_name.keys()))
if missing:
    print(f"Warning: {len(missing)} summary File Name entries have no materials CSV. Example: {missing[:6]}")

# ---------------------------
# 3) Fit scalers using the available (matched) per-surface files
# ---------------------------
num_radius = []
num_thickness = []
num_semid = []
material_set = set()
for base, path in surface_files_by_name.items():
    df = read_csv_tolerant(path)
    for c in SURFACE_NUM_COLS:
        if c in df.columns:
            vals = pd.to_numeric(df[c], errors="coerce").replace([np.inf, -np.inf], np.nan).dropna().astype(float).tolist()
            if vals:
                if c == "Radius": num_radius.extend(vals)
                elif c == "Thickness": num_thickness.extend(vals)
                else: num_semid.extend(vals)
    if MAT_COL in df.columns:
        material_set.update(df[MAT_COL].astype(str).fillna("None").unique())

# fit scalers (fallback to dummy if empty)
radius_scaler = MinMaxScaler().fit(np.array(num_radius).reshape(-1,1)) if len(num_radius)>0 else None
thickness_scaler = MinMaxScaler().fit(np.array(num_thickness).reshape(-1,1)) if len(num_thickness)>0 else None
semid_scaler = MinMaxScaler().fit(np.array(num_semid).reshape(-1,1)) if len(num_semid)>0 else None

global_list = [info["globals"].astype(float) for info in lens_info.values() if info is not None]
global_scaler = MinMaxScaler().fit(np.vstack(global_list)) if len(global_list) > 0 else None

# material encoder (deterministic numeric feature)
material_list = sorted(list(material_set)) if material_set else ["None"]
mat_encoder = LabelEncoder()
mat_encoder.fit(material_list)

# save scalers & encoders (attention-specific)
joblib.dump({
    "radius_scaler": radius_scaler,
    "thickness_scaler": thickness_scaler,
    "semid_scaler": semid_scaler,
    "global_scaler": global_scaler,
    "mat_encoder": mat_encoder
}, os.path.join(OUT_DIR, "scalers_attention.joblib"))
print("Saved attention scalers to", os.path.join(OUT_DIR, "scalers_attention.joblib"))

# ---------------------------
# 4) Build per-lens pooled vectors using deterministic attention
# ---------------------------
# We'll create token vectors: [r_scaled, t_scaled, s_scaled, mat_norm]
# Attention logits = token_vector dot dataset_mean_query (mean token vector across all tokens)
# Attention weights = softmax(logits) per-lens
pooled_list = []
targets_list = []
fnames_list = []
all_token_vectors = []  # for computing dataset query

# first pass: gather token vectors per-lens, store temporarily
per_lens_tokens = {}
for fname in allowed_bases:
    if fname not in surface_files_by_name:
        # no per-surface csv for this summary entry; skip
        continue
    path = surface_files_by_name[fname]
    df = read_csv_tolerant(path)

    # coerce numeric columns
    r = pd.to_numeric(df["Radius"], errors="coerce").replace([np.inf,-np.inf], np.nan).fillna(0.0).astype(float).values if "Radius" in df.columns else np.zeros((0,))
    t = pd.to_numeric(df["Thickness"], errors="coerce").replace([np.inf,-np.inf], np.nan).fillna(0.0).astype(float).values if "Thickness" in df.columns else np.zeros((0,))
    s = pd.to_numeric(df["SemiDiameter"], errors="coerce").replace([np.inf,-np.inf], np.nan).fillna(0.0).astype(float).values if "SemiDiameter" in df.columns else np.zeros((0,))

    # scale (if scalers exist)
    def _scale_vec(arr, scaler):
        if scaler is None or arr.size==0:
            return arr.astype(float)
        arr2 = np.array(arr, dtype=float)
        # clip to scaler range to be robust
        try:
            lo = float(scaler.data_min_[0]); hi = float(scaler.data_max_[0])
            arr2 = np.clip(arr2, lo, hi)
        except Exception:
            pass
        return scaler.transform(arr2.reshape(-1,1)).reshape(-1)

    r_s = _scale_vec(r, radius_scaler)
    t_s = _scale_vec(t, thickness_scaler)
    s_s = _scale_vec(s, semid_scaler)

    # material numeric feature
    if MAT_COL in df.columns:
        mats = df[MAT_COL].astype(str).fillna("None").values
    else:
        mats = np.array(["None"] * max(1, len(r_s)))
    mats_idx = mat_encoder.transform(np.array([m if m in mat_encoder.classes_ else "None" for m in mats], dtype=object))
    # normalize mat index to [0,1]
    if len(mat_encoder.classes_) > 1:
        mats_norm = mats_idx.astype(float) / (len(mat_encoder.classes_)-1)
    else:
        mats_norm = mats_idx.astype(float) * 0.0

    # make token vectors (N,4). If token lengths mismatch, align to min length
    L = min(len(r_s), len(t_s), len(s_s), len(mats_norm))
    if L == 0:
        # skip lenses with no valid surface tokens
        continue
    toks = np.stack([r_s[:L], t_s[:L], s_s[:L], mats_norm[:L]], axis=1)
    per_lens_tokens[fname] = toks
    all_token_vectors.append(toks.reshape(-1, toks.shape[-1]))

# concat all token vectors to compute dataset mean query
if len(all_token_vectors) == 0:
    raise SystemExit("No token vectors found (no matched per-surface CSVs with numeric tokens). Aborting.")
all_tok_cat = np.vstack(all_token_vectors)
query_vec = np.mean(all_tok_cat, axis=0)  # deterministic query
# if query is zero vector (rare), fallback to uniform query
if np.allclose(query_vec, 0.0):
    query_vec = np.ones_like(query_vec)

# now compute pooled vector per lens
for fname, toks in per_lens_tokens.items():
    # compute logits as dot(toks, query)
    logits = toks.dot(query_vec)
    weights = softmax_1d(logits)
    pooled = (weights[:, None] * toks).sum(axis=0)
    pooled_list.append(pooled.astype(float))
    targets_list.append(lens_info[fname]["targets"].astype(float))
    fnames_list.append(fname)

pooled_arr = np.vstack(pooled_list)        # N x D_pooled (D_pooled == 4)
targets_arr = np.vstack(targets_list)      # N x T
fnames_arr = np.array(fnames_list, dtype=object)

print("Computed pooled vectors for", pooled_arr.shape[0], "lenses. pooled dim:", pooled_arr.shape[1])

# ---------------------------
# 5) optional: add global features to pooled vector (diag,f#,efl scaled)
# ---------------------------
# We'll append the scaled globals to pooled vector so downstream code can use both.
globals_scaled_list = []
for fname in fnames_list:
    g = lens_info[fname]["globals"].astype(float).reshape(1,-1)
    if global_scaler is not None:
        g_s = global_scaler.transform(g).reshape(-1)
    else:
        g_s = g.reshape(-1)
    globals_scaled_list.append(g_s)
globals_scaled_arr = np.vstack(globals_scaled_list)
# concat to pooled features
pooled_plus_globals = np.concatenate([pooled_arr, globals_scaled_arr], axis=1)  # N x (4+3)

# ---------------------------
# 6) split train/val using same deterministic split
# ---------------------------
train_idx, val_idx = train_test_split(np.arange(len(fnames_arr)), test_size=TEST_SIZE, random_state=RANDOM_SEED, shuffle=True)

pooled_train = pooled_plus_globals[train_idx]
targets_train = targets_arr[train_idx]
fnames_train = fnames_arr[train_idx]

pooled_val = pooled_plus_globals[val_idx]
targets_val = targets_arr[val_idx]
fnames_val = fnames_arr[val_idx]

# ---------------------------
# 7) save artifacts (attention-specific names)
# ---------------------------
np.savez(os.path.join(OUT_DIR, "all_pooled.npz"),
         pooled=pooled_plus_globals, targets=targets_arr, fnames=fnames_arr)

np.savez(os.path.join(OUT_DIR, "pooled_train.npz"),
         pooled=pooled_train, targets=targets_train, fnames=fnames_train)

np.savez(os.path.join(OUT_DIR, "pooled_val.npz"),
         pooled=pooled_val, targets=targets_val, fnames=fnames_val)

# Save a small metadata file too
meta = {
    "pooled_dim": pooled_plus_globals.shape[1],
    "target_cols": TARGET_COLS,
    "surface_token_cols": SURFACE_NUM_COLS + [MAT_COL],
    "n_lenses": len(fnames_arr)
}
joblib.dump(meta, os.path.join(OUT_DIR, "meta_attention.joblib"))

print("Saved attention pooling artifacts to", OUT_DIR)
