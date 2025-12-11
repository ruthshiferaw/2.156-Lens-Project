#!/usr/bin/env python3
"""
predict_from_surface_csv.py

Usage:
    python predict_from_surface_csv.py /path/to/lens_surface_file.csv

Assumptions:
 - artifacts/ contains:
     - best_surface_transformer.pth
     - target_scaler.joblib (optional)
     - radius_scaler.joblib
     - thickness_scaler.joblib
     - semid_scaler.joblib
     - global_scaler.joblib
     - mat_encoder.joblib
 - The per-surface CSV file has columns: "Radius", "Thickness", "SemiDiameter", "Material"
 - The filename (without extension) contains the globals as the last three underscore-separated tokens:
     e.g. SOME_NAME_12.34_2.8_50.0.csv  -> diag=12.34, fnum=2.8, efl=50.0
"""
import os
import sys
import io
import argparse
import json

import numpy as np
import pandas as pd
import joblib

import torch
import torch.nn as nn

# --------- constants (match training) ----------
BATCH_SIZE = 1
EMBED_DIM = 256
NUM_HEADS = 16
NUM_LAYERS = 8
MLP_HIDDEN = 256
DROPOUT = 0.1
MAX_SURFACES = 60
TARGET_COLS = [
    "Tan Shift", "Sag Shift",
    "Long_0.4861", "Long_0.5876", "Long_0.6563",
    "Poly",
    "RMS_0.4861", "RMS_0.5876", "RMS_0.6563",
    "Effective F/#"
]

DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# --------- helper: tolerant CSV read ----------
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

# --------- model class (must match training) ----------
class SurfaceTransformerModel(nn.Module):
    def __init__(self, n_materials, embed_dim=EMBED_DIM, num_heads=NUM_HEADS,
                 num_layers=NUM_LAYERS, mlp_hidden=MLP_HIDDEN, dropout=DROPOUT, num_targets=len(TARGET_COLS)):
        super().__init__()
        self.embed_dim = embed_dim
        self.num_proj = nn.Linear(3, embed_dim//2)
        self.mat_emb = nn.Embedding(n_materials, embed_dim//2)
        self.combine_ln = nn.LayerNorm(embed_dim)
        self.cls_token = nn.Parameter(torch.randn(1, 1, embed_dim))
        self.pos_emb = nn.Parameter(torch.randn(1, MAX_SURFACES+1, embed_dim))
        encoder_layer = nn.TransformerEncoderLayer(d_model=embed_dim, nhead=num_heads,
                                                   dim_feedforward=embed_dim*4, dropout=dropout, activation='gelu')
        self.transformer = nn.TransformerEncoder(encoder_layer, num_layers=num_layers)
        self.global_proj = nn.Linear(3, embed_dim//2)
        self.head = nn.Sequential(
            nn.LayerNorm(embed_dim + embed_dim//2),
            nn.Linear(embed_dim + embed_dim//2, mlp_hidden),
            nn.GELU(),
            nn.Dropout(dropout),
            nn.Linear(mlp_hidden, mlp_hidden//2),
            nn.GELU(),
            nn.Dropout(dropout),
            nn.Linear(mlp_hidden//2, num_targets)
        )

    def forward(self, radius, thickness, semi, mat_idx, mask, globals_in):
        B, L = radius.shape
        num_feats = torch.stack([radius, thickness, semi], dim=-1)   # (B,L,3)
        num_proj = self.num_proj(num_feats)                         # (B,L,embed_dim//2)
        mat_e = self.mat_emb(mat_idx)                               # (B,L,embed_dim//2)
        token_emb = torch.cat([num_proj, mat_e], dim=-1)            # (B,L,embed_dim)
        token_emb = self.combine_ln(token_emb)
        cls_tokens = self.cls_token.expand(B, -1, -1)
        x = torch.cat([cls_tokens, token_emb], dim=1)               # (B, L+1, embed_dim)
        pos = self.pos_emb[:, :L+1, :]
        x = x + pos
        x = x.transpose(0,1)   # transformer expects (S, B, E)
        src_key_padding_mask = ~mask
        cls_col = torch.zeros((B,1), dtype=torch.bool, device=mask.device)
        src_key_padding_mask = torch.cat([cls_col, src_key_padding_mask], dim=1)
        out = self.transformer(x, src_key_padding_mask=src_key_padding_mask)
        out = out.transpose(0,1)
        cls_out = out[:,0,:]                # (B, embed_dim)
        g_proj = self.global_proj(globals_in)  # (B, embed_dim//2)
        combined = torch.cat([cls_out, g_proj], dim=-1)
        preds = self.head(combined)
        return preds

# --------- parse globals from filename ----------
def parse_globals_from_filename(path):
    base = os.path.splitext(os.path.basename(path))[0]
    parts = base.split("_")
    if len(parts) >= 3:
        # last three parts are diag, f#, efl
        try:
            diag = float(parts[-3])
            fnum = float(parts[-2])
            efl = float(parts[-1])
            return np.array([diag, fnum, efl], dtype=float)
        except Exception:
            pass
    # fallback zeros
    return np.array([0.0, 0.0, 0.0], dtype=float)

# --------- prepare input tensors ----------
def prepare_tensors(df, mat_encoder, radius_scaler, thickness_scaler, semid_scaler, global_scaler):
    # required columns
    for c in ["Radius", "Thickness", "SemiDiameter", "Material"]:
        if c not in df.columns:
            raise RuntimeError(f"Input CSV missing required column: {c}")

    r_raw = pd.to_numeric(df["Radius"], errors="raise").astype(float).values
    t_raw = pd.to_numeric(df["Thickness"], errors="raise").astype(float).values
    s_raw = pd.to_numeric(df["SemiDiameter"], errors="raise").astype(float).values
    mats = df["Material"].astype(str).fillna("None").values

    # map unknown materials -> "None" if present, otherwise map to 0 (first class)
    classes = list(mat_encoder.classes_)
    mats_safe = []
    for m in mats:
        if m in mat_encoder.classes_:
            mats_safe.append(m)
        elif "None" in mat_encoder.classes_:
            mats_safe.append("None")
        else:
            # map to first class
            mats_safe.append(mat_encoder.classes_[0])
    mats_idx = mat_encoder.transform(np.array(mats_safe)).astype(np.int64)

    def _scale(arr, scaler):
        arr2 = np.array(arr, dtype=float)
        if scaler is None:
            return arr2.astype(np.float32)
        try:
            lo = float(scaler.data_min_[0]); hi = float(scaler.data_max_[0])
            arr2 = np.clip(arr2, lo, hi)
        except Exception:
            pass
        return scaler.transform(arr2.reshape(-1,1)).reshape(-1).astype(np.float32)

    r_scaled = _scale(r_raw, radius_scaler)
    t_scaled = _scale(t_raw, thickness_scaler)
    s_scaled = _scale(s_raw, semid_scaler)

    n = len(r_scaled)
    L = min(n, MAX_SURFACES)
    # pad or truncate
    radii = np.zeros((1, MAX_SURFACES), dtype=np.float32)
    thicks = np.zeros((1, MAX_SURFACES), dtype=np.float32)
    semis = np.zeros((1, MAX_SURFACES), dtype=np.float32)
    mats_arr = np.zeros((1, MAX_SURFACES), dtype=np.int64)
    mask = np.zeros((1, MAX_SURFACES), dtype=np.bool_)

    radii[0, :L] = r_scaled[:L]
    thicks[0, :L] = t_scaled[:L]
    semis[0, :L] = s_scaled[:L]
    mats_arr[0, :L] = mats_idx[:L]
    mask[0, :L] = True

    return (torch.tensor(radii, device=DEVICE),
            torch.tensor(thicks, device=DEVICE),
            torch.tensor(semis, device=DEVICE),
            torch.tensor(mats_arr, device=DEVICE),
            torch.tensor(mask, device=DEVICE))

def main():
    parser = argparse.ArgumentParser(description="Predict lens performance from a per-surface CSV using saved transformer artifacts.")
    parser.add_argument("csv", help="Path to per-surface CSV file")
    parser.add_argument("--artifacts", default="artifacts", help="Folder containing saved artifacts (default: artifacts/)")
    parser.add_argument("--save-out", default=False, type=bool, help="Save JSON output to artifacts/prediction_<basename>.json")
    args = parser.parse_args()

    csv_path = args.csv
    art_dir = args.artifacts

    if not os.path.isfile(csv_path):
        print("CSV file not found:", csv_path)
        sys.exit(1)

    # load artifacts
    def load_optional(path):
        return joblib.load(path) if os.path.exists(path) else None

    target_scaler = load_optional(os.path.join(art_dir, "target_scaler.joblib"))
    radius_scaler = load_optional(os.path.join(art_dir, "radius_scaler.joblib"))
    thickness_scaler = load_optional(os.path.join(art_dir, "thickness_scaler.joblib"))
    semid_scaler = load_optional(os.path.join(art_dir, "semid_scaler.joblib"))
    global_scaler = load_optional(os.path.join(art_dir, "global_scaler.joblib"))
    mat_encoder = load_optional(os.path.join(art_dir, "mat_encoder.joblib"))

    if mat_encoder is None:
        print("mat_encoder.joblib not found in", art_dir)
        sys.exit(1)

    # load model weights
    model_path = os.path.join(art_dir, "best_surface_transformer.pth")
    if not os.path.exists(model_path):
        # fallback name used in training script
        alt = "best_surface_transformer_no_impute.pth"
        if os.path.exists(alt):
            model_path = alt
        else:
            print("Model weights not found at", model_path)
            sys.exit(1)

    n_materials = len(mat_encoder.classes_) if hasattr(mat_encoder, "classes_") else 1
    model = SurfaceTransformerModel(n_materials=n_materials).to(DEVICE)
    state = torch.load(model_path, map_location=DEVICE)
    try:
        model.load_state_dict(state)
    except Exception:
        # sometimes state_dict is nested under 'model' or similar; try to be flexible
        if isinstance(state, dict) and "state_dict" in state:
            model.load_state_dict(state["state_dict"])
        else:
            model.load_state_dict(state)

    model.eval()

    # read csv
    df = read_csv_tolerant(csv_path)

    # parse globals from filename
    globals_raw = parse_globals_from_filename(csv_path).reshape(1,-1)
    if global_scaler is not None:
        try:
            globals_scaled = global_scaler.transform(globals_raw).astype(np.float32).reshape(1,-1)
        except Exception:
            globals_scaled = globals_raw.astype(np.float32).reshape(1,-1)
    else:
        globals_scaled = globals_raw.astype(np.float32).reshape(1,-1)

    # prepare input tensors
    r_t, t_t, s_t, mats_t, mask_t = prepare_tensors(df, mat_encoder,
                                                   radius_scaler, thickness_scaler, semid_scaler, global_scaler)

    globals_t = torch.tensor(globals_scaled, device=DEVICE)

    # forward
    with torch.no_grad():
        preds = model(r_t, t_t, s_t, mats_t, mask_t, globals_t)   # shape (1, num_targets)
        preds_np = preds.cpu().numpy().reshape(-1)

    # inverse transform if scaler available
    if target_scaler is not None:
        try:
            preds_orig = target_scaler.inverse_transform(preds_np.reshape(1,-1)).reshape(-1)
        except Exception:
            preds_orig = preds_np.copy()
    else:
        preds_orig = preds_np.copy()

    # Build output dict
    out = {k: float(v) for k,v in zip(TARGET_COLS, preds_orig)}
    out_meta = {
        "input_csv": os.path.abspath(csv_path),
        "globals_parsed": globals_raw.flatten().tolist(),
        "artifacts_dir": os.path.abspath(art_dir)
    }
    result = {"meta": out_meta, "predictions": out}

    # print results nicely
    print("Predictions for:", csv_path)
    for k,v in out.items():
        print(f"  {k}: {v:.6g}")

    # save as json
    if args.save_out:
        os.makedirs(art_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(csv_path))[0]
        out_path = os.path.join(art_dir, f"prediction_{base}.json")
        with open(out_path, "w") as fh:
            json.dump(result, fh, indent=2)
        print("Saved prediction JSON to", out_path)

if __name__ == "__main__":
    main()


#WITH NORMALIZATION:

# #!/usr/bin/env python3
# """
# PREDICT.py - prediction script that unnormalizes outputs to original units.

# Usage:
#   python "PREDICT.py" "<path/to/per-surface.csv>" [--artifacts PATH] [--summary-original PATH]

# Notes:
#  - Primary method to unnormalize uses the original summary CSV (default:
#      Prime Lenses + Data/CSVExports/file_lens_summary.csv)
#    The script computes per-target min/max there and inverts normalized [0,1] -> original.
#  - If that file / columns aren't available, it will try to load artifacts/target_scaler.joblib
#    and use its inverse_transform.
#  - If neither are available, it will return raw model outputs and warn.
# """
# import os
# import sys
# import io
# import argparse
# import json
# import traceback

# import numpy as np
# import pandas as pd
# import joblib
# import torch
# import torch.nn as nn

# # ---------------- constants (match training) ----------------
# EMBED_DIM = 256
# NUM_HEADS = 16
# NUM_LAYERS = 8
# MLP_HIDDEN = 256
# DROPOUT = 0.1
# MAX_SURFACES = 60
# TARGET_COLS = [
#     "Tan Shift", "Sag Shift",
#     "Long_0.4861", "Long_0.5876", "Long_0.6563",
#     "Poly",
#     "RMS_0.4861", "RMS_0.5876", "RMS_0.6563",
#     "Effective F/#"
# ]
# # default path to the original (non-normalized) summary CSV used by robust_normalize_summary.py
# DEFAULT_SUMMARY_ORIGINAL = r"Prime Lenses + Data/CSVExports/file_lens_summary.csv"

# DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# # ---------------- tolerant CSV read ----------------
# def read_csv_tolerant(path):
#     encs = ["utf-8", "utf-8-sig", "utf-16", "latin1", "cp1252"]
#     for enc in encs:
#         try:
#             return pd.read_csv(path, encoding=enc)
#         except Exception:
#             continue
#     with open(path, "rb") as fh:
#         text = fh.read().decode("utf-8", errors="replace")
#     return pd.read_csv(io.StringIO(text))

# # ---------------- model class (must match training) ----------------
# class SurfaceTransformerModel(nn.Module):
#     def __init__(self, n_materials, embed_dim=EMBED_DIM, num_heads=NUM_HEADS,
#                  num_layers=NUM_LAYERS, mlp_hidden=MLP_HIDDEN, dropout=DROPOUT, num_targets=len(TARGET_COLS)):
#         super().__init__()
#         self.embed_dim = embed_dim
#         self.num_proj = nn.Linear(3, embed_dim//2)
#         self.mat_emb = nn.Embedding(n_materials, embed_dim//2)
#         self.combine_ln = nn.LayerNorm(embed_dim)
#         self.cls_token = nn.Parameter(torch.randn(1, 1, embed_dim))
#         self.pos_emb = nn.Parameter(torch.randn(1, MAX_SURFACES+1, embed_dim))
#         encoder_layer = nn.TransformerEncoderLayer(d_model=embed_dim, nhead=num_heads,
#                                                    dim_feedforward=embed_dim*4, dropout=dropout, activation='gelu')
#         self.transformer = nn.TransformerEncoder(encoder_layer, num_layers=num_layers)
#         self.global_proj = nn.Linear(3, embed_dim//2)
#         self.head = nn.Sequential(
#             nn.LayerNorm(embed_dim + embed_dim//2),
#             nn.Linear(embed_dim + embed_dim//2, mlp_hidden),
#             nn.GELU(),
#             nn.Dropout(dropout),
#             nn.Linear(mlp_hidden, mlp_hidden//2),
#             nn.GELU(),
#             nn.Dropout(dropout),
#             nn.Linear(mlp_hidden//2, num_targets)
#         )

#     def forward(self, radius, thickness, semi, mat_idx, mask, globals_in):
#         B, L = radius.shape
#         num_feats = torch.stack([radius, thickness, semi], dim=-1)   # (B,L,3)
#         num_proj = self.num_proj(num_feats)                         # (B,L,embed_dim//2)
#         mat_e = self.mat_emb(mat_idx)                               # (B,L,embed_dim//2)
#         token_emb = torch.cat([num_proj, mat_e], dim=-1)            # (B,L,embed_dim)
#         token_emb = self.combine_ln(token_emb)
#         cls_tokens = self.cls_token.expand(B, -1, -1)
#         x = torch.cat([cls_tokens, token_emb], dim=1)               # (B, L+1, embed_dim)
#         pos = self.pos_emb[:, :L+1, :]
#         x = x + pos
#         x = x.transpose(0,1)   # transformer expects (S, B, E)
#         src_key_padding_mask = ~mask
#         cls_col = torch.zeros((B,1), dtype=torch.bool, device=mask.device)
#         src_key_padding_mask = torch.cat([cls_col, src_key_padding_mask], dim=1)
#         out = self.transformer(x, src_key_padding_mask=src_key_padding_mask)
#         out = out.transpose(0,1)
#         cls_out = out[:,0,:]
#         g_proj = self.global_proj(globals_in)
#         combined = torch.cat([cls_out, g_proj], dim=-1)
#         preds = self.head(combined)
#         return preds

# # ---------------- parse globals from filename ----------------
# def parse_globals_from_filename(path):
#     base = os.path.splitext(os.path.basename(path))[0]
#     parts = base.split("_")
#     if len(parts) >= 3:
#         try:
#             return np.array([float(parts[-3]), float(parts[-2]), float(parts[-1])], dtype=float)
#         except Exception:
#             pass
#     return np.array([0.0, 0.0, 0.0], dtype=float)

# # ---------------- prepare tensors ----------------
# def prepare_tensors(df, mat_encoder, radius_scaler, thickness_scaler, semid_scaler):
#     for c in ["Radius","Thickness","SemiDiameter","Material"]:
#         if c not in df.columns:
#             raise RuntimeError(f"Input CSV missing required column: {c}")
#     r_raw = pd.to_numeric(df["Radius"], errors="raise").astype(float).values
#     t_raw = pd.to_numeric(df["Thickness"], errors="raise").astype(float).values
#     s_raw = pd.to_numeric(df["SemiDiameter"], errors="raise").astype(float).values
#     mats = df["Material"].astype(str).fillna("None").values

#     # safe mapping of materials
#     mats_safe = []
#     for m in mats:
#         if m in mat_encoder.classes_:
#             mats_safe.append(m)
#         elif "None" in mat_encoder.classes_:
#             mats_safe.append("None")
#         else:
#             mats_safe.append(mat_encoder.classes_[0])
#     mats_idx = mat_encoder.transform(np.array(mats_safe)).astype(np.int64)

#     def _scale(arr, scaler):
#         arr2 = np.array(arr, dtype=float)
#         if scaler is None:
#             return arr2.astype(np.float32)
#         try:
#             lo = float(scaler.data_min_[0]); hi = float(scaler.data_max_[0])
#             arr2 = np.clip(arr2, lo, hi)
#         except Exception:
#             pass
#         return scaler.transform(arr2.reshape(-1,1)).reshape(-1).astype(np.float32)

#     r_scaled = _scale(r_raw, radius_scaler)
#     t_scaled = _scale(t_raw, thickness_scaler)
#     s_scaled = _scale(s_raw, semid_scaler)

#     n = len(r_scaled)
#     L = min(n, MAX_SURFACES)
#     radii = np.zeros((1, MAX_SURFACES), dtype=np.float32)
#     thicks = np.zeros((1, MAX_SURFACES), dtype=np.float32)
#     semis = np.zeros((1, MAX_SURFACES), dtype=np.float32)
#     mats_arr = np.zeros((1, MAX_SURFACES), dtype=np.int64)
#     mask = np.zeros((1, MAX_SURFACES), dtype=np.bool_)

#     radii[0,:L] = r_scaled[:L]
#     thicks[0,:L] = t_scaled[:L]
#     semis[0,:L] = s_scaled[:L]
#     mats_arr[0,:L] = mats_idx[:L]
#     mask[0,:L] = True

#     return (torch.tensor(radii, device=DEVICE),
#             torch.tensor(thicks, device=DEVICE),
#             torch.tensor(semis, device=DEVICE),
#             torch.tensor(mats_arr, device=DEVICE),
#             torch.tensor(mask, device=DEVICE))

# # ---------------- helper to compute original min/max for targets ----------------
# def compute_target_min_max_from_original_summary(summary_path, targets):
#     """
#     Return dict: {target: (min, max)} computed from summary_path.
#     Only uses finite numeric values; if a column missing or no finite values, it's not included.
#     """
#     if not os.path.exists(summary_path):
#         return {}
#     try:
#         df = read_csv_tolerant(summary_path)
#     except Exception:
#         return {}
#     mm = {}
#     for t in targets:
#         if t in df.columns:
#             vals = pd.to_numeric(df[t], errors='coerce').values
#             finite = vals[np.isfinite(vals)]
#             if finite.size > 0:
#                 mm[t] = (float(np.nanmin(finite)), float(np.nanmax(finite)))
#     return mm

# # ---------------- main ----------------
# def main():
#     parser = argparse.ArgumentParser(description="Predict lens performance and unnormalize outputs to original units.")
#     parser.add_argument("csv", help="per-surface csv path")
#     parser.add_argument("--artifacts", default="artifacts", help="artifacts folder (default: artifacts/)")
#     parser.add_argument("--summary-original", default=DEFAULT_SUMMARY_ORIGINAL,
#                         help=f"path to original (non-normalized) summary CSV (default: {DEFAULT_SUMMARY_ORIGINAL})")
#     parser.add_argument("--no-save", action="store_true", help="do not write JSON output")
#     args = parser.parse_args()

#     csv_path = args.csv
#     art_dir = args.artifacts
#     summary_original = args.summary_original

#     try:
#         if not os.path.isfile(csv_path):
#             raise FileNotFoundError(f"Input CSV not found: {csv_path}")

#         # load artifacts (optional)
#         def load_opt(fname):
#             p = os.path.join(art_dir, fname)
#             return joblib.load(p) if os.path.exists(p) else None

#         print("Loading artifacts from:", os.path.abspath(art_dir))
#         target_scaler = load_opt("target_scaler.joblib")
#         radius_scaler = load_opt("radius_scaler.joblib")
#         thickness_scaler = load_opt("thickness_scaler.joblib")
#         semid_scaler = load_opt("semid_scaler.joblib")
#         global_scaler = load_opt("global_scaler.joblib")
#         mat_encoder = load_opt("mat_encoder.joblib")

#         print("Artifacts status:")
#         for k,v in [("target_scaler", target_scaler), ("radius_scaler", radius_scaler),
#                     ("thickness_scaler", thickness_scaler), ("semid_scaler", semid_scaler),
#                     ("global_scaler", global_scaler), ("mat_encoder", mat_encoder)]:
#             print(f"  {k}: {'found' if v is not None else 'MISSING'}")

#         if mat_encoder is None:
#             raise RuntimeError("mat_encoder.joblib missing - cannot proceed")

#         # locate model file
#         candidates = [
#             os.path.join(art_dir, "best_surface_transformer.pth"),
#             os.path.join(art_dir, "best_surface_transformer_no_impute.pth"),
#             "best_surface_transformer.pth",
#             "best_surface_transformer_no_impute.pth"
#         ]
#         model_path = next((p for p in candidates if p and os.path.exists(p)), None)
#         if model_path is None:
#             raise RuntimeError("Model weights not found. Searched: " + ", ".join(candidates))

#         print("Loading model weights from:", model_path)
#         n_materials = len(mat_encoder.classes_) if hasattr(mat_encoder, "classes_") else 1
#         model = SurfaceTransformerModel(n_materials=n_materials).to(DEVICE)
#         state = torch.load(model_path, map_location=DEVICE)
#         try:
#             model.load_state_dict(state)
#         except Exception:
#             if isinstance(state, dict) and "state_dict" in state:
#                 model.load_state_dict(state["state_dict"])
#             else:
#                 model.load_state_dict(state)
#         model.eval()

#         # read per-surface csv
#         print("Reading input CSV:", csv_path)
#         df = read_csv_tolerant(csv_path)
#         print("Rows in input CSV:", len(df))

#         # build globals and scale with global_scaler if available
#         globals_raw = parse_globals_from_filename(csv_path).reshape(1,-1)
#         if global_scaler is not None:
#             try:
#                 globals_scaled = global_scaler.transform(globals_raw).astype(np.float32).reshape(1,-1)
#             except Exception:
#                 globals_scaled = globals_raw.astype(np.float32).reshape(1,-1)
#         else:
#             globals_scaled = globals_raw.astype(np.float32).reshape(1,-1)
#         globals_t = torch.tensor(globals_scaled, device=DEVICE)

#         # prepare tokens
#         r_t, t_t, s_t, mats_t, mask_t = prepare_tensors(df, mat_encoder, radius_scaler, thickness_scaler, semid_scaler)

#         # forward pass
#         with torch.no_grad():
#             preds = model(r_t, t_t, s_t, mats_t, mask_t, globals_t)  # shape (1, num_targets)
#             preds_np = preds.cpu().numpy().reshape(-1)

#         # Now unnormalize preds:
#         # Primary method: compute per-target min/max from original (non-normalized) summary csv.
#         target_minmax = compute_target_min_max_from_original_summary(summary_original, TARGET_COLS)
#         preds_un = None

#         if len(target_minmax) == len(TARGET_COLS):
#             # all targets have min/max -> invert min-max
#             preds_un = np.zeros_like(preds_np, dtype=float)
#             for i, tname in enumerate(TARGET_COLS):
#                 mn, mx = target_minmax[tname]
#                 # assume model output is normalized in [0,1]; clip just in case
#                 v_norm = float(np.clip(preds_np[i], 0.0, 1.0))
#                 preds_un[i] = v_norm * (mx - mn) + mn
#             print("Unnormalized predictions via original summary min/max (preferred).")
#         else:
#             # partial or missing original summary info -> fallback to target_scaler (if present)
#             if target_scaler is not None:
#                 try:
#                     preds_un = target_scaler.inverse_transform(preds_np.reshape(1,-1)).reshape(-1)
#                     print("Unnormalized predictions via target_scaler.joblib (fallback).")
#                 except Exception:
#                     preds_un = None

#         if preds_un is None:
#             # last-resort: return raw model outputs and warn
#             print("WARNING: Could not unnormalize predictions (missing original summary columns and target_scaler).")
#             preds_un = preds_np.copy()

#         # Build output mapping
#         out = {k: float(v) for k,v in zip(TARGET_COLS, preds_un)}
#         out_meta = {
#             "input_csv": os.path.abspath(csv_path),
#             "globals_parsed": globals_raw.flatten().tolist(),
#             "artifacts_dir": os.path.abspath(art_dir),
#             "used_unnormalization_method": ("original_summary_minmax" if len(target_minmax)==len(TARGET_COLS)
#                                            else ("target_scaler" if target_scaler is not None else "none"))
#         }
#         result = {"meta": out_meta, "predictions": out}

#         # print
#         print("\nPredictions (in original units):")
#         for k,v in out.items():
#             print(f"  {k}: {v:.6g}")

#         # save
#         if not args.no_save:
#             os.makedirs(art_dir, exist_ok=True)
#             base = os.path.splitext(os.path.basename(csv_path))[0]
#             out_path = os.path.join(art_dir, f"prediction_{base}.json")
#             with open(out_path, "w") as fh:
#                 json.dump(result, fh, indent=2)
#             print("\nSaved JSON ->", out_path)

#     except Exception as e:
#         print("\nERROR during prediction:")
#         print(str(e))
#         print("\nTraceback (full):")
#         traceback.print_exc()
#         sys.exit(1)

# if __name__ == "__main__":
#     main()
