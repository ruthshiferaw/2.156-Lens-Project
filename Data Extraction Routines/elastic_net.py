#!/usr/bin/env python3
"""
transformer_no_impute_final_with_elasticnet.py

Same as your transformer_no_impute_final.py but uses Elastic Net (L1 + L2) regularization
instead of Dropout in the MLP head. The Elastic Net penalty is added to the training loss only.
"""
import os
import glob
import io
import random
from collections import defaultdict

import numpy as np
import pandas as pd

import torch
import torch.nn as nn
from torch.utils.data import Dataset, DataLoader
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import MinMaxScaler, LabelEncoder

# ---------------------------
# CONFIG / HYPERPARAMS
# ---------------------------
CLEANED_SUMMARY_CSV = r"Prime Lenses + Data/CSVExports/file_lens_summary_normalized_imputed_mean.csv"
MATERIALS_CLEANED_FOLDER = r"Prime Lenses + Data/LensDataExportsRenamedMaterialsCleaned"

BATCH_SIZE = 32
EMBED_DIM = 256
NUM_HEADS = 8
NUM_LAYERS = 8
MLP_HIDDEN = 256
# DROPOUT no longer used for regularization; keep param for compatibility but not applied
DROPOUT = 0.0
LR = 1e-4
WEIGHT_DECAY = 0.0  # set to 0 because we compute L2 manually (avoid double-counting)
NUM_EPOCHS = 60
DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")
MAX_SURFACES = 60
RANDOM_SEED = 42

# Elastic Net hyperparams (new)
ELASTIC_LAMBDA = 1e-4  # overall regularization strength — tune this
L1_RATIO = 0.5         # fraction of L1 in elastic net (0 = pure L2, 1 = pure L1)

torch.manual_seed(RANDOM_SEED)
np.random.seed(RANDOM_SEED)
random.seed(RANDOM_SEED)

TARGET_COLS = [
    "Tan Shift", "Sag Shift",
    "Long_0.4861", "Long_0.5876", "Long_0.6563",
    "Poly",
    "RMS_0.4861", "RMS_0.5876", "RMS_0.6563",
    "Effective F/#"
]

# ---------------------------
# Helpers: robust CSV read (tries several encodings)
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

def assert_finite_array(arr, name_hint="array", fname=None):
    arr = np.asarray(arr)
    if not np.isfinite(arr).all():
        where_bad = np.where(~np.isfinite(arr))[0][:8].tolist()
        extra = f" in file {fname}" if fname is not None else ""
        raise RuntimeError(f"Non-finite values detected in {name_hint}{extra}. Example indices: {where_bad}")

# ---------------------------
# 1) Load cleaned summary CSV (must be pre-imputed/no NaN/Inf)
# ---------------------------
if not os.path.isfile(CLEANED_SUMMARY_CSV):
    raise SystemExit(f"Cleaned summary CSV not found: {CLEANED_SUMMARY_CSV}\nRun the imputer/cleaner script first.")

summary = read_csv_tolerant(CLEANED_SUMMARY_CSV)
print("Loaded summary shape:", summary.shape)

# Validate target columns exist
missing_targets = [c for c in TARGET_COLS if c not in summary.columns]
if missing_targets:
    raise SystemExit(f"Missing target columns in cleaned summary CSV: {missing_targets}")

# coerce and validate numeric finiteness for target columns
summary[TARGET_COLS] = summary[TARGET_COLS].apply(pd.to_numeric, errors="raise")
nan_counts = summary[TARGET_COLS].isna().sum().to_dict()
if any(v>0 for v in nan_counts.values()):
    raise SystemExit(f"Found NaNs in target columns of cleaned summary CSV: {nan_counts}")
if np.isinf(summary[TARGET_COLS].values).any():
    raise SystemExit("Found inf/-inf in target columns of cleaned summary CSV. Please clean before running this script.")

# Ensure File Name exists
if "File Name" not in summary.columns:
    raise SystemExit("Cleaned summary missing 'File Name' column.")

# Build lens_info mapping
lens_info = {}
for _, row in summary.iterrows():
    fname = str(row["File Name"])
    parts = fname.split("_")
    try:
        diag = float(parts[-3])
        fnum = float(parts[-2])
        efl = float(parts[-1])
    except Exception:
        diag, fnum, efl = 0.0, 0.0, 0.0
    targets = row[TARGET_COLS].values.astype(float)
    lens_info[fname] = {"targets": targets, "globals": np.array([diag, fnum, efl], dtype=float), "summary_row": row}

# ---------------------------
# 2) Load cleaned material CSVs (no cleaning here) and validate
# ---------------------------
if not os.path.isdir(MATERIALS_CLEANED_FOLDER):
    raise SystemExit(f"Materials cleaned folder not found: {MATERIALS_CLEANED_FOLDER}. Create it with the cleaning script.")

surface_files = sorted(glob.glob(os.path.join(MATERIALS_CLEANED_FOLDER, "*.csv")))
if len(surface_files) == 0:
    raise SystemExit(f"No CSV files found in {MATERIALS_CLEANED_FOLDER}")

# map base -> path for all csvs in folder
all_surface_files_by_base = {os.path.splitext(os.path.basename(p))[0]: p for p in surface_files}

# build set of allowed bases from the cleaned summary 'File Name' column (exact match, no extension)
allowed_bases = set(str(fn).strip() for fn in summary["File Name"].astype(str).values)

# Filter the available surface files to only those whose base is present in allowed_bases
surface_files_by_name = {base: path for base, path in all_surface_files_by_base.items() if base in allowed_bases}

# Diagnostics: which summary entries are missing a corresponding per-surface CSV
missing_from_materials = sorted(list(allowed_bases - set(surface_files_by_name.keys())))
extra_materials = sorted(list(set(all_surface_files_by_base.keys()) - set(surface_files_by_name.keys())))

print(f"Found {len(surface_files)} CSVs in materials folder; {len(surface_files_by_name)} match entries in cleaned summary.")
if len(missing_from_materials) > 0:
    print(f"Warning: {len(missing_from_materials)} summary 'File Name' entries have no matching materials CSV. Example missing (first 10):")
    print(" ", missing_from_materials[:10])
else:
    print("All summary File Name entries have a matching materials CSV (good).")

# collect material vocab and numeric pools
material_set = set()
num_radius, num_thickness, num_semid = [], [], []

for base, path in surface_files_by_name.items():
    df = read_csv_tolerant(path)

    # required columns exist?
    for c in ["Radius", "Thickness", "SemiDiameter", "Material"]:
        if c not in df.columns:
            raise SystemExit(f"Surface file {path} missing required column '{c}'. The file should be cleaned first.")

    # coerce numeric and assert no NaN/Inf in numeric columns
    df["Radius"] = pd.to_numeric(df["Radius"], errors="raise").astype(float)
    df["Thickness"] = pd.to_numeric(df["Thickness"], errors="raise").astype(float)
    df["SemiDiameter"] = pd.to_numeric(df["SemiDiameter"], errors="raise").astype(float)

    # validate finiteness
    if not np.isfinite(df["Radius"].values).all():
        raise SystemExit(f"Non-finite Radius values found in {path}. Clean the file first.")
    if not np.isfinite(df["Thickness"].values).all():
        raise SystemExit(f"Non-finite Thickness values found in {path}. Clean the file first.")
    if not np.isfinite(df["SemiDiameter"].values).all():
        raise SystemExit(f"Non-finite SemiDiameter values found in {path}. Clean the file first.")

    # collect pools for scaler fitting
    num_radius.extend(df["Radius"].astype(float).tolist())
    num_thickness.extend(df["Thickness"].astype(float).tolist())
    num_semid.extend(df["SemiDiameter"].astype(float).tolist())

    # materials (strings)
    material_set.update(df["Material"].astype(str).fillna("None").unique())

material_set = sorted(list(material_set)) if material_set else ["None"]
mat_encoder = LabelEncoder()
mat_encoder.fit(material_set)

print("Found materials vocab size:", len(material_set))
print("Collected numeric counts -> radius:", len(num_radius),
      "thickness:", len(num_thickness), "semi:", len(num_semid))

# ---------------------------
# 3) Fit scalers on finite data (no imputation)
radius_scaler = MinMaxScaler().fit(np.array(num_radius).reshape(-1,1)) if len(num_radius)>0 else None
thickness_scaler = MinMaxScaler().fit(np.array(num_thickness).reshape(-1,1)) if len(num_thickness)>0 else None
semid_scaler = MinMaxScaler().fit(np.array(num_semid).reshape(-1,1)) if len(num_semid)>0 else None

# globals scaler: build array from lens_info globals (they should be numeric and finite)
global_list = [info["globals"].astype(float) for info in lens_info.values()]
if len(global_list) > 0:
    globals_arr = np.vstack(global_list).astype(float)
    if not np.isfinite(globals_arr).all():
        raise SystemExit("Found non-finite values in parsed global features (diag/f# / efl). Clean summary or parsing.")
    global_scaler = MinMaxScaler().fit(globals_arr)
else:
    global_scaler = None

# NOTE: do NOT fit target_scaler here to avoid leakage. We'll fit it on TRAIN only after the split.
target_scaler = None

print("Fitted token & global scalers (target scaler will be fit on TRAIN only later).")

# ---------------------------
# 4) Dataset (NO imputation here)
class LensDataset(Dataset):
    def __init__(self, lens_info, surface_files_by_name, mat_encoder,
                 radius_scaler, thickness_scaler, semid_scaler, global_scaler, target_scaler,
                 max_surfaces=MAX_SURFACES):
        self.keys = [k for k in lens_info.keys() if os.path.splitext(k)[0] in surface_files_by_name]
        self.surface_files = surface_files_by_name
        self.lens_info = lens_info
        self.mat_encoder = mat_encoder
        self.radius_scaler = radius_scaler
        self.thickness_scaler = thickness_scaler
        self.semid_scaler = semid_scaler
        self.global_scaler = global_scaler
        self.target_scaler = target_scaler
        self.max_surfaces = max_surfaces

    def __len__(self):
        return len(self.keys)

    def __getitem__(self, idx):
        fname = self.keys[idx]
        base = os.path.splitext(fname)[0]
        path = self.surface_files[base]
        df = read_csv_tolerant(path)

        # ensure numeric columns are finite (no cleaning here)
        r_raw = pd.to_numeric(df["Radius"], errors="raise").astype(float).values
        t_raw = pd.to_numeric(df["Thickness"], errors="raise").astype(float).values
        s_raw = pd.to_numeric(df["SemiDiameter"], errors="raise").astype(float).values

        if not np.isfinite(r_raw).all() or not np.isfinite(t_raw).all() or not np.isfinite(s_raw).all():
            raise RuntimeError(f"Surface file {path} contains non-finite numeric entries. Please clean it first.")

        # materials
        mats = df["Material"].astype(str).fillna("None").values
        mats = np.array([m if m in mat_encoder.classes_ else "None" for m in mats], dtype=object)
        mats_idx = mat_encoder.transform(mats).astype(np.int64)

        # scale numeric token columns using fitted scalers (scalers were fit on finite values)
        def _scale(arr, scaler):
            if scaler is None:
                return arr.astype(np.float32)
            arr2 = np.array(arr, dtype=float)
            # clip to fitted range for safety
            try:
                lo = float(scaler.data_min_[0]); hi = float(scaler.data_max_[0])
                arr2 = np.clip(arr2, lo, hi)
            except Exception:
                pass
            return scaler.transform(arr2.reshape(-1,1)).reshape(-1).astype(np.float32)

        r_scaled = _scale(r_raw, self.radius_scaler)
        t_scaled = _scale(t_raw, self.thickness_scaler)
        s_scaled = _scale(s_raw, self.semid_scaler)

        tokens = {
            "radius": r_scaled,
            "thickness": t_scaled,
            "semi": s_scaled,
            "mat_idx": mats_idx,
            "n_surfaces": len(r_scaled)
        }

        # globals & targets (already clean)
        globals_raw = self.lens_info[fname]["globals"].astype(float).reshape(1,-1)
        globals_scaled = (self.global_scaler.transform(globals_raw).reshape(-1).astype(np.float32)
                          if self.global_scaler is not None else globals_raw.reshape(-1).astype(np.float32))
        targets = self.lens_info[fname]["targets"].reshape(1,-1)
        # if target_scaler was provided, transform; otherwise return raw targets
        if self.target_scaler is None:
            targets_scaled = targets.reshape(-1).astype(np.float32)
        else:
            targets_scaled = self.target_scaler.transform(targets).reshape(-1).astype(np.float32)

        return tokens, globals_scaled, targets_scaled, fname

def collate_fn(batch):
    batch_size = len(batch)
    max_len = min(MAX_SURFACES, max(x[0]["n_surfaces"] for x in batch))

    radii = np.zeros((batch_size, max_len), dtype=np.float32)
    thicks = np.zeros((batch_size, max_len), dtype=np.float32)
    semis = np.zeros((batch_size, max_len), dtype=np.float32)
    mats = np.zeros((batch_size, max_len), dtype=np.int64)
    mask = np.zeros((batch_size, max_len), dtype=np.bool_)

    globals_batch = np.zeros((batch_size, globals_arr.shape[1] if 'globals_arr' in globals() else 3), dtype=np.float32)
    targets_batch = np.zeros((batch_size, len(TARGET_COLS)), dtype=np.float32)
    fnames = []

    for i, (tokens, g, targ, fname) in enumerate(batch):
        L = min(tokens["n_surfaces"], max_len)
        radii[i, :L] = tokens["radius"][:L]
        thicks[i, :L] = tokens["thickness"][:L]
        semis[i, :L] = tokens["semi"][:L]
        mats[i, :L] = tokens["mat_idx"][:L]
        mask[i, :L] = True
        globals_batch[i] = g
        targets_batch[i] = targ
        fnames.append(fname)

    return {
        "radius": torch.tensor(radii, device=DEVICE),
        "thickness": torch.tensor(thicks, device=DEVICE),
        "semi": torch.tensor(semis, device=DEVICE),
        "mat_idx": torch.tensor(mats, device=DEVICE),
        "mask": torch.tensor(mask, device=DEVICE),
        "globals": torch.tensor(globals_batch, device=DEVICE),
        "targets": torch.tensor(targets_batch, device=DEVICE),
        "fnames": fnames
    }

# ---------------------------
# Model definition (same architecture but Dropout replaced/removed from head)
class SurfaceTransformerModel(nn.Module):
    def __init__(self, n_materials, embed_dim=EMBED_DIM, num_heads=NUM_HEADS, num_layers=NUM_LAYERS, mlp_hidden=MLP_HIDDEN, dropout=DROPOUT, num_targets=len(TARGET_COLS)):
        super().__init__()
        self.embed_dim = embed_dim
        self.num_proj = nn.Linear(3, embed_dim//2)
        self.mat_emb = nn.Embedding(n_materials, embed_dim//2)
        self.combine_ln = nn.LayerNorm(embed_dim)
        self.cls_token = nn.Parameter(torch.randn(1, 1, embed_dim))
        self.pos_emb = nn.Parameter(torch.randn(1, MAX_SURFACES+1, embed_dim))
        encoder_layer = nn.TransformerEncoderLayer(d_model=embed_dim, nhead=num_heads, dim_feedforward=embed_dim*4, dropout=dropout, activation='gelu')
        self.transformer = nn.TransformerEncoder(encoder_layer, num_layers=num_layers)
        self.global_proj = nn.Linear(3, embed_dim//2)
        # Dropout layers removed — replaced with Identity so head shape remains the same
        self.head = nn.Sequential(
            nn.LayerNorm(embed_dim + embed_dim//2),
            nn.Linear(embed_dim + embed_dim//2, mlp_hidden),
            nn.GELU(),
            nn.Identity(),
            nn.Linear(mlp_hidden, mlp_hidden//2),
            nn.GELU(),
            nn.Identity(),
            nn.Linear(mlp_hidden//2, num_targets)
        )

    def forward(self, radius, thickness, semi, mat_idx, mask, globals_in):
        B, L = radius.shape
        num_feats = torch.stack([radius, thickness, semi], dim=-1)
        num_proj = self.num_proj(num_feats)
        mat_e = self.mat_emb(mat_idx)
        token_emb = torch.cat([num_proj, mat_e], dim=-1)
        token_emb = self.combine_ln(token_emb)
        cls_tokens = self.cls_token.expand(B, -1, -1)
        x = torch.cat([cls_tokens, token_emb], dim=1)
        pos = self.pos_emb[:, :L+1, :]
        x = x + pos
        x = x.transpose(0,1)
        src_key_padding_mask = ~mask
        cls_col = torch.zeros((B,1), dtype=torch.bool, device=mask.device)
        src_key_padding_mask = torch.cat([cls_col, src_key_padding_mask], dim=1)
        out = self.transformer(x, src_key_padding_mask=src_key_padding_mask)
        out = out.transpose(0,1)
        cls_out = out[:,0,:]
        g_proj = self.global_proj(globals_in)
        combined = torch.cat([cls_out, g_proj], dim=-1)
        preds = self.head(combined)
        return preds

# ---------------------------
# Prepare data / dataloaders
all_keys = list(lens_info.keys())
all_keys = [k for k in all_keys if os.path.splitext(k)[0] in surface_files_by_name]
if len(all_keys) == 0:
    raise SystemExit("No lens entries found with matching surface CSVs in cleaned materials folder.")

train_keys, val_keys = train_test_split(all_keys, test_size=0.15, random_state=RANDOM_SEED)
train_info = {k: lens_info[k] for k in train_keys}
val_info = {k: lens_info[k] for k in val_keys}

# Fit target_scaler on TRAIN targets only (no leakage)
train_target_vals = np.vstack([train_info[k]["targets"] for k in train_keys]) if len(train_keys)>0 else np.zeros((0, len(TARGET_COLS)))
if train_target_vals.size == 0:
    target_scaler = None
else:
    target_scaler = MinMaxScaler().fit(train_target_vals)
# Save that scaler
import joblib
os.makedirs("artifacts", exist_ok=True)
if target_scaler is not None:
    joblib.dump(target_scaler, os.path.join("artifacts","target_scaler.joblib"))
    print("Saved target_scaler fitted on TRAIN to artifacts/target_scaler.joblib")
else:
    print("No train targets found; target_scaler is None.")

train_dataset = LensDataset(train_info, surface_files_by_name, mat_encoder, radius_scaler, thickness_scaler, semid_scaler, global_scaler, target_scaler)
val_dataset = LensDataset(val_info, surface_files_by_name, mat_encoder, radius_scaler, thickness_scaler, semid_scaler, global_scaler, target_scaler)

train_loader = DataLoader(train_dataset, batch_size=BATCH_SIZE, shuffle=True, collate_fn=collate_fn)
val_loader = DataLoader(val_dataset, batch_size=BATCH_SIZE, shuffle=False, collate_fn=collate_fn)

# ---------------------------
# Model init / optimizer / loss
n_materials = len(material_set) if len(material_set) > 0 else 1
model = SurfaceTransformerModel(n_materials=n_materials).to(DEVICE)

optimizer = torch.optim.AdamW(model.parameters(), lr=LR, weight_decay=WEIGHT_DECAY)
criterion = nn.MSELoss(reduction="none")
target_weights = torch.ones(len(TARGET_COLS), device=DEVICE)

scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, mode='min', factor=0.5, patience=4)

def tensor_has_bad(tensor):
    a = tensor.detach().cpu()
    return torch.isnan(a).any().item() or torch.isinf(a).any().item()

def compute_elastic_net_penalty(model, lmbda=ELASTIC_LAMBDA, l1_ratio=L1_RATIO):
    """Compute L1 and L2 (squared) sums over model parameters (excluding biases / scalars)."""
    l1 = torch.tensor(0.0, device=DEVICE)
    l2 = torch.tensor(0.0, device=DEVICE)
    for p in model.parameters():
        if not p.requires_grad:
            continue
        # exclude biases / 0-dim params from regularization
        if p.dim() <= 1:
            continue
        l1 = l1 + p.abs().sum()
        l2 = l2 + (p * p).sum()
    # Elastic net: lmbda * ( l1_ratio * L1 + (1-l1_ratio) * 0.5 * L2 )
    return lmbda * (l1_ratio * l1 + (1.0 - l1_ratio) * 0.5 * l2)

def evaluate(model, loader):
    model.eval()
    total_loss = 0.0
    total_n = 0
    with torch.no_grad():
        for batch in loader:
            preds = model(batch["radius"], batch["thickness"], batch["semi"], batch["mat_idx"], batch["mask"], batch["globals"])
            loss_mat = criterion(preds, batch["targets"])
            weighted = loss_mat * target_weights.unsqueeze(0)
            loss = weighted.mean()
            total_loss += float(loss.item()) * preds.size(0)
            total_n += preds.size(0)
    return total_loss / total_n if total_n>0 else float("nan")

# ---------------------------
# Train loop (add elastic net penalty to training loss only)
best_val = float('inf')

# track losses per epoch
train_losses_per_epoch = []
val_losses_per_epoch = []

for epoch in range(1, NUM_EPOCHS+1):
    model.train()
    running_loss = 0.0
    n_samples = 0
    for batch in train_loader:
        # data checks
        bad_slots = [k for k in ["radius","thickness","semi","globals","targets"] if tensor_has_bad(batch[k])]
        if bad_slots:
            print("Found NaN/Inf in data batch slots:", bad_slots)
            for i,fname in enumerate(batch["fnames"]):
                print(" Problem sample:", fname)
            raise RuntimeError("NaN/Inf detected in batch - inputs are not clean.")

        preds = model(batch["radius"], batch["thickness"], batch["semi"], batch["mat_idx"], batch["mask"], batch["globals"])
        if tensor_has_bad(preds):
            raise RuntimeError("Model produced non-finite predictions; aborting.")

        loss_mat = criterion(preds, batch["targets"])
        weighted = loss_mat * target_weights.unsqueeze(0)
        base_loss = weighted.mean()

        # compute elastic net penalty (training only)
        reg = compute_elastic_net_penalty(model, lmbda=ELASTIC_LAMBDA, l1_ratio=L1_RATIO)
        loss = base_loss + reg

        if not torch.isfinite(loss):
            raise RuntimeError("Non-finite loss detected; aborting.")

        optimizer.zero_grad()
        loss.backward()
        torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
        optimizer.step()

        running_loss += float(base_loss.item()) * preds.size(0)  # record base loss (without reg) for logging comparability
        n_samples += preds.size(0)

    train_loss = running_loss / n_samples if n_samples>0 else float("nan")
    val_loss = evaluate(model, val_loader)

    # append losses for plotting later
    train_losses_per_epoch.append(train_loss)
    val_losses_per_epoch.append(val_loss)

    print(f"Epoch {epoch}/{NUM_EPOCHS} TrainLoss={train_loss:.6f} ValLoss={val_loss:.6f} (reg={float(reg.item()):.6g})")

    if val_loss < best_val:
        best_val = val_loss
        torch.save(model.state_dict(), "best_surface_transformer_no_impute_elasticnet.pth")
        print("Saved best model.")

print("Training complete. Best val loss:", best_val)

# final scheduler step on last validation loss
val_loss = evaluate(model, val_loader)
try:
    scheduler.step(val_loss)
except Exception:
    pass

# ---------------------------
# Save artifacts (scalers + encoders + model + histories + preds)
import joblib, os
os.makedirs("artifacts", exist_ok=True)
torch.save(model.state_dict(), "artifacts/best_surface_transformer_elasticnet.pth")
joblib.dump(target_scaler, "artifacts/target_scaler.joblib")
joblib.dump(radius_scaler, "artifacts/radius_scaler.joblib")
joblib.dump(thickness_scaler, "artifacts/thickness_scaler.joblib")
joblib.dump(semid_scaler, "artifacts/semid_scaler.joblib")
joblib.dump(global_scaler, "artifacts/global_scaler.joblib")
joblib.dump(mat_encoder, "artifacts/mat_encoder.joblib")

# helper to collect preds/targets across a loader and inverse-transform
def collect_preds_targets(loader):
    preds_list, targs_list = [], []
    fnames = []
    model.eval()
    with torch.no_grad():
        for batch in loader:
            preds = model(batch["radius"], batch["thickness"], batch["semi"],
                          batch["mat_idx"], batch["mask"], batch["globals"])
            preds_list.append(preds.cpu().numpy())
            targs_list.append(batch["targets"].cpu().numpy())
            fnames.extend(batch["fnames"])
    if len(preds_list) == 0:
        return np.zeros((0, len(TARGET_COLS))), np.zeros((0, len(TARGET_COLS))), fnames
    preds_all = np.vstack(preds_list)
    targs_all = np.vstack(targs_list)
    # inverse-transform if possible
    try:
        preds_orig = target_scaler.inverse_transform(preds_all)
        targs_orig = target_scaler.inverse_transform(targs_all)
    except Exception:
        preds_orig = preds_all.copy()
        targs_orig = targs_all.copy()
    return preds_orig, targs_orig, fnames

# collect and save training & validation preds/targets
preds_train, targs_train, train_fnames = collect_preds_targets(train_loader)
preds_val, targs_val, val_fnames = collect_preds_targets(val_loader)

np.savez(os.path.join("artifacts", "train_history.npz"),
         train_losses=np.array(train_losses_per_epoch),
         val_losses=np.array(val_losses_per_epoch))
np.savez(os.path.join("artifacts", "val_predictions.npz"),
         preds=preds_val, targs=targs_val, fnames=np.array(val_fnames, dtype=object))
np.savez(os.path.join("artifacts", "train_predictions.npz"),
         preds=preds_train, targs=targs_train, fnames=np.array(train_fnames, dtype=object))

print("Saved artifacts to ./artifacts/: scalers, encoders, model, histories, and predictions.")
