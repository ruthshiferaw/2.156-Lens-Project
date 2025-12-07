#!/usr/bin/env python3
"""
compare_attention_plots.py

Produces 3 diagnostics comparing attention-pooled features between train and val:
  1) Attention weight distribution (histograms) -- requires per-sample 'attn' arrays
  2) PCA scatter of pooled features (2D) colored by split
  3) Correlation heatmaps between pooled PCA components and targets (side-by-side train vs val)

Usage:
  python compare_attention_plots.py
  OR
  python compare_attention_plots.py <train_npz> <val_npz>

Artifact expectations (saved by attention pooling step):
  - artifacts/attention_pooled_train.npz
  - artifacts/attention_pooled_val.npz

Each .npz should contain at least:
  - 'pooled' (N x D) or 'pooled_feats' : pooled feature vectors
  - 'targets' (N x T) or 'targs'        : target values aligned to pooled rows
Optional:
  - 'attn' : object array/list of per-sample attention weight arrays
  - 'fnames': names
"""

import os
import sys
import numpy as np
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
from scipy.stats import pearsonr
import math

# Paths (defaults)
ART_DIR = "artifacts"
DEFAULT_TRAIN = os.path.join(ART_DIR, "pooled_train.npz")
DEFAULT_VAL   = os.path.join(ART_DIR, "pooled_val.npz")
OUT_DIR = os.path.join(ART_DIR, "attention_diagnostics")
os.makedirs(OUT_DIR, exist_ok=True)

def load_pooled(path):
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    d = np.load(path, allow_pickle=True)
    # pooled key variants
    pooled = None
    for k in ("pooled","pooled_feats","pooled_features"):
        if k in d:
            pooled = d[k]
            break
    # attention arrays (optional)
    attn = None
    for k in ("attn","attentions","attention"):
        if k in d:
            attn = d[k]
            break
    # targets
    targets = None
    for k in ("targets","targs","y"):
        if k in d:
            targets = d[k]
            break
    fnames = d["fnames"] if "fnames" in d else None
    return {"pooled": pooled, "attn": attn, "targets": targets, "fnames": fnames, "raw": d}

def ensure_2d(a, name):
    a = np.asarray(a)
    if a.ndim == 1:
        a = a.reshape(-1, 1)
    if a.ndim != 2:
        raise ValueError(f"{name} must be 2D array-like (got shape {a.shape})")
    return a

def safe_concat_attn(attn_obj):
    # attn_obj expected to be sequence-like length N of 1D arrays
    if attn_obj is None:
        return None
    try:
        return [np.asarray(x).astype(float) for x in attn_obj]
    except Exception:
        return None

def plot_attention_distribution(attn_train_list, attn_val_list, out_png):
    """
    Flatten per-sample attention arrays to single pool of weights for distribution comparison.
    We'll also compute per-sample max and entropy as diagnostics.
    """
    import numpy as np
    # flatten
    all_train = np.concatenate([a for a in attn_train_list if len(a)>0]) if (attn_train_list and any(len(a)>0 for a in attn_train_list)) else np.array([])
    all_val   = np.concatenate([a for a in attn_val_list   if len(a)>0]) if (attn_val_list   and any(len(a)>0 for a in attn_val_list))   else np.array([])

    # per-sample max-weight & entropy
    def per_sample_stats(attn_list):
        max_w = []
        ent = []
        for a in attn_list:
            if len(a)==0:
                max_w.append(np.nan); ent.append(np.nan); continue
            p = a / (a.sum() + 1e-12)
            max_w.append(float(np.max(p)))
            # entropy
            ent.append(float(-np.sum(np.where(p>0, p*np.log(p), 0.0))))
        return np.array(max_w), np.array(ent)

    train_max, train_ent = per_sample_stats(attn_train_list) if attn_train_list else (np.array([]), np.array([]))
    val_max, val_ent       = per_sample_stats(attn_val_list) if attn_val_list else (np.array([]), np.array([]))

    fig, axes = plt.subplots(2, 2, figsize=(12,8))
    ax0, ax1, ax2, ax3 = axes.flatten()

    # histogram of raw weights
    bins = np.linspace(0, 1, 101)
    if all_train.size>0:
        ax0.hist(all_train, bins=bins, alpha=0.5, label=f"train (n={len(attn_train_list)})", density=True)
    if all_val.size>0:
        ax0.hist(all_val, bins=bins, alpha=0.5, label=f"val (n={len(attn_val_list)})", density=True)
    ax0.set_title("Attention weight distribution (flattened across surfaces)")
    ax0.set_xlabel("Attention weight")
    ax0.set_ylabel("Density")
    ax0.legend()

    # per-sample max weight
    if train_max.size>0:
        ax1.hist(train_max[~np.isnan(train_max)], bins=bins, alpha=0.6, label="train", density=True)
    if val_max.size>0:
        ax1.hist(val_max[~np.isnan(val_max)], bins=bins, alpha=0.6, label="val", density=True)
    ax1.set_title("Per-sample max attention weight distribution")
    ax1.set_xlabel("max attention")
    ax1.legend()

    # per-sample entropy
    ent_bins = np.linspace(0, np.nanmax(np.concatenate([train_ent[~np.isnan(train_ent)] if train_ent.size else np.array([]),
                                                         val_ent[~np.isnan(val_ent)] if val_ent.size else np.array([])]))+1e-6, 80) if (train_ent.size or val_ent.size) else np.linspace(0,1,10)
    if train_ent.size>0:
        ax2.hist(train_ent[~np.isnan(train_ent)], bins=ent_bins, alpha=0.6, label="train", density=True)
    if val_ent.size>0:
        ax2.hist(val_ent[~np.isnan(val_ent)], bins=ent_bins, alpha=0.6, label="val", density=True)
    ax2.set_title("Per-sample attention entropy (higher = uniform)")
    ax2.set_xlabel("Entropy")
    ax2.legend()

    # boxplot summary of max weights
    labels = []
    data = []
    if train_max.size>0:
        labels.append("train")
        data.append(train_max[~np.isnan(train_max)])
    if val_max.size>0:
        labels.append("val")
        data.append(val_max[~np.isnan(val_max)])
    if data:
        ax3.boxplot(data, labels=labels, showfliers=False)
        ax3.set_title("Boxplot: per-sample max attention weight")
    else:
        ax3.text(0.5,0.5,"No per-sample stats available", ha='center')

    fig.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.show()

def plot_pca_scatter(pooled_train, pooled_val, labels_train=None, labels_val=None, out_png=None):
    # stack to fit PCA consistently
    X = np.vstack([pooled_train, pooled_val])
    pca = PCA(n_components=2)
    X2 = pca.fit_transform(X)
    n1 = pooled_train.shape[0]
    Xtr = X2[:n1]; Xval = X2[n1:]
    fig, ax = plt.subplots(figsize=(8,6))
    ax.scatter(Xtr[:,0], Xtr[:,1], s=10, alpha=0.6, label=f"train (n={Xtr.shape[0]})")
    ax.scatter(Xval[:,0], Xval[:,1], s=12, alpha=0.8, label=f"val (n={Xval.shape[0]})", marker='^')
    ax.set_title("PCA (2D) of pooled embeddings — train vs val")
    ax.set_xlabel("PC1")
    ax.set_ylabel("PC2")
    ax.legend()
    ax.grid(alpha=0.15)
    fig.tight_layout()
    if out_png:
        fig.savefig(out_png, dpi=150)
    plt.show()
    return pca

def plot_corr_heatmaps(pooled_train, targets_train, pooled_val, targets_val, target_col_names=None, out_png=None, n_components=16):
    # We'll PCA-reduce pooled to n_components for interpretability, compute pearson corr matrix (PCs x targets)
    pca = PCA(n_components=min(n_components, pooled_train.shape[1], pooled_val.shape[1]))
    # fit on train pooled, transform train & val for comparable components
    pca.fit(pooled_train)
    Ttr = pca.transform(pooled_train)
    Tval = pca.transform(pooled_val)

    # ensure targets 2D
    targets_train = ensure_2d(targets_train, "targets_train")
    targets_val   = ensure_2d(targets_val, "targets_val")
    num_pc = Ttr.shape[1]
    num_targets = targets_train.shape[1]

    # compute correlation matrices (num_pc x num_targets)
    corr_tr = np.zeros((num_pc, num_targets))
    corr_val = np.zeros((num_pc, num_targets))
    for i in range(num_pc):
        for j in range(num_targets):
            # handle constant arrays with try/except
            try:
                corr_tr[i,j] = pearsonr(Ttr[:,i], targets_train[:,j])[0]
            except Exception:
                corr_tr[i,j] = np.nan
            try:
                corr_val[i,j] = pearsonr(Tval[:,i], targets_val[:,j])[0]
            except Exception:
                corr_val[i,j] = np.nan

    # Plot side-by-side heatmaps
    fig, axes = plt.subplots(1,2, figsize=(14, max(6, num_pc*0.25)))
    vmax = np.nanmax(np.abs(np.concatenate([corr_tr.flatten(), corr_val.flatten()])))
    vmax = max(vmax, 1e-6)
    im0 = axes[0].imshow(corr_tr, vmin=-vmax, vmax=vmax, aspect='auto', cmap='RdBu_r')
    axes[0].set_title("Train: corr(PC_k, target_j)")
    axes[0].set_xlabel("Targets")
    axes[0].set_ylabel("PCA components (pooled)")
    im1 = axes[1].imshow(corr_val, vmin=-vmax, vmax=vmax, aspect='auto', cmap='RdBu_r')
    axes[1].set_title("Val: corr(PC_k, target_j)")
    axes[1].set_xlabel("Targets")
    # set xticks labels
    if target_col_names is None:
        xlabels = [f"t{j}" for j in range(num_targets)]
    else:
        xlabels = target_col_names
    axes[0].set_xticks(np.arange(len(xlabels))); axes[0].set_xticklabels(xlabels, rotation=45, ha='right')
    axes[1].set_xticks(np.arange(len(xlabels))); axes[1].set_xticklabels(xlabels, rotation=45, ha='right')
    axes[0].set_yticks(np.arange(num_pc)); axes[0].set_yticklabels([f"PC{i+1}" for i in range(num_pc)])
    axes[1].set_yticks(np.arange(num_pc)); axes[1].set_yticklabels([f"PC{i+1}" for i in range(num_pc)])
    # colorbar
    # Add a new axes for the colorbar: [left, bottom, width, height]
    cbar_ax = fig.add_axes([0.92, 0.15, 0.02, 0.7])  
    cbar = fig.colorbar(im1, cax=cbar_ax)

    # fig.tight_layout()
    if out_png:
        fig.savefig(out_png, dpi=150)
    plt.show()
    return pca, corr_tr, corr_val

def main(train_path, val_path):
    print("Loading artifacts:")
    print("  train:", train_path)
    print("  val:  ", val_path)
    tr = load_pooled(train_path)
    va = load_pooled(val_path)

    # unpack pooled features & targets
    if tr["pooled"] is None or va["pooled"] is None:
        raise RuntimeError("Both train and val artifacts must contain pooled feature arrays under key 'pooled' or 'pooled_feats'.")

    pooled_tr = ensure_2d(tr["pooled"], "pooled_train").astype(float)
    pooled_va = ensure_2d(va["pooled"], "pooled_val").astype(float)

    # targets: if missing, we can't do correlation heatmap
    if tr["targets"] is None or va["targets"] is None:
        print("Warning: one of the artifacts has no 'targets' array. Correlation heatmap will be skipped.")
        targets_tr = None; targets_va = None
    else:
        targets_tr = ensure_2d(tr["targets"], "targets_train").astype(float)
        targets_va = ensure_2d(va["targets"], "targets_val").astype(float)
        if targets_tr.shape[1] != targets_va.shape[1]:
            print("Warning: train/val targets have differing number of target columns. Truncating to min columns.")
            m = min(targets_tr.shape[1], targets_va.shape[1])
            targets_tr = targets_tr[:, :m]; targets_va = targets_va[:, :m]

    # attention lists
    attn_tr = safe_concat_attn(tr["attn"])
    attn_va = safe_concat_attn(va["attn"])

    # Plot 1: attention distribution if available
    if attn_tr is None or attn_va is None:
        print("No attention arrays found in one of the artifacts — skipping attention distribution plot.")
    else:
        out1 = os.path.join(OUT_DIR, "attention_distribution_train_vs_val.png")
        plot_attention_distribution(attn_tr, attn_va, out1)
        print("Saved attention distribution plot to:", out1)

    # Plot 2: PCA scatter of pooled embeddings
    out2 = os.path.join(OUT_DIR, "pooled_pca_train_vs_val.png")
    _pca = plot_pca_scatter(pooled_tr, pooled_va, out_png=out2)
    print("Saved pooled PCA scatter to:", out2)

    # Plot 3: Correlation heatmaps (PCA comps vs targets)
    if targets_tr is None or targets_va is None:
        print("Skipping correlation heatmaps because targets are missing.")
    else:
        # derive target column names if present in the artifact metadata
        target_names = None
        # try to get from train raw file if available as list
        if hasattr(tr["raw"], "files") and "target_names" in tr["raw"]:
            try:
                target_names = list(tr["raw"]["target_names"])
            except Exception:
                target_names = None
        # fallback to generic names
        if target_names is None:
            target_names = [f"t{i}" for i in range(targets_tr.shape[1])]
        out3 = os.path.join(OUT_DIR, "pooled_corr_train_vs_val.png")
        plot_corr_heatmaps(pooled_tr, targets_tr, pooled_va, targets_va, target_col_names=target_names, out_png=out3, n_components=16)
        print("Saved pooled->target correlation heatmaps to:", out3)

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        train_path = sys.argv[1]; val_path = sys.argv[2]
    else:
        train_path = DEFAULT_TRAIN; val_path = DEFAULT_VAL
    main(train_path, val_path)
