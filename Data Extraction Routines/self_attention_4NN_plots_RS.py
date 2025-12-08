#!/usr/bin/env python3
"""
plot_multi_or_single.py

Plotter that supports:
 - old single-model artifacts:
     artifacts/train_history.npz
     artifacts/val_predictions.npz (contains preds,targs,fnames)
 - new multi-model artifacts created by the multi-model trainer:
     artifacts/train_history__<group>.npz
     artifacts/val_predictions__<group>.npz (contains preds,targs,fnames,cols)

It will produce for each model (or the single model):
  1) loss curve (train+val)
  2) val Pred vs Actual matrix
  3) train Pred vs Actual matrix (if available)
Plots are saved to artifacts/attention_diagnostics/.
"""
import os, glob
import numpy as np
import matplotlib.pyplot as plt
import math
from sklearn.metrics import r2_score as _r2_score

ART_DIR = "artifacts"
OUT_DIR = os.path.join(ART_DIR, "attention_diagnostics")
os.makedirs(OUT_DIR, exist_ok=True)

# helper: safe load npz and handle missing keys
def load_npz(path):
    if not os.path.exists(path):
        return None
    try:
        return np.load(path, allow_pickle=True)
    except Exception as e:
        print("Failed to load", path, ":", e)
        return None

def save_and_show(fig, out_path):
    try:
        fig.savefig(out_path, dpi=150, bbox_inches="tight")
        print("Saved:", out_path)
    except Exception as e:
        print("Failed to save:", out_path, ":", e)
    plt.close(fig)

def plot_loss_curve(train_losses, val_losses, out_png, title="Loss per epoch"):
    fig = plt.figure(figsize=(8,5))
    epochs = np.arange(1, len(train_losses)+1)
    plt.plot(epochs, train_losses, marker='o', label='Train Loss')
    plt.plot(epochs, val_losses, marker='o', label='Val Loss')
    plt.xlabel("Epoch"); plt.ylabel("Loss"); plt.title(title)
    plt.legend(); plt.grid(alpha=0.25); plt.tight_layout()
    save_and_show(fig, out_png)

def plot_pred_vs_actual_matrix(preds_all, targs_all, target_names, out_png, title_prefix=""):
    preds_all = np.asarray(preds_all)
    targs_all = np.asarray(targs_all)
    if preds_all.size == 0 or targs_all.size == 0:
        print("Empty preds or targs; skipping", out_png)
        return
    # ensure shapes
    if preds_all.ndim == 1:
        preds_all = preds_all.reshape(-1, 1)
    if targs_all.ndim == 1:
        targs_all = targs_all.reshape(-1, 1)

    num_targets = preds_all.shape[1]
    # prefer 2x6-ish layouts for <=12 targets
    if num_targets <= 12:
        rows, cols = 2, min(6, max(1, math.ceil(num_targets/2)))
    else:
        cols = math.ceil(math.sqrt(num_targets)); rows = math.ceil(num_targets/cols)

    fig, axes = plt.subplots(rows, cols, figsize=(4*cols, 4*rows))
    axes = np.array(axes).reshape(-1)

    r2s_local = []
    for i in range(num_targets):
        try:
            r2 = float(_r2_score(targs_all[:, i], preds_all[:, i]))
        except Exception:
            r2 = float('nan')
        r2s_local.append(r2)

    for i in range(num_targets):
        ax = axes[i]
        # scatter - handle NaN gracefully
        ax.scatter(targs_all[:, i], preds_all[:, i], s=8, alpha=0.5)
        # compute min/max safely
        mn = float(np.nanmin(np.stack([targs_all[:, i], preds_all[:, i]])))
        mx = float(np.nanmax(np.stack([targs_all[:, i], preds_all[:, i]])))
        if mn == mx:
            mx = mn + 1.0
        ax.plot([mn, mx], [mn, mx], linestyle="--", linewidth=1)
        name = target_names[i] if target_names is not None and i < len(target_names) else f"t{i}"
        ax.set_title(f"{name}", fontsize=10)
        ax.text(0.02, 0.95, f"R² = {r2s_local[i]:.3f}", transform=ax.transAxes,
                fontsize=9, verticalalignment='top',
                bbox=dict(boxstyle="round,pad=0.2", facecolor="white", alpha=0.6, edgecolor='none'))
        ax.set_xlabel("Actual"); ax.set_ylabel("Predicted")

    # remove unused subplots
    for j in range(num_targets, len(axes)):
        fig.delaxes(axes[j])

    fig.suptitle(title_prefix, fontsize=14)
    fig.tight_layout(rect=[0, 0.03, 1, 0.97])
    fig.subplots_adjust(hspace=0.6, wspace=0.3)
    save_and_show(fig, out_png)

# function to handle a single model artifact set
def handle_model(base_name, hist_path, val_path, train_path=None):
    # try to load history
    hist = load_npz(hist_path)
    if hist is None:
        print(f"No history found at {hist_path}; skipping loss plot for {base_name}")
    else:
        # robust lookup for train/val losses
        train_losses = None
        val_losses = None
        for k in ("train_losses","train_loss","train"):
            if k in hist:
                train_losses = hist[k]
                break
        for k in ("val_losses","val_loss","val"):
            if k in hist:
                val_losses = hist[k]
                break

        if train_losses is None or val_losses is None:
            print("History file found but missing expected keys (train_losses/val_losses).")
        else:
            plot_loss_curve(np.asarray(train_losses), np.asarray(val_losses),
                            os.path.join(OUT_DIR, f"{base_name}_loss_curve.png"),
                            title=f"{base_name} Loss per epoch")

    # load val preds
    val = load_npz(val_path)
    if val is None:
        print(f"No val predictions found at {val_path}; skipping pred vs actual for {base_name}")
        return

    # try to get preds, targs, and column names robustly
    preds = None
    targs = None
    fnames = None
    cols = None

    for k in ("preds","pred","y_pred","predictions"):
        if k in val:
            preds = val[k]; break
    for k in ("targs","targ","targets","y_true"):
        if k in val:
            targs = val[k]; break
    if "fnames" in val: fnames = val["fnames"]
    if "cols" in val:
        cols = list(val["cols"])
    elif hasattr(val, "files") and "cols" in val.files:
        cols = list(val["cols"])

    # If cols not in val artifact, try to use a global fallback (user-provided)
    if cols is None:
        # try to infer number of targets and create generic names
        if preds is not None:
            preds_arr = np.asarray(preds)
            n = preds_arr.shape[1] if preds_arr.ndim==2 else 1
            cols = [f"t{i}" for i in range(n)]
        elif targs is not None:
            targs_arr = np.asarray(targs)
            n = targs_arr.shape[1] if targs_arr.ndim==2 else 1
            cols = [f"t{i}" for i in range(n)]
        else:
            cols = []

    # Plot val Pred vs Actual
    if preds is not None and targs is not None:
        # convert to arrays (safe)
        preds_arr = np.asarray(preds)
        targs_arr = np.asarray(targs)
        plot_pred_vs_actual_matrix(preds_arr, targs_arr,
                                   target_names=cols,
                                   out_png=os.path.join(OUT_DIR, f"{base_name}_val_pred_vs_actual.png"),
                                   title_prefix=f"{base_name}: Validation Pred vs Actual")
    else:
        print(f"Val artifact {val_path} missing preds/targs arrays; skipping pred vs actual for {base_name}")

    # Plot train Pred vs Actual if available
    if train_path and os.path.exists(train_path):
        tr = load_npz(train_path)
        if tr is not None:
            # robustly fetch preds and targs without using `or` on arrays
            tpreds = None
            tttargs = None
            for k in ("preds","pred","y_pred","predictions"):
                if k in tr:
                    tpreds = tr[k]; break
            for k in ("targs","targ","targets","y_true"):
                if k in tr:
                    tttargs = tr[k]; break

            if tpreds is not None and tttargs is not None:
                plot_pred_vs_actual_matrix(np.asarray(tpreds), np.asarray(tttargs),
                                           target_names=cols,
                                           out_png=os.path.join(OUT_DIR, f"{base_name}_train_pred_vs_actual.png"),
                                           title_prefix=f"{base_name}: Train Pred vs Actual")
            else:
                print("Train predictions file exists but missing preds/targs; skipping train plot.")
        else:
            print("Failed to load train predictions npz.")

# MAIN: detect grouped artifacts
group_val_files = sorted(glob.glob(os.path.join(ART_DIR, "val_predictions__*.npz")))
group_hist_files = { os.path.basename(p).split("__",1)[1].rsplit(".npz",1)[0]: p for p in glob.glob(os.path.join(ART_DIR, "train_history__*.npz")) }

if group_val_files:
    print("Detected grouped artifacts. Will plot each group:")
    for vpath in group_val_files:
        base = os.path.basename(vpath).replace("val_predictions__","").rsplit(".npz",1)[0]
        hist_path = os.path.join(ART_DIR, f"train_history__{base}.npz")
        train_path = os.path.join(ART_DIR, f"train_predictions__{base}.npz")
        print(" - group:", base)
        handle_model(base_name=base, hist_path=hist_path, val_path=vpath, train_path=train_path)
else:
    # fallback to single-model artifact names
    single_hist = os.path.join(ART_DIR, "train_history.npz")
    single_val  = os.path.join(ART_DIR, "val_predictions.npz")
    single_train = os.path.join(ART_DIR, "train_predictions.npz")
    if os.path.exists(single_val):
        print("No grouped artifacts found. Using single-model artifacts.")
        handle_model(base_name="single_model", hist_path=single_hist, val_path=single_val, train_path=single_train)
    else:
        print("No artifacts found to plot. Expected grouped val_predictions__<group>.npz or val_predictions.npz")