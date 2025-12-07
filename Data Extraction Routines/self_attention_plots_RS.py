#!/usr/bin/env python3
import os
import numpy as np
import joblib
import matplotlib.pyplot as plt
import math
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.metrics import r2_score as _r2_score  # avoid possible shadowing

ART_DIR = "artifacts"
OUT_DIR = os.path.join(ART_DIR, "attention_diagnostics")
os.makedirs(OUT_DIR, exist_ok=True)

HIST_PATH = os.path.join(ART_DIR, "train_history.npz")
VAL_PATH  = os.path.join(ART_DIR, "val_predictions.npz")
TARGET_SCALER_PATH = os.path.join(ART_DIR, "target_scaler.joblib")

# ---- load saved things ----
if not os.path.exists(HIST_PATH) or not os.path.exists(VAL_PATH):
    raise SystemExit("Missing artifacts: ensure train script wrote artifacts/train_history.npz and artifacts/val_predictions.npz")

hist = np.load(HIST_PATH)
train_losses = hist["train_losses"]
val_losses = hist["val_losses"]

val = np.load(VAL_PATH, allow_pickle=True)
preds_orig = val["preds"]
targs_orig = val["targs"]
fnames = val["fnames"]

# optionally load target_scaler if preds/targs were saved scaled (not required if saved original units)
# target_scaler = joblib.load(TARGET_SCALER_PATH) if os.path.exists(TARGET_SCALER_PATH) else None

TARGET_COLS = [
    "Tan Shift", "Sag Shift",
    "Long_0.4861", "Long_0.5876", "Long_0.6563",
    "Poly",
    "RMS_0.4861", "RMS_0.5876", "RMS_0.6563",
    "Rel. Ill", "Effective F/#"
]

# ---------- helper: save+show utility ----------
def save_and_show(fig, out_path, dpi=150):
    try:
        fig.savefig(out_path, dpi=dpi, bbox_inches="tight")
        print("Saved plot to:", out_path)
    except Exception as e:
        print("Warning: failed to save", out_path, ":", e)
    plt.show()

# ---- 1) Loss curve (train + val on same plot) ----
fig1 = plt.figure(figsize=(8,5))
epochs = np.arange(1, len(train_losses)+1)
plt.plot(epochs, train_losses, marker='o', label='Train Loss')
plt.plot(epochs, val_losses, marker='o', label='Val Loss')
plt.xlabel("Epoch")
plt.ylabel("Loss")
plt.title("Training & Validation Loss per Epoch")
plt.legend()
plt.grid(alpha=0.25)
plt.tight_layout()
loss_path = os.path.join(OUT_DIR, "loss_curve.png")
save_and_show(fig1, loss_path)
plt.close(fig1)

# ---- helper to draw matrix of pred vs actuals (2x6 layout if <=12 targets) ----
def plot_pred_vs_actual_matrix(preds_all, targs_all, title_prefix="", out_png=None):
    num_targets = preds_all.shape[1]
    # prefer 2x6 layout when <=12 targets
    if num_targets <= 12:
        rows, cols = 2, 5
    else:
        cols = math.ceil(math.sqrt(num_targets))
        rows = math.ceil(num_targets / cols)

    fig, axes = plt.subplots(rows, cols, figsize=(4*cols, 4*rows))
    axes = np.array(axes).reshape(-1)

    r2s_local = []
    for i in range(num_targets):
        try:
            r2 = _r2_score(targs_all[:, i], preds_all[:, i])
        except Exception:
            r2 = float('nan')
        r2s_local.append(r2)

    for i in range(num_targets):
        ax = axes[i]
        ax.scatter(targs_all[:, i], preds_all[:, i], s=8, alpha=0.5)
        mn = min(float(np.nanmin(targs_all[:, i])), float(np.nanmin(preds_all[:, i])))
        mx = max(float(np.nanmax(targs_all[:, i])), float(np.nanmax(preds_all[:, i])))
        if mn == mx:
            # avoid degenerate line
            mx = mn + 1.0
        ax.plot([mn, mx], [mn, mx], linestyle="--", linewidth=1)
        ax.set_title(f"{TARGET_COLS[i]}", fontsize=10)
        # text-only R^2 in corner:
        ax.text(0.02, 0.95, f"R² = {r2s_local[i]:.3f}", transform=ax.transAxes,
                fontsize=9, verticalalignment='top',
                bbox=dict(boxstyle="round,pad=0.2", facecolor="white", alpha=0.6, edgecolor='none'))
        ax.set_xlabel("Actual")
        ax.set_ylabel("Predicted")

    for j in range(num_targets, len(axes)):
        fig.delaxes(axes[j])

    fig.suptitle(title_prefix, fontsize=14)
    # first tighten outer layout, then add spacing between subplots to avoid overlap
    fig.tight_layout(rect=[0, 0.03, 1, 0.97])
    fig.subplots_adjust(hspace=0.6, wspace=0.3)

    if out_png:
        try:
            fig.savefig(out_png, dpi=150, bbox_inches="tight")
            print("Saved plot to:", out_png)
        except Exception as e:
            print("Warning: failed to save", out_png, ":", e)

    plt.show()
    return fig

# ---- 2) Validation predicted vs actual (matrix) ----
val_out = os.path.join(OUT_DIR, "val_pred_vs_actual.png")
plot_pred_vs_actual_matrix(preds_orig, targs_orig, title_prefix="Validation: Predicted vs Actual", out_png=val_out)

# ---- 3) Training predicted vs actual (OPTIONAL) ----
# If you saved training-set preds & targets, plot them as well. Otherwise skip.
TRAIN_PRED_PATH = os.path.join(ART_DIR, "train_predictions.npz")
if os.path.exists(TRAIN_PRED_PATH):
    t = np.load(TRAIN_PRED_PATH, allow_pickle=True)
    train_preds = t["preds"]; train_targs = t["targs"]
    train_out = os.path.join(OUT_DIR, "train_pred_vs_actual.png")
    plot_pred_vs_actual_matrix(train_preds, train_targs, title_prefix="Train: Predicted vs Actual", out_png=train_out)
else:
    print("No train_predictions.npz found — skipping train pred vs actual plot.")
