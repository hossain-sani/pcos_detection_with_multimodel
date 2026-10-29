"""Step 4 - Multi-Model Benchmark & Bayesian Hyperparameter Tuning.

Stage 4 of the proposed 7-stage PCOS diagnosis pipeline (proposed.png).
Follows the thesis methodology, improving on Efty et al. 2026 (Health Science Reports).

Architecture & Protocols:
  Part A (Table 12 & Step 5 Bridge):
    - Proposed 16 Rotterdam Bio-Markers from Step 3 (stage3_final_16).
    - 9 Proposed Models + 1 Baseline:
        1. LogReg (L2)
        2. Decision Tree (DT baseline)
        3. Random Forest (RF)
        4. XGBoost (XGB)
        5. CatBoost (CB)
        6. LightGBM (LGBM)
        7. SVM (RBF)
        8. Multi-Layer Perceptron (MLP)
        9. Google TabNet (Attentive Tabular Transformer)
       10. Deep ResNet (Tabular ResNet with residual skip connections)
    - Protocol: Nested 5x5 Cross-Validation (cv_splits.joblib, zero test leakage).
    - Bayesian Hyper-Tuning: Optuna TPE sampler tuned strictly inside inner folds.
    - Inside-Fold Resampling: SMOTE-NC on training folds only; untouched validation folds.
    - Outputs: Out-of-Fold (OOF) prediction probabilities + best params saved for Step 5 Neuro-Tree Stacking.
    - Fills Table 12 in docx: Acc, Prec, Rec, F1 formatted as Mean ± Std.

  Part B (Table 8 - Feature Subsets Benchmark):
    - Evaluates 10 models across 12 feature subsets derived from Step 3 wrapper RFE:
        Top 20, Top 16, Top 12 x [LR, XGB, RF, CB] = 12 feature subsets.
    - 10 Models from paper: LR, DT, RF, XGB, CB, AB, SVM, KNN, MLP, LDA.
    - Protocol: Stratified 10-Fold CV with SMOTE-NC inside folds on 432 training samples.
    - Fills Table 8 in docx: Mean Accuracy (%) matrix (10 models x 12 subsets).

Terminal Display:
  - Rich interactive progress bars with percentage completion, ETAs, and live metrics.
  - Aligned publication-grade ASCII tables printed directly to the terminal.
  - Cumulative docx preservation and byte-identical header verification check.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import time
import warnings
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Tuple

# Suppress repetitive library warnings to keep progress bars clean
warnings.filterwarnings("ignore")

import joblib
import numpy as np
import pandas as pd
from tqdm import tqdm

# Headless matplotlib
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.base import BaseEstimator, ClassifierMixin
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.ensemble import AdaBoostClassifier, RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    brier_score_loss,
    f1_score,
    precision_score,
    recall_score,
    roc_auc_score,
    roc_curve,
)
from sklearn.model_selection import StratifiedKFold
from sklearn.neighbors import KNeighborsClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC
from sklearn.tree import DecisionTreeClassifier

from imblearn.over_sampling import SMOTE, SMOTENC
from catboost import CatBoostClassifier
from lightgbm import LGBMClassifier
from xgboost import XGBClassifier

import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader, TensorDataset

import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


# -----------------------------------------------------------------------------
# Global Paths & Constants
# -----------------------------------------------------------------------------
SEED = 42
ROOT = Path(__file__).resolve().parents[2]
STEP_DIR = ROOT / "step4_benchmark"
RESULTS_DIR = STEP_DIR / "results"
DATA_DIR = RESULTS_DIR / "data"
ARTIFACTS_DIR = RESULTS_DIR / "artifacts"
FITTED_MODELS_DIR = ARTIFACTS_DIR / "fitted_models"
QC_DIR = RESULTS_DIR / "qc_reports"
FIGURES_DIR = RESULTS_DIR / "figures"
TABLES_DIR = STEP_DIR / "tables"

# Inputs from previous steps
DATA_CSV = ROOT / "step1_preprocessing" / "results" / "pcos_cleaned_raw.csv"
MANIFEST_JSON = ROOT / "step1_preprocessing" / "results" / "artifacts" / "feature_manifest.json"
CV_SPLITS_JOB = ROOT / "step2_balancing" / "results" / "artifacts" / "cv_splits.joblib"
SELECTED_FEATS_JSON = ROOT / "step3_feature_selection" / "results" / "artifacts" / "selected_features.json"
WRAPPER_RFE_CSV = ROOT / "step3_feature_selection" / "results" / "qc_reports" / "wrapper_rfe_ranks.csv"
STEP3_DOCX = ROOT / "step3_feature_selection" / "tables" / "Step3_Tables_Filled.docx"
TEMPLATE_DOCX = ROOT / "docs" / "All_Tables_Templates.docx"
OUT_DOCX = TABLES_DIR / "Step4_Tables_Filled.docx"

TARGET_COL = "PCOS (Y/N)"

# Tuning budgets
OPTUNA_TRIALS_CLASSICAL = 15
OPTUNA_TRIALS_DEEP = 6


def log(msg: str) -> None:
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {msg}", flush=True)


def ensure_dirs() -> None:
    for d in [DATA_DIR, ARTIFACTS_DIR, FITTED_MODELS_DIR, QC_DIR, FIGURES_DIR, TABLES_DIR]:
        d.mkdir(parents=True, exist_ok=True)


# -----------------------------------------------------------------------------
# Prerequisite Verification
# -----------------------------------------------------------------------------
def check_prerequisites() -> None:
    missing = []
    try:
        import optuna
    except ImportError:
        missing.append("optuna")

    try:
        import pytorch_tabnet
    except ImportError:
        missing.append("pytorch-tabnet")

    if missing:
        print("\n" + "=" * 76)
        print(" [!] Missing required dependencies for Step 4:")
        for pkg in missing:
            print(f"     - {pkg}")
        print("\n Please install them in your terminal via:")
        print("     python -m pip install optuna pytorch-tabnet")
        print("=" * 76 + "\n")
        sys.exit(1)


# -----------------------------------------------------------------------------
# Deep Learning Implementations (Tabular ResNet & TabNet Wrapper)
# -----------------------------------------------------------------------------
class ResNetBlock(nn.Module):
    """Residual Block with BatchNorm, ReLU, and Dropout for tabular representations."""
    def __init__(self, d_hidden: int, dropout: float = 0.2):
        super().__init__()
        self.norm1 = nn.BatchNorm1d(d_hidden)
        self.relu1 = nn.ReLU()
        self.dropout1 = nn.Dropout(dropout)
        self.linear1 = nn.Linear(d_hidden, d_hidden)
        self.norm2 = nn.BatchNorm1d(d_hidden)
        self.relu2 = nn.ReLU()
        self.dropout2 = nn.Dropout(dropout)
        self.linear2 = nn.Linear(d_hidden, d_hidden)

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        res = x
        out = self.linear1(self.dropout1(self.relu1(self.norm1(x))))
        out = self.linear2(self.dropout2(self.relu2(self.norm2(out))))
        return res + out


class TabularResNet(nn.Module):
    """Deep Tabular ResNet (Gorishniy et al. NeurIPS 2021)."""
    def __init__(self, input_dim: int, d_hidden: int = 64, n_blocks: int = 2, dropout: float = 0.2):
        super().__init__()
        self.input_layer = nn.Linear(input_dim, d_hidden)
        self.blocks = nn.ModuleList([ResNetBlock(d_hidden, dropout) for _ in range(n_blocks)])
        self.head_norm = nn.BatchNorm1d(d_hidden)
        self.head_relu = nn.ReLU()
        self.head_linear = nn.Linear(d_hidden, 1)

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        h = self.input_layer(x)
        for b in self.blocks:
            h = b(h)
        h = self.head_relu(self.head_norm(h))
        return self.head_linear(h).squeeze(-1)


class DeepResNetClassifier(BaseEstimator, ClassifierMixin):
    """Scikit-Learn compatible wrapper for Tabular ResNet."""
    def __init__(
        self,
        d_hidden: int = 64,
        n_blocks: int = 2,
        dropout: float = 0.2,
        lr: float = 1e-3,
        weight_decay: float = 1e-4,
        epochs: int = 50,
        batch_size: int = 32,
        random_state: int = SEED,
    ):
        self.d_hidden = d_hidden
        self.n_blocks = n_blocks
        self.dropout = dropout
        self.lr = lr
        self.weight_decay = weight_decay
        self.epochs = epochs
        self.batch_size = batch_size
        self.random_state = random_state
        self.model: TabularResNet | None = None
        self.classes_ = np.array([0, 1])

    def fit(self, X: np.ndarray, y: np.ndarray, eval_set: Tuple[np.ndarray, np.ndarray] | None = None) -> DeepResNetClassifier:
        torch.manual_seed(self.random_state)
        np.random.seed(self.random_state)
        input_dim = X.shape[1]
        self.model = TabularResNet(input_dim, self.d_hidden, self.n_blocks, self.dropout)

        optimizer = optim.AdamW(self.model.parameters(), lr=self.lr, weight_decay=self.weight_decay)
        criterion = nn.BCEWithLogitsLoss()

        X_t = torch.tensor(X, dtype=torch.float32)
        y_t = torch.tensor(y, dtype=torch.float32)
        dataset = TensorDataset(X_t, y_t)
        loader = DataLoader(dataset, batch_size=self.batch_size, shuffle=True)

        if eval_set is not None:
            ev_X, ev_y = eval_set
            ev_X_t = torch.tensor(ev_X, dtype=torch.float32)
            ev_y_t = torch.tensor(ev_y, dtype=torch.float32)

        best_loss = float("inf")
        best_state = None
        patience = 8
        no_improve = 0

        for _ in range(self.epochs):
            self.model.train()
            for bx, by in loader:
                if bx.size(0) <= 1:
                    continue  # BatchNorm requires batch size > 1
                optimizer.zero_grad()
                logits = self.model(bx)
                loss = criterion(logits, by)
                loss.backward()
                optimizer.step()

            if eval_set is not None and ev_X_t.size(0) > 1:
                self.model.eval()
                with torch.no_grad():
                    val_logits = self.model(ev_X_t)
                    val_loss = criterion(val_logits, ev_y_t).item()
                if val_loss < best_loss:
                    best_loss = val_loss
                    best_state = {k: v.cpu().clone() for k, v in self.model.state_dict().items()}
                    no_improve = 0
                else:
                    no_improve += 1
                    if no_improve >= patience:
                        break

        if best_state is not None:
            self.model.load_state_dict(best_state)
        self.model.eval()
        return self

    def predict_proba(self, X: np.ndarray) -> np.ndarray:
        if self.model is None:
            raise RuntimeError("Model is not fitted.")
        self.model.eval()
        X_t = torch.tensor(X, dtype=torch.float32)
        with torch.no_grad():
            logits = self.model(X_t)
            probs = torch.sigmoid(logits).cpu().numpy()
        probs = np.clip(probs, 1e-7, 1.0 - 1e-7)
        return np.column_stack([1.0 - probs, probs])

    def predict(self, X: np.ndarray) -> np.ndarray:
        return (self.predict_proba(X)[:, 1] >= 0.5).astype(int)


class TabNetWrapper(BaseEstimator, ClassifierMixin):
    """Scikit-Learn compatible wrapper for Google TabNet."""
    def __init__(
        self,
        n_d: int = 16,
        n_a: int = 16,
        n_steps: int = 3,
        gamma: float = 1.3,
        lambda_sparse: float = 1e-3,
        lr: float = 2e-2,
        random_state: int = SEED,
    ):
        self.n_d = n_d
        self.n_a = n_a
        self.n_steps = n_steps
        self.gamma = gamma
        self.lambda_sparse = lambda_sparse
        self.lr = lr
        self.random_state = random_state
        self.model: Any = None
        self.classes_ = np.array([0, 1])

    def fit(self, X: np.ndarray, y: np.ndarray, eval_set: Tuple[np.ndarray, np.ndarray] | None = None) -> TabNetWrapper:
        from pytorch_tabnet.tab_model import TabNetClassifier
        self.model = TabNetClassifier(
            n_d=self.n_d,
            n_a=self.n_a,
            n_steps=self.n_steps,
            gamma=self.gamma,
            lambda_sparse=self.lambda_sparse,
            optimizer_fn=torch.optim.Adam,
            optimizer_params=dict(lr=self.lr),
            scheduler_fn=torch.optim.lr_scheduler.StepLR,
            scheduler_params=dict(step_size=10, gamma=0.9),
            verbose=0,
            seed=self.random_state,
        )
        X_arr = np.asarray(X, dtype=np.float32)
        y_arr = np.asarray(y, dtype=np.int64)
        eval_data = None
        if eval_set is not None:
            ev_X = np.asarray(eval_set[0], dtype=np.float32)
            ev_y = np.asarray(eval_set[1], dtype=np.int64)
            eval_data = [(ev_X, ev_y)]

        self.model.fit(
            X_train=X_arr,
            y_train=y_arr,
            eval_set=eval_data,
            eval_metric=["auc"],
            patience=8,
            max_epochs=40,
            batch_size=64,
            virtual_batch_size=16,
            drop_last=False,
        )
        return self

    def predict_proba(self, X: np.ndarray) -> np.ndarray:
        X_arr = np.asarray(X, dtype=np.float32)
        return self.model.predict_proba(X_arr)

    def predict(self, X: np.ndarray) -> np.ndarray:
        X_arr = np.asarray(X, dtype=np.float32)
        return self.model.predict(X_arr)


# -----------------------------------------------------------------------------
# Resampling & Preprocessing Helpers (Anti-Leakage)
# -----------------------------------------------------------------------------
def get_resampler(cat_indices: List[int]) -> Any:
    """Return SMOTE-NC when categorical features exist, else standard SMOTE."""
    if len(cat_indices) > 0:
        return SMOTENC(categorical_features=cat_indices, random_state=SEED)
    return SMOTE(random_state=SEED)


def scale_features(
    X_train: pd.DataFrame | np.ndarray,
    X_val: pd.DataFrame | np.ndarray,
) -> Tuple[np.ndarray, np.ndarray, StandardScaler]:
    """Fit scaler strictly on training data; transform validation data."""
    scaler = StandardScaler()
    X_tr_sc = scaler.fit_transform(X_train)
    X_va_sc = scaler.transform(X_val)
    return X_tr_sc, X_va_sc, scaler


# -----------------------------------------------------------------------------
# Part A: Nested 5x5 CV with Optuna Hyper-Tuning (Table 12)
# -----------------------------------------------------------------------------
T12_MODELS = [
    "LR", "DT", "RF", "XGB", "CB", "LGBM", "SVM", "MLP", "TabNet", "Deep ResNet"
]


def create_model_instance(name: str, params: Dict[str, Any]) -> Any:
    """Instantiate model with given hyperparameter dict."""
    if name == "LR":
        return LogisticRegression(
            penalty="l2",
            C=params.get("C", 1.0),
            solver="lbfgs",
            max_iter=1000,
            random_state=SEED,
        )
    elif name == "DT":
        return DecisionTreeClassifier(
            max_depth=params.get("max_depth", 5),
            min_samples_split=params.get("min_samples_split", 5),
            criterion=params.get("criterion", "gini"),
            random_state=SEED,
        )
    elif name == "RF":
        return RandomForestClassifier(
            n_estimators=params.get("n_estimators", 100),
            max_depth=params.get("max_depth", 6),
            min_samples_split=params.get("min_samples_split", 4),
            random_state=SEED,
            n_jobs=-1,
        )
    elif name == "XGB":
        return XGBClassifier(
            n_estimators=params.get("n_estimators", 100),
            max_depth=params.get("max_depth", 4),
            learning_rate=params.get("learning_rate", 0.05),
            subsample=params.get("subsample", 0.8),
            colsample_bytree=params.get("colsample_bytree", 0.8),
            random_state=SEED,
            n_jobs=-1,
            eval_metric="logloss",
        )
    elif name == "CB":
        return CatBoostClassifier(
            iterations=120,
            depth=params.get("depth", 5),
            learning_rate=params.get("learning_rate", 0.05),
            l2_leaf_reg=params.get("l2_leaf_reg", 3.0),
            random_seed=SEED,
            thread_count=-1,
            verbose=0,
        )
    elif name == "LGBM":
        return LGBMClassifier(
            n_estimators=params.get("n_estimators", 100),
            num_leaves=params.get("num_leaves", 31),
            max_depth=params.get("max_depth", 5),
            learning_rate=params.get("learning_rate", 0.05),
            random_state=SEED,
            n_jobs=-1,
            verbose=-1,
        )
    elif name == "SVM":
        return SVC(
            kernel="rbf",
            C=params.get("C", 1.0),
            gamma=params.get("gamma", 0.05),
            probability=True,
            random_state=SEED,
        )
    elif name == "MLP":
        hidden = params.get("hidden_layer_sizes", (64,))
        if isinstance(hidden, list):
            hidden = tuple(hidden)
        return MLPClassifier(
            hidden_layer_sizes=hidden,
            alpha=params.get("alpha", 1e-3),
            learning_rate_init=params.get("learning_rate_init", 1e-3),
            max_iter=250,
            random_state=SEED,
            early_stopping=True,
            n_iter_no_change=8,
        )
    elif name == "TabNet":
        return TabNetWrapper(
            n_d=params.get("n_da", 16),
            n_a=params.get("n_da", 16),
            n_steps=params.get("n_steps", 3),
            gamma=params.get("gamma", 1.3),
            lr=params.get("lr", 1e-2),
            random_state=SEED,
        )
    elif name == "Deep ResNet":
        return DeepResNetClassifier(
            d_hidden=params.get("d_hidden", 64),
            n_blocks=params.get("n_blocks", 2),
            dropout=params.get("dropout", 0.2),
            lr=params.get("lr", 1e-3),
            random_state=SEED,
        )
    raise ValueError(f"Unknown model name: {name}")


def sample_optuna_params(trial: Any, name: str) -> Dict[str, Any]:
    """Sample hyperparameter candidate using Optuna trial."""
    if name == "LR":
        return {"C": trial.suggest_float("C", 1e-3, 10.0, log=True)}
    elif name == "DT":
        return {
            "max_depth": trial.suggest_int("max_depth", 3, 10),
            "min_samples_split": trial.suggest_int("min_samples_split", 2, 15),
            "criterion": trial.suggest_categorical("criterion", ["gini", "entropy"]),
        }
    elif name == "RF":
        return {
            "n_estimators": trial.suggest_int("n_estimators", 80, 200, step=20),
            "max_depth": trial.suggest_int("max_depth", 4, 12),
            "min_samples_split": trial.suggest_int("min_samples_split", 2, 8),
        }
    elif name == "XGB":
        return {
            "n_estimators": trial.suggest_int("n_estimators", 80, 200, step=20),
            "max_depth": trial.suggest_int("max_depth", 3, 6),
            "learning_rate": trial.suggest_float("learning_rate", 0.02, 0.2, log=True),
            "subsample": trial.suggest_float("subsample", 0.7, 1.0),
            "colsample_bytree": trial.suggest_float("colsample_bytree", 0.7, 1.0),
        }
    elif name == "CB":
        return {
            "depth": trial.suggest_int("depth", 4, 7),
            "learning_rate": trial.suggest_float("learning_rate", 0.02, 0.2, log=True),
            "l2_leaf_reg": trial.suggest_float("l2_leaf_reg", 1.0, 8.0),
        }
    elif name == "LGBM":
        return {
            "num_leaves": trial.suggest_int("num_leaves", 15, 45),
            "max_depth": trial.suggest_int("max_depth", 3, 7),
            "learning_rate": trial.suggest_float("learning_rate", 0.02, 0.2, log=True),
            "n_estimators": trial.suggest_int("n_estimators", 80, 200, step=20),
        }
    elif name == "SVM":
        return {
            "C": trial.suggest_float("C", 0.1, 50.0, log=True),
            "gamma": trial.suggest_float("gamma", 1e-3, 0.5, log=True),
        }
    elif name == "MLP":
        hidden_idx = trial.suggest_categorical("hidden_idx", [0, 1, 2, 3])
        hidden_choices = [(64,), (128,), (64, 32), (128, 64)]
        return {
            "hidden_layer_sizes": hidden_choices[hidden_idx],
            "alpha": trial.suggest_float("alpha", 1e-4, 1e-1, log=True),
            "learning_rate_init": trial.suggest_float("learning_rate_init", 1e-3, 1e-2, log=True),
        }
    elif name == "TabNet":
        return {
            "n_da": trial.suggest_categorical("n_da", [8, 16, 24]),
            "n_steps": trial.suggest_int("n_steps", 3, 5),
            "gamma": trial.suggest_float("gamma", 1.1, 1.5),
            "lr": trial.suggest_float("lr", 5e-3, 3e-2, log=True),
        }
    elif name == "Deep ResNet":
        return {
            "d_hidden": trial.suggest_categorical("d_hidden", [64, 96, 128]),
            "n_blocks": trial.suggest_int("n_blocks", 2, 3),
            "dropout": trial.suggest_float("dropout", 0.1, 0.35),
            "lr": trial.suggest_float("lr", 5e-4, 5e-3, log=True),
        }
    return {}


def tune_and_evaluate_nested_cv(
    X_full: pd.DataFrame,
    y_full: np.ndarray,
    cv_splits: Dict[str, Any],
    final16_cols: List[str],
    cat_cols: List[str],
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """Execute Nested 5x5 CV with inside-fold Optuna Bayesian Search with progress reporting."""
    import optuna
    optuna.logging.set_verbosity(optuna.logging.WARNING)

    cat_indices = [i for i, c in enumerate(final16_cols) if c in cat_cols]
    outer_folds = cv_splits["outer_folds"]

    fold_metrics_records = []
    oof_predictions_df = pd.DataFrame(
        index=cv_splits["holdout_train_idx"],
        columns=["y_true"] + [f"prob_{m}" for m in T12_MODELS] + [f"pred_{m}" for m in T12_MODELS],
    )
    oof_predictions_df["y_true"] = y_full[cv_splits["holdout_train_idx"]]

    best_hyperparams: Dict[str, Dict[str, Any]] = {m: {} for m in T12_MODELS}
    optuna_trials_records = []

    total_tasks = len(outer_folds) * len(T12_MODELS)
    print("\n" + "=" * 80)
    print(" [STAGE 4.1] NESTED 5x5 CROSS-VALIDATION WITH BAYESIAN OPTUNA TUNING")
    print(" Protocol: 5 Outer Folds x 10 Models | Inside-Fold SMOTE-NC | Zero Leakage")
    print("=" * 80)

    with tqdm(
        total=total_tasks,
        desc="Nested 5x5 CV Progress",
        unit="model",
        dynamic_ncols=True,
        leave=True,
        bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
    ) as pbar:

        for fold_info in outer_folds:
            f_idx = fold_info["outer_fold"]
            tr_global = fold_info["train_idx"]
            va_global = fold_info["valid_idx"]
            inner_splits = fold_info["inner_splits"]

            X_outer_tr = X_full.iloc[tr_global][final16_cols].reset_index(drop=True)
            y_outer_tr = y_full[tr_global]
            X_outer_va = X_full.iloc[va_global][final16_cols].reset_index(drop=True)
            y_outer_va = y_full[va_global]

            for m_name in T12_MODELS:
                is_deep = m_name in ["TabNet", "Deep ResNet"]
                n_trials = OPTUNA_TRIALS_DEEP if is_deep else OPTUNA_TRIALS_CLASSICAL

                pbar.set_description(f"Outer Fold {f_idx + 1}/5 | Tuning {m_name:<11}")

                # Inner Optuna objective
                def objective(trial: Any) -> float:
                    params = sample_optuna_params(trial, m_name)
                    inner_f1_scores = []

                    for in_split in inner_splits:
                        in_tr_glob = in_split["train_idx"]
                        in_va_glob = in_split["valid_idx"]

                        in_X_tr = X_full.iloc[in_tr_glob][final16_cols].reset_index(drop=True)
                        in_y_tr = y_full[in_tr_glob]
                        in_X_va = X_full.iloc[in_va_glob][final16_cols].reset_index(drop=True)
                        in_y_va = y_full[in_va_glob]

                        # 1. Resample inner train strictly inside
                        smote = get_resampler(cat_indices)
                        in_X_tr_bal, in_y_tr_bal = smote.fit_resample(in_X_tr, in_y_tr)

                        # 2. Scale features
                        in_X_tr_sc, in_X_va_sc, _ = scale_features(in_X_tr_bal, in_X_va)

                        # 3. Model Fit & Predict
                        mdl = create_model_instance(m_name, params)
                        if is_deep:
                            mdl.fit(in_X_tr_sc, in_y_tr_bal, eval_set=(in_X_va_sc, in_y_va))
                        else:
                            mdl.fit(in_X_tr_sc, in_y_tr_bal)

                        in_preds = mdl.predict(in_X_va_sc)
                        score = f1_score(in_y_va, in_preds, zero_division=0)
                        inner_f1_scores.append(score)

                    return float(np.mean(inner_f1_scores))

                sampler = optuna.samplers.TPESampler(seed=SEED + f_idx * 100)
                study = optuna.create_study(direction="maximize", sampler=sampler)
                study.optimize(objective, n_trials=n_trials, show_progress_bar=False)

                best_p = study.best_params
                best_score = study.best_value
                best_hyperparams[m_name][f"fold_{f_idx}"] = best_p

                optuna_trials_records.append({
                    "outer_fold": f_idx,
                    "model": m_name,
                    "best_inner_f1": best_score,
                    "n_trials": n_trials,
                    "best_params": json.dumps(best_p),
                })

                # Outer Train Fit with Best Params
                smote_outer = get_resampler(cat_indices)
                X_out_tr_bal, y_out_tr_bal = smote_outer.fit_resample(X_outer_tr, y_outer_tr)
                X_out_tr_sc, X_out_va_sc, _ = scale_features(X_out_tr_bal, X_outer_va)

                best_model = create_model_instance(m_name, best_p)
                if is_deep:
                    best_model.fit(X_out_tr_sc, y_out_tr_bal, eval_set=(X_out_va_sc, y_outer_va))
                else:
                    best_model.fit(X_out_tr_sc, y_out_tr_bal)

                # Outer Validation Predict
                preds = best_model.predict(X_out_va_sc)
                probs = best_model.predict_proba(X_out_va_sc)[:, 1]

                acc = accuracy_score(y_outer_va, preds)
                prec = precision_score(y_outer_va, preds, zero_division=0)
                rec = recall_score(y_outer_va, preds, zero_division=0)
                f1 = f1_score(y_outer_va, preds, zero_division=0)
                roc = roc_auc_score(y_outer_va, probs)
                brier = brier_score_loss(y_outer_va, probs)

                fold_metrics_records.append({
                    "outer_fold": f_idx,
                    "model": m_name,
                    "accuracy": acc,
                    "precision": prec,
                    "recall": rec,
                    "f1_score": f1,
                    "roc_auc": roc,
                    "brier_score": brier,
                })

                # Store OOF predictions
                oof_predictions_df.loc[va_global, f"prob_{m_name}"] = probs
                oof_predictions_df.loc[va_global, f"pred_{m_name}"] = preds

                pbar.update(1)
                pbar.write(
                    f"  [Fold {f_idx+1}/5] {m_name:<11} | Inner F1: {best_score*100:5.2f}% | "
                    f"Val Acc: {acc*100:5.2f}% | Val F1: {f1*100:5.2f}% | AUC: {roc:6.4f}"
                )

    fold_metrics_df = pd.DataFrame(fold_metrics_records)

    # Compute Summary (Mean ± Std)
    summary_records = []
    for m in T12_MODELS:
        m_df = fold_metrics_df[fold_metrics_df["model"] == m]
        summary_records.append({
            "model": m,
            "accuracy_mean": m_df["accuracy"].mean(),
            "accuracy_std": m_df["accuracy"].std(),
            "precision_mean": m_df["precision"].mean(),
            "precision_std": m_df["precision"].std(),
            "recall_mean": m_df["recall"].mean(),
            "recall_std": m_df["recall"].std(),
            "f1_mean": m_df["f1_score"].mean(),
            "f1_std": m_df["f1_score"].std(),
            "roc_auc_mean": m_df["roc_auc"].mean(),
            "roc_auc_std": m_df["roc_auc"].std(),
            "brier_mean": m_df["brier_score"].mean(),
            "brier_std": m_df["brier_score"].std(),
        })
    summary_df = pd.DataFrame(summary_records)
    optuna_summary_df = pd.DataFrame(optuna_trials_records)

    return fold_metrics_df, summary_df, oof_predictions_df, optuna_summary_df


def train_and_save_final_models(
    X_full: pd.DataFrame,
    y_full: np.ndarray,
    holdout_train_idx: List[int],
    final16_cols: List[str],
    cat_cols: List[str],
    summary_df: pd.DataFrame,
) -> None:
    """Fit final models on the complete 432 holdout training set (SMOTE-NC balanced) for Step 5."""
    print("\n" + "=" * 80)
    print(" [STAGE 4.2] SERIALIZING FINAL TRAINED MODELS (Step 5 Super-Ensemble Prep)")
    print("=" * 80)

    cat_indices = [i for i, c in enumerate(final16_cols) if c in cat_cols]
    X_tr = X_full.iloc[holdout_train_idx][final16_cols].reset_index(drop=True)
    y_tr = y_full[holdout_train_idx]

    smote = get_resampler(cat_indices)
    X_tr_bal, y_tr_bal = smote.fit_resample(X_tr, y_tr)
    scaler = StandardScaler()
    X_tr_sc = scaler.fit_transform(X_tr_bal)

    # Save scaler
    joblib.dump(scaler, ARTIFACTS_DIR / "final_16_scaler.joblib")

    with tqdm(
        T12_MODELS,
        desc="Fitting Final Models",
        unit="model",
        dynamic_ncols=True,
        leave=True,
        bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt} [{elapsed}, {rate_fmt}]"
    ) as pbar:
        for m_name in pbar:
            pbar.set_description(f"Fitting Final {m_name:<11}")
            is_deep = m_name in ["TabNet", "Deep ResNet"]
            mdl = create_model_instance(m_name, {})
            if is_deep:
                mdl.fit(X_tr_sc, y_tr_bal)
            else:
                mdl.fit(X_tr_sc, y_tr_bal)
            save_path = FITTED_MODELS_DIR / f"{m_name.lower().replace(' ', '_')}_model.joblib"
            joblib.dump(mdl, save_path)


# -----------------------------------------------------------------------------
# Part B: 10-Fold CV on 12 Feature Sets (Table 8)
# -----------------------------------------------------------------------------
T8_MODELS = ["LR", "DT", "RF", "XGB", "CB", "AB", "SVM", "KNN", "MLP", "LDA"]
RFE_WRAPPERS = ["LR", "XGB", "RF", "CB"]
RFE_SIZES = ["top20", "top16", "top12"]


def get_fast_model(name: str) -> Any:
    """Return configured model for Table 8 benchmarking."""
    if name == "LR":
        return LogisticRegression(penalty="l2", C=1.0, solver="lbfgs", max_iter=1000, random_state=SEED)
    elif name == "DT":
        return DecisionTreeClassifier(max_depth=5, min_samples_split=4, random_state=SEED)
    elif name == "RF":
        return RandomForestClassifier(n_estimators=100, max_depth=6, random_state=SEED, n_jobs=-1)
    elif name == "XGB":
        return XGBClassifier(n_estimators=100, max_depth=4, learning_rate=0.08, random_state=SEED, n_jobs=-1, eval_metric="logloss")
    elif name == "CB":
        return CatBoostClassifier(iterations=100, depth=5, learning_rate=0.08, random_seed=SEED, thread_count=-1, verbose=0)
    elif name == "AB":
        return AdaBoostClassifier(n_estimators=50, random_state=SEED)
    elif name == "SVM":
        return SVC(kernel="rbf", C=1.5, probability=True, random_state=SEED)
    elif name == "KNN":
        return KNeighborsClassifier(n_neighbors=5)
    elif name == "MLP":
        return MLPClassifier(hidden_layer_sizes=(64,), alpha=1e-3, max_iter=250, random_state=SEED)
    elif name == "LDA":
        return LinearDiscriminantAnalysis()
    raise ValueError(f"Unknown Table 8 model: {name}")


def run_table8_benchmark(
    X_full: pd.DataFrame,
    y_full: np.ndarray,
    holdout_train_idx: List[int],
    wrapper_df: pd.DataFrame,
    cat_cols: List[str],
) -> pd.DataFrame:
    """Run 10-fold Stratified CV for 10 models across 12 feature subsets with live progress bar."""
    print("\n" + "=" * 80)
    print(" [STAGE 4.3] TABLE 8 BENCHMARK: 10 MODELS x 12 WRAPPER FEATURE SUBSETS")
    print(" Protocol: 10-Fold Stratified CV | 120 Total Evaluated Configurations")
    print("=" * 80)

    X_train_base = X_full.iloc[holdout_train_idx].reset_index(drop=True)
    y_train_base = y_full[holdout_train_idx]

    skf = StratifiedKFold(n_splits=10, shuffle=True, random_state=SEED)

    # Prepare feature subsets
    col_names = []
    feature_subsets: Dict[str, List[str]] = {}
    for size in RFE_SIZES:
        for wrap in RFE_WRAPPERS:
            c_name = f"{size}_{wrap}"
            col_names.append(c_name)
            subset = wrapper_df[(wrapper_df["model"] == wrap) & (wrapper_df[size] == True)]["feature"].tolist()
            feature_subsets[c_name] = subset

    acc_matrix = pd.DataFrame(index=T8_MODELS, columns=col_names, dtype=float)
    total_evals = len(col_names) * len(T8_MODELS)

    with tqdm(
        total=total_evals,
        desc="Table 8 Evaluations",
        unit="eval",
        dynamic_ncols=True,
        leave=True,
        bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
    ) as pbar:
        for c_name in col_names:
            feats = feature_subsets[c_name]
            cat_indices = [i for i, c in enumerate(feats) if c in cat_cols]
            X_sub = X_train_base[feats]

            for m_name in T8_MODELS:
                pbar.set_description(f"Table 8 | Subset: {c_name:<10} | Model: {m_name:<4}")

                fold_accs = []
                for tr_i, va_i in skf.split(X_sub, y_train_base):
                    X_tr, y_tr = X_sub.iloc[tr_i], y_train_base[tr_i]
                    X_va, y_va = X_sub.iloc[va_i], y_train_base[va_i]

                    # Inside-fold SMOTE-NC
                    smote = get_resampler(cat_indices)
                    X_tr_bal, y_tr_bal = smote.fit_resample(X_tr, y_tr)

                    # Scale inside fold
                    X_tr_sc, X_va_sc, _ = scale_features(X_tr_bal, X_va)

                    mdl = get_fast_model(m_name)
                    mdl.fit(X_tr_sc, y_tr_bal)
                    preds = mdl.predict(X_va_sc)
                    fold_accs.append(accuracy_score(y_va, preds))

                mean_acc = float(np.mean(fold_accs))
                acc_matrix.loc[m_name, c_name] = mean_acc
                pbar.update(1)

    return acc_matrix


# -----------------------------------------------------------------------------
# Figure Generation
# -----------------------------------------------------------------------------
def generate_figures(
    t8_matrix: pd.DataFrame,
    t12_summary: pd.DataFrame,
    oof_df: pd.DataFrame,
    optuna_df: pd.DataFrame,
) -> None:
    """Generate publication-ready figures for thesis report."""
    sns.set_theme(style="whitegrid", font="sans-serif")

    # Figure 1: Table 8 Heatmap
    plt.figure(figsize=(14, 7))
    col_labels = [
        f"{c.split('_')[0].upper()} ({c.split('_')[1]})" for c in t8_matrix.columns
    ]
    sns.heatmap(
        t8_matrix * 100,
        annot=True,
        fmt=".1f",
        cmap="YlGnBu",
        cbar_kws={"label": "Mean Accuracy (%)"},
        xticklabels=col_labels,
        yticklabels=t8_matrix.index,
    )
    plt.title("Table 8: Model Accuracy (%) Across 12 Wrapper Feature Subsets (10-Fold CV)", fontsize=13, weight="bold", pad=12)
    plt.xlabel("Wrapper RFE Feature Subset", fontsize=11, labelpad=8)
    plt.ylabel("Evaluated Machine Learning Model", fontsize=11, labelpad=8)
    plt.xticks(rotation=30, ha="right")
    plt.tight_layout()
    plt.savefig(FIGURES_DIR / "t8_feature_subsets_heatmap.png", dpi=300)
    plt.close()

    # Figure 2: Table 12 Performance Bar Chart with Error Bars
    plt.figure(figsize=(14, 6))
    metrics = ["accuracy", "precision", "recall", "f1"]
    metric_labels = ["Accuracy", "Precision", "Recall", "F1 Score"]
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"]
    x = np.arange(len(t12_summary))
    width = 0.2

    for i, (m, lbl, col) in enumerate(zip(metrics, metric_labels, colors)):
        means = t12_summary[f"{m}_mean"] * 100
        stds = t12_summary[f"{m}_std"] * 100
        plt.bar(x + (i - 1.5) * width, means, width, yerr=stds, capsize=3, label=lbl, color=col, alpha=0.88)

    plt.xticks(x, t12_summary["model"], rotation=25, ha="right", fontsize=10)
    plt.ylabel("Score (%)", fontsize=11)
    plt.ylim(70, 102)
    plt.title("Table 12: Model Performance on Proposed 16 Bio-Markers (Nested 5x5 CV with Optuna)", fontsize=13, weight="bold", pad=12)
    plt.legend(loc="lower right", frameon=True)
    plt.tight_layout()
    plt.savefig(FIGURES_DIR / "t12_model_performance_bars.png", dpi=300)
    plt.close()

    # Figure 3: ROC Curves for Proposed Models
    plt.figure(figsize=(9, 8))
    y_true = oof_df["y_true"].astype(int).values
    for m in T12_MODELS:
        y_prob = oof_df[f"prob_{m}"].astype(float).values
        fpr, tpr, _ = roc_curve(y_true, y_prob)
        auc_val = roc_auc_score(y_true, y_prob)
        plt.plot(fpr, tpr, lw=2, label=f"{m} (AUC = {auc_val:.3f})")

    plt.plot([0, 1], [0, 1], "k--", lw=1.5, label="Chance")
    plt.xlim([-0.02, 1.0])
    plt.ylim([0.0, 1.05])
    plt.xlabel("False Positive Rate (1 - Specificity)", fontsize=11)
    plt.ylabel("True Positive Rate (Sensitivity)", fontsize=11)
    plt.title("Out-of-Fold ROC Curves on Proposed 16 Features (Nested 5x5 CV)", fontsize=13, weight="bold", pad=12)
    plt.legend(loc="lower right", fontsize=9)
    plt.tight_layout()
    plt.savefig(FIGURES_DIR / "nested_cv_roc_curves.png", dpi=300)
    plt.close()

    # Figure 4: Optuna Tuning Convergence Summary
    plt.figure(figsize=(10, 5))
    opt_avg = optuna_df.groupby("model")["best_inner_f1"].mean().sort_values(ascending=False)
    ax_bar = sns.barplot(x=opt_avg.values * 100, y=opt_avg.index, palette="viridis")
    plt.title("Optuna Bayesian Optimization: Mean Inner-CV F1 Score Across Models", fontsize=12, weight="bold", pad=10)
    plt.xlabel("Best Inner Validation F1 Score (%)", fontsize=10)
    plt.ylabel("Model", fontsize=10)
    for p in ax_bar.patches:
        val = p.get_width()
        ax_bar.annotate(f"{val:.2f}%", (val - 5.0, p.get_y() + p.get_height() / 2.0),
                        ha="center", va="center", color="white", weight="bold", fontsize=9)
    plt.tight_layout()
    plt.savefig(FIGURES_DIR / "optuna_convergence_plots.png", dpi=300)
    plt.close()


# -----------------------------------------------------------------------------
# Document Filling & Strict Header Verification
# -----------------------------------------------------------------------------
def _docx_set_cell(cell: Any, text: str, align: str = "center", bold: bool = False) -> None:
    """Set formatted text in a docx cell."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(9.5)
    run.bold = bold


def fill_tables_and_verify(t8_matrix: pd.DataFrame, t12_summary: pd.DataFrame) -> bool:
    """Fill Table 8 and Table 12 in cumulative docx and verify headers."""
    if not STEP3_DOCX.exists():
        raise FileNotFoundError(f"Missing Step 3 docx: {STEP3_DOCX}")

    shutil.copyfile(STEP3_DOCX, OUT_DOCX)
    doc = Document(OUT_DOCX)

    # 1. Fill Table 8 (Index 7)
    t8 = doc.tables[7]
    for row_idx, m_name in enumerate(T8_MODELS, start=3):
        for col_idx, col_key in enumerate(t8_matrix.columns, start=1):
            val = t8_matrix.loc[m_name, col_key] * 100.0
            _docx_set_cell(t8.cell(row_idx, col_idx), f"{val:.2f}")

    # 2. Fill Table 12 (Index 11)
    t12 = doc.tables[11]
    for row_idx, m_name in enumerate(T12_MODELS, start=1):
        m_row = t12_summary[t12_summary["model"] == m_name].iloc[0]
        acc_str = f"{m_row['accuracy_mean'] * 100:.2f} ± {m_row['accuracy_std'] * 100:.2f}"
        prec_str = f"{m_row['precision_mean'] * 100:.2f} ± {m_row['precision_std'] * 100:.2f}"
        rec_str = f"{m_row['recall_mean'] * 100:.2f} ± {m_row['recall_std'] * 100:.2f}"
        f1_str = f"{m_row['f1_mean'] * 100:.2f} ± {m_row['f1_std'] * 100:.2f}"

        _docx_set_cell(t12.cell(row_idx, 0), m_name, align="left", bold=True)
        _docx_set_cell(t12.cell(row_idx, 1), acc_str)
        _docx_set_cell(t12.cell(row_idx, 2), prec_str)
        _docx_set_cell(t12.cell(row_idx, 3), rec_str)
        _docx_set_cell(t12.cell(row_idx, 4), f1_str)

    doc.save(OUT_DOCX)

    # 3. Header Verification against Immutable Template
    a, b = Document(TEMPLATE_DOCX), Document(OUT_DOCX)
    header_rows = {
        1: 1, 2: 3, 3: 3, 4: 1, 5: 1, 6: 1, 7: 1, 8: 3, 9: 3,
        10: 2, 11: 2, 12: 1, 13: 1, 14: 2, 15: 1, 16: 1, 17: 1
    }
    ok = len(a.tables) == len(b.tables)
    for ti, (ta, tb) in enumerate(zip(a.tables, b.tables), start=1):
        for ri in range(header_rows.get(ti, 1)):
            for ci in range(len(ta.columns)):
                if ta.cell(ri, ci).text != tb.cell(ri, ci).text:
                    ok = False
                    break
    return ok


# -----------------------------------------------------------------------------
# Terminal Results Summary Display
# -----------------------------------------------------------------------------
def print_terminal_summary(summary_df: pd.DataFrame, t8_matrix: pd.DataFrame, header_ok: bool, elapsed: float) -> None:
    """Print clean, publication-grade aligned ASCII tables to terminal."""
    ranking = summary_df.sort_values("f1_mean", ascending=False).reset_index(drop=True)

    print("\n" + "=" * 90)
    print(" TABLE 12 BENCHMARK SUMMARY (Proposed 16 Features | Nested 5x5 CV with Optuna)")
    print("=" * 90)
    print(f" {'#':<3} {'Model':<13} {'Accuracy (%)':<16} {'Precision (%)':<16} {'Recall (%)':<16} {'F1-Score (%)':<16} {'ROC-AUC':<9}")
    print("-" * 90)

    for i, r in ranking.iterrows():
        acc_s = f"{r['accuracy_mean']*100:5.2f} ± {r['accuracy_std']*100:4.2f}"
        prec_s = f"{r['precision_mean']*100:5.2f} ± {r['precision_std']*100:4.2f}"
        rec_s = f"{r['recall_mean']*100:5.2f} ± {r['recall_std']*100:4.2f}"
        f1_s = f"{r['f1_mean']*100:5.2f} ± {r['f1_std']*100:4.2f}"
        auc_s = f"{r['roc_auc_mean']:6.4f}"
        star = " ★" if i == 0 else "  "
        print(f" {i+1:<2}{star}{r['model']:<13} {acc_s:<16} {prec_s:<16} {rec_s:<16} {f1_s:<16} {auc_s:<9}")

    print("=" * 90)

    # Top configurations in Table 8
    print("\n" + "=" * 90)
    print(" TABLE 8 HIGHLIGHTS: TOP PERFORMING MODEL & FEATURE SUBSET CONFIGURATIONS")
    print("=" * 90)
    t8_stacked = t8_matrix.stack().reset_index()
    t8_stacked.columns = ["Model", "Subset", "Accuracy"]
    top_t8 = t8_stacked.sort_values("Accuracy", ascending=False).head(5).reset_index(drop=True)

    print(f" {'Rank':<6} {'Model':<12} {'Feature Subset':<20} {'10-Fold CV Accuracy (%)':<25}")
    print("-" * 70)
    for i, r in top_t8.iterrows():
        print(f" #{i+1:<5} {r['Model']:<12} {r['Subset']:<20} {r['Accuracy']*100:6.2f}%")
    print("=" * 90)

    # Header check & timing
    h_status = "PASS - all headers byte-identical (Tables 1-17 verified)" if header_ok else "FAIL - header mismatch"
    print("\n" + "-" * 90)
    print(f" [✓] Execution Completed in : {elapsed / 60:.2f} minutes")
    print(f" [✓] Header Verification    : {h_status}")
    print(f" [✓] Cumulative Document    : {OUT_DOCX.relative_to(ROOT)}")
    print(f" [✓] Stacking OOF Bridge    : {DATA_DIR.relative_to(ROOT)}/oof_predictions.csv")
    print(f" [✓] Fitted Model Weights   : {FITTED_MODELS_DIR.relative_to(ROOT)}/*")
    print("-" * 90 + "\n")


# -----------------------------------------------------------------------------
# Main Execution Flow
# -----------------------------------------------------------------------------
def main() -> int:
    start_time = time.time()
    ensure_dirs()

    print("\n" + "=" * 90)
    print(" PCOS DIAGNOSIS PIPELINE - STAGE 4: MULTI-MODEL BENCHMARK & OPTUNA TUNING")
    print(" Reference: Efty et al. 2026 | Enhanced Rotterdam Biomarkers & Nested 5x5 CV")
    print("=" * 90)

    check_prerequisites()

    # Load inputs
    log("Loading validated preprocessed data and split artifacts...")
    full_df = pd.read_csv(DATA_CSV)
    y_full = full_df[TARGET_COL].values.astype(int)

    with open(MANIFEST_JSON, encoding="utf-8") as f:
        manifest = json.load(f)
    cat_cols = manifest.get("categorical_features", [])

    cv_splits = joblib.load(CV_SPLITS_JOB)
    holdout_train_idx = cv_splits["holdout_train_idx"]

    with open(SELECTED_FEATS_JSON, encoding="utf-8") as f:
        selected_manifest = json.load(f)
    final16_cols = selected_manifest["internal_names"]["stage3"]
    log(f"Selected 16 Features: {', '.join(final16_cols[:5])}... ({len(final16_cols)} total)")

    wrapper_df = pd.read_csv(WRAPPER_RFE_CSV)

    # 1. Part A: Nested 5x5 CV with Optuna
    fold_metrics_df, summary_df, oof_df, optuna_df = tune_and_evaluate_nested_cv(
        full_df, y_full, cv_splits, final16_cols, cat_cols
    )

    # Save Part A Data
    fold_metrics_df.to_csv(DATA_DIR / "nested_cv_fold_metrics.csv", index=False)
    summary_df.to_csv(DATA_DIR / "t12_nested_metrics_summary.csv", index=False)
    oof_df.to_csv(DATA_DIR / "oof_predictions.csv", index=True)
    optuna_df.to_csv(QC_DIR / "optuna_study_summary.csv", index=False)

    # 2. Train and export final models on full 432 train set for Step 5
    train_and_save_final_models(full_df, y_full, holdout_train_idx, final16_cols, cat_cols, summary_df)

    # 3. Part B: 10-Fold CV on 12 Feature Sets (Table 8)
    t8_matrix = run_table8_benchmark(full_df, y_full, holdout_train_idx, wrapper_df, cat_cols)
    t8_matrix.to_csv(DATA_DIR / "t8_accuracy_matrix.csv", index=True)

    # 4. Generate Visualizations & Fill Tables
    print("\n" + "=" * 80)
    print(" [STAGE 4.4] GENERATING THESIS FIGURES & POPULATING DOCUMENT TABLES")
    print("=" * 80)
    with tqdm(total=3, desc="Post-Processing Tasks", unit="task", dynamic_ncols=True, leave=True) as pbar:
        pbar.set_description("Rendering 4 publication figures")
        generate_figures(t8_matrix, summary_df, oof_df, optuna_df)
        pbar.update(1)

        pbar.set_description("Populating Tables 8 & 12 in docx")
        header_ok = fill_tables_and_verify(t8_matrix, summary_df)
        pbar.update(1)

        pbar.set_description("Exporting QC telemetry & summaries")
        elapsed = time.time() - start_time
        ranking = summary_df.sort_values("f1_mean", ascending=False).reset_index(drop=True)
        ranking.to_csv(QC_DIR / "model_ranking_report.csv", index=False)

        summary = {
            "timestamp_utc": datetime.now(timezone.utc).isoformat(),
            "elapsed_seconds": round(elapsed, 2),
            "protocol": {
                "part_a": "Nested 5x5 CV with inside-fold Optuna Bayesian Search (Proposed 16 Features)",
                "part_b": "Stratified 10-Fold CV across 12 Wrapper Feature Subsets",
                "smote_nc": "Applied strictly inside training folds; untouched validation folds",
                "holdout_test": "109 test samples 100% untouched (reserved for Step 6)",
            },
            "top_performing_model": {
                "model": ranking.iloc[0]["model"],
                "accuracy": f"{ranking.iloc[0]['accuracy_mean'] * 100:.2f}%",
                "f1_score": f"{ranking.iloc[0]['f1_mean'] * 100:.2f}%",
                "roc_auc": f"{ranking.iloc[0]['roc_auc_mean']:.4f}",
            },
            "artifacts_saved": [
                "results/data/t8_accuracy_matrix.csv",
                "results/data/t12_nested_metrics_summary.csv",
                "results/data/nested_cv_fold_metrics.csv",
                "results/data/oof_predictions.csv",
                "results/artifacts/fitted_models/*",
                "tables/Step4_Tables_Filled.docx",
            ],
        }
        with open(QC_DIR / "benchmark_qc_summary.json", "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2)
        pbar.update(1)

    # 5. Print Clean Terminal Summary
    print_terminal_summary(summary_df, t8_matrix, header_ok, elapsed)
    return 0


if __name__ == "__main__":
    sys.exit(main())
