"""Step 2: Data Splitting & Balancing (Nested 5x5 CV + SMOTE-NC).

Implements the proposed leakage-free protocol:
  - untouched stratified 80/20 holdout validation set (Table 2 protocol)
  - on the training partition: nested 5x5 stratified CV
    (outer 5 folds for evaluation, inner 5 folds for tuning)
  - SMOTE-NC applied ONLY inside training folds (zero data leakage)

Single-run script: executing this file regenerates all Step-2 outputs
(results, figures, QC reports, filled Table-2 docx).
"""

from __future__ import annotations

import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

import joblib
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from imblearn.over_sampling import SMOTENC
from sklearn.model_selection import StratifiedKFold, train_test_split

SEED = 42
TEST_SIZE = 0.20
N_OUTER = 5
N_INNER = 5

ROOT = Path(__file__).resolve().parents[2]
STEP1_DIR = ROOT / "step1_preprocessing"
STEP_DIR = ROOT / "step2_balancing"
RESULTS_DIR = STEP_DIR / "results"
ARTIFACTS_DIR = RESULTS_DIR / "artifacts"
QC_DIR = RESULTS_DIR / "qc_reports"
FIG_DIR = RESULTS_DIR / "figures"
BAL_DIR = RESULTS_DIR / "balanced_folds"
SPLIT_DIR = RESULTS_DIR / "split_data"
TABLES_DIR = STEP_DIR / "tables"

STEP1_RAW_CSV = STEP1_DIR / "results" / "pcos_cleaned_raw.csv"
STEP1_MANIFEST = STEP1_DIR / "results" / "artifacts" / "feature_manifest.json"
STEP1_DOCX = STEP1_DIR / "tables" / "Step1_Tables_Filled.docx"
TEMPLATE_DOCX = ROOT / "docs" / "All_Tables_Templates.docx"
OUT_DOCX = TABLES_DIR / "Step2_Tables_Filled.docx"

TARGET = "PCOS (Y/N)"

DOCX_FONT = "Calibri"
DOCX_SIZE = 9


def log(msg: str) -> None:
    print(f"[step2] {msg}", flush=True)


def ensure_dirs() -> None:
    for path in (RESULTS_DIR, ARTIFACTS_DIR, QC_DIR, FIG_DIR, BAL_DIR, SPLIT_DIR, TABLES_DIR):
        path.mkdir(parents=True, exist_ok=True)


def load_data() -> tuple[pd.DataFrame, list[str], list[str]]:
    df = pd.read_csv(STEP1_RAW_CSV)
    manifest = json.loads(STEP1_MANIFEST.read_text(encoding="utf-8"))
    feature_cols = manifest["feature_columns"]
    categorical = manifest["categorical_features"]
    log(f"Loaded {df.shape[0]} rows x {df.shape[1]} cols from Step 1")
    return df, feature_cols, categorical


def make_holdout(df: pd.DataFrame) -> tuple[np.ndarray, np.ndarray]:
    idx = np.arange(len(df))
    train_idx, test_idx = train_test_split(
        idx, test_size=TEST_SIZE, random_state=SEED, stratify=df[TARGET].values
    )
    return np.sort(train_idx), np.sort(test_idx)


def make_nested_folds(
    y_train: np.ndarray, train_idx: np.ndarray
) -> list[dict]:
    outer = StratifiedKFold(n_splits=N_OUTER, shuffle=True, random_state=SEED)
    folds = []
    for k, (o_tr_rel, o_te_rel) in enumerate(outer.split(np.zeros(len(y_train)), y_train)):
        o_tr = train_idx[o_tr_rel]
        o_te = train_idx[o_te_rel]
        y_outer_tr = y_train[o_tr_rel]
        inner = StratifiedKFold(n_splits=N_INNER, shuffle=True, random_state=SEED + k)
        inner_splits = []
        for j, (i_tr_rel, i_va_rel) in enumerate(
            inner.split(np.zeros(len(y_outer_tr)), y_outer_tr)
        ):
            inner_splits.append({
                "inner_fold": j,
                "train_idx": o_tr[i_tr_rel].tolist(),
                "valid_idx": o_tr[i_va_rel].tolist(),
            })
        folds.append({
            "outer_fold": k,
            "train_idx": o_tr.tolist(),
            "valid_idx": o_te.tolist(),
            "inner_splits": inner_splits,
        })
    return folds


def apply_smote_nc(
    X: pd.DataFrame, y: pd.Series, categorical_cols: list[str], seed: int
) -> tuple[pd.DataFrame, pd.Series, int]:
    cat_idx = [X.columns.get_loc(c) for c in categorical_cols if c in X.columns]
    sampler = SMOTENC(
        categorical_features=cat_idx,
        random_state=seed,
        k_neighbors=5,
        sampling_strategy="auto",
    )
    X_res, y_res = sampler.fit_resample(X, y)
    n_synthetic = len(X_res) - len(X)
    return X_res, y_res, n_synthetic


def balancing_report_row(
    name: str, y_before: pd.Series, y_after: pd.Series, n_syn: int
) -> dict:
    return {
        "split": name,
        "n_before": int(len(y_before)),
        "neg_before": int((y_before == 0).sum()),
        "pos_before": int((y_before == 1).sum()),
        "n_after": int(len(y_after)),
        "neg_after": int((y_after == 0).sum()),
        "pos_after": int((y_after == 1).sum()),
        "n_synthetic": int(n_syn),
    }


def make_figures(holdout_before: dict, holdout_after: dict, fold_rows: list[dict]) -> None:
    sns.set_theme(style="whitegrid", context="notebook")

    fig, axes = plt.subplots(1, 2, figsize=(10, 4.5))
    for ax, (title, d) in zip(axes, [("Before SMOTE-NC", holdout_before), ("After SMOTE-NC (train only)", holdout_after)]):
        labels = ["Non-PCOS (0)", "PCOS (1)"]
        values = [d["neg"], d["pos"]]
        sns.barplot(x=labels, y=values, hue=labels, palette=["#4C72B0", "#DD8452"],
                    legend=False, ax=ax)
        ax.set_title(title)
        ax.set_ylabel("Patients")
        for i, v in enumerate(values):
            ax.text(i, v + 3, str(v), ha="center", fontsize=10)
    fig.suptitle("Holdout train partition: class balance (Table 2 protocol)")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "class_balance_before_after.png", dpi=150)
    plt.close(fig)

    folds_df = pd.DataFrame(fold_rows)
    fig, ax = plt.subplots(figsize=(9, 4.5))
    x = np.arange(N_OUTER)
    w = 0.35
    ax.bar(x - w / 2, folds_df["pos_before"], w, label="PCOS before", color="#DD8452")
    ax.bar(x + w / 2, folds_df["pos_after"], w, label="PCOS after (inner-train)", color="#55A868")
    for i, (b, a) in enumerate(zip(folds_df["pos_before"], folds_df["pos_after"])):
        ax.text(i - w / 2, b + 1, str(b), ha="center", fontsize=8)
        ax.text(i + w / 2, a + 1, str(a), ha="center", fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels([f"Fold {i}" for i in x])
    ax.set_ylabel("PCOS samples")
    ax.set_title("Outer CV folds: minority count before/after SMOTE-NC")
    ax.legend()
    fig.tight_layout()
    fig.savefig(FIG_DIR / "per_fold_balance.png", dpi=150)
    plt.close(fig)


def fill_table2(doc: Document, before: dict, after: dict) -> None:
    t = doc.tables[1]

    def set_cell(cell, text: str) -> None:
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(str(text))
        run.font.size = Pt(DOCX_SIZE)
        run.font.name = DOCX_FONT
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), DOCX_FONT)

    vals = [
        f"{before['neg']}/{before['pos']}",
        f"{before['test_neg']}/{before['test_pos']}",
        f"{before['neg'] + before['test_neg']}/{before['pos'] + before['test_pos']}",
        f"{after['neg']}/{after['pos']}",
        f"{after['test_neg']}/{after['test_pos']}",
        f"{after['neg'] + after['test_neg']}/{after['pos'] + after['test_pos']}",
    ]
    for c, v in enumerate(vals):
        set_cell(t.cell(3, c), v)


def verify_headers(src: Path, dst: Path) -> bool:
    a, b = Document(src), Document(dst)
    header_rows = {1: 1, 2: 3, 3: 3, 4: 1, 5: 1, 6: 1, 7: 1, 8: 3, 9: 3,
                   10: 2, 11: 2, 12: 1, 13: 1, 14: 2, 15: 1, 16: 1, 17: 1}
    ok = len(a.tables) == len(b.tables)
    for ti, (ta, tb) in enumerate(zip(a.tables, b.tables), start=1):
        for ri in range(header_rows.get(ti, 1)):
            for ci in range(len(ta.columns)):
                if ta.cell(ri, ci).text != tb.cell(ri, ci).text:
                    print(f"  header mismatch T{ti} R{ri}C{ci}")
                    ok = False
    return ok


def main() -> None:
    ensure_dirs()
    df, feature_cols, categorical_cols = load_data()
    X_all = df[feature_cols]
    y_all = df[TARGET].astype(int)

    train_idx, test_idx = make_holdout(df)
    X_train, y_train = X_all.iloc[train_idx], y_all.iloc[train_idx]
    X_test, y_test = X_all.iloc[test_idx], y_all.iloc[test_idx]
    log(f"Holdout split: train={len(train_idx)} test={len(test_idx)}")

    folds = make_nested_folds(y_train.values, train_idx)
    log(f"Nested CV: {N_OUTER} outer x {N_INNER} inner stratified folds")

    df.iloc[train_idx].to_csv(SPLIT_DIR / "holdout_train.csv", index=False)
    df.iloc[test_idx].to_csv(SPLIT_DIR / "holdout_test.csv", index=False)
    for fold in folds:
        k = fold["outer_fold"]
        df.iloc[fold["train_idx"]].to_csv(SPLIT_DIR / f"outer_fold_{k}_train.csv", index=False)
        df.iloc[fold["valid_idx"]].to_csv(SPLIT_DIR / f"outer_fold_{k}_valid.csv", index=False)
    n_split_files = 2 + 2 * len(folds)
    log(f"Saved {n_split_files} updated split CSVs -> {SPLIT_DIR.name}/")

    report_rows = []

    X_h, y_h, n_syn_h = apply_smote_nc(X_train, y_train, categorical_cols, SEED)
    report_rows.append(balancing_report_row("holdout_train", y_train, y_h, n_syn_h))
    X_h.to_csv(BAL_DIR / "holdout_train_balanced.csv", index=False)
    y_h.rename(TARGET).to_csv(BAL_DIR / "holdout_train_balanced_target.csv", index=False)

    fold_fig_rows = []
    for fold in folds:
        k = fold["outer_fold"]
        o_tr = np.array(fold["train_idx"])
        o_te = np.array(fold["valid_idx"])
        y_o_tr = y_all.iloc[o_tr]
        X_o_tr = X_all.iloc[o_tr]
        X_b, y_b, n_syn = apply_smote_nc(
            X_o_tr, y_o_tr, categorical_cols, SEED + k
        )
        row = balancing_report_row(f"outer_fold_{k}_train", y_o_tr, y_b, n_syn)
        report_rows.append(row)
        fold_fig_rows.append({
            "fold": k,
            "pos_before": int((y_o_tr == 1).sum()),
            "pos_after": int((y_b == 1).sum()),
        })
        X_b.to_csv(BAL_DIR / f"outer_fold_{k}_train_balanced.csv", index=False)
        y_b.rename(TARGET).to_csv(
            BAL_DIR / f"outer_fold_{k}_train_balanced_target.csv", index=False
        )
        for inner in fold["inner_splits"]:
            y_inner_tr = y_all.iloc[np.array(inner["train_idx"])]
            report_rows.append({
                "split": f"outer_{k}_inner_{inner['inner_fold']}_train",
                "n_before": int(len(y_inner_tr)),
                "neg_before": int((y_inner_tr == 0).sum()),
                "pos_before": int((y_inner_tr == 1).sum()),
                "n_after": np.nan,
                "neg_after": np.nan,
                "pos_after": np.nan,
                "n_synthetic": np.nan,
            })
        log(f"Fold {k}: train {len(o_tr)} -> balanced {len(X_b)} (+{n_syn} synthetic), valid {len(o_te)}")

    report = pd.DataFrame(report_rows)
    report.to_csv(QC_DIR / "balancing_report.csv", index=False)

    holdout_row = report[report["split"] == "holdout_train"].iloc[0]
    holdout_before = {
        "neg": int(holdout_row["neg_before"]), "pos": int(holdout_row["pos_before"]),
        "test_neg": int((y_test == 0).sum()), "test_pos": int((y_test == 1).sum()),
    }
    holdout_after = {
        "neg": int(holdout_row["neg_after"]), "pos": int(holdout_row["pos_after"]),
        "test_neg": holdout_before["test_neg"], "test_pos": holdout_before["test_pos"],
    }

    make_figures(holdout_before, holdout_after, fold_fig_rows)

    assignments = pd.DataFrame({
        "row_id": np.arange(len(df)),
        "Patient_File_No": df["Patient_File_No"] if "Patient_File_No" in df.columns else np.arange(len(df)),
        "holdout_split": "train",
        "outer_fold": -1,
    })
    assignments.iloc[test_idx, assignments.columns.get_loc("holdout_split")] = "test"
    for fold in folds:
        assignments.iloc[fold["valid_idx"], assignments.columns.get_loc("outer_fold")] = fold["outer_fold"]
    assignments.to_csv(ARTIFACTS_DIR / "fold_assignments.csv", index=False)

    joblib.dump(
        {"holdout_train_idx": train_idx.tolist(), "holdout_test_idx": test_idx.tolist(),
         "outer_folds": folds, "seed": SEED},
        ARTIFACTS_DIR / "cv_splits.joblib",
    )

    TABLES_DIR.mkdir(parents=True, exist_ok=True)
    src_docx = STEP1_DOCX if STEP1_DOCX.exists() else TEMPLATE_DOCX
    shutil.copyfile(src_docx, OUT_DOCX)
    doc = Document(OUT_DOCX)
    fill_table2(doc, holdout_before, holdout_after)
    doc.save(OUT_DOCX)

    t2_csv = pd.DataFrame([
        {"Section": "Before oversampling", "Training": f"{holdout_before['neg']}/{holdout_before['pos']}",
         "Testing": f"{holdout_before['test_neg']}/{holdout_before['test_pos']}",
         "Total": f"{holdout_before['neg'] + holdout_before['test_neg']}/{holdout_before['pos'] + holdout_before['test_pos']}"},
        {"Section": "After oversampling", "Training": f"{holdout_after['neg']}/{holdout_after['pos']}",
         "Testing": f"{holdout_after['test_neg']}/{holdout_after['test_pos']}",
         "Total": f"{holdout_after['neg'] + holdout_after['test_neg']}/{holdout_after['pos'] + holdout_after['test_pos']}"},
    ])
    t2_csv.to_csv(TABLES_DIR / "Table2_Dataset_Distribution.csv", index=False)
    log(f"Table 2 filled (cells show Neg/Pos counts): {OUT_DOCX.name}")

    ok = verify_headers(TEMPLATE_DOCX, OUT_DOCX)
    log("Header verification: " + ("PASS" if ok else "FAIL"))

    summary = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "protocol": {
            "holdout": f"stratified {1 - TEST_SIZE:.0%}/{TEST_SIZE:.0%}, seed={SEED}",
            "nested_cv": f"{N_OUTER}x{N_INNER} stratified, seed={SEED}",
            "sampler": "SMOTE-NC (categorical_features preserved), inside-fold only",
            "validation_test": "untouched, never resampled",
        },
        "holdout": {
            "train": len(train_idx), "test": len(test_idx),
            "train_before": holdout_before, "train_after": holdout_after,
        },
        "outer_folds": [
            {"fold": f["outer_fold"], "train": len(f["train_idx"]),
             "valid": len(f["valid_idx"]), "inner_splits": len(f["inner_splits"])}
            for f in folds
        ],
        "balanced_files": sorted(p.name for p in BAL_DIR.glob("*.csv")),
        "split_data_files": sorted(p.name for p in SPLIT_DIR.glob("*.csv")),
    }
    (QC_DIR / "balancing_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print()
    log("DONE")
    log(f"Holdout train before: Non-PCOS {holdout_before['neg']} / PCOS {holdout_before['pos']}")
    log(f"Holdout train after : Non-PCOS {holdout_after['neg']} / PCOS {holdout_after['pos']} (+{int(holdout_row['n_synthetic'])} synthetic)")
    log(f"Test untouched      : Non-PCOS {holdout_before['test_neg']} / PCOS {holdout_before['test_pos']}")
    log(f"Split data CSVs     : {SPLIT_DIR.relative_to(ROOT)} ({n_split_files} files)")
    log(f"Outputs: {RESULTS_DIR.relative_to(ROOT)} + {TABLES_DIR.relative_to(ROOT)}")


if __name__ == "__main__":
    sys.exit(main())
