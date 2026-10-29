"""Step 1: Pre-processing & Feature Engineering for PCOS prediction.

Follows the proposed tri-stage methodology (Stage 1) aligned with
Efty et al. 2026 (Health Science Reports) leakage-free protocol:
  - clean -> impute -> cap outliers -> engineer 22 Rotterdam biomarkers
  - 80/20 stratified split FIRST, all transforms fit on train only
  - Robust + Quantile normalization (train-fitted)
Outputs two CSVs (raw-cleaned + scaled) plus QC reports, figures and artifacts.
"""

from __future__ import annotations

import json
import platform
import re
import shutil
import sys
from datetime import datetime, timezone
from importlib import metadata as importlib_metadata
from pathlib import Path

import joblib
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import psutil
import seaborn as sns
import sklearn
import statsmodels
import statsmodels.api as sm
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from scipy import stats
from sklearn.impute import SimpleImputer
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import QuantileTransformer, RobustScaler

SEED = 42
TEST_SIZE = 0.20
CAP_LOW = 0.01
CAP_HIGH = 0.99
SKEW_THRESHOLD = 1.0
TOP_N_TABLE5 = 16

ROOT = Path(__file__).resolve().parents[2]
DATA_PATH = ROOT / "data" / "PCOS_data.csv"
STEP_DIR = ROOT / "step1_preprocessing"
RESULTS_DIR = STEP_DIR / "results"
ARTIFACTS_DIR = RESULTS_DIR / "artifacts"
QC_DIR = RESULTS_DIR / "qc_reports"
FIG_DIR = RESULTS_DIR / "figures"
TABLES_DIR = STEP_DIR / "tables"
TEMPLATE_DOCX = ROOT / "docs" / "All_Tables_Templates.docx"
OUT_DOCX = TABLES_DIR / "Step1_Tables_Filled.docx"

TARGET = "PCOS (Y/N)"
DROP_COLS = ["Sl. No", "Patient File No.", "Unnamed: 44"]
KEEP_META = "Patient File No."

BIOMARKER_SPECS = [
    ("PCOM_Flag", "binary", "int((Follicle L >= 12) | (Follicle R >= 12))",
     "Rotterdam 2018: >=12 follicles per ovary on ultrasound"),
    ("OligoAnovulation_Flag", "binary", "int(Cycle(R/I) != 2)",
     "Rotterdam: irregular/absent ovulation (Cycle code 2 = regular)"),
    ("Clinical_Hyperandrogenism_Flag", "binary", "int(hair growth OR skin darkening)",
     "Rotterdam: clinical hyperandrogenism (hirsutism / acanthosis signs)"),
    ("LH_FSH_Ratio", "continuous", "LH(mIU/mL) / FSH(mIU/mL)",
     "Classic PCOS marker; >2 suggests LH excess (inverse of raw FSH/LH)"),
    ("Androgen_Score", "ordinal04",
     "hair growth + skin darkening + pimples + int(LH/FSH >= 2)",
     "Composite clinical+biochemical androgen excess score (0-4)"),
    ("Follicle_Total", "continuous", "Follicle No. (L) + Follicle No. (R)",
     "Total antral follicle count (Rotterdam PCOM volume proxy)"),
    ("Follicle_Avg", "continuous", "(Follicle No. (L) + Follicle No. (R)) / 2",
     "Mean bilateral follicle count"),
    ("Follicle_Asymmetry", "continuous", "abs(Follicle No. (L) - Follicle No. (R))",
     "Ovarian asymmetry; elevated in unilateral PCOM"),
    ("Ovarian_Volume_Proxy", "continuous",
     "(FolL + FolR) * (AvgFsizeL + AvgFsizeR) / 2",
     "Total follicular load = count x mean cyst diameter proxy"),
    ("WHR_Recomputed", "continuous", "Waist(inch) / Hip(inch)",
     "Recomputed waist-to-hip ratio (central obesity marker)"),
    ("WHR_Abn_Flag", "binary", "int(WHR_Recomputed > 0.85)",
     "WHO central obesity threshold for women (WHR > 0.85)"),
    ("BMI_Category", "ordinal04",
     "0:<18.5, 1:18.5-24.9, 2:25-29.9, 3:30-34.9, 4:>=35",
     "WHO BMI classes (ordinal encoding for dose-response)"),
    ("TSH_Abn_Flag", "binary", "int((TSH < 0.4) | (TSH > 4.0))",
     "Thyroid dysfunction overlay (common PCOS differential)"),
    ("Hb_Abn_Flag", "binary", "int(Hb < 12.0)",
     "Anemia flag (Hb < 12 g/dL, WHO women threshold)"),
    ("RBS_Diabetic_Flag", "binary", "int(RBS >= 126)",
     "Diabetes-range random blood sugar (ADA >= 126 mg/dL)"),
    ("VitD_Deficient_Flag", "binary", "int(Vit D3 < 20)",
     "Vitamin D deficiency (< 20 ng/mL, Endocrine Society)"),
    ("Pulse_Elevated_Flag", "binary", "int(Pulse rate > 100)",
     "Tachycardia overlay"),
    ("Cycle_Length_Abn_Flag", "binary", "int((Cycle length < 3) | (Cycle length > 7))",
     "Abnormal menstrual bleeding duration (dataset records 0-12 days)"),
    ("AMH_High_Flag", "binary", "int(AMH > 6.8)",
     "Elevated AMH (ovarian reserve / PCOM biochemical proxy)"),
    ("Metabolic_Risk_Score", "ordinal04",
     "int(BMI >= 25) + WHR_Abn_Flag + RBS_Diabetic_Flag + Weight gain",
     "Cumulative metabolic dysfunction score (0-4)"),
    ("AMH_x_Follicle", "continuous", "AMH * Follicle_Total",
     "Hormone x morphology interaction (AMH drives follicle arrest)"),
    ("Age_x_BMI", "continuous", "Age * BMI",
     "Age-adiposity interaction for metabolic PCOS phenotype"),
]

BIOMARKER_NAMES = [b[0] for b in BIOMARKER_SPECS]
BIOMARKER_TYPES = {b[0]: b[1] for b in BIOMARKER_SPECS}

RAW_TO_DISPLAY = {
    "Age (yrs)": "Age (years)",
    "Weight (Kg)": "Weight (kg)",
    "Height(Cm)": "Height (cm)",
    "Pulse rate(bpm)": "Pulse rate (bpm)",
    "Hb(g/dl)": "Hb (g/dL)",
    "I   beta-HCG(mIU/mL)": "I beta-HCG (mIU/mL)",
    "II    beta-HCG(mIU/mL)": "II beta-HCG (mIU/mL)",
    "FSH(mIU/mL)": "FSH (mIU/mL)",
    "LH(mIU/mL)": "LH (mIU/mL)",
    "FSH/LH": "FSH/LH",
    "Hip(inch)": "Hip (inch)",
    "Waist(inch)": "Waist (inch)",
    "Waist:Hip Ratio": "Waist:Hip Ratio",
    "TSH (mIU/L)": "TSH (mIU/L)",
    "AMH(ng/mL)": "AMH (ng/mL)",
    "PRL(ng/mL)": "PRL (ng/mL)",
    "Vit D3 (ng/mL)": "Vit D3 (ng/mL)",
    "PRG(ng/mL)": "PRG (ng/mL)",
    "RBS(mg/dl)": "RBS (mg/dL)",
    "BP _Systolic (mmHg)": "BP systolic (mmHg)",
    "BP _Diastolic (mmHg)": "BP diastolic (mmHg)",
    "Follicle No. (L)": "Follicle (L)",
    "Follicle No. (R)": "Follicle (R)",
    "Avg. F size (L) (mm)": "Avg. F size (L) (mm)",
    "Avg. F size (R) (mm)": "Avg. F size (R) (mm)",
    "Endometrium (mm)": "Endometrium (mm)",
    "RR (breaths/min)": "RR (breaths/min)",
    "Cycle(R/I)": "Cycle (R/I)",
    "Cycle length(days)": "Cycle length (days)",
    "Marraige Status (Yrs)": "Marriage status (years)",
    "No. of abortions": "No. of abortions",
}

FLAG_LIKE = {
    "Blood Group", "Cycle (R/I)", "Pregnant(Y/N)", "Weight gain(Y/N)",
    "hair growth(Y/N)", "Skin darkening (Y/N)", "Hair loss(Y/N)",
    "Pimples(Y/N)", "Fast food (Y/N)", "Reg.Exercise(Y/N)",
}

BINARY_TARGET = TARGET


def log(msg: str) -> None:
    print(f"[step1] {msg}", flush=True)


def lib_version(name: str) -> str:
    try:
        return importlib_metadata.version(name)
    except importlib_metadata.PackageNotFoundError:
        return "n/a"


def ensure_dirs() -> None:
    for path in (RESULTS_DIR, ARTIFACTS_DIR, QC_DIR, FIG_DIR):
        path.mkdir(parents=True, exist_ok=True)


def load_and_clean() -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(DATA_PATH)
    df.columns = [str(c).strip() for c in df.columns]
    missing_before = df.isna().sum().rename("missing_before").to_frame()
    missing_before["dtype_before"] = df.dtypes.astype(str).values
    missing_before["n_unique"] = [df[c].nunique(dropna=True) for c in df.columns]

    df = df.drop(columns=[c for c in DROP_COLS if c in df.columns])

    hcg2 = "II    beta-HCG(mIU/mL)"
    if hcg2 in df.columns:
        df[hcg2] = (
            df[hcg2]
            .astype(str)
            .str.replace(r"(\d)\.+$", r"\1", regex=True)
            .str.strip()
        )
        df[hcg2] = pd.to_numeric(df[hcg2], errors="coerce")

    for col in df.columns:
        if col == TARGET:
            continue
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df[TARGET] = pd.to_numeric(df[TARGET], errors="coerce").astype(int)
    log(f"Loaded {df.shape[0]} rows x {df.shape[1]} cols after cleaning")
    return df, missing_before


def build_feature_lists(df: pd.DataFrame) -> tuple[list[str], list[str], list[str]]:
    feature_cols = [c for c in df.columns if c not in (TARGET, KEEP_META)]
    engineered = [c for c in BIOMARKER_NAMES if c in df.columns]
    continuous = []
    categorical = []
    for col in feature_cols:
        if col in engineered:
            btype = BIOMARKER_TYPES[col]
            if btype in ("binary", "ordinal04"):
                categorical.append(col)
            else:
                continuous.append(col)
        elif col in FLAG_LIKE:
            categorical.append(col)
        else:
            n_unique = df[col].nunique(dropna=True)
            if n_unique <= 2:
                categorical.append(col)
            else:
                continuous.append(col)
    return feature_cols, continuous, categorical


def engineer_biomarkers(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    fol_l = out["Follicle No. (L)"]
    fol_r = out["Follicle No. (R)"]
    size_l = out["Avg. F size (L) (mm)"]
    size_r = out["Avg. F size (R) (mm)"]
    hair = out["hair growth(Y/N)"].fillna(0)
    skin = out["Skin darkening (Y/N)"].fillna(0)
    pimp = out["Pimples(Y/N)"].fillna(0)
    weight_gain = out["Weight gain(Y/N)"].fillna(0)
    fsh = out["FSH(mIU/mL)"].replace(0, np.nan)
    lh = out["LH(mIU/mL)"]
    ratio = (lh / fsh).replace([np.inf, -np.inf], np.nan)
    whr = out["Waist(inch)"] / out["Hip(inch)"].replace(0, np.nan)
    amh = out["AMH(ng/mL)"]
    cycle_abn = ((out["Cycle length(days)"] < 3) | (out["Cycle length(days)"] > 7)).astype(int)

    out["PCOM_Flag"] = ((fol_l >= 12) | (fol_r >= 12)).astype(int)
    out["OligoAnovulation_Flag"] = (out["Cycle(R/I)"] != 2).astype(int)
    out["Clinical_Hyperandrogenism_Flag"] = ((hair > 0) | (skin > 0)).astype(int)
    out["LH_FSH_Ratio"] = ratio
    out["Androgen_Score"] = (
        hair.astype(int) + skin.astype(int) + pimp.astype(int) + (ratio >= 2).astype(int)
    )
    out["Follicle_Total"] = fol_l + fol_r
    out["Follicle_Avg"] = (fol_l + fol_r) / 2.0
    out["Follicle_Asymmetry"] = (fol_l - fol_r).abs()
    out["Ovarian_Volume_Proxy"] = (fol_l + fol_r) * (size_l + size_r) / 2.0
    out["WHR_Recomputed"] = whr
    out["WHR_Abn_Flag"] = (whr > 0.85).astype(int)
    out["BMI_Category"] = pd.cut(
        out["BMI"],
        bins=[-np.inf, 18.5, 25, 30, 35, np.inf],
        labels=[0, 1, 2, 3, 4],
        right=False,
    ).astype(float)
    out["TSH_Abn_Flag"] = ((out["TSH (mIU/L)"] < 0.4) | (out["TSH (mIU/L)"] > 4.0)).astype(int)
    out["Hb_Abn_Flag"] = (out["Hb(g/dl)"] < 12.0).astype(int)
    out["RBS_Diabetic_Flag"] = (out["RBS(mg/dl)"] >= 126).astype(int)
    out["VitD_Deficient_Flag"] = (out["Vit D3 (ng/mL)"] < 20).astype(int)
    out["Pulse_Elevated_Flag"] = (out["Pulse rate(bpm)"] > 100).astype(int)
    out["Cycle_Length_Abn_Flag"] = cycle_abn
    out["AMH_High_Flag"] = (amh > 6.8).astype(int)
    out["Metabolic_Risk_Score"] = (
        (out["BMI"] >= 25).astype(int)
        + out["WHR_Abn_Flag"]
        + out["RBS_Diabetic_Flag"]
        + weight_gain.astype(int)
    )
    out["AMH_x_Follicle"] = amh * out["Follicle_Total"]
    out["Age_x_BMI"] = out["Age (yrs)"] * out["BMI"]
    return out


def make_split(df: pd.DataFrame) -> tuple[np.ndarray, np.ndarray]:
    idx = np.arange(len(df))
    train_idx, test_idx = train_test_split(
        idx,
        test_size=TEST_SIZE,
        random_state=SEED,
        stratify=df[TARGET].values,
    )
    return np.sort(train_idx), np.sort(test_idx)


def cap_apply(s: pd.Series, lo, hi) -> pd.Series:
    return s.clip(lo, hi)


def build_figures(df: pd.DataFrame, imputed: pd.DataFrame, capped: pd.DataFrame,
                  feature_cols: list[str], top_features: list[str]) -> None:
    sns.set_theme(style="whitegrid", context="notebook")

    fig, ax = plt.subplots(figsize=(6, 4))
    counts = df[TARGET].value_counts().sort_index()
    sns.barplot(x=counts.index.astype(str), y=counts.values, hue=counts.index.astype(str),
                palette=["#4C72B0", "#DD8452"], legend=False, ax=ax)
    ax.set_title("Class balance before balancing")
    ax.set_xlabel("PCOS (Y/N)")
    ax.set_ylabel("Patients")
    for i, v in enumerate(counts.values):
        ax.text(i, v + 3, str(v), ha="center", fontsize=10)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "class_balance.png", dpi=150)
    plt.close(fig)

    miss = df.isna().sum()
    miss = miss[miss > 0]
    fig, ax = plt.subplots(figsize=(7, 3.5))
    if len(miss):
        sns.barplot(x=miss.index, y=miss.values, color="#C44E52", ax=ax)
        ax.set_title("Missing values before imputation")
        ax.tick_params(axis="x", rotation=30)
    else:
        ax.text(0.5, 0.5, "No missing values", ha="center", va="center")
        ax.set_title("Missing values before imputation")
        ax.axis("off")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "missingness_before.png", dpi=150)
    plt.close(fig)

    skewed = (
        imputed[list(set(feature_cols) & set(imputed.columns))]
        .apply(lambda s: s.skew() if s.nunique() > 2 else np.nan)
        .abs()
        .sort_values(ascending=False)
        .head(4)
        .index.tolist()
    )
    fig, axes = plt.subplots(2, 2, figsize=(10, 7))
    for ax, col in zip(axes.ravel(), skewed):
        sns.histplot(imputed[col], kde=True, color="#4C72B0", ax=ax, label="Before capping", stat="density")
        sns.histplot(capped[col], kde=True, color="#DD8452", ax=ax, label="After capping", stat="density")
        ax.set_title(col)
        ax.legend(fontsize=8)
    fig.suptitle("Outlier capping (1st-99th pct) on most skewed features", y=1.01)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "outlier_capping_before_after.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    corr_cols = top_features + [TARGET]
    corr = imputed[corr_cols].corr()
    fig, ax = plt.subplots(figsize=(12, 10))
    sns.heatmap(corr, cmap="RdBu_r", center=0, annot=False, fmt=".2f",
                linewidths=0.4, ax=ax, cbar_kws={"shrink": 0.8})
    ax.set_title("Correlation heatmap: top features + target")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "correlation_heatmap.png", dpi=150)
    plt.close(fig)


def table5_stats(df: pd.DataFrame, top_features: list[str]) -> pd.DataFrame:
    y = df[TARGET].astype(int)
    rows = []
    for feat in top_features:
        x = pd.to_numeric(df[feat], errors="coerce")
        mask = x.notna() & y.notna()
        xv, yv = x[mask].astype(float), y[mask].astype(int)
        z = (xv - xv.mean()) / (xv.std(ddof=0) if xv.std(ddof=0) else 1.0)
        try:
            model = sm.Logit(yv, sm.add_constant(z)).fit(disp=0)
            beta = float(model.params.iloc[1])
            ci = model.conf_int(alpha=0.05)
            lo, hi = float(ci.iloc[1, 0]), float(ci.iloc[1, 1])
            pval = float(model.pvalues.iloc[1])
        except Exception:
            beta, lo, hi, pval = np.nan, np.nan, np.nan, np.nan
        corr = float(np.corrcoef(xv, yv)[0, 1]) if xv.std() > 0 else np.nan
        rows.append({
            "Feature": RAW_TO_DISPLAY.get(feat, feat),
            "β": beta,
            "β-CI-lower": lo,
            "β-CI-upper": hi,
            "Y-correlation": corr,
            "p": pval,
            "Y1-mean": float(xv[yv == 1].mean()),
            "Y0-mean": float(xv[yv == 0].mean()),
            "All mean": float(xv.mean()),
            "All Std": float(xv.std(ddof=0)),
        })
    return pd.DataFrame(rows)


DOCX_FONT = "Calibri"
DOCX_SIZE = 9


def _docx_set_cell(cell, text: str, align: str = "center") -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(str(text))
    run.bold = False
    run.font.size = Pt(DOCX_SIZE)
    run.font.name = DOCX_FONT
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), DOCX_FONT)


def _fmt4(v: float) -> str:
    return "" if pd.isna(v) else f"{v:.4f}"


def _fmt_p(v: float) -> str:
    if pd.isna(v):
        return ""
    return f"{v:.2e}" if v < 0.001 else f"{v:.4f}"


def build_tables_docx(t5: pd.DataFrame, manifest: dict, split_info: dict) -> None:
    TABLES_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(TEMPLATE_DOCX, OUT_DOCX)
    doc = Document(OUT_DOCX)

    table5 = doc.tables[4]
    assert len(table5.rows) - 1 == len(t5)
    for i, row in t5.iterrows():
        r = i + 1
        _docx_set_cell(table5.cell(r, 0), row["Feature"], align="left")
        _docx_set_cell(table5.cell(r, 1), _fmt4(row["β"]))
        _docx_set_cell(table5.cell(r, 2), _fmt4(row["β-CI-lower"]))
        _docx_set_cell(table5.cell(r, 3), _fmt4(row["β-CI-upper"]))
        _docx_set_cell(table5.cell(r, 4), _fmt4(row["Y-correlation"]))
        _docx_set_cell(table5.cell(r, 5), _fmt_p(row["p"]))
        _docx_set_cell(table5.cell(r, 6), f"{row['Y1-mean']:.3f}")
        _docx_set_cell(table5.cell(r, 7), f"{row['Y0-mean']:.3f}")
        _docx_set_cell(table5.cell(r, 8), f"{row['All mean']:.3f}")
        _docx_set_cell(table5.cell(r, 9), f"{row['All Std']:.3f}")

    libs = [
        ("pandas", lib_version("pandas")),
        ("numpy", lib_version("numpy")),
        ("scikit-learn", lib_version("scikit-learn")),
        ("imbalanced-learn", lib_version("imbalanced-learn")),
        ("matplotlib", lib_version("matplotlib")),
        ("seaborn", lib_version("seaborn")),
        ("LightGBM", lib_version("lightgbm")),
        ("XGBoost", lib_version("xgboost")),
        ("SHAP", lib_version("shap")),
        ("SciPy", lib_version("scipy")),
        ("statsmodels", lib_version("statsmodels")),
        ("LIME", lib_version("lime")),
    ]
    table6 = doc.tables[5]
    for i, (lib, ver) in enumerate(libs):
        r = 1 + (i % 6)
        c = 0 if i < 6 else 2
        _docx_set_cell(table6.cell(r, c), lib, align="left")
        _docx_set_cell(table6.cell(r, c + 1), ver)
    pd.DataFrame(libs, columns=["Library", "Version"]).to_csv(
        TABLES_DIR / "Table6_Libraries.csv", index=False
    )

    cpu = platform.processor() or platform.machine()
    try:
        import torch
        gpu = torch.cuda.get_device_name(0) if torch.cuda.is_available() else "None (CPU-only PyTorch build)"
    except Exception:
        gpu = "None detected"
    env_rows = [
        ("CPU", f"{cpu} ({psutil.cpu_count(logical=True)} logical cores)"),
        ("GPU", gpu),
        ("RAM", f"{psutil.virtual_memory().total / 1e9:.1f} GB"),
        ("Experimental tools",
         f"Python {platform.python_version()}; scikit-learn {lib_version('scikit-learn')}; statsmodels {statsmodels.__version__}"),
    ]
    table7 = doc.tables[6]
    for i, (res, detail) in enumerate(env_rows, start=1):
        _docx_set_cell(table7.cell(i, 0), res, align="left")
        _docx_set_cell(table7.cell(i, 1), detail, align="left")
    pd.DataFrame(env_rows, columns=["Resources", "Details"]).to_csv(
        TABLES_DIR / "Table7_Environment.csv", index=False
    )

    doc.save(OUT_DOCX)
    log(f"Tables docx saved: {OUT_DOCX.name} (T5+T6+T7 filled)")

    a, b = Document(TEMPLATE_DOCX), Document(OUT_DOCX)
    header_rows = {1: 1, 2: 3, 3: 3, 4: 1, 5: 1, 6: 1, 7: 1, 8: 3, 9: 3,
                   10: 2, 11: 2, 12: 1, 13: 1, 14: 2, 15: 1, 16: 1, 17: 1}
    ok = len(a.tables) == len(b.tables)
    for ti, (ta, tb) in enumerate(zip(a.tables, b.tables), start=1):
        for ri in range(header_rows.get(ti, 1)):
            for ci in range(len(ta.columns)):
                if ta.cell(ri, ci).text != tb.cell(ri, ci).text:
                    ok = False
    log("Header verification: " + ("PASS - all headers byte-identical" if ok else "FAIL"))


def main() -> None:
    ensure_dirs()
    df_raw, missing_before = load_and_clean()

    feature_cols_all, _, _ = build_feature_lists(df_raw)
    df_eng = engineer_biomarkers(df_raw)

    feature_cols, continuous, categorical = build_feature_lists(df_eng)
    log(f"Features: {len(feature_cols)} (continuous={len(continuous)}, categorical={len(categorical)})")
    log(f"Engineered biomarkers present: {sum(c in df_eng.columns for c in BIOMARKER_NAMES)}/22")

    train_idx, test_idx = make_split(df_eng)
    y_train = df_eng[TARGET].iloc[train_idx]

    imputer = SimpleImputer(strategy="median")
    imputer.fit(df_eng.iloc[train_idx][feature_cols])
    imputed_full = df_eng.copy()
    imputed_full[feature_cols] = imputer.transform(df_eng[feature_cols])
    log(f"Imputed with train medians | train={len(train_idx)} test={len(test_idx)}")

    capper = {}
    capped_full = imputed_full.copy()
    outlier_rows = []
    for col in continuous:
        lo = float(imputed_full.iloc[train_idx][col].quantile(CAP_LOW))
        hi = float(imputed_full.iloc[train_idx][col].quantile(CAP_HIGH))
        capper[col] = (lo, hi)
        before = capped_full[col]
        clipped_total = int(((before < lo) | (before > hi)).sum())
        clipped_train = int(((imputed_full.iloc[train_idx][col] < lo) |
                             (imputed_full.iloc[train_idx][col] > hi)).sum())
        clipped_test = int(((imputed_full.iloc[test_idx][col] < lo) |
                            (imputed_full.iloc[test_idx][col] > hi)).sum())
        capped_full[col] = cap_apply(capped_full[col], lo, hi)
        outlier_rows.append({
            "feature": col, "p1_train": lo, "p99_train": hi,
            "n_clipped_train": clipped_train, "n_clipped_test": clipped_test,
            "n_clipped_total": clipped_total,
        })
    log(f"Capped {len(continuous)} continuous features at {CAP_LOW:.0%}-{CAP_HIGH:.0%} (train quantiles)")

    skewed_feats, robust_feats = [], []
    for col in continuous:
        skew = float(imputed_full.iloc[train_idx][col].skew())
        if abs(skew) >= SKEW_THRESHOLD:
            skewed_feats.append(col)
        else:
            robust_feats.append(col)

    n_quantiles = min(300, len(train_idx))
    qt = QuantileTransformer(
        output_distribution="normal", n_quantiles=n_quantiles, random_state=SEED
    )
    rs = RobustScaler()
    scaled_full = capped_full.copy()
    if skewed_feats:
        qt.fit(capped_full.iloc[train_idx][skewed_feats])
        scaled_full[skewed_feats] = qt.transform(capped_full[skewed_feats])
    if robust_feats:
        rs.fit(capped_full.iloc[train_idx][robust_feats])
        scaled_full[robust_feats] = rs.transform(capped_full[robust_feats])
    log(f"Normalization: Quantile(n={len(skewed_feats)}) + Robust(n={len(robust_feats)}), fit on train")

    split_series = pd.Series("train", index=df_eng.index, dtype=object)
    split_series.iloc[test_idx] = "test"

    meta_cols = [c for c in (KEEP_META,) if c in capped_full.columns]
    raw_out = capped_full.copy()
    raw_out.insert(0, "Patient_File_No", raw_out[KEEP_META] if KEEP_META in raw_out.columns else np.arange(len(raw_out)))
    if KEEP_META in raw_out.columns:
        raw_out = raw_out.drop(columns=[KEEP_META])
    raw_out.insert(1, "Split", split_series.values)
    raw_out[TARGET] = capped_full[TARGET].values
    raw_path = RESULTS_DIR / "pcos_cleaned_raw.csv"
    raw_out.to_csv(raw_path, index=False)

    scaled_out = scaled_full.copy()
    scaled_out.insert(0, "Patient_File_No", raw_out["Patient_File_No"].values)
    scaled_out.insert(1, "Split", split_series.values)
    scaled_out[TARGET] = capped_full[TARGET].values
    scaled_path = RESULTS_DIR / "pcos_cleaned_scaled.csv"
    scaled_out.to_csv(scaled_path, index=False)
    log(f"Saved {raw_path.name} and {scaled_path.name} ({raw_out.shape[0]} rows each)")

    stats_df = capped_full.copy()
    y = stats_df[TARGET]
    corrs = {}
    for col in feature_cols:
        x = stats_df[col].astype(float)
        if x.std(ddof=0) > 0:
            corrs[col] = abs(float(np.corrcoef(x, y)[0, 1]))
        else:
            corrs[col] = 0.0
    top_features = sorted(corrs.keys(), key=lambda c: (-corrs[c], c))[:TOP_N_TABLE5]
    t5 = table5_stats(stats_df, top_features)
    t5_csv = STEP_DIR / "tables" / "Table5_Statistical_Summary.csv"
    t5.to_csv(t5_csv, index=False, float_format="%.6g")

    joblib.dump(imputer, ARTIFACTS_DIR / "imputer.joblib")
    joblib.dump(capper, ARTIFACTS_DIR / "outlier_capper.joblib")
    joblib.dump({"quantile": qt, "robust": rs}, ARTIFACTS_DIR / "scalers.joblib")
    joblib.dump(
        {"train_idx": train_idx.tolist(), "test_idx": test_idx.tolist(), "seed": SEED},
        ARTIFACTS_DIR / "split_indices.joblib",
    )

    manifest = {
        "seed": SEED,
        "test_size": TEST_SIZE,
        "target": TARGET,
        "n_raw_features": len(feature_cols),
        "feature_columns": feature_cols,
        "continuous_features": continuous,
        "categorical_features": categorical,
        "quantile_features": skewed_feats,
        "robust_features": robust_feats,
        "biomarker_names": BIOMARKER_NAMES,
        "top16_table5": top_features,
        "display_names": RAW_TO_DISPLAY,
    }
    with open(ARTIFACTS_DIR / "feature_manifest.json", "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=2)

    missing_after = capped_full[feature_cols].isna().sum()
    miss_df = missing_before.join(
        missing_after.rename("missing_after"), how="left"
    ).fillna({"missing_after": 0}).astype({"missing_after": int})
    miss_df["imputation_method"] = [
        "median (train-fitted)" if m > 0 else "none" for m in miss_df["missing_before"]
    ]
    miss_df.index.name = "column"
    miss_df.to_csv(QC_DIR / "missing_before_after.csv", float_format="%.4f")

    pd.DataFrame(outlier_rows).to_csv(QC_DIR / "outlier_report.csv", index=False, float_format="%.4f")

    eng_rows = []
    for i, (name, btype, formula, rationale) in enumerate(BIOMARKER_SPECS, 1):
        eng_rows.append({
            "no": i, "feature": name, "type": btype,
            "formula": formula, "rationale": rationale,
            "nonzero_rate": float((capped_full[name] != 0).mean()) if name in capped_full else np.nan,
            "y_corr": float(np.corrcoef(capped_full[name].astype(float), y)[0, 1])
            if name in capped_full and capped_full[name].std() > 0 else np.nan,
        })
    pd.DataFrame(eng_rows).to_csv(QC_DIR / "feature_engineering_log.csv", index=False, float_format="%.4f")

    build_figures(df_raw, imputed_full, capped_full, feature_cols, top_features)

    summary = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "dataset": {
            "rows": int(len(capped_full)),
            "raw_columns": 45,
            "features_after_engineering": len(feature_cols),
            "biomarkers_added": len(BIOMARKER_NAMES),
            "target_counts": {str(k): int(v) for k, v in df_raw[TARGET].value_counts().sort_index().items()},
        },
        "split": {
            "method": "stratified 80/20",
            "seed": SEED,
            "train": int(len(train_idx)),
            "test": int(len(test_idx)),
            "train_target_counts": {str(k): int(v) for k, v in y_train.value_counts().sort_index().items()},
        },
        "preprocessing": {
            "imputation": "median (train-fitted SimpleImputer)",
            "outlier_capping": f"train {CAP_LOW:.0%}/{CAP_HIGH:.0%} quantiles on {len(continuous)} continuous features",
            "quantile_transformed": len(skewed_feats),
            "robust_scaled": len(robust_feats),
            "skew_threshold": SKEW_THRESHOLD,
        },
        "top16_table5": top_features,
        "outputs": {
            "raw_csv": str(raw_path.relative_to(ROOT)),
            "scaled_csv": str(scaled_path.relative_to(ROOT)),
            "table5_csv": str(t5_csv.relative_to(ROOT)),
        },
        "environment": {
            "python": platform.python_version(),
            "platform": platform.platform(),
            "cpu": platform.processor() or platform.machine(),
            "cores_logical": psutil.cpu_count(logical=True),
            "ram_gb": round(psutil.virtual_memory().total / 1e9, 1),
            "pandas": pd.__version__,
            "numpy": np.__version__,
            "scikit_learn": sklearn.__version__,
            "statsmodels": lib_version("statsmodels"),
        },
    }
    with open(QC_DIR / "preprocessing_summary.json", "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)

    build_tables_docx(t5, manifest, {"train_idx": train_idx, "test_idx": test_idx})

    print()
    log("DONE")
    log(f"Rows: {len(capped_full)} | features: {len(feature_cols)} (+{len(BIOMARKER_NAMES)} biomarkers)")
    log(f"Train/test: {len(train_idx)}/{len(test_idx)} | missing after: {int(missing_after.sum())}")
    log(f"NaN/inf in raw CSV feature cols: {int(raw_out[feature_cols].isna().sum().sum() + np.isinf(raw_out[feature_cols].select_dtypes('number')).sum().sum())}")
    log(f"Top-16 for Table 5: {top_features}")
    print(json.dumps(summary["dataset"], indent=2))


if __name__ == "__main__":
    sys.exit(main())
