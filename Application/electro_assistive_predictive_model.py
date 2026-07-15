#!/usr/bin/env python3
"""ELECTRO assistive and predictive machine-learning pipeline.

Architecture
------------
1. Elastic-net logistic regression baseline.
2. Complexity/monotonicity-constrained histogram gradient-boosted trees.
3. Experimental complexity-constrained deep/cascade forest.
4. Constrained regression surrogates mapping design variables to simulation
   responses (E/D/P/Q metrics).
5. Counterfactual recommendation search with explicit geometry/physics checks.

The script intentionally works with a flexible ELECTRO CSV schema.  It creates
an editable JSON configuration and a schema template, then stores the exact
resolved schema in the trained bundle.

Examples
--------
  python electro_assistive_predictive_model.py template --directory electro_template
  python electro_assistive_predictive_model.py train --data simulations.csv --output-dir results
  python electro_assistive_predictive_model.py predict --bundle results/electro_model.joblib --data new.csv
  python electro_assistive_predictive_model.py recommend --bundle results/electro_model.joblib \
      --data candidate.csv --row-index 0 --output results/recommendations.csv
"""

from __future__ import annotations

import argparse
import copy
import json
import math
import re
import sys
import time
import warnings
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

import joblib
import numpy as np
import pandas as pd
from scipy.special import expit
from sklearn.base import BaseEstimator, ClassifierMixin, clone
from sklearn.calibration import CalibratedClassifierCV
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import ExtraTreesClassifier, HistGradientBoostingClassifier
from sklearn.ensemble import HistGradientBoostingRegressor, RandomForestClassifier
from sklearn.impute import SimpleImputer
from sklearn.inspection import permutation_importance
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    average_precision_score,
    balanced_accuracy_score,
    brier_score_loss,
    confusion_matrix,
    f1_score,
    log_loss,
    matthews_corrcoef,
    precision_score,
    recall_score,
    roc_auc_score,
)
from sklearn.model_selection import GroupKFold, StratifiedGroupKFold, StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, RobustScaler

# Give custom classes a stable import path even when this file is executed as a
# script.  This makes joblib bundles loadable from notebooks and other programs.
sys.modules.setdefault("electro_assistive_predictive_model", sys.modules[__name__])

RANDOM_STATE = 42
EPS = 1e-9


def canonical(name: str) -> str:
    """Normalize column names while preserving familiar snake_case names."""
    text = str(name).strip().lower()
    text = re.sub(r"[%/]+", "_percent_", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


LABEL_ALIASES = {"pass_fail", "passfail", "label", "target", "passed", "is_pass"}
GROUP_ALIASES = {
    "bushing_id", "design_id", "model_id", "geometry_id", "part_number", "bushing_design_id"
}

# Geometry/design variables that recommendation search may change.  Bounds are
# deliberately local percentages until engineering absolute limits are supplied.
DEFAULT_MUTABLE_RULES: dict[str, dict[str, Any]] = {
    "conductor_diameter_mm": {"relative_low": 0.85, "relative_high": 1.15, "monotonic": 1},
    "shield_diameter_mm": {"relative_low": 0.85, "relative_high": 1.15, "monotonic": 0},
    "shield_length_mm": {"relative_low": 0.80, "relative_high": 1.20, "monotonic": 0},
    "shield_top_position_mm": {"relative_low": 0.90, "relative_high": 1.10, "monotonic": 0},
    "shield_bottom_position_mm": {"relative_low": 0.90, "relative_high": 1.10, "monotonic": 0},
    "top_bulb_distance_to_nearest_shed_mm": {"relative_low": 0.75, "relative_high": 1.25, "monotonic": 0},
    "bottom_bulb_distance_to_nearest_shed_mm": {"relative_low": 0.75, "relative_high": 1.25, "monotonic": 0},
    "shell_mean_diameter_mm": {"relative_low": 0.90, "relative_high": 1.10, "monotonic": 0},
    "creepage_distance_mm": {"relative_low": 0.90, "relative_high": 1.15, "monotonic": 1},
}

# Aliases allow the script to accept both the original human-readable vector and
# the later snake_case automated vector.
FEATURE_ALIASES = {
    "conductor_diameter": "conductor_diameter_mm",
    "shield_diameter": "shield_diameter_mm",
    "shield_length": "shield_length_mm",
    "outer_shell_mean_diameter": "shell_mean_diameter_mm",
    "outer_epoxy_mean_diameter": "shell_mean_diameter_mm",
    "creepage_distance": "creepage_distance_mm",
    "bil_voltage": "bil_voltage_kv",
    "simulation_voltage": "simulation_voltage_kv",
    "top_bulb_distance_to_nearest_shed": "top_bulb_distance_to_nearest_shed_mm",
    "bottom_bulb_distance_to_nearest_shed": "bottom_bulb_distance_to_nearest_shed_mm",
}

DEFAULT_RESPONSE_PATTERNS = [
    r"(^|_)global_[edpq]_(max|mean|p95|auc)$",
    r"(^|_)(conductor|shield|shell|top|bottom).*_[edpq]_(max|mean|p95|auc)$",
    r"surface_(bound_)?q", r"bound_charge", r"charge_(max|min|mean|p95|auc)",
]

TEMPLATE_COLUMNS = [
    "bushing_id", "simulation_id", "simulation_mode", "impulse_polarity",
    "simulation_voltage_kv", "bil_voltage_kv", "conductor_diameter_mm",
    "conductor_length_mm", "conductor_material", "shield_present",
    "shield_diameter_mm", "shield_length_mm", "shield_top_position_mm",
    "shield_bottom_position_mm", "shield_material",
    "top_bulb_distance_to_nearest_shed_mm", "bottom_bulb_distance_to_nearest_shed_mm",
    "shell_mean_diameter_mm", "shell_material", "creepage_distance_mm",
    "global_E_max", "global_E_mean", "global_E_p95", "global_E_auc",
    "global_E_peak_d_percent", "E_curve1_max", "E_curve1_mean", "E_curve1_p95",
    "E_curve1_auc", "E_curve1_peak_d_percent", "E_curve2_max", "E_curve2_mean",
    "E_curve2_p95", "E_curve2_auc", "E_curve2_peak_d_percent",
    "conductor_E_max", "conductor_E_mean", "conductor_E_p95", "conductor_E_auc",
    "conductor_E_peak_d_percent", "conductor_E_peak_zone_id",
    "conductor_E_max_over_global_E_max", "conductor_E_auc_over_global_E_auc",
    "shield_E_max", "shield_E_mean", "shield_E_p95", "shield_E_auc",
    "top_shed_E_max", "bottom_shed_E_max",
    "global_D_max", "global_P_max", "global_surface_bound_Q_max",
    "global_surface_bound_Q_min", "global_surface_bound_Q_abs_max",
    "conductor_surface_bound_Q_abs_max", "shield_surface_bound_Q_abs_max",
    "pass_fail",
]


@dataclass
class ElectroConfig:
    label_column: str = "pass_fail"
    group_column: str = "bushing_id"
    positive_label: Any = 1
    exclude_columns: list[str] = field(default_factory=lambda: ["simulation_id", "source_file", "notes"])
    immutable_columns: list[str] = field(default_factory=lambda: [
        "bushing_id", "simulation_id", "simulation_mode", "impulse_polarity",
        "simulation_voltage_kv", "bil_voltage_kv", "conductor_material",
        "shield_material", "shell_material", "pass_fail"
    ])
    mutable_rules: dict[str, dict[str, Any]] = field(default_factory=lambda: copy.deepcopy(DEFAULT_MUTABLE_RULES))
    response_patterns: list[str] = field(default_factory=lambda: list(DEFAULT_RESPONSE_PATTERNS))
    monotonic_constraints: dict[str, int] = field(default_factory=lambda: {
        "creepage_distance_mm": 1,
        "conductor_diameter_mm": 0,
        # Add only defensible directions.  1 raises pass probability; -1 lowers it.
    })
    required_clearance_mm: float = 1.0
    shell_wall_min_mm: float = 1.0
    recommendation_target_probability: float = 0.80
    max_relative_change_per_feature: float = 0.25
    max_features_changed: int = 3
    cv_folds: int = 5
    random_state: int = RANDOM_STATE
    future_optimizer: str = "reserved: Bayesian optimization or genetic algorithm after surrogate validation"


def load_config(path: str | None) -> ElectroConfig:
    cfg = ElectroConfig()
    if not path:
        return cfg
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    for key, value in raw.items():
        if not hasattr(cfg, key):
            warnings.warn(f"Ignoring unknown configuration key: {key}")
        else:
            setattr(cfg, key, value)
    cfg.label_column = canonical(cfg.label_column)
    cfg.group_column = canonical(cfg.group_column)
    cfg.exclude_columns = [canonical(x) for x in cfg.exclude_columns]
    cfg.immutable_columns = [canonical(x) for x in cfg.immutable_columns]
    cfg.mutable_rules = {canonical(k): v for k, v in cfg.mutable_rules.items()}
    cfg.monotonic_constraints = {canonical(k): int(v) for k, v in cfg.monotonic_constraints.items()}
    return cfg


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    renamed = {}
    seen: set[str] = set()
    for col in out.columns:
        new = FEATURE_ALIASES.get(canonical(col), canonical(col))
        if new in seen:
            raise ValueError(f"Columns collide after normalization: {col!r} -> {new!r}")
        renamed[col] = new
        seen.add(new)
    out = out.rename(columns=renamed)
    # Normalize common boolean/category encodings without modifying arbitrary text.
    for col in out.select_dtypes(include=["object", "string"]).columns:
        out[col] = out[col].astype("string").str.strip()
        out[col] = out[col].replace({"": pd.NA, "N/A": pd.NA, "n/a": pd.NA})
    return out


def resolve_column(columns: Sequence[str], requested: str, aliases: set[str]) -> str | None:
    requested = canonical(requested)
    if requested in columns:
        return requested
    candidates = [c for c in columns if canonical(c) in aliases]
    return candidates[0] if len(candidates) == 1 else None


def encode_label(series: pd.Series, positive_label: Any = 1) -> np.ndarray:
    if series.isna().any():
        raise ValueError("Labeled training rows contain missing pass/fail values.")
    text = series.astype(str).str.strip().str.lower()
    mapping = {
        "pass": 1, "passed": 1, "true": 1, "yes": 1, "y": 1, "1": 1, "1.0": 1,
        "fail": 0, "failed": 0, "false": 0, "no": 0, "n": 0, "0": 0, "0.0": 0,
    }
    if set(text.unique()).issubset(mapping):
        return text.map(mapping).to_numpy(dtype=int)
    unique = list(pd.unique(series))
    if len(unique) != 2:
        raise ValueError(f"Pass/fail target must contain exactly two classes; found {unique}")
    return (series == positive_label).astype(int).to_numpy()


def infer_feature_types(X: pd.DataFrame) -> tuple[list[str], list[str]]:
    categorical: list[str] = []
    numeric: list[str] = []
    for col in X.columns:
        s = X[col]
        if pd.api.types.is_bool_dtype(s) or isinstance(s.dtype, pd.CategoricalDtype):
            categorical.append(col)
        elif pd.api.types.is_numeric_dtype(s):
            numeric.append(col)
        else:
            converted = pd.to_numeric(s, errors="coerce")
            if converted.notna().mean() >= 0.90:
                X[col] = converted
                numeric.append(col)
            else:
                categorical.append(col)
    return numeric, categorical


def make_preprocessor(numeric: list[str], categorical: list[str], scale_numeric: bool) -> ColumnTransformer:
    num_steps: list[tuple[str, Any]] = [("impute", SimpleImputer(strategy="median", add_indicator=True))]
    if scale_numeric:
        num_steps.append(("scale", RobustScaler()))
    cat_pipe = Pipeline([
        ("impute", SimpleImputer(strategy="most_frequent")),
        ("onehot", OneHotEncoder(handle_unknown="ignore", sparse_output=False, min_frequency=2)),
    ])
    return ColumnTransformer([
        ("numeric", Pipeline(num_steps), numeric),
        ("categorical", cat_pipe, categorical),
    ], remainder="drop", verbose_feature_names_out=True)


def transformed_monotonic_vector(preprocessor: ColumnTransformer, feature_constraints: dict[str, int]) -> list[int]:
    names = preprocessor.get_feature_names_out()
    result: list[int] = []
    for name in names:
        direction = 0
        if name.startswith("numeric__"):
            raw = name.removeprefix("numeric__")
            # Missing indicators never receive monotonic constraints.
            if not raw.startswith("missingindicator_"):
                direction = int(feature_constraints.get(raw, 0))
        result.append(direction)
    return result


class CascadeForestClassifier(ClassifierMixin, BaseEstimator):
    """Small-data cascade forest with out-of-fold probability augmentation.

    Constraints are explicit limits on depth, leaves, estimators and cascade
    layers.  This is an experimental challenger, not an implementation of a
    specific proprietary "deep forest" package.
    """

    def __init__(self, n_layers: int = 3, n_estimators: int = 250, max_depth: int = 7,
                 min_samples_leaf: int = 4, inner_folds: int = 3,
                 min_improvement: float = 0.002, random_state: int = RANDOM_STATE):
        self.n_layers = n_layers
        self.n_estimators = n_estimators
        self.max_depth = max_depth
        self.min_samples_leaf = min_samples_leaf
        self.inner_folds = inner_folds
        self.min_improvement = min_improvement
        self.random_state = random_state

    def _pair(self, layer: int):
        common = dict(
            n_estimators=self.n_estimators, max_depth=self.max_depth,
            min_samples_leaf=self.min_samples_leaf, class_weight="balanced_subsample",
            max_features="sqrt", n_jobs=-1, random_state=self.random_state + layer,
        )
        return RandomForestClassifier(**common), ExtraTreesClassifier(**common)

    def fit(self, X, y):
        X = np.asarray(X, dtype=float)
        y = np.asarray(y, dtype=int)
        self.classes_ = np.array([0, 1])
        self.layers_: list[tuple[Any, Any]] = []
        current = X
        class_counts = np.bincount(y, minlength=2)
        folds = max(2, min(self.inner_folds, int(class_counts.min())))
        cv = StratifiedKFold(folds, shuffle=True, random_state=self.random_state)
        best_loss = np.inf
        for layer in range(self.n_layers):
            oof = np.zeros((len(y), 4), dtype=float)
            for train, valid in cv.split(current, y):
                rf, et = self._pair(layer)
                rf.fit(current[train], y[train]); et.fit(current[train], y[train])
                oof[valid, :2] = rf.predict_proba(current[valid])
                oof[valid, 2:] = et.predict_proba(current[valid])
            probability = (oof[:, 1] + oof[:, 3]) / 2
            loss = log_loss(y, np.clip(probability, 1e-6, 1 - 1e-6))
            if layer > 0 and best_loss - loss < self.min_improvement:
                break
            best_loss = min(best_loss, loss)
            rf, et = self._pair(layer)
            rf.fit(current, y); et.fit(current, y)
            self.layers_.append((rf, et))
            current = np.column_stack([current, oof])
        self.n_layers_fitted_ = len(self.layers_)
        return self

    def predict_proba(self, X):
        current = np.asarray(X, dtype=float)
        final = None
        for rf, et in self.layers_:
            p_rf, p_et = rf.predict_proba(current), et.predict_proba(current)
            final = (p_rf + p_et) / 2
            current = np.column_stack([current, p_rf, p_et])
        return final

    def predict(self, X):
        return (self.predict_proba(X)[:, 1] >= 0.5).astype(int)


def safe_splits(y: np.ndarray, groups: np.ndarray | None, requested: int, seed: int):
    min_class = int(np.bincount(y, minlength=2).min())
    if min_class < 2:
        raise ValueError("At least two labeled examples of each class are required.")
    if groups is not None and len(np.unique(groups)) >= 3:
        group_count = len(np.unique(groups))
        n = max(2, min(requested, min_class, group_count))
        splitter = StratifiedGroupKFold(n_splits=n, shuffle=True, random_state=seed)
        splits = list(splitter.split(np.zeros(len(y)), y, groups))
        # Some tiny grouped data sets cannot put both labels in each fold.
        if all(len(np.unique(y[tr])) == 2 and len(np.unique(y[va])) == 2 for tr, va in splits):
            return splits, f"StratifiedGroupKFold({n})"
        n = max(2, min(requested, group_count))
        splitter2 = GroupKFold(n_splits=n)
        splits = list(splitter2.split(np.zeros(len(y)), y, groups))
        if all(len(np.unique(y[tr])) == 2 for tr, _ in splits):
            return splits, f"GroupKFold({n})"
        warnings.warn("Groups could not produce valid class-balanced training folds; using stratified row CV.")
    n = max(2, min(requested, min_class))
    splitter = StratifiedKFold(n_splits=n, shuffle=True, random_state=seed)
    return list(splitter.split(np.zeros(len(y)), y)), f"StratifiedKFold({n})"


def candidate_models(numeric: list[str], categorical: list[str], constraints: dict[str, int], seed: int):
    # The two settings per family expose useful bias/variance choices without an
    # expensive search that would be unjustified at ~100 simulations.
    candidates: dict[str, Any] = {}
    for strength, l1 in [(0.10, 0.15), (1.0, 0.50), (10.0, 0.80)]:
        name = f"elastic_net_logistic_C{strength:g}_l1{l1:g}"
        candidates[name] = Pipeline([
            ("preprocess", make_preprocessor(numeric, categorical, True)),
            ("model", LogisticRegression(
                penalty="elasticnet", solver="saga", C=strength, l1_ratio=l1,
                class_weight="balanced", max_iter=10000, random_state=seed,
            )),
        ])

    for leaf, l2 in [(8, 1.0), (15, 3.0), (25, 5.0)]:
        pre = make_preprocessor(numeric, categorical, False)
        # Fit a disposable preprocessor so transformed feature names and the
        # corresponding constraint vector are known inside each outer fold.
        # A wrapper below replaces the model constraint after preprocessing.
        model = HistGradientBoostingClassifier(
            learning_rate=0.05, max_iter=250, max_leaf_nodes=leaf, max_depth=5,
            min_samples_leaf=max(4, leaf // 2), l2_regularization=l2,
            early_stopping=True, validation_fraction=0.15, n_iter_no_change=20,
            random_state=seed,
        )
        candidates[f"constrained_gbt_leaves{leaf}_l2{l2:g}"] = ConstraintAwarePipeline(pre, model, constraints)

    for depth, leaf in [(5, 5), (8, 4)]:
        candidates[f"constrained_deep_forest_depth{depth}_leaf{leaf}"] = Pipeline([
            ("preprocess", make_preprocessor(numeric, categorical, False)),
            ("model", CascadeForestClassifier(
                n_layers=3, n_estimators=220, max_depth=depth,
                min_samples_leaf=leaf, inner_folds=3, random_state=seed,
            )),
        ])
    return candidates


class ConstraintAwarePipeline(BaseEstimator, ClassifierMixin):
    """Fits preprocessing first, then applies named monotonic GBT constraints."""

    def __init__(self, preprocess, model, feature_constraints):
        self.preprocess = preprocess
        self.model = model
        self.feature_constraints = feature_constraints

    def fit(self, X, y):
        self.preprocess_ = clone(self.preprocess)
        Xt = self.preprocess_.fit_transform(X, y)
        vector = transformed_monotonic_vector(self.preprocess_, self.feature_constraints)
        self.model_ = clone(self.model).set_params(monotonic_cst=vector)
        self.model_.fit(Xt, y)
        self.classes_ = np.array([0, 1])
        return self

    def predict_proba(self, X):
        return self.model_.predict_proba(self.preprocess_.transform(X))

    def predict(self, X):
        return (self.predict_proba(X)[:, 1] >= 0.5).astype(int)


def metric_row(y: np.ndarray, p: np.ndarray, threshold: float = 0.5) -> dict[str, float]:
    p = np.clip(np.asarray(p, dtype=float), 1e-7, 1 - 1e-7)
    pred = (p >= threshold).astype(int)
    tn, fp, fn, tp = confusion_matrix(y, pred, labels=[0, 1]).ravel()
    return {
        "roc_auc": roc_auc_score(y, p),
        "pr_auc": average_precision_score(y, p),
        "log_loss": log_loss(y, p),
        "brier": brier_score_loss(y, p),
        "accuracy": accuracy_score(y, pred),
        "balanced_accuracy": balanced_accuracy_score(y, pred),
        "precision_pass": precision_score(y, pred, zero_division=0),
        "recall_pass": recall_score(y, pred, zero_division=0),
        "recall_fail": recall_score(1 - y, 1 - pred, zero_division=0),
        "f1": f1_score(y, pred, zero_division=0),
        "mcc": matthews_corrcoef(y, pred),
        "tn": int(tn), "fp": int(fp), "fn": int(fn), "tp": int(tp),
    }


def choose_threshold(y: np.ndarray, p: np.ndarray) -> float:
    # Prefer failure sensitivity when balanced accuracy is tied.
    best = (-np.inf, -np.inf, 0.5)
    for threshold in np.linspace(0.10, 0.90, 161):
        pred = (p >= threshold).astype(int)
        score = balanced_accuracy_score(y, pred)
        fail_recall = recall_score(1 - y, 1 - pred, zero_division=0)
        if (score, fail_recall) > best[:2]:
            best = (score, fail_recall, float(threshold))
    return best[2]


def bootstrap_intervals(y: np.ndarray, p: np.ndarray, groups: np.ndarray | None,
                        iterations: int = 500, seed: int = RANDOM_STATE):
    rng = np.random.default_rng(seed)
    values = {"roc_auc": [], "pr_auc": [], "balanced_accuracy": [], "brier": []}
    if groups is None:
        units = np.arange(len(y)); unit_rows = {i: np.array([i]) for i in units}
    else:
        units = np.unique(groups); unit_rows = {u: np.flatnonzero(groups == u) for u in units}
    for _ in range(iterations):
        sampled = rng.choice(units, size=len(units), replace=True)
        idx = np.concatenate([unit_rows[u] for u in sampled])
        if len(np.unique(y[idx])) < 2:
            continue
        m = metric_row(y[idx], p[idx])
        for key in values:
            values[key].append(m[key])
    return {key: [float(np.quantile(v, 0.025)), float(np.quantile(v, 0.975))]
            for key, v in values.items() if v}


def cross_validated_predictions(models: dict[str, Any], X: pd.DataFrame, y: np.ndarray,
                                splits, verbose: bool = True):
    predictions: dict[str, np.ndarray] = {}
    errors: dict[str, str] = {}
    for model_index, (name, estimator) in enumerate(models.items(), 1):
        if verbose:
            print(f"[{model_index}/{len(models)}] Evaluating {name} ...", flush=True)
        oof = np.full(len(y), np.nan)
        try:
            for fold, (train, valid) in enumerate(splits):
                fitted = clone(estimator)
                fitted.fit(X.iloc[train], y[train])
                oof[valid] = fitted.predict_proba(X.iloc[valid])[:, 1]
            if np.isnan(oof).any():
                raise RuntimeError("OOF prediction coverage is incomplete")
            predictions[name] = oof
        except Exception as exc:
            errors[name] = f"{type(exc).__name__}: {exc}"
            warnings.warn(f"{name} failed and will be skipped: {errors[name]}")
    if not predictions:
        raise RuntimeError(f"Every predictive model failed: {errors}")
    return predictions, errors


def add_hybrid_candidates(predictions: dict[str, np.ndarray]) -> dict[str, np.ndarray]:
    out = dict(predictions)
    baseline = max((k for k in predictions if k.startswith("elastic_net")),
                   key=lambda k: np.var(predictions[k]), default=None)
    gbt = max((k for k in predictions if k.startswith("constrained_gbt")),
              key=lambda k: np.var(predictions[k]), default=None)
    forest = max((k for k in predictions if k.startswith("constrained_deep")),
                 key=lambda k: np.var(predictions[k]), default=None)
    if baseline and gbt:
        for weight in (0.25, 0.50, 0.75):
            out[f"hybrid_logistic_{1-weight:.2f}_gbt_{weight:.2f}"] = (
                (1 - weight) * predictions[baseline] + weight * predictions[gbt]
            )
    if baseline and gbt and forest:
        out["hybrid_three_model_equal"] = (
            predictions[baseline] + predictions[gbt] + predictions[forest]
        ) / 3
    return out


def family(name: str) -> str:
    if name.startswith("elastic_net"): return "predictive_baseline"
    if name.startswith("constrained_gbt"): return "primary_nonlinear"
    if name.startswith("constrained_deep"): return "experimental_challenger"
    return "hybrid"


def select_models(y: np.ndarray, predictions: dict[str, np.ndarray], groups=None):
    rows = []
    for name, p in predictions.items():
        threshold = choose_threshold(y, p)
        row = {"model": name, "family": family(name), "threshold": threshold}
        row.update(metric_row(y, p, threshold))
        # Ranking rewards discrimination and calibration; PR-AUC is emphasized
        # because pass/fail data may be imbalanced.
        row["selection_score"] = (
            0.40 * row["pr_auc"] + 0.25 * row["roc_auc"] +
            0.20 * row["balanced_accuracy"] + 0.15 * (1 - row["brier"])
        )
        rows.append(row)
    metrics = pd.DataFrame(rows).sort_values(
        ["selection_score", "pr_auc", "brier"], ascending=[False, False, True]
    ).reset_index(drop=True)
    best_name = str(metrics.iloc[0]["model"])
    intervals = bootstrap_intervals(y, predictions[best_name], groups)
    return metrics, best_name, intervals


class ProbabilityEnsemble(ClassifierMixin, BaseEstimator):
    def __init__(self, components: list[tuple[Any, float]], threshold: float, name: str):
        self.components = components
        total = sum(w for _, w in components)
        self.weights = [w / total for _, w in components]
        self.threshold = float(threshold)
        self.name = name
        self.classes_ = np.array([0, 1])

    def fit(self, X, y=None):
        """Compatibility method: components are deliberately pre-fitted."""
        return self

    def predict_proba(self, X):
        p = sum(w * model.predict_proba(X)[:, 1]
                for (model, _), w in zip(self.components, self.weights))
        return np.column_stack([1 - p, p])

    def predict(self, X):
        return (self.predict_proba(X)[:, 1] >= self.threshold).astype(int)


def parse_hybrid(name: str, base_models: dict[str, Any], best_by_family: dict[str, str], threshold: float):
    if name.startswith("hybrid_logistic"):
        match = re.search(r"logistic_([0-9.]+)_gbt_([0-9.]+)", name)
        w1, w2 = float(match.group(1)), float(match.group(2))
        return ProbabilityEnsemble([
            (base_models[best_by_family["predictive_baseline"]], w1),
            (base_models[best_by_family["primary_nonlinear"]], w2),
        ], threshold, name)
    if name == "hybrid_three_model_equal":
        return ProbabilityEnsemble([
            (base_models[best_by_family[f]], 1.0) for f in
            ("predictive_baseline", "primary_nonlinear", "experimental_challenger")
        ], threshold, name)
    return base_models[name]


def match_response_columns(columns: Iterable[str], patterns: Sequence[str]) -> list[str]:
    compiled = [re.compile(p, re.I) for p in patterns]
    return [c for c in columns if any(p.search(c) for p in compiled)]


class SurrogateSet:
    """One constrained GBT regressor per post-simulation response."""
    def __init__(self, models: dict[str, Any], feature_columns: list[str], response_columns: list[str],
                 validation: pd.DataFrame):
        self.models = models
        self.feature_columns = feature_columns
        self.response_columns = response_columns
        self.validation = validation

    def predict(self, X: pd.DataFrame) -> pd.DataFrame:
        X = X.reindex(columns=self.feature_columns)
        return pd.DataFrame({name: model.predict(X) for name, model in self.models.items()}, index=X.index)


# Stable class locations for portable joblib serialization.
for _portable_class in (
    ElectroConfig, CascadeForestClassifier, ConstraintAwarePipeline,
    ProbabilityEnsemble, SurrogateSet,
):
    _portable_class.__module__ = "electro_assistive_predictive_model"


def fit_surrogates(df: pd.DataFrame, predictor_features: list[str], response_columns: list[str],
                   cfg: ElectroConfig, splits) -> SurrogateSet | None:
    # Surrogate inputs exclude simulation responses and identifiers.  They retain
    # operating conditions plus geometry/material information.
    response_set = set(response_columns)
    surrogate_features = [c for c in predictor_features if c not in response_set]
    if not response_columns or len(surrogate_features) < 2:
        return None
    X = df[surrogate_features].copy()
    numeric, categorical = infer_feature_types(X)
    validation_rows = []
    models: dict[str, Any] = {}
    for response in response_columns:
        y = pd.to_numeric(df[response], errors="coerce")
        known = y.notna().to_numpy()
        if known.sum() < max(20, 4 * max(1, len(surrogate_features) // 10)) or y[known].nunique() < 4:
            continue
        Xk, yk = X.loc[known], y.loc[known].to_numpy(float)
        # Separate folds are used because a response may have missing rows.
        folds = min(cfg.cv_folds, max(2, len(yk) // 10))
        cv = StratifiedKFold(n_splits=folds, shuffle=True, random_state=cfg.random_state)
        # Quantile bins preserve response range more reliably than plain KFold.
        bins = pd.qcut(yk, q=min(folds, max(2, len(np.unique(yk)))), labels=False, duplicates="drop")
        oof = np.full(len(yk), np.nan)
        for tr, va in cv.split(Xk, bins):
            pre = make_preprocessor(numeric, categorical, False)
            Xt = pre.fit_transform(Xk.iloc[tr])
            vector = transformed_monotonic_vector(pre, cfg.monotonic_constraints)
            reg = HistGradientBoostingRegressor(
                learning_rate=0.05, max_iter=250, max_leaf_nodes=15, max_depth=5,
                min_samples_leaf=max(5, min(15, len(tr) // 6)), l2_regularization=3.0,
                monotonic_cst=vector, early_stopping=True, random_state=cfg.random_state,
            )
            reg.fit(Xt, yk[tr]); oof[va] = reg.predict(pre.transform(Xk.iloc[va]))
        rmse = float(np.sqrt(np.mean((yk - oof) ** 2)))
        mae = float(np.mean(np.abs(yk - oof)))
        baseline_mae = float(np.mean(np.abs(yk - np.median(yk))))
        r2 = float(1 - np.sum((yk - oof) ** 2) / max(EPS, np.sum((yk - yk.mean()) ** 2)))
        validation_rows.append({
            "response": response, "n": len(yk), "mae": mae, "rmse": rmse,
            "r2": r2, "median_baseline_mae": baseline_mae,
            "skill_vs_median": 1 - mae / max(EPS, baseline_mae),
            "recommendation_ready": bool(r2 > 0 and mae < baseline_mae),
        })
        pre = make_preprocessor(numeric, categorical, False)
        Xt = pre.fit_transform(Xk)
        vector = transformed_monotonic_vector(pre, cfg.monotonic_constraints)
        reg = HistGradientBoostingRegressor(
            learning_rate=0.05, max_iter=250, max_leaf_nodes=15, max_depth=5,
            min_samples_leaf=max(5, min(15, len(yk) // 6)), l2_regularization=3.0,
            monotonic_cst=vector, early_stopping=True, random_state=cfg.random_state,
        )
        reg.fit(Xt, yk)
        models[response] = Pipeline([("preprocess", pre), ("model", reg)])
    if not models:
        return None
    validation = pd.DataFrame(validation_rows).sort_values("skill_vs_median", ascending=False)
    return SurrogateSet(models, surrogate_features, list(models), validation)


def check_physics(row: pd.Series, cfg: ElectroConfig) -> tuple[bool, list[str]]:
    reasons: list[str] = []
    def value(name):
        return pd.to_numeric(pd.Series([row.get(name, np.nan)]), errors="coerce").iloc[0]
    positive_names = [name for name in cfg.mutable_rules if name in row.index]
    for name in positive_names:
        v = value(name)
        if pd.notna(v) and v <= 0:
            reasons.append(f"{name} must be positive")
    conductor, shield = value("conductor_diameter_mm"), value("shield_diameter_mm")
    shell = value("shell_mean_diameter_mm")
    if pd.notna(conductor) and pd.notna(shield) and shield < conductor + 2 * cfg.required_clearance_mm:
        reasons.append("shield diameter violates conductor radial clearance")
    inner = shield if pd.notna(shield) else conductor
    if pd.notna(inner) and pd.notna(shell) and shell < inner + 2 * cfg.shell_wall_min_mm:
        reasons.append("shell diameter violates minimum wall/clearance")
    shield_len, conductor_len = value("shield_length_mm"), value("conductor_length_mm")
    if pd.notna(shield_len) and pd.notna(conductor_len) and shield_len >= conductor_len:
        reasons.append("shield length must be shorter than conductor length")
    return not reasons, reasons


def local_bounds(row: pd.Series, df_reference: pd.DataFrame, cfg: ElectroConfig):
    result = {}
    for name, rule in cfg.mutable_rules.items():
        if name not in row or name not in df_reference:
            continue
        current = pd.to_numeric(pd.Series([row[name]]), errors="coerce").iloc[0]
        observed = pd.to_numeric(df_reference[name], errors="coerce").dropna()
        if pd.isna(current) or observed.empty:
            continue
        low = rule.get("absolute_low", current * rule.get("relative_low", 1 - cfg.max_relative_change_per_feature))
        high = rule.get("absolute_high", current * rule.get("relative_high", 1 + cfg.max_relative_change_per_feature))
        # Do not extrapolate beyond a modest 5% margin of observed training data.
        span = max(EPS, observed.max() - observed.min())
        low = max(float(low), float(observed.min() - 0.05 * span))
        high = min(float(high), float(observed.max() + 0.05 * span))
        if high > low:
            result[name] = (low, high, float(current), float(max(EPS, observed.std())))
    return result


def generate_counterfactuals(row: pd.Series, bounds, n: int, max_changed: int, seed: int):
    rng = np.random.default_rng(seed)
    names = list(bounds)
    candidates = []
    for _ in range(n):
        new = row.copy()
        k = int(rng.integers(1, min(max_changed, len(names)) + 1))
        chosen = rng.choice(names, size=k, replace=False)
        for name in chosen:
            low, high, current, _ = bounds[name]
            # Triangular sampling focuses on smaller, manufacturable changes.
            new[name] = rng.triangular(low, current, high)
        candidates.append(new)
    return pd.DataFrame(candidates)


def recommend(bundle: dict[str, Any], row: pd.Series, n_candidates: int, n_results: int):
    cfg: ElectroConfig = bundle["config"]
    row = normalize_frame(pd.DataFrame([row])).iloc[0]
    features = bundle["feature_columns"]
    for col in features:
        if col not in row:
            row[col] = np.nan
    bounds = local_bounds(row, bundle["training_reference"], cfg)
    if not bounds:
        raise ValueError("No configured mutable numeric feature is present in both this row and training data.")
    raw = generate_counterfactuals(row, bounds, n_candidates, cfg.max_features_changed, cfg.random_state)
    valid_mask, violation_text = [], []
    for _, candidate in raw.iterrows():
        valid, reasons = check_physics(candidate, cfg)
        valid_mask.append(valid); violation_text.append("; ".join(reasons))
    candidates = raw.loc[valid_mask].copy()
    if candidates.empty:
        raise RuntimeError("Every candidate violated configured geometry/physics constraints.")

    surrogate: SurrogateSet | None = bundle.get("surrogates")
    trusted_responses: list[str] = []
    if surrogate is not None:
        ready = surrogate.validation.set_index("response")["recommendation_ready"].to_dict()
        trusted_responses = [c for c in surrogate.response_columns if ready.get(c, False)]
        if trusted_responses:
            response_predictions = surrogate.predict(candidates)
            for col in trusted_responses:
                candidates[col] = response_predictions[col]

    model = bundle["selected_model"]
    base_probability = float(model.predict_proba(pd.DataFrame([row]).reindex(columns=features))[:, 1][0])
    candidates["predicted_pass_probability"] = model.predict_proba(candidates.reindex(columns=features))[:, 1]
    candidates["probability_improvement"] = candidates["predicted_pass_probability"] - base_probability
    changed_lists, costs = [], []
    for _, candidate in candidates.iterrows():
        changes, cost = [], 0.0
        for name, (_, _, current, scale) in bounds.items():
            delta = float(candidate[name]) - current
            if abs(delta) > 1e-8:
                changes.append(name)
                cost += abs(delta) / scale
        changed_lists.append(", ".join(changes)); costs.append(cost)
    candidates["changed_features"] = changed_lists
    candidates["normalized_change_cost"] = costs
    # Pareto-like utility: high pass probability, small design change.  Failure
    # probability is never treated as an acceptable optimization tradeoff.
    candidates["recommendation_score"] = (
        candidates["predicted_pass_probability"] - 0.025 * candidates["normalized_change_cost"]
    )
    candidates["meets_probability_target"] = (
        candidates["predicted_pass_probability"] >= cfg.recommendation_target_probability
    )
    candidates["base_pass_probability"] = base_probability
    candidates["surrogate_updated_responses"] = ", ".join(trusted_responses)
    candidates["recommendation_status"] = np.where(
        candidates["probability_improvement"] > 0,
        "candidate_improvement_not_engineering_approval", "no_predicted_improvement"
    )
    return candidates.sort_values(
        ["meets_probability_target", "recommendation_score", "normalized_change_cost"],
        ascending=[False, False, True],
    ).head(n_results)


def feature_importance(model, X: pd.DataFrame, y: np.ndarray, seed: int) -> pd.DataFrame:
    try:
        sample_n = min(500, len(X))
        result = permutation_importance(
            model, X.iloc[:sample_n], y[:sample_n], scoring="average_precision",
            n_repeats=15, random_state=seed, n_jobs=-1,
        )
        return pd.DataFrame({
            "feature": X.columns, "importance_mean": result.importances_mean,
            "importance_std": result.importances_std,
        }).sort_values("importance_mean", ascending=False)
    except Exception as exc:
        warnings.warn(f"Permutation importance unavailable: {exc}")
        return pd.DataFrame(columns=["feature", "importance_mean", "importance_std"])


def train(args):
    cfg = load_config(args.config)
    output = Path(args.output_dir); output.mkdir(parents=True, exist_ok=True)
    df = normalize_frame(pd.read_csv(args.data))
    label = resolve_column(df.columns, cfg.label_column, LABEL_ALIASES)
    if not label:
        raise ValueError(f"Could not resolve pass/fail label. Columns: {list(df.columns)}")
    group = resolve_column(df.columns, cfg.group_column, GROUP_ALIASES)
    labeled = df[df[label].notna()].reset_index(drop=True)
    unlabeled_count = int(df[label].isna().sum())
    y = encode_label(labeled[label], cfg.positive_label)
    if len(y) < 20:
        warnings.warn("Fewer than 20 labeled simulations: metrics are smoke tests, not evidence of generalization.")
    if min(np.bincount(y, minlength=2)) < 5:
        warnings.warn("One class has fewer than five rows; uncertainty will be very high.")

    excluded = set(cfg.exclude_columns) | {label}
    if group:
        excluded.add(group)  # group IDs are never predictive inputs
    feature_columns = [c for c in labeled.columns if c not in excluded and not labeled[c].isna().all()]
    # Constant columns cannot teach relationships and can destabilize tiny data.
    constant = [c for c in feature_columns if labeled[c].nunique(dropna=True) <= 1]
    feature_columns = [c for c in feature_columns if c not in constant]
    X = labeled[feature_columns].copy()
    numeric, categorical = infer_feature_types(X)
    groups = labeled[group].fillna("__missing_group__").astype(str).to_numpy() if group else None
    splits, cv_name = safe_splits(y, groups, cfg.cv_folds, cfg.random_state)

    print(f"Training rows: {len(labeled)} ({int(y.sum())} pass, {int((1-y).sum())} fail)")
    print(f"Features: {len(feature_columns)} ({len(numeric)} numeric, {len(categorical)} categorical)")
    print(f"Validation: {cv_name}; unlabeled rows retained outside classifier: {unlabeled_count}")
    candidates = candidate_models(numeric, categorical, cfg.monotonic_constraints, cfg.random_state)
    base_predictions, errors = cross_validated_predictions(candidates, X, y, splits)

    # Choose the strongest variant inside each architecture family before hybridizing.
    base_metrics, _, _ = select_models(y, base_predictions, groups)
    best_by_family = {}
    for fam in ("predictive_baseline", "primary_nonlinear", "experimental_challenger"):
        subset = base_metrics[base_metrics.family == fam]
        if not subset.empty:
            best_by_family[fam] = str(subset.iloc[0].model)
    chosen_family_predictions = {name: base_predictions[name] for name in best_by_family.values()}
    all_predictions = add_hybrid_candidates(chosen_family_predictions)
    metrics, best_name, intervals = select_models(y, all_predictions, groups)

    # Fit selected variants and all architecture winners on all labeled data.
    fitted_base = {}
    for name in set(best_by_family.values()):
        fitted_base[name] = clone(candidates[name]).fit(X, y)
    best_threshold = float(metrics.loc[metrics.model == best_name, "threshold"].iloc[0])
    selected_model = parse_hybrid(best_name, fitted_base, best_by_family, best_threshold)
    if not best_name.startswith("hybrid"):
        selected_model.threshold = best_threshold

    response_columns = match_response_columns(feature_columns, cfg.response_patterns)
    print(f"Fitting assistive regression surrogates for {len(response_columns)} candidate responses ...")
    # Rows without pass/fail labels cannot train the classifier, but their known
    # simulation responses remain valid supervised labels for assistive surrogates.
    surrogate = fit_surrogates(df, feature_columns, response_columns, cfg, splits)

    importance = feature_importance(selected_model, X, y, cfg.random_state)
    oof = pd.DataFrame({"actual_pass_fail": y})
    if group:
        oof[group] = labeled[group].values
    for name, values in all_predictions.items():
        oof[f"probability__{name}"] = values

    training_reference = df.reindex(columns=feature_columns).copy()
    bundle = {
        "format_version": 1,
        "created_unix": time.time(),
        "config": cfg,
        "label_column": label,
        "group_column": group,
        "feature_columns": feature_columns,
        "numeric_columns": numeric,
        "categorical_columns": categorical,
        "constant_columns_removed": constant,
        "selected_model_name": best_name,
        "selected_threshold": best_threshold,
        "selected_model": selected_model,
        "architecture_winners": fitted_base,
        "best_by_family": best_by_family,
        "surrogates": surrogate,
        "training_reference": training_reference,
        "class_counts": {"fail": int((1-y).sum()), "pass": int(y.sum())},
        "cv_method": cv_name,
    }
    joblib.dump(bundle, output / "electro_model.joblib", compress=3)
    metrics.to_csv(output / "model_metrics.csv", index=False)
    base_metrics.to_csv(output / "all_candidate_metrics.csv", index=False)
    oof.to_csv(output / "oof_predictions.csv", index=False)
    importance.to_csv(output / "permutation_importance.csv", index=False)
    if surrogate is not None:
        surrogate.validation.to_csv(output / "surrogate_metrics.csv", index=False)

    report = {
        "selected_model": best_name,
        "selected_threshold": best_threshold,
        "selected_model_metrics": metrics.iloc[0].to_dict(),
        "bootstrap_95_percent_intervals": intervals,
        "architecture_winners": best_by_family,
        "validation": cv_name,
        "labeled_rows": len(labeled), "unlabeled_rows": unlabeled_count,
        "class_counts": bundle["class_counts"],
        "feature_count": len(feature_columns),
        "constant_columns_removed": constant,
        "surface_bound_charge_features": [c for c in feature_columns if re.search(r"(surface.*q|bound.*charge)", c, re.I)],
        "surrogate_count": 0 if surrogate is None else len(surrogate.models),
        "failed_candidates": errors,
        "important_interpretation": [
            "Model ranking is based on grouped out-of-fold predictions where feasible.",
            "Hybrid weights and operating threshold are exploratory because they are selected from the same OOF predictions.",
            "Counterfactual recommendations are decision support, not causal proof or engineering approval.",
            "Only surrogate responses outperforming a median baseline are propagated into recommendations.",
            cfg.future_optimizer,
        ],
    }
    (output / "training_report.json").write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    print("\nModel ranking (out-of-fold):")
    print(metrics[["model", "family", "selection_score", "pr_auc", "roc_auc", "balanced_accuracy", "brier", "threshold"]].to_string(index=False))
    print(f"\nSelected: {best_name} at pass threshold {best_threshold:.3f}")
    print(f"Saved training artifacts to {output.resolve()}")


def predict_command(args):
    bundle = joblib.load(args.bundle)
    df = normalize_frame(pd.read_csv(args.data))
    features = bundle["feature_columns"]
    missing = [c for c in features if c not in df]
    for col in missing:
        df[col] = np.nan
    probability = bundle["selected_model"].predict_proba(df.reindex(columns=features))[:, 1]
    threshold = bundle["selected_threshold"]
    result = df.copy()
    result["predicted_pass_probability"] = probability
    result["predicted_pass_fail"] = (probability >= threshold).astype(int)
    result["model_name"] = bundle["selected_model_name"]
    result["decision_threshold"] = threshold
    output = Path(args.output or "electro_predictions.csv")
    result.to_csv(output, index=False)
    print(result[["predicted_pass_probability", "predicted_pass_fail"]].to_string(index=False))
    print(f"Saved {output.resolve()}")


def recommend_command(args):
    bundle = joblib.load(args.bundle)
    df = normalize_frame(pd.read_csv(args.data))
    if args.row_index < 0 or args.row_index >= len(df):
        raise IndexError(f"row-index {args.row_index} is outside 0..{len(df)-1}")
    result = recommend(bundle, df.iloc[args.row_index], args.candidates, args.results)
    output = Path(args.output or "electro_recommendations.csv")
    result.to_csv(output, index=False)
    display = ["predicted_pass_probability", "probability_improvement", "changed_features",
               "normalized_change_cost", "meets_probability_target", "recommendation_status"]
    print(result[display].to_string(index=False))
    print(f"Saved {output.resolve()}")


def template_command(args):
    directory = Path(args.directory); directory.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(columns=TEMPLATE_COLUMNS).to_csv(directory / "electro_input_template.csv", index=False)
    cfg = ElectroConfig()
    (directory / "electro_model_config.json").write_text(
        json.dumps(asdict(cfg), indent=2), encoding="utf-8"
    )
    print(f"Created template and editable configuration in {directory.resolve()}")


def build_parser():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("template", help="create an ELECTRO CSV template and editable constraints")
    p.add_argument("--directory", default="electro_template")
    p.set_defaults(func=template_command)
    p = sub.add_parser("train", help="train, compare and save predictive/assistive models")
    p.add_argument("--data", required=True, help="CSV containing labeled ELECTRO simulations")
    p.add_argument("--config", help="optional JSON generated by the template command")
    p.add_argument("--output-dir", default="electro_results")
    p.set_defaults(func=train)
    p = sub.add_parser("predict", help="predict pass/fail probability")
    p.add_argument("--bundle", required=True, help="trained electro_model.joblib")
    p.add_argument("--data", required=True)
    p.add_argument("--output")
    p.set_defaults(func=predict_command)
    p = sub.add_parser("recommend", help="physics-constrained counterfactual search")
    p.add_argument("--bundle", required=True)
    p.add_argument("--data", required=True, help="CSV containing the design row")
    p.add_argument("--row-index", type=int, default=0)
    p.add_argument("--candidates", type=int, default=5000)
    p.add_argument("--results", type=int, default=10)
    p.add_argument("--output")
    p.set_defaults(func=recommend_command)
    return parser


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        if getattr(args, "debug", False):
            raise
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
