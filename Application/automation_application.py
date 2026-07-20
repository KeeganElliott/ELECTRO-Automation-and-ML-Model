from __future__ import annotations

"""
ELECTRO internal data-collection application.

Data architecture
-----------------
1. design_profiles.csv          one authoritative row per design/variant
2. simulation_records.csv       one authoritative row per simulation
3. electro_master_input_vector          merged audit/research table
4. electro_expanded_model_ready_vector  former full research learning table
5. electro_model_ready_input_vector     fixed 80-feature initial ML table

Generated/imported artifacts are organized as:
    ELECTRO_Data_Collected/<active_design_id>/
"""

import csv
import gc
import importlib
import importlib.util
import subprocess
import json
import os
import re
import shutil
import sys
import traceback
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import numpy as np
    import pandas as pd
except ImportError as exc:
    raise SystemExit("Install numpy, pandas, and openpyxl: py -m pip install numpy pandas openpyxl") from exc

try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None
    filedialog = None


# Packages are checked against the exact interpreter running this application.
# This prevents `py -m pip` from installing into a different Python environment.
REQUIRED_PACKAGES = {
    "numpy": "numpy",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "matplotlib": "matplotlib",
    "scipy": "scipy",
    "pywt": "PyWavelets",
    "win32com.client": "pywin32",
    "pdfplumber": "pdfplumber",
    "fitz": "pymupdf",
    "PIL": "pillow",
    "pytesseract": "pytesseract",
}

def ensure_runtime_dependencies() -> None:
    missing = []
    failures = []
    for import_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ModuleNotFoundError as exc:
            if exc.name == import_name.split(".")[0]:
                missing.append(pip_name)
            else:
                failures.append((import_name, repr(exc)))
        except Exception as exc:
            failures.append((import_name, repr(exc)))

    if missing:
        print("\nInstalling missing packages into the Python interpreter running this app:")
        print(sys.executable)
        print("Packages:", ", ".join(sorted(set(missing))))
        cmd = [sys.executable, "-m", "pip", "install", "--upgrade", *sorted(set(missing))]
        result = subprocess.run(cmd, check=False)
        if result.returncode != 0:
            raise RuntimeError(
                "Automatic dependency installation failed. Run this exact command in PowerShell:\n"
                + f'"{sys.executable}" -m pip install --upgrade ' + " ".join(sorted(set(missing)))
            )
        importlib.invalidate_caches()

    # Recheck and expose the true underlying exception.
    final_failures = []
    for import_name in REQUIRED_PACKAGES:
        try:
            importlib.import_module(import_name)
        except Exception as exc:
            final_failures.append((import_name, repr(exc)))
    if final_failures:
        details = "\n".join(f"  {name}: {error}" for name, error in final_failures)
        raise RuntimeError(
            "One or more dependencies are installed but cannot be imported by the current Python interpreter.\n"
            f"Interpreter: {sys.executable}\n{details}\n\n"
            "Repair using this interpreter, not a generic `py` command."
        )

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "ELECTRO_Data_Collected"
DATA_DIR.mkdir(parents=True, exist_ok=True)

DESIGN_DB_CSV = DATA_DIR / "design_profiles.csv"
SIMULATION_DB_CSV = DATA_DIR / "simulation_records.csv"
STATE_JSON = DATA_DIR / "application_state.json"
MASTER_CSV = DATA_DIR / "electro_master_input_vector.csv"
MASTER_XLSX = DATA_DIR / "electro_master_input_vector.xlsx"
EXPANDED_MODEL_READY_CSV = DATA_DIR / "electro_expanded_model_ready_input_vector.csv"
EXPANDED_MODEL_READY_XLSX = DATA_DIR / "electro_expanded_model_ready_input_vector.xlsx"
MODEL_READY_CSV = DATA_DIR / "electro_model_ready_input_vector.csv"
MODEL_READY_XLSX = DATA_DIR / "electro_model_ready_input_vector.xlsx"

SCRIPT_CANDIDATES = {
    "solidworks": ["solidworks_geometry.py"],
    "electro_geometry": ["electro_geometry.py"],
    "creepage": ["creepage.py", "pdf.py"],
    "tier1": ["tier1.py"],
    "simulation": ["electro_automation.py"],
}

# Option 13 uses the same COM program identifier as the companion ELECTRO
# scripts whenever it can discover one. The fallback is the identifier used by
# current ELECTRO Python automation examples. Override it without editing this
# file by setting ELECTRO_COM_PROGID in Windows before launching the app.
DEFAULT_ELECTRO_COM_PROGID = "Electro.Application"

# Current shed-removal creepage calculator. These constants are specific to
# the 20-700-000 design and must not be generalized to another design family.
SHED_CALCULATION_DESIGN = "20-700-000"
EXTERNAL_BASE_CREEPAGE_MM = 1736.0
INTERNAL_BASE_CREEPAGE_MM = 1397.0
EXTERNAL_CREEPAGE_LOSS_PER_SHED_MM = 67.0
INTERNAL_CREEPAGE_LOSS_PER_SHED_MM = 37.0
MIN_EXTERNAL_CREEPAGE_MM = 1473.2
MIN_INTERNAL_CREEPAGE_MM = 1168.4
MAX_EXTERNAL_SHEDS_REMOVED = int(
    (EXTERNAL_BASE_CREEPAGE_MM - MIN_EXTERNAL_CREEPAGE_MM)
    / EXTERNAL_CREEPAGE_LOSS_PER_SHED_MM
)
MAX_INTERNAL_SHEDS_REMOVED = int(
    (INTERNAL_BASE_CREEPAGE_MM - MIN_INTERNAL_CREEPAGE_MM)
    / INTERNAL_CREEPAGE_LOSS_PER_SHED_MM
)
MM_PER_INCH = 25.4

# Provenance, paths, and free text are excluded from the learning matrix.
# design_id and simulation_id are retained as non-predictive metadata so the
# training script can group by design and trace individual simulations.
# pass_fail_label is retained for human readability, while pass_fail_code is
# the numeric target used by the classifier.
MODEL_EXCLUDE_EXACT = {
    "row_id", "simulation_record_id",
    "created_at", "updated_at", "collected_at", "source_file",
    "electro_model_name", "notes",
    "drawing_pdf_file", "drawing_pdf_path", "electro_export_path",
    "solidworks_output", "electro_geometry_output", "creepage_output",
    "solidworks_assembly", "conductor_component", "shield_diameter_component",
    "shield_length_component",
    # Retained in the design profile for correction auditing, but excluded
    # from learning matrices so only the corrected canonical values are used.
    "shield_diameter_segment_raw_mm", "shell_mean_diameter_segment_raw_mm",
    "diameter_correction_mm", "diameter_correction_sign",
    "diameter_correction_conductor_nominal_mm",
    "diameter_correction_conductor_nominal_label",
    "diameter_calculation_method",
}
MODEL_EXCLUDE_SUBSTRINGS = ("_path", "_file", "_output", "_error", "context", "keyword")

# Legacy moderate reduction for the renamed expanded model-ready vector only.
# The compact vector below uses its own fixed 80-feature schema, while the audit
# master and Tier 1 outputs retain every extracted feature.
MODEL_EXCLUDE_EXACT.update({
    # Export/sampling metadata rather than bushing physics.
    "num_total_points",
    "num_E_points",
    "n_zones",

    # Fixed outer zoning boundaries. Interior boundaries remain available.
    "zone_edge_0_percent",
    "zone_edge_4_percent",

    # Processing-status flags rather than physical predictors.
    "dsp_features_enabled",
    "dsp_wavelet_available",
    "dsp_fft_available",

    # Procedural/constant descriptors in the current workflow.
    "zoning_method",
    "zoning_field",
    "tier1_field",
    "zoning_field_E",
    "tier1_field_E",

    # Duplicate unit representations; millimetres remain authoritative.
    "top_creepage_distance_in",
    "bottom_creepage_distance_in",
})

# Regex patterns for moderately redundant columns. Zone-specific maxima,
# means, p95 values, AUC values, normalized ratios, peak distances, FFT/PSD,
# and wavelet-energy descriptors are intentionally preserved.
MODEL_EXCLUDE_REGEX = [
    # Absolute coordinates depend on model origin and are less transferable.
    re.compile(r"_peak_[xyz]$", re.IGNORECASE),

    # A zone-specific peak is already known to belong to the zone named in
    # the column itself. Global and whole-curve peak-zone IDs are retained.
    re.compile(r"^zone\d+_.*_peak_zone_id$", re.IGNORECASE),
    re.compile(r"^E_curve\d+_zone\d+_.*_peak_zone_id$", re.IGNORECASE),

    # Zone start/end fields duplicate the retained interior zone boundaries.
    re.compile(r"^zone\d+_.*_(start|end)_percent$", re.IGNORECASE),

    # Sampling interval used by the FFT, not a bushing characteristic.
    re.compile(r"_fft_dx_percent$", re.IGNORECASE),

    # Keep wavelet/FFT peak, minimum, and maximum magnitudes because they may
    # describe localized spatial shape. Remove only the duplicated derivative
    # descriptor; the FFT derivative feature remains available.
    re.compile(r"_wavelet_max_abs_dE_dd$", re.IGNORECASE),

    # Wavelength is the reciprocal of retained dominant spatial frequency.
    re.compile(r"_fft_dominant_wavelength_percent$", re.IGNORECASE),
]

MODEL_CATEGORICAL_COLUMNS = {
    "simulation_type", "voltage_polarity", "transient_waveform_name",
    "dominant_E_max_zone_label", "dominant_E_auc_zone_label",
}

# Retained in the model-ready file for grouping, traceability, and human review.
# These columns must be excluded from X by the training script.
MODEL_METADATA_COLUMNS = [
    "design_id",
    "simulation_id",
    "pass_fail_label",
]

# Fixed first-stage model schema. These are the 75 actual predictor dimensions;
# grouping/record metadata, the human label, and pass_fail_code are additional
# non-predictor columns. The expanded vector remains available for later feature
# studies after substantially more simulations have been collected.
COMPACT_MODEL_FEATURES = [
    # Operating condition and design ratings (5)
    "simulation_voltage_kv",
    "voltage_rating_kv",
    "bil_voltage_kv",
    "simulation_type_static",
    "voltage_polarity_negative",

    # Geometry (18 numeric/binary). Material selections are invariant across
    # designs and therefore excluded: constants cannot improve predictions.
    "shield_present",
    "conductor_diameter_mm",
    "conductor_length_mm",
    "shield_diameter_mm",
    "shield_length_mm",
    "shell_mean_diameter_mm",
    "top_creepage_distance_mm",
    "bottom_creepage_distance_mm",
    "total_creepage_distance_mm",
    "top_bulb_distance_to_nearest_shed_mm",
    "bottom_bulb_distance_to_nearest_shed_mm",
    "top_shed_outward_delta_y_mm",
    "bottom_shed_outward_delta_y_mm",
    "top_shed_outward_delta_x_mm",
    "bottom_shed_outward_delta_x_mm",
    "conductor_to_shield_radial_clearance_mm",
    "shield_to_shell_radial_clearance_mm",
    "shield_length_over_conductor_length",

    # Compact E-stress set (25), emphasizing top/bottom flashover zones
    "global_E_max",
    "global_E_mean",
    "global_E_p95",
    "global_E_auc",
    "global_E_peak_d_percent",
    "E_curve1_max",
    "E_curve1_p95",
    "E_curve2_max",
    "E_curve2_p95",
    "conductor_E_max",
    "conductor_E_p95",
    "conductor_E_max_over_global_E_max",
    "shield_E_max",
    "shield_E_p95",
    "shield_E_max_over_global_E_max",
    "top_stress_E_max",
    "top_stress_E_mean",
    "top_stress_E_p95",
    "top_stress_E_auc",
    "top_stress_E_max_over_global_E_max",
    "bottom_stress_E_max",
    "bottom_stress_E_mean",
    "bottom_stress_E_p95",
    "bottom_stress_E_auc",
    "bottom_stress_E_max_over_global_E_max",

    # DSP/spatial-shape set (12), including top/bottom localization
    "conductor_E_wavelet_high_freq_energy_ratio",
    "conductor_E_wavelet_peak_power_spatial_freq_cyc_per_mm",
    "shield_E_wavelet_high_freq_energy_ratio",
    "shield_E_wavelet_peak_power_spatial_freq_cyc_per_mm",
    "conductor_E_fft_max_abs_dE_dd",
    "conductor_E_fft_dominant_spatial_freq_cyc_per_mm",
    "shield_E_fft_max_abs_dE_dd",
    "shield_E_fft_dominant_spatial_freq_cyc_per_mm",
    "top_stress_E_wavelet_high_freq_energy_ratio",
    "top_stress_E_fft_max_abs_dE_dd",
    "bottom_stress_E_wavelet_high_freq_energy_ratio",
    "bottom_stress_E_fft_max_abs_dE_dd",

    # Surface-bound Q-stress set (15), emphasizing top/bottom flashover zones
    "global_surface_bound_Q_max",
    "global_surface_bound_Q_min",
    "global_surface_bound_Q_abs_max",
    "global_surface_bound_Q_p95_abs",
    "global_surface_bound_Q_auc_abs",
    "conductor_surface_bound_Q_abs_max",
    "shield_surface_bound_Q_abs_max",
    "top_stress_surface_bound_Q_abs_max",
    "top_stress_surface_bound_Q_p95_abs",
    "top_stress_surface_bound_Q_auc_abs",
    "top_stress_surface_bound_Q_abs_max_over_global_surface_bound_Q_abs_max",
    "bottom_stress_surface_bound_Q_abs_max",
    "bottom_stress_surface_bound_Q_p95_abs",
    "bottom_stress_surface_bound_Q_auc_abs",
    "bottom_stress_surface_bound_Q_abs_max_over_global_surface_bound_Q_abs_max",
]

assert len(COMPACT_MODEL_FEATURES) == 75, "Compact model schema must remain exactly 75 features."

# Design-level model inputs that may be manually overridden. Simulation
# conditions, E/Q stress results, DSP features, labels, derived ratios, and
# provenance fields are intentionally not editable through this menu.
MANUALLY_EDITABLE_DESIGN_FEATURES = [
    ("conductor_diameter_mm", "Conductor diameter (mm)"),
    ("conductor_length_mm", "Conductor length (mm)"),
    ("shield_diameter_mm", "Shield diameter (mm)"),
    ("shield_length_mm", "Shield length (mm)"),
    ("shell_mean_diameter_mm", "Shell mean diameter (mm)"),
    ("top_creepage_distance_mm", "Top/external creepage distance (mm)"),
    ("bottom_creepage_distance_mm", "Bottom/internal creepage distance (mm)"),
    ("top_bulb_distance_to_nearest_shed_mm", "Top bulb distance to nearest shed (mm)"),
    ("bottom_bulb_distance_to_nearest_shed_mm", "Bottom bulb distance to nearest shed (mm)"),
    ("top_shed_outward_delta_x_mm", "Top shed outward delta X (mm)"),
    ("top_shed_outward_delta_y_mm", "Top shed outward delta Y (mm)"),
    ("bottom_shed_outward_delta_x_mm", "Bottom shed outward delta X (mm)"),
    ("bottom_shed_outward_delta_y_mm", "Bottom shed outward delta Y (mm)"),
    ("voltage_rating_kv", "Regular/rated operating voltage (kV)"),
    ("bil_voltage_kv", "Rated BIL voltage (kV)"),
]

# Current and legacy collection scripts have used several equivalent names.
# The compact vector exposes only the canonical name on the left. Matching is
# case-insensitive, so historical E/Q capitalization remains usable.
COMPACT_FEATURE_ALIASES = {
    "voltage_rating_kv": [
        "rated_voltage_kv", "regular_voltage_rating_kv", "nominal_voltage_kv",
    ],
    "bil_voltage_kv": [
        "rated_bil_voltage_kv", "basic_impulse_level_kv", "bil_rating_kv",
    ],
    "shell_mean_diameter_mm": [
        "outer_shell_mean_diameter_mm", "shell_diameter_mm", "shell_diameter_electro_units",
    ],
    "conductor_diameter_mm": ["conductor_diameter_electro_units"],
    "conductor_length_mm": ["conductor_length_electro_units"],
    "shield_diameter_mm": ["shield_diameter_electro_units"],
    "shield_length_mm": ["shield_length_electro_units"],
    "top_bulb_distance_to_nearest_shed_mm": [
        "top_bulb_to_nearest_shed_mm", "top_bulb_shed_distance_mm",
        "upper_shield_to_shed_distance", "upper_shield_to_shed_distance_mm",
    ],
    "bottom_bulb_distance_to_nearest_shed_mm": [
        "bottom_bulb_to_nearest_shed_mm", "bottom_bulb_shed_distance_mm",
        "lower_shield_to_shed_distance", "lower_shield_to_shed_distance_mm",
    ],
    "top_shed_outward_delta_y_mm": [
        "top_outward_delta_y_mm", "upper_shed_outward_delta_y_mm",
        "upper_shield_to_shed_outward_delta_y",
        "upper_shield_to_shed_outward_delta_y_mm",
    ],
    "bottom_shed_outward_delta_y_mm": [
        "bottom_outward_delta_y_mm", "lower_shed_outward_delta_y_mm",
        "lower_shield_to_shed_outward_delta_y",
        "lower_shield_to_shed_outward_delta_y_mm",
    ],
    "top_shed_outward_delta_x_mm": [
        "top_outward_delta_x_mm", "upper_shed_outward_delta_x_mm",
        "upper_shield_to_shed_outward_delta_x",
        "upper_shield_to_shed_outward_delta_x_mm",
    ],
    "bottom_shed_outward_delta_x_mm": [
        "bottom_outward_delta_x_mm", "lower_shed_outward_delta_x_mm",
        "lower_shield_to_shed_outward_delta_x",
        "lower_shield_to_shed_outward_delta_x_mm",
    ],
    "global_surface_bound_Q_max": ["global_Q_max", "global_surface_bound_q_max"],
    "global_surface_bound_Q_min": ["global_Q_min", "global_surface_bound_q_min"],
    "global_surface_bound_Q_abs_max": [
        "global_Q_abs_max", "global_Q_max_abs", "global_surface_bound_q_abs_max",
    ],
    "global_surface_bound_Q_p95_abs": [
        "global_Q_p95_abs", "global_Q_abs_p95", "global_surface_bound_q_p95_abs",
    ],
    "global_surface_bound_Q_auc_abs": [
        "global_Q_auc_abs", "global_Q_abs_auc", "global_surface_bound_q_auc_abs",
    ],
    "conductor_surface_bound_Q_abs_max": [
        "conductor_Q_abs_max", "conductor_Q_max_abs", "conductor_surface_bound_q_abs_max",
    ],
    "shield_surface_bound_Q_abs_max": [
        "shield_Q_abs_max", "shield_Q_max_abs", "shield_surface_bound_q_abs_max",
    ],
    "top_stress_surface_bound_Q_abs_max": [
        "top_stress_Q_abs_max", "top_stress_Q_max_abs", "top_stress_surface_bound_q_abs_max",
    ],
    "bottom_stress_surface_bound_Q_abs_max": [
        "bottom_stress_Q_abs_max", "bottom_stress_Q_max_abs", "bottom_stress_surface_bound_q_abs_max",
    ],
}

COMPACT_CATEGORICAL_FEATURES: set[str] = set()

COMPACT_DERIVED_FEATURES = {
    "simulation_type_static", "voltage_polarity_negative", "shield_present",
    "total_creepage_distance_mm", "conductor_to_shield_radial_clearance_mm",
    "shield_to_shell_radial_clearance_mm", "shield_length_over_conductor_length",
}


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def sanitize_name(text: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text).strip()).strip("._")
    return safe or "unnamed_design"


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, str):
        text = value.strip()
        if text.lower() in {"", "n/a", "na", "none", "nan"}:
            return None
        try:
            return float(text)
        except ValueError:
            return text
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value



PASS_FAIL_TO_CODE = {
    "pass": 1,
    "passed": 1,
    "p": 1,
    "1": 1,
    "1.0": 1,
    "true": 1,
    "yes": 1,
    "fail": 0,
    "failed": 0,
    "f": 0,
    "0": 0,
    "0.0": 0,
    "false": 0,
    "no": 0,
}

PASS_FAIL_FROM_CODE = {
    1: "pass",
    0: "fail",
}


def normalize_pass_fail_pair(record: dict[str, Any]) -> dict[str, Any]:
    """
    Keep the human-readable and numeric pass/fail fields synchronized.

    Stored values:
        pass_fail_label = "pass", "fail", or None
        pass_fail_code  = 1, 0, or None

    A blank/NaN value means the simulation is unlabeled. Conflicting label/code
    pairs are rejected rather than silently corrupting the target.
    """
    output = dict(record)

    raw_label = output.get("pass_fail_label")
    raw_code = output.get("pass_fail_code")

    label: str | None = None
    code: int | None = None

    if raw_label is not None and not (
        isinstance(raw_label, float) and pd.isna(raw_label)
    ):
        label_text = str(raw_label).strip().lower()
        if label_text not in {"", "n/a", "na", "none", "nan", "unknown"}:
            if label_text not in PASS_FAIL_TO_CODE:
                raise ValueError(
                    "Unsupported pass/fail label "
                    f"{raw_label!r}. Use Pass, Fail, or N/A."
                )
            code = PASS_FAIL_TO_CODE[label_text]
            label = PASS_FAIL_FROM_CODE[code]

    if raw_code is not None and not (
        isinstance(raw_code, float) and pd.isna(raw_code)
    ):
        numeric_code = pd.to_numeric(
            pd.Series([raw_code]),
            errors="coerce",
        ).iloc[0]

        if pd.isna(numeric_code):
            raise ValueError(
                f"pass_fail_code must be 0, 1, or blank; received {raw_code!r}."
            )

        if float(numeric_code) not in {0.0, 1.0}:
            raise ValueError(
                f"pass_fail_code must be 0, 1, or blank; received {raw_code!r}."
            )
        numeric_code = int(numeric_code)

        if code is not None and numeric_code != code:
            raise ValueError(
                "pass_fail_label and pass_fail_code disagree: "
                f"{raw_label!r} versus {raw_code!r}."
            )

        code = numeric_code
        label = PASS_FAIL_FROM_CODE[code]

    output["pass_fail_label"] = label
    output["pass_fail_code"] = code
    return output


def normalize_pass_fail_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize existing simulation rows during rebuild/migration."""
    if df.empty:
        return df.copy()

    normalized_rows = [
        normalize_pass_fail_pair(
            {key: clean_value(value) for key, value in row.items()}
        )
        for row in df.to_dict(orient="records")
    ]
    return pd.DataFrame(normalized_rows)


def load_state() -> dict[str, Any]:
    if not STATE_JSON.exists():
        return {"active_design_id": None}
    try:
        data = json.loads(STATE_JSON.read_text(encoding="utf-8"))
        data.setdefault("active_design_id", None)
        return data
    except Exception:
        return {"active_design_id": None}


def save_state(state: dict[str, Any]) -> None:
    temp = STATE_JSON.with_suffix(".tmp")
    temp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    temp.replace(STATE_JSON)


def load_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    return pd.read_csv(path) if path.exists() else pd.DataFrame(columns=columns)


def load_design_db() -> pd.DataFrame:
    return load_csv(DESIGN_DB_CSV, ["design_id", "created_at", "updated_at"])


def load_simulation_db() -> pd.DataFrame:
    return load_csv(SIMULATION_DB_CSV, ["simulation_record_id", "design_id", "simulation_id", "collected_at"])


def save_design_db(df: pd.DataFrame) -> None:
    if "design_id" not in df.columns:
        raise RuntimeError("Design database is missing design_id.")
    df.drop_duplicates(subset=["design_id"], keep="last").to_csv(DESIGN_DB_CSV, index=False)


def design_folder(design_id: str) -> Path:
    folder = DATA_DIR / sanitize_name(design_id)
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def resolve_script(key: str) -> Path:
    for name in SCRIPT_CANDIDATES[key]:
        path = APP_DIR / name
        if path.exists():
            return path
    # Flexible fallback handles filename suffixes such as "(2)".
    stem_terms = {
        "solidworks": "solidworks_geometry",
        "electro_geometry": "electro_geometry",
        "creepage": "creepage",
        "tier1": "tier1",
        "simulation": "electro_automation",
    }
    matches = sorted(APP_DIR.glob(f"{stem_terms[key]}*.py"))
    if matches:
        return matches[-1]
    raise FileNotFoundError(f"No compatible {key} module was found in {APP_DIR}")


@contextmanager
def working_directory(path: Path):
    old = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(old)


def run_module_main(
    key: str,
    cwd: Path,
    env: dict[str, str] | None = None,
    main_kwargs: dict[str, Any] | None = None,
) -> Any:
    """
    Import and run a companion module inside the current Python process.

    ``main_kwargs`` are passed directly to the companion module's ``main``
    function. This is the preferred path for tier1.py because the already
    selected ELECTRO export can be supplied explicitly without opening a
    second file-selection dialog.

    ``env`` is retained for backwards compatibility with companion modules
    that still read environment variables.
    """
    script = resolve_script(key)
    old_env: dict[str, str | None] = {}

    if env:
        for name, value in env.items():
            old_env[name] = os.environ.get(name)
            os.environ[name] = str(value)

    unique_name = f"electro_internal_{key}_{uuid.uuid4().hex}"

    try:
        with working_directory(cwd):
            spec = importlib.util.spec_from_file_location(unique_name, script)

            if spec is None or spec.loader is None:
                raise RuntimeError(
                    f"Unable to import companion module:\n{script}"
                )

            module = importlib.util.module_from_spec(spec)
            sys.modules[unique_name] = module
            spec.loader.exec_module(module)

            main_fn = getattr(module, "main", None)
            if not callable(main_fn):
                raise RuntimeError(
                    f"{script.name} does not expose a callable main()."
                )

            return main_fn(**(main_kwargs or {}))

    finally:
        sys.modules.pop(unique_name, None)

        for name, old_value in old_env.items():
            if old_value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = old_value


def prompt_text(label: str, default: str = "", allow_blank: bool = True) -> str:
    suffix = f" [{default}]" if default else ""
    while True:
        value = input(f"{label}{suffix}: ").strip()
        if value:
            return value
        if default:
            return default
        if allow_blank:
            return ""
        print("A value is required.")


def prompt_optional_float(label: str, default: float | None = None) -> float | None:
    shown = "N/A" if default is None else str(default)
    while True:
        value = input(f"{label} [{shown}]: ").strip()
        if not value:
            return default
        if value.lower() in {"n/a", "na", "none"}:
            return None
        try:
            return float(value)
        except ValueError:
            print("Enter a number or N/A.")


def prompt_nonnegative_int(
    label: str,
    default: int = 0,
    maximum: int | None = None,
) -> int:
    """Prompt for a whole-number count that cannot be negative."""
    while True:
        value = input(f"{label} [{default}]: ").strip()
        if not value:
            return int(default)
        try:
            parsed = int(value)
        except ValueError:
            print("Enter a whole number greater than or equal to zero.")
            continue
        if parsed < 0:
            print("The number of removed sheds cannot be negative.")
            continue
        if maximum is not None and parsed > maximum:
            print(f"Enter a whole number from 0 through {maximum}.")
            continue
        return parsed


def choose_file(title: str, patterns: list[tuple[str, str]]) -> Path | None:
    if tk is None:
        value = prompt_text(title + " (full path)")
        return Path(value).expanduser().resolve() if value else None
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True); root.update()
    selected = filedialog.askopenfilename(parent=root, title=title, filetypes=patterns)
    root.destroy()
    return Path(selected).resolve() if selected else None


def get_active_design(state: dict[str, Any]) -> tuple[str, dict[str, Any]]:
    design_id = state.get("active_design_id")
    if not design_id:
        raise RuntimeError("No active design. Create or select one first.")
    df = load_design_db()
    match = df[df["design_id"].astype(str) == str(design_id)]
    if match.empty:
        raise RuntimeError(f"Active design '{design_id}' is missing from design_profiles.csv.")
    return str(design_id), {k: clean_value(v) for k, v in match.iloc[-1].to_dict().items()}


def upsert_design(design: dict[str, Any]) -> None:
    design_id = str(design["design_id"])
    df = load_design_db()
    df = df[df["design_id"].astype(str) != design_id] if not df.empty else df
    df = pd.concat([df, pd.DataFrame([design])], ignore_index=True, sort=False)
    save_design_db(df)
    design_folder(design_id)
    rebuild_outputs()


def append_simulation_record(row: dict[str, Any]) -> None:
    old = load_simulation_db()
    cleaned = {k: clean_value(v) for k, v in row.items()}
    cleaned = normalize_pass_fail_pair(cleaned)
    new = pd.DataFrame([cleaned])
    pd.concat([old, new], ignore_index=True, sort=False).to_csv(
        SIMULATION_DB_CSV,
        index=False,
    )
    rebuild_outputs()


def should_exclude_model_column(column_name: str) -> bool:
    """Return True when a column belongs in the audit data, not the ML matrix."""
    if column_name in MODEL_EXCLUDE_EXACT:
        return True

    lowered = column_name.lower()
    if any(token in lowered for token in MODEL_EXCLUDE_SUBSTRINGS):
        return True

    return any(pattern.search(column_name) for pattern in MODEL_EXCLUDE_REGEX)


def _normalize_region_label(value: Any) -> str | None:
    """Convert a user zone label into a stable model-feature prefix."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    text = sanitize_name(str(value).strip().lower())
    if not text or text in {"unknown", "ignore", "n_a", "na", "none", "nan"}:
        return None
    return text


def _zone_label_columns(columns: list[str]) -> dict[int, str]:
    """Return {zone_number: label_column_name} for zone1_label, zone2_label, ..."""
    output: dict[int, str] = {}
    for column in columns:
        match = re.fullmatch(r"zone(\d+)_label", column, re.IGNORECASE)
        if match:
            output[int(match.group(1))] = column
    return output


def _semantic_feature_name(column: str, zone_id: int, region: str) -> str | None:
    """
    Convert zone-number-dependent names into stable physical-region names.

    Examples
    --------
    zone1_top_stress_E_max -> top_stress_E_max
    E_curve1_zone1_top_stress_max -> E_curve1_top_stress_max
    zone3_bottom_stress_E_fft_total_psd_energy
        -> bottom_stress_E_fft_total_psd_energy
    """
    escaped_region = re.escape(region)
    patterns = (
        # Aggregate zone features.
        re.compile(
            rf"^zone{zone_id}_{escaped_region}_((?:E|Q)_.+|surface_bound_Q_.+)$",
            re.IGNORECASE,
        ),
        # Per-curve zone features.
        re.compile(
            rf"^(E_curve\d+)_zone{zone_id}_{escaped_region}_(.+)$",
            re.IGNORECASE,
        ),
    )

    aggregate_match = patterns[0].match(column)
    if aggregate_match:
        return f"{region}_{aggregate_match.group(1)}"

    curve_match = patterns[1].match(column)
    if curve_match:
        return f"{curve_match.group(1)}_{region}_{curve_match.group(2)}"

    return None


def _add_semantic_region_boundaries(
    output_row: dict[str, Any],
    source_row: pd.Series,
    zone_labels: dict[int, str],
) -> None:
    """
    Replace anonymous zone edges with boundaries tied to physical regions.

    For a row whose labels are:
        zone1=bottom_stress, zone2=conductor,
        zone3=top_stress, zone4=shield

    the output contains:
        bottom_stress_start_percent / end_percent / width_percent
        conductor_start_percent / end_percent / width_percent
        top_stress_start_percent / end_percent / width_percent
        shield_start_percent / end_percent / width_percent
    """
    for zone_id, region in zone_labels.items():
        start = pd.to_numeric(
            pd.Series([source_row.get(f"zone_edge_{zone_id - 1}_percent")]),
            errors="coerce",
        ).iloc[0]
        end = pd.to_numeric(
            pd.Series([source_row.get(f"zone_edge_{zone_id}_percent")]),
            errors="coerce",
        ).iloc[0]

        if pd.notna(start):
            output_row[f"{region}_start_percent"] = float(start)
        if pd.notna(end):
            output_row[f"{region}_end_percent"] = float(end)
        if pd.notna(start) and pd.notna(end):
            output_row[f"{region}_width_percent"] = float(end - start)


def _map_peak_zone_to_region(
    output_row: dict[str, Any],
    source_row: pd.Series,
    source_column: str,
    output_column: str,
    zone_labels: dict[int, str],
) -> None:
    """Translate a temporary numeric peak-zone ID into a physical region label."""
    raw = pd.to_numeric(pd.Series([source_row.get(source_column)]), errors="coerce").iloc[0]
    if pd.isna(raw):
        return

    region = zone_labels.get(int(raw))
    if region:
        output_row[output_column] = region


def semanticize_zone_features(master: pd.DataFrame) -> pd.DataFrame:
    """
    Build stable physical-region features before model-ready filtering.

    The audit master remains unchanged. Only the model-ready DataFrame is
    transformed. Zone numbers are temporary segmentation identifiers; physical
    labels such as top_stress, conductor, bottom_stress, and shield become the
    stable feature identities.
    """
    if master.empty:
        return master.copy()

    zone_label_cols = _zone_label_columns(list(master.columns))
    if not zone_label_cols:
        # Backward-compatible fallback for legacy data lacking zone label columns.
        return master.copy()

    semantic_rows: list[dict[str, Any]] = []

    for row_index, source_row in master.iterrows():
        zone_labels: dict[int, str] = {}
        used_regions: set[str] = set()

        for zone_id, label_column in sorted(zone_label_cols.items()):
            region = _normalize_region_label(source_row.get(label_column))
            if region is None:
                continue
            if region in used_regions:
                raise ValueError(
                    "Duplicate physical zone label in one simulation row: "
                    f"'{region}' appears more than once at row {row_index}. "
                    "Each physical region must be assigned to only one zone."
                )
            used_regions.add(region)
            zone_labels[zone_id] = region

        output_row: dict[str, Any] = {}

        # Copy all non-zone-specific fields. Raw zone labels, raw zone edges,
        # and temporary peak-zone IDs are handled separately below.
        for column, value in source_row.items():
            if re.fullmatch(r"zone\d+_label", column, re.IGNORECASE):
                continue
            if re.fullmatch(r"zone_edge_\d+_percent", column, re.IGNORECASE):
                continue
            if column in {
                "global_E_peak_zone_id",
                "E_curve1_peak_zone_id",
                "E_curve2_peak_zone_id",
            }:
                continue

            semantic_name = None
            for zone_id, region in zone_labels.items():
                semantic_name = _semantic_feature_name(column, zone_id, region)
                if semantic_name:
                    break

            if semantic_name:
                # A region feature may only be assigned once in a row.
                if semantic_name in output_row and pd.notna(output_row[semantic_name]):
                    raise ValueError(
                        f"Semantic feature collision for '{semantic_name}' at row {row_index}."
                    )
                output_row[semantic_name] = value
            else:
                output_row[column] = value

        _add_semantic_region_boundaries(output_row, source_row, zone_labels)

        _map_peak_zone_to_region(
            output_row,
            source_row,
            "global_E_peak_zone_id",
            "global_E_peak_region",
            zone_labels,
        )
        _map_peak_zone_to_region(
            output_row,
            source_row,
            "E_curve1_peak_zone_id",
            "E_curve1_peak_region",
            zone_labels,
        )
        _map_peak_zone_to_region(
            output_row,
            source_row,
            "E_curve2_peak_zone_id",
            "E_curve2_peak_region",
            zone_labels,
        )

        semantic_rows.append(output_row)

    return pd.DataFrame(semantic_rows)


def build_expanded_model_ready(master: pd.DataFrame) -> pd.DataFrame:
    """
    Create the former broad learning matrix using physical-region names.

    The audit master remains unchanged and retains temporary zone numbers.
    The model-ready table translates zone-number-dependent features into stable
    physical identities such as top_stress, conductor, bottom_stress, and shield.
    """
    if master.empty:
        return pd.DataFrame()

    semantic_master = semanticize_zone_features(master)

    keep = [
        col
        for col in semantic_master.columns
        if not should_exclude_model_column(col)
    ]
    learning = semantic_master[keep].copy()

    # Preserve only the intended human-readable pass/fail label. Other internal
    # label columns remain audit-only. Semantic peak-region fields are retained
    # and one-hot encoded below.
    label_cols = [
        c
        for c in learning.columns
        if c.endswith("_label") and c != "pass_fail_label"
    ]
    learning = learning.drop(columns=label_cols, errors="ignore")

    metadata_columns = [
        c for c in MODEL_METADATA_COLUMNS if c in learning.columns
    ]

    semantic_categorical = {
        "global_E_peak_region",
        "E_curve1_peak_region",
        "E_curve2_peak_region",
    }
    categorical = [
        c
        for c in learning.columns
        if c in MODEL_CATEGORICAL_COLUMNS or c in semantic_categorical
    ]

    for col in learning.columns:
        if (
            col in categorical
            or col in metadata_columns
            or col == "pass_fail_code"
        ):
            continue
        learning[col] = pd.to_numeric(learning[col], errors="coerce")

    if categorical:
        learning[categorical] = (
            learning[categorical]
            .fillna("unknown")
            .astype(str)
        )
        learning = pd.get_dummies(
            learning,
            columns=categorical,
            prefix=categorical,
            dtype=int,
        )

    # Remove columns that contain no usable learning information.
    learning = learning.dropna(axis=1, how="all")

    # Keep a stable feature schema even when the current dataset is small.
    # Constant-column removal should be performed later inside the training
    # pipeline, not while the master model-ready file is being generated.
    target_name = "pass_fail_code"

    if target_name in learning.columns:
        target = pd.to_numeric(learning.pop(target_name), errors="coerce")
        learning[target_name] = target

    # Keep grouping/traceability metadata first and the target last. The model
    # training script must exclude simulation_id and pass_fail_label from X and
    # use design_id only as the grouped-validation key.
    ordered_metadata = [
        c for c in MODEL_METADATA_COLUMNS if c in learning.columns
    ]
    ordered_features = [
        c
        for c in learning.columns
        if c not in ordered_metadata and c != target_name
    ]
    ordered_columns = ordered_metadata + ordered_features
    if target_name in learning.columns:
        ordered_columns.append(target_name)

    return learning.reindex(columns=ordered_columns)


def _canonical_lookup(columns: list[str]) -> dict[str, str]:
    """Case-insensitive column lookup without changing research/audit names."""
    return {str(column).strip().lower(): column for column in columns}


def _coalesced_series(
    frame: pd.DataFrame,
    canonical_name: str,
    aliases: list[str] | None = None,
) -> pd.Series:
    """Return the first nonblank value across a canonical field and aliases."""
    lookup = _canonical_lookup(list(frame.columns))
    candidates = [canonical_name, *(aliases or [])]
    present = [lookup[name.lower()] for name in candidates if name.lower() in lookup]
    if not present:
        return pd.Series(pd.NA, index=frame.index, dtype="object")

    result = frame[present[0]].copy()
    for column in present[1:]:
        result = result.combine_first(frame[column])
    return result


def _numeric_series(frame: pd.DataFrame, name: str) -> pd.Series:
    return pd.to_numeric(
        _coalesced_series(frame, name, COMPACT_FEATURE_ALIASES.get(name)),
        errors="coerce",
    )


def _safe_ratio(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    denominator = pd.to_numeric(denominator, errors="coerce")
    numerator = pd.to_numeric(numerator, errors="coerce")
    return numerator.div(denominator.where(denominator.abs() > 1e-12))


def standardize_master_measurements(master: pd.DataFrame) -> pd.DataFrame:
    """
    Coalesce historical geometry/rating names into the canonical mm/kV names.

    Source databases remain auditable, while every generated input vector uses
    explicit millimetre column names. No inch representation is introduced.
    """
    if master.empty:
        return master.copy()

    output = master.copy()
    canonical_fields = {
        "conductor_diameter_mm", "conductor_length_mm", "shield_diameter_mm",
        "shield_length_mm", "shell_mean_diameter_mm", "bushing_inside_diameter_mm",
        "top_bulb_distance_to_nearest_shed_mm", "bottom_bulb_distance_to_nearest_shed_mm",
        "top_shed_outward_delta_y_mm", "bottom_shed_outward_delta_y_mm",
        "top_shed_outward_delta_x_mm", "bottom_shed_outward_delta_x_mm",
        "voltage_rating_kv", "bil_voltage_kv",
    }
    for name in canonical_fields:
        output[name] = _coalesced_series(
            output,
            name,
            COMPACT_FEATURE_ALIASES.get(name),
        )

    # Remove only known duplicate measurement representations from generated
    # vectors. Original source CSVs and archived extractor outputs are untouched.
    duplicate_aliases = {
        alias
        for name in canonical_fields
        for alias in COMPACT_FEATURE_ALIASES.get(name, [])
        if alias != name
    }
    duplicate_aliases.update({
        "top_creepage_distance_in", "bottom_creepage_distance_in",
    })
    output = output.drop(columns=[c for c in duplicate_aliases if c in output], errors="ignore")
    return output


def build_compact_model_ready(master: pd.DataFrame) -> pd.DataFrame:
    """Build the fixed 75-feature initial ML vector plus metadata and target."""
    if master.empty:
        columns = MODEL_METADATA_COLUMNS + COMPACT_MODEL_FEATURES + ["pass_fail_code"]
        return pd.DataFrame(columns=columns)

    semantic = semanticize_zone_features(standardize_master_measurements(master))
    compact = pd.DataFrame(index=semantic.index)

    # Non-predictive grouping/traceability fields are retained for safe training.
    for column in MODEL_METADATA_COLUMNS:
        compact[column] = _coalesced_series(semantic, column)

    # Copy canonical/raw features first. Derived fields are replaced below.
    for feature in COMPACT_MODEL_FEATURES:
        compact[feature] = _coalesced_series(
            semantic,
            feature,
            COMPACT_FEATURE_ALIASES.get(feature),
        )

    sim_type = _coalesced_series(semantic, "simulation_type").astype("string").str.lower()
    existing_static = pd.to_numeric(compact["simulation_type_static"], errors="coerce")
    derived_static = sim_type.str.startswith("stat", na=False).astype(int)
    compact["simulation_type_static"] = derived_static.where(sim_type.notna(), existing_static)

    polarity = _coalesced_series(semantic, "voltage_polarity").astype("string").str.lower()
    existing_negative = pd.to_numeric(
        _coalesced_series(
            semantic,
            "voltage_polarity_negative",
            ["voltage_polarity_-"],
        ),
        errors="coerce",
    )
    derived_negative = polarity.isin({
        "-", "negative", "neg", "-1", "minus",
    }).astype(int)
    compact["voltage_polarity_negative"] = derived_negative.where(
        polarity.notna(),
        existing_negative,
    )

    conductor_diameter = _numeric_series(semantic, "conductor_diameter_mm")
    shield_diameter = _numeric_series(semantic, "shield_diameter_mm")
    shield_length = _numeric_series(semantic, "shield_length_mm")
    conductor_length = _numeric_series(semantic, "conductor_length_mm")
    shell_diameter = _numeric_series(semantic, "shell_mean_diameter_mm")
    top_creepage = _numeric_series(semantic, "top_creepage_distance_mm")
    bottom_creepage = _numeric_series(semantic, "bottom_creepage_distance_mm")

    explicit_shield = _coalesced_series(semantic, "shield_present")
    explicit_shield_text = explicit_shield.astype("string").str.lower()
    compact["shield_present"] = np.where(
        explicit_shield_text.isin({"0", "false", "no", "n"}),
        0,
        np.where(
            explicit_shield_text.isin({"1", "true", "yes", "y"}),
            1,
            shield_diameter.notna().astype(int),
        ),
    )
    compact["total_creepage_distance_mm"] = top_creepage.add(bottom_creepage, fill_value=0)
    compact.loc[top_creepage.isna() & bottom_creepage.isna(), "total_creepage_distance_mm"] = np.nan
    compact["conductor_to_shield_radial_clearance_mm"] = (shield_diameter - conductor_diameter) / 2.0
    compact["shield_to_shell_radial_clearance_mm"] = (shell_diameter - shield_diameter) / 2.0
    compact["shield_length_over_conductor_length"] = _safe_ratio(shield_length, conductor_length)

    # All retained compact features are numeric/binary.
    for feature in COMPACT_MODEL_FEATURES:
        if feature not in COMPACT_CATEGORICAL_FEATURES:
            compact[feature] = pd.to_numeric(compact[feature], errors="coerce")
        else:
            compact[feature] = compact[feature].astype("string")

    compact["pass_fail_code"] = pd.to_numeric(
        _coalesced_series(semantic, "pass_fail_code"),
        errors="coerce",
    )
    ordered = MODEL_METADATA_COLUMNS + COMPACT_MODEL_FEATURES + ["pass_fail_code"]
    return compact.reindex(columns=ordered)


def compact_feature_dictionary() -> pd.DataFrame:
    """Describe roles so the training pipeline cannot confuse IDs or labels."""
    mutable_geometry = {
        "conductor_diameter_mm", "shield_diameter_mm", "shield_length_mm",
        "shell_mean_diameter_mm", "top_creepage_distance_mm",
        "bottom_creepage_distance_mm", "top_bulb_distance_to_nearest_shed_mm",
        "bottom_bulb_distance_to_nearest_shed_mm", "top_shed_outward_delta_y_mm",
        "bottom_shed_outward_delta_y_mm", "top_shed_outward_delta_x_mm",
        "bottom_shed_outward_delta_x_mm",
    }
    rows = []
    for column in MODEL_METADATA_COLUMNS + COMPACT_MODEL_FEATURES + ["pass_fail_code"]:
        if column == "design_id":
            role, family = "group_key", "metadata"
        elif column == "simulation_id":
            role, family = "record_id", "metadata"
        elif column == "pass_fail_label":
            role, family = "human_label_excluded", "target_metadata"
        elif column == "pass_fail_code":
            role, family = "classification_target", "target"
        elif column in COMPACT_DERIVED_FEATURES:
            role, family = "derived_model_feature", "derived"
        else:
            role = "model_feature"
            if "surface_bound_Q" in column:
                family = "q_stress"
            elif "wavelet" in column or "fft" in column:
                family = "dsp"
            elif "_E_" in column or column.startswith("E_curve") or column.startswith("global_E"):
                family = "e_stress"
            elif column.endswith("_kv") or column.startswith("simulation_type") or column.startswith("voltage_polarity"):
                family = "operating_condition"
            else:
                family = "geometry"
        if family in {"e_stress", "q_stress"}:
            surrogate_use = "candidate_surrogate_target"
        elif family == "dsp":
            surrogate_use = "postsimulation_only_not_surrogate_input"
        elif column in COMPACT_MODEL_FEATURES:
            surrogate_use = "surrogate_input"
        else:
            surrogate_use = "excluded"

        if column in {"top_creepage_distance_mm", "bottom_creepage_distance_mm"}:
            collection_source = "pdf.py extraction or 20-700-000 shed calculation"
        elif column in {"voltage_rating_kv", "bil_voltage_kv"}:
            collection_source = "pdf.py drawing extraction"
        elif column in {
            "conductor_diameter_mm", "conductor_length_mm", "shield_diameter_mm",
            "shield_length_mm", "shell_mean_diameter_mm",
            "top_bulb_distance_to_nearest_shed_mm", "bottom_bulb_distance_to_nearest_shed_mm",
            "top_shed_outward_delta_y_mm", "bottom_shed_outward_delta_y_mm",
            "top_shed_outward_delta_x_mm", "bottom_shed_outward_delta_x_mm",
        }:
            collection_source = (
                "electro_geometry.py segment extraction; optional 20-700-000 "
                "shield/shell diameter correction"
            )
        elif family in {"e_stress", "q_stress", "dsp"}:
            collection_source = "tier1.py simulation export analysis"
        elif role == "derived_model_feature":
            collection_source = "automation_application.py derived"
        elif family == "operating_condition":
            collection_source = "simulation metadata"
        else:
            collection_source = "centralized application"

        rows.append({
            "column_name": column,
            "role": role,
            "feature_family": family,
            "included_in_predictor_X": column in COMPACT_MODEL_FEATURES,
            "surrogate_use": surrogate_use,
            "counterfactual_mutable": column in mutable_geometry,
            "collection_source": collection_source,
        })
    return pd.DataFrame(rows)


def rebuild_outputs() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    designs = load_design_db()
    simulations = normalize_pass_fail_dataframe(load_simulation_db())

    # Persist synchronized pass/fail fields for legacy rows as part of rebuild.
    if not simulations.empty:
        simulations.to_csv(SIMULATION_DB_CSV, index=False)

    if simulations.empty or designs.empty:
        master = simulations.copy()
    else:
        design_view = designs.drop(columns=["created_at", "updated_at"], errors="ignore")
        master = simulations.merge(design_view, on="design_id", how="left", validate="many_to_one")
    master = standardize_master_measurements(master)
    master.to_csv(MASTER_CSV, index=False)

    expanded_model_ready = build_expanded_model_ready(master)
    model_ready = build_compact_model_ready(master)
    expanded_model_ready.to_csv(EXPANDED_MODEL_READY_CSV, index=False)
    model_ready.to_csv(MODEL_READY_CSV, index=False)

    try:
        with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl") as writer:
            master.to_excel(writer, index=False, sheet_name="Master_Input_Vector")
            designs.to_excel(writer, index=False, sheet_name="Design_Profiles")
            simulations.to_excel(writer, index=False, sheet_name="Simulation_Records")
        with pd.ExcelWriter(EXPANDED_MODEL_READY_XLSX, engine="openpyxl") as writer:
            expanded_model_ready.to_excel(writer, index=False, sheet_name="Expanded_Model_Ready_Vector")
            pd.DataFrame({
                "column_name": expanded_model_ready.columns,
                "role": [
                    "target"
                    if c == "pass_fail_code"
                    else "group_key"
                    if c == "design_id"
                    else "record_id"
                    if c == "simulation_id"
                    else "human_label_excluded"
                    if c == "pass_fail_label"
                    else "candidate_model_feature"
                    for c in expanded_model_ready.columns
                ],
            }).to_excel(
                writer,
                index=False,
                sheet_name="Expanded_Feature_Dictionary",
            )
        with pd.ExcelWriter(MODEL_READY_XLSX, engine="openpyxl") as writer:
            model_ready.to_excel(writer, index=False, sheet_name="Model_Ready_Vector")
            compact_feature_dictionary().to_excel(
                writer,
                index=False,
                sheet_name="Feature_Dictionary",
            )
    except Exception as exc:
        print(f"Warning: Excel output could not be written: {exc}")
    return master, expanded_model_ready, model_ready


def create_design(state: dict[str, Any]) -> None:
    design_id = prompt_text("Design ID / variant name", allow_blank=False)
    df = load_design_db()
    if not df.empty and design_id in df["design_id"].astype(str).tolist():
        raise RuntimeError(f"Design '{design_id}' already exists.")
    upsert_design({"design_id": design_id, "created_at": now_iso(), "updated_at": now_iso()})
    state["active_design_id"] = design_id
    save_state(state)
    print(f"Created and selected: {design_id}\nFolder: {design_folder(design_id)}")


def select_design(state: dict[str, Any]) -> None:
    df = load_design_db()
    ids = sorted(df["design_id"].dropna().astype(str).unique()) if not df.empty else []
    if not ids:
        print("No designs exist."); return
    for i, design_id in enumerate(ids, 1):
        print(f"  {i}. {design_id}{' *' if design_id == state.get('active_design_id') else ''}")
    value = prompt_text("Select number or exact ID", allow_blank=False)
    selected = ids[int(value)-1] if value.isdigit() and 1 <= int(value) <= len(ids) else value
    if selected not in ids:
        raise RuntimeError(f"Unknown design: {selected}")
    state["active_design_id"] = selected
    save_state(state)
    design_folder(selected)
    print(f"Active design: {selected}")


def read_solidworks_output(path: Path) -> dict[str, Any]:
    row = pd.read_csv(path).iloc[-1]
    return {
        "solidworks_assembly": clean_value(row.get("Assembly")),
        "conductor_component": clean_value(row.get("Conductor Component")),
        "conductor_diameter_mm": clean_value(row.get("Conductor Diameter mm")),
        "shield_diameter_component": clean_value(row.get("Shield Diameter Component")),
        "shield_diameter_mm": clean_value(row.get("Shield Diameter mm")),
        "shield_length_component": clean_value(row.get("Shield Length Component")),
        "shield_length_mm": clean_value(row.get("Shield Length mm")),
    }


def collect_solidworks_geometry(state: dict[str, Any]) -> None:
    design_id, design = get_active_design(state); folder = design_folder(design_id)
    run_module_main("solidworks", folder)
    out = folder / "solidworks_extracted_features.csv"
    if not out.exists(): raise RuntimeError(f"Expected output missing: {out}")
    design.update(read_solidworks_output(out)); design["solidworks_output"] = str(out); design["updated_at"] = now_iso()
    upsert_design(design)


def read_electro_geometry_output(path: Path) -> dict[str, Any]:
    """Read canonical or legacy ELECTRO geometry feature/value CSV output."""
    features: dict[str, Any] = {}
    name_map = {
        # Persisted ELECTRO segment references.
        "origin_segment_id": "origin_segment_id",
        "conductor_segment_id": "conductor_segment_id",
        "shield_segment_id": "shield_segment_id",
        "largest_diameter_shell_segment_id": "largest_diameter_shell_segment_id",
        "top_shed_segment_id": "top_shed_segment_id",
        "bottom_shed_segment_id": "bottom_shed_segment_id",

        # Legacy base geometry names.
        "conductor_diameter": "conductor_diameter_mm",
        "conductor_length": "conductor_length_mm",
        "shield_diameter": "shield_diameter_mm",
        "shield_length": "shield_length_mm",
        "shell_diameter": "shell_mean_diameter_mm",

        # Legacy upper/lower physical relationship names.
        "upper_shield_to_shed_outward_delta_y": "top_shed_outward_delta_y_mm",
        "lower_shield_to_shed_outward_delta_y": "bottom_shed_outward_delta_y_mm",
        "upper_shield_to_shed_distance": "top_bulb_distance_to_nearest_shed_mm",
        "lower_shield_to_shed_distance": "bottom_bulb_distance_to_nearest_shed_mm",
        "upper_shield_to_shed_global_delta_x": "top_shed_global_delta_x_mm",
        "upper_shield_to_shed_global_delta_y": "top_shed_global_delta_y_mm",
        "upper_shield_to_shed_outward_delta_x": "top_shed_outward_delta_x_mm",
        "lower_shield_to_shed_global_delta_x": "bottom_shed_global_delta_x_mm",
        "lower_shield_to_shed_global_delta_y": "bottom_shed_global_delta_y_mm",
        "lower_shield_to_shed_outward_delta_x": "bottom_shed_outward_delta_x_mm",
        "shield_shed_alignment_tolerance": "shield_shed_alignment_tolerance_mm",
    }
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in list(csv.reader(f))[1:]:
            if not row: break
            if len(row) >= 2:
                # ELECTRO geometry is standardized to millimetres throughout
                # the application and downstream ML pipeline.
                source_name = row[0].strip()
                mapped = name_map.get(source_name, source_name)
                features[mapped] = clean_value(row[1])
    return features


def collect_electro_geometry_from_segments(
    state: dict[str, Any],
    apply_20_700_diameter_correction: bool = False,
) -> None:
    design_id, design = get_active_design(state); folder = design_folder(design_id)
    if apply_20_700_diameter_correction:
        print(
            "\nWARNING: The shield/shell diameter correction was developed for "
            f"20-700-000. Active design: '{design_id}'. Verify that the formula "
            "is appropriate before using the result."
        )

    run_module_main(
        "electro_geometry",
        folder,
        main_kwargs={
            "apply_20_700_diameter_correction": apply_20_700_diameter_correction,
        },
    )
    out = folder / "electro_simple_ref_features.csv"
    if not out.exists(): raise RuntimeError(f"Expected output missing: {out}")
    design.update(read_electro_geometry_output(out)); design["electro_geometry_output"] = str(out); design["updated_at"] = now_iso()
    upsert_design(design)


def collect_electro_geometry(state: dict[str, Any]) -> None:
    """Route option 4 to ordinary or 20-700-000-corrected extraction."""
    print("\nELECTRO segment-geometry workflow")
    print("---------------------------------")
    print("1. Standard segment-geometry extraction")
    print("2. 20-700-000 extraction with corrected shield and shell diameters")
    print("0. Cancel")

    while True:
        choice = input("Choose 1, 2, or 0: ").strip()
        if choice == "1":
            collect_electro_geometry_from_segments(state, False)
            return
        if choice == "2":
            collect_electro_geometry_from_segments(state, True)
            return
        if choice == "0":
            print("ELECTRO geometry workflow cancelled.")
            return
        print("Invalid selection. Enter 1, 2, or 0.")


def _row_first_value(row: pd.Series, *names: str) -> Any:
    """Return the first usable value from alternate extractor column names."""
    for name in names:
        value = clean_value(row.get(name))
        if value is not None:
            return value
    return None


def collect_creepage_from_pdf(state: dict[str, Any]) -> None:
    """Run the existing PDF creepage and voltage-rating extraction workflow."""
    design_id, design = get_active_design(state); folder = design_folder(design_id)
    returned_output = run_module_main("creepage", folder)
    output_candidates = [
        folder / "creepage_distance_results.csv",
        folder / "drawing_extraction_results.csv",  # legacy pdf.py output
    ]
    if returned_output:
        returned_path = Path(str(returned_output))
        if not returned_path.is_absolute():
            returned_path = folder / returned_path
        output_candidates.insert(0, returned_path)
    existing_outputs = [path for path in output_candidates if path.exists()]
    if not existing_outputs:
        expected = "\n".join(str(path) for path in output_candidates)
        raise RuntimeError(f"Drawing extractor completed but no compatible output was found:\n{expected}")
    out = max(existing_outputs, key=lambda path: path.stat().st_mtime)
    row = pd.read_csv(out).iloc[-1]
    design.update({
        "drawing_pdf_file": clean_value(row.get("pdf_file")),
        "drawing_pdf_path": clean_value(row.get("pdf_path")),
        "top_creepage_distance_mm": clean_value(row.get("top_creepage_distance_mm")),
        "top_creepage_distance_in": clean_value(row.get("top_creepage_distance_in")),
        "bottom_creepage_distance_mm": clean_value(row.get("bottom_creepage_distance_mm")),
        "bottom_creepage_distance_in": clean_value(row.get("bottom_creepage_distance_in")),
        "voltage_rating_kv": _row_first_value(
            row,
            "voltage_rating_kv", "rated_voltage_kv", "regular_voltage_rating_kv",
        ),
        "bil_voltage_kv": _row_first_value(
            row,
            "bil_voltage_kv", "rated_bil_voltage_kv", "basic_impulse_level_kv",
        ),
        "creepage_output": str(out), "updated_at": now_iso(),
    })
    # Preserve the selected source PDF in the design folder when accessible.
    src = clean_value(row.get("pdf_path"))
    if src and Path(str(src)).exists():
        copied = folder / Path(str(src)).name
        if Path(str(src)).resolve() != copied.resolve(): shutil.copy2(src, copied)
        design["drawing_pdf_archived_path"] = str(copied)
    upsert_design(design)


def calculate_20_700_000_creepage_values(
    external_sheds_removed: int,
    internal_sheds_removed: int,
) -> dict[str, float | int]:
    """Calculate remaining creepage for the 20-700-000 shed geometry."""
    if external_sheds_removed < 0 or internal_sheds_removed < 0:
        raise ValueError("Removed-shed counts must be greater than or equal to zero.")

    external_mm = (
        EXTERNAL_BASE_CREEPAGE_MM
        - external_sheds_removed * EXTERNAL_CREEPAGE_LOSS_PER_SHED_MM
    )
    internal_mm = (
        INTERNAL_BASE_CREEPAGE_MM
        - internal_sheds_removed * INTERNAL_CREEPAGE_LOSS_PER_SHED_MM
    )
    if external_mm < MIN_EXTERNAL_CREEPAGE_MM:
        raise ValueError(
            f"External creepage would be {external_mm:.1f} mm, below the "
            f"{MIN_EXTERNAL_CREEPAGE_MM:.1f} mm minimum. Remove fewer external sheds."
        )
    if internal_mm < MIN_INTERNAL_CREEPAGE_MM:
        raise ValueError(
            f"Internal creepage would be {internal_mm:.1f} mm, below the "
            f"{MIN_INTERNAL_CREEPAGE_MM:.1f} mm minimum. Remove fewer internal sheds."
        )

    return {
        "external_sheds_removed": int(external_sheds_removed),
        "internal_sheds_removed": int(internal_sheds_removed),
        "external_creepage_distance_mm": float(external_mm),
        "internal_creepage_distance_mm": float(internal_mm),
        # The existing model-ready schema uses top/bottom. For this design,
        # external maps to the 1736 mm/top path and internal to 1397 mm/bottom.
        "top_creepage_distance_mm": float(external_mm),
        "bottom_creepage_distance_mm": float(internal_mm),
        "top_creepage_distance_in": float(external_mm / MM_PER_INCH),
        "bottom_creepage_distance_in": float(internal_mm / MM_PER_INCH),
    }


def calculate_20_700_000_creepage(state: dict[str, Any]) -> None:
    """Apply the 20-700-000 shed formulas to the active design after warning."""
    design_id, design = get_active_design(state)
    print("\n20-700-000 shed-removal creepage calculator")
    print("------------------------------------------")
    print(
        "WARNING: These shed-removal creepage formulas were developed for "
        f"20-700-000. Active design: '{design_id}'. Verify that the formulas "
        "are appropriate before using the result."
    )
    print(
        "External creepage = 1736 mm - external sheds removed x 67 mm\n"
        "Internal creepage = 1397 mm - internal sheds removed x 37 mm\n"
        "Minimum remaining external creepage: 1473.2 mm\n"
        "Minimum remaining internal creepage: 1168.4 mm"
    )

    external_removed = prompt_nonnegative_int(
        f"Number of external sheds removed (0-{MAX_EXTERNAL_SHEDS_REMOVED})",
        0,
        MAX_EXTERNAL_SHEDS_REMOVED,
    )
    internal_removed = prompt_nonnegative_int(
        f"Number of internal sheds removed (0-{MAX_INTERNAL_SHEDS_REMOVED})",
        0,
        MAX_INTERNAL_SHEDS_REMOVED,
    )
    calculated = calculate_20_700_000_creepage_values(
        external_removed,
        internal_removed,
    )

    print("\nCalculated remaining creepage")
    print(
        f"  External/top: {calculated['external_creepage_distance_mm']:.3f} mm "
        f"({calculated['top_creepage_distance_in']:.4f} in)"
    )
    print(
        f"  Internal/bottom: {calculated['internal_creepage_distance_mm']:.3f} mm "
        f"({calculated['bottom_creepage_distance_in']:.4f} in)"
    )
    print(
        f"  Total: "
        f"{calculated['external_creepage_distance_mm'] + calculated['internal_creepage_distance_mm']:.3f} mm"
    )

    if not confirm_yes_no("Save these calculated creepage values", default=True):
        print("Calculated values were not saved.")
        return

    design.update(calculated)
    design.update({
        "creepage_calculation_method": "20-700-000 shed-removal calculation",
        "creepage_calculation_design_basis": SHED_CALCULATION_DESIGN,
        "creepage_output": "calculated from removed shed counts",
        "updated_at": now_iso(),
    })
    upsert_design(design)
    print("Calculated creepage values saved to the active design profile.")


def collect_creepage(state: dict[str, Any]) -> None:
    """Route option 5 to PDF extraction or the design-specific calculator."""
    print("\nCreepage and voltage-rating workflow")
    print("------------------------------------")
    print("1. PDF extraction: creepage distances and voltage ratings")
    print("2. Creepage calculation from removed sheds (20-700-000 formula basis)")
    print("0. Cancel")

    while True:
        choice = input("Choose 1, 2, or 0: ").strip()
        if choice == "1":
            collect_creepage_from_pdf(state)
            return
        if choice == "2":
            calculate_20_700_000_creepage(state)
            return
        if choice == "0":
            print("Creepage workflow cancelled.")
            return
        print("Invalid selection. Enter 1, 2, or 0.")


def edit_manual_design_features(state: dict[str, Any]) -> None:
    design_id, design = get_active_design(state)
    legacy_defaults = {
        "shell_mean_diameter_mm": design.get("outer_shell_mean_diameter_mm"),
        "voltage_rating_kv": design.get("rated_voltage_kv"),
    }
    changed = False

    while True:
        print(f"\nManually editable design inputs for '{design_id}'")
        print("Simulation-derived E/Q/DSP data and simulation metadata are excluded.")
        print("Select a variable by number; derived clearances and ratios rebuild automatically.")
        print("-" * 78)
        for index, (key, label) in enumerate(MANUALLY_EDITABLE_DESIGN_FEATURES, 1):
            current = clean_value(design.get(key))
            if current is None:
                current = clean_value(legacy_defaults.get(key))
            shown = "N/A" if current is None else current
            print(f"{index:>2}. {label} [{key}] = {shown}")
        print(" 0. Finish and save changes")

        choice = input(
            f"Choose 0-{len(MANUALLY_EDITABLE_DESIGN_FEATURES)}: "
        ).strip()
        if choice == "0":
            if changed:
                design["updated_at"] = now_iso()
                upsert_design(design)
                print("Manual design-input changes saved; model-ready vectors rebuilt.")
            else:
                print("No manual design-input changes were made.")
            return

        try:
            selected_index = int(choice)
        except ValueError:
            print("Invalid selection. Enter one of the listed numbers.")
            continue
        if not 1 <= selected_index <= len(MANUALLY_EDITABLE_DESIGN_FEATURES):
            print("Invalid selection. Enter one of the listed numbers.")
            continue

        key, label = MANUALLY_EDITABLE_DESIGN_FEATURES[selected_index - 1]
        current = clean_value(design.get(key))
        if current is None:
            current = clean_value(legacy_defaults.get(key))
        design[key] = prompt_optional_float(label, current)
        changed = True

        if key == "top_creepage_distance_mm":
            value = clean_value(design[key])
            design["top_creepage_distance_in"] = (
                None if value is None else float(value) / MM_PER_INCH
            )
            design["creepage_calculation_method"] = "manual override"
        elif key == "bottom_creepage_distance_mm":
            value = clean_value(design[key])
            design["bottom_creepage_distance_in"] = (
                None if value is None else float(value) / MM_PER_INCH
            )
            design["creepage_calculation_method"] = "manual override"
        elif key in {"shield_diameter_mm", "shell_mean_diameter_mm"}:
            design["diameter_calculation_method"] = "manual override"


def collect_simulation_metadata() -> dict[str, Any]:
    sim_type = prompt_text("Simulation type (static/transient)", "static").lower()
    data = {
        "simulation_id": prompt_text("Simulation ID", f"sim_{datetime.now():%Y%m%d_%H%M%S}"),
        "simulation_type": sim_type,
        "simulation_voltage_kv": prompt_optional_float("Applied voltage (kV)"),
        "voltage_polarity": prompt_text("Voltage polarity (+/-/N/A)", "N/A"),
        "electro_model_name": prompt_text("ELECTRO model name", ""),
        "notes": prompt_text("Simulation notes", ""),
    }
    if sim_type.startswith("trans"):
        data.update({
            "transient_waveform_name": prompt_text("Waveform/source name", "impulse"),
            "bil_front_time_us": prompt_optional_float("Impulse front time (us)", 1.2),
            "bil_time_to_half_us": prompt_optional_float("Impulse time to half (us)", 50.0),
        })
    return data


def process_export_and_append(state: dict[str, Any]) -> None:
    design_id, _ = get_active_design(state)
    folder = design_folder(design_id)

    # This must be the only ELECTRO-export file-selection dialog.
    selected = choose_file(
        "Select ELECTRO graph export",
        [
            ("ELECTRO exports", "*.csv *.txt"),
            ("CSV files", "*.csv"),
            ("Text files", "*.txt"),
            ("All files", "*.*"),
        ],
    )

    if selected is None:
        print("No ELECTRO export selected.")
        return

    selected = selected.resolve()

    # Preserve the original export inside the active design folder.
    archived = folder / selected.name

    if selected != archived.resolve():
        shutil.copy2(selected, archived)

    archived = archived.resolve()

    print(f"\nSelected ELECTRO export:\n{selected}")
    print(f"Working copy stored at:\n{archived}")

    metadata = collect_simulation_metadata()

    # Pass the already-selected file directly into tier1.main().
    # This prevents tier1.py from opening a second file-selection dialog.
    tier1_features = run_module_main(
        "tier1",
        folder,
        main_kwargs={
            "export_file": archived,
            "output_dir": folder,
        },
    )

    tier1_output = folder / "tier1_input_vector.csv"

    # Prefer the feature dictionary returned directly by tier1.main().
    if isinstance(tier1_features, dict):
        features = {
            key: clean_value(value)
            for key, value in tier1_features.items()
        }
    else:
        # Compatibility fallback for an older tier1.py that only writes CSV.
        if not tier1_output.exists():
            raise RuntimeError(
                "Tier 1 analysis finished, but it returned no features and "
                f"the expected output was not created:\n{tier1_output}"
            )

        tier1_df = pd.read_csv(tier1_output)
        if tier1_df.empty:
            raise RuntimeError(
                f"Tier 1 output contains no feature rows:\n{tier1_output}"
            )

        features = {
            key: clean_value(value)
            for key, value in tier1_df.iloc[-1].to_dict().items()
        }

    row = {
        "simulation_record_id": str(uuid.uuid4()),
        "collected_at": now_iso(),
        "design_id": design_id,
        **metadata,
        **features,
        "electro_export_path": str(archived),
    }

    append_simulation_record(row)

    print("\nSimulation record appended successfully.")
    print(f"Design artifacts:\n{folder}")



def confirm_yes_no(prompt: str, default: bool = False) -> bool:
    """Return True only when the user explicitly confirms the action."""
    suffix = "Y/n" if default else "y/N"
    value = input(f"{prompt} [{suffix}]: " ).strip().lower()
    if not value:
        return default
    return value in {"y", "yes"}


def backup_database_file(path: Path) -> Path | None:
    """Create a timestamped backup before a destructive database edit."""
    if not path.exists():
        return None
    backup_dir = DATA_DIR / "Database_Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backup_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, backup)
    return backup


def simulation_display_value(value: Any) -> str:
    cleaned = clean_value(value)
    return "N/A" if cleaned is None else str(cleaned)


def print_simulation_table(df: pd.DataFrame) -> None:
    if df.empty:
        print("No simulation records are available.")
        return

    print("\nAvailable simulation records")
    print("-" * 118)
    print(
        f"{'#':<4}{'Design':<26}{'Simulation ID':<25}"
        f"{'Type':<13}{'Voltage (kV)':<15}{'Polarity':<10}{'Label':<10}{'Collected'}"
    )
    print("-" * 118)

    display = df.reset_index(drop=True)
    for index, row in display.iterrows():
        print(
            f"{index + 1:<4}"
            f"{simulation_display_value(row.get('design_id'))[:24]:<26}"
            f"{simulation_display_value(row.get('simulation_id'))[:23]:<25}"
            f"{simulation_display_value(row.get('simulation_type'))[:11]:<13}"
            f"{simulation_display_value(row.get('simulation_voltage_kv'))[:13]:<15}"
            f"{simulation_display_value(row.get('voltage_polarity'))[:8]:<10}"
            f"{simulation_display_value(row.get('pass_fail_label'))[:8]:<10}"
            f"{simulation_display_value(row.get('collected_at'))}"
        )


def choose_simulation_record(df: pd.DataFrame) -> pd.Series | None:
    if df.empty:
        print("No simulation records are available.")
        return None

    print_simulation_table(df)
    value = input("Select a record number, or press Enter to cancel: " ).strip()
    if not value:
        return None
    if not value.isdigit() or not 1 <= int(value) <= len(df):
        print("Invalid record number.")
        return None
    return df.reset_index(drop=True).iloc[int(value) - 1]


def remove_simulation_record(record_id: str) -> None:
    simulations = load_simulation_db()
    if simulations.empty or "simulation_record_id" not in simulations.columns:
        raise RuntimeError("No simulation database records are available.")

    mask = simulations["simulation_record_id"].astype(str) == str(record_id)
    if not mask.any():
        raise RuntimeError("The selected simulation record no longer exists.")

    backup = backup_database_file(SIMULATION_DB_CSV)
    simulations.loc[~mask].to_csv(SIMULATION_DB_CSV, index=False)
    rebuild_outputs()
    print("Simulation record removed and vectors rebuilt.")
    if backup is not None:
        print(f"Backup created: {backup}")


def delete_simulation_interactive() -> None:
    simulations = load_simulation_db()
    selected = choose_simulation_record(simulations)
    if selected is None:
        return

    record_id = simulation_display_value(selected.get("simulation_record_id"))
    print("\nSelected simulation")
    print(f"  Design: {simulation_display_value(selected.get('design_id'))}")
    print(f"  Simulation ID: {simulation_display_value(selected.get('simulation_id'))}")
    print(f"  Type: {simulation_display_value(selected.get('simulation_type'))}")
    print(f"  Voltage: {simulation_display_value(selected.get('simulation_voltage_kv'))} kV")
    print(f"  Label: {simulation_display_value(selected.get('pass_fail_label'))}")

    if not confirm_yes_no("Permanently remove this simulation from the source database"):
        print("Deletion cancelled.")
        return
    remove_simulation_record(record_id)


def undo_last_simulation_append() -> None:
    simulations = load_simulation_db()
    if simulations.empty:
        print("No simulation records are available.")
        return

    if "collected_at" in simulations.columns:
        parsed = pd.to_datetime(simulations["collected_at"], errors="coerce", utc=True)
        if parsed.notna().any():
            selected_index = parsed.idxmax()
        else:
            selected_index = simulations.index[-1]
    else:
        selected_index = simulations.index[-1]

    selected = simulations.loc[selected_index]
    print("\nMost recently appended simulation")
    print(f"  Design: {simulation_display_value(selected.get('design_id'))}")
    print(f"  Simulation ID: {simulation_display_value(selected.get('simulation_id'))}")
    print(f"  Collected: {simulation_display_value(selected.get('collected_at'))}")

    if not confirm_yes_no("Undo this append"):
        print("Undo cancelled.")
        return
    remove_simulation_record(str(selected.get("simulation_record_id")))


def delete_design_interactive(state: dict[str, Any]) -> None:
    designs = load_design_db()
    if designs.empty:
        print("No designs are available.")
        return

    ids = sorted(designs["design_id"].dropna().astype(str).unique())
    print("\nDesigns")
    for index, design_id in enumerate(ids, 1):
        sim_count = 0
        simulations = load_simulation_db()
        if not simulations.empty and "design_id" in simulations.columns:
            sim_count = int((simulations["design_id"].astype(str) == design_id).sum())
        marker = " *active" if design_id == state.get("active_design_id") else ""
        print(f"  {index}. {design_id} ({sim_count} simulations){marker}")

    value = input("Select a design number, or press Enter to cancel: " ).strip()
    if not value:
        return
    if not value.isdigit() or not 1 <= int(value) <= len(ids):
        print("Invalid design number.")
        return
    design_id = ids[int(value) - 1]

    simulations = load_simulation_db()
    sim_count = 0
    if not simulations.empty and "design_id" in simulations.columns:
        sim_count = int((simulations["design_id"].astype(str) == design_id).sum())

    print("\nWARNING")
    print(f"Deleting '{design_id}' will remove 1 design and {sim_count} simulation record(s).")
    print("The generated audit, expanded, and compact model-ready vectors will then be rebuilt.")
    if not confirm_yes_no("Continue with design deletion"):
        print("Deletion cancelled.")
        return

    design_backup = backup_database_file(DESIGN_DB_CSV)
    sim_backup = backup_database_file(SIMULATION_DB_CSV)

    designs = designs[designs["design_id"].astype(str) != design_id]
    save_design_db(designs)

    if not simulations.empty and "design_id" in simulations.columns:
        simulations = simulations[simulations["design_id"].astype(str) != design_id]
        simulations.to_csv(SIMULATION_DB_CSV, index=False)

    if state.get("active_design_id") == design_id:
        state["active_design_id"] = None
        save_state(state)

    folder = DATA_DIR / sanitize_name(design_id)
    if folder.exists() and confirm_yes_no("Also delete the design artifact folder", default=False):
        shutil.rmtree(folder)
        print(f"Deleted artifact folder: {folder}")
    elif folder.exists():
        print(f"Artifact folder retained: {folder}")

    rebuild_outputs()
    print("Design, related simulations, and generated vectors were updated.")
    for backup in (design_backup, sim_backup):
        if backup is not None:
            print(f"Backup created: {backup}")


def search_simulations() -> None:
    simulations = load_simulation_db()
    if simulations.empty:
        print("No simulation records are available.")
        return
    query = input("Search design ID, simulation ID, type, voltage, notes, or label: " ).strip().lower()
    if not query:
        return
    searchable = simulations.fillna("").astype(str)
    mask = searchable.apply(lambda column: column.str.lower().str.contains(query, regex=False)).any(axis=1)
    results = simulations.loc[mask].reset_index(drop=True)
    if results.empty:
        print("No matching simulation records found.")
        return
    print_simulation_table(results)


def manage_dataset(state: dict[str, Any]) -> None:
    while True:
        print("\n" + "=" * 72)
        print("DATASET MANAGEMENT")
        print("=" * 72)
        print("1. Delete one simulation record")
        print("2. Undo most recent simulation append")
        print("3. Delete a design/variant and its simulations")
        print("4. Search simulation records")
        print("5. Rebuild audit, expanded, and compact model-ready vectors")
        print("0. Return to main menu")
        choice = input("Choose an option: " ).strip()

        if choice == "1":
            delete_simulation_interactive()
        elif choice == "2":
            undo_last_simulation_append()
        elif choice == "3":
            delete_design_interactive(state)
        elif choice == "4":
            search_simulations()
        elif choice == "5":
            rebuild_outputs()
            print("Outputs rebuilt.")
        elif choice == "0":
            return
        else:
            print("Invalid selection.")

def show_active_design(state: dict[str, Any]) -> None:
    design_id, design = get_active_design(state)

    print(
        f"\nActive design: {design_id}\n"
        f"Folder: {design_folder(design_id)}\n"
        + "-" * 72
    )

    segment_fields = [
        ("origin_segment_id", "Origin"),
        ("conductor_segment_id", "Conductor"),
        ("shield_segment_id", "Shield"),
        ("largest_diameter_shell_segment_id", "Largest-diameter shell"),
        ("top_shed_segment_id", "Top shed"),
        ("bottom_shed_segment_id", "Bottom shed"),
    ]

    print("\nSaved ELECTRO segment IDs")
    print("--------------------------")
    segment_values_present = False

    for key, label in segment_fields:
        value = clean_value(design.get(key))
        if value is not None:
            segment_values_present = True
            # CSV loading may represent integer IDs as floats, such as 422.0.
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            print(f"{label}: {value}")

    if not segment_values_present:
        print(
            "No segment IDs have been saved for this design yet. "
            "Run option 4 to collect ELECTRO segment geometry."
        )

    print("\nAll active-design fields")
    print("------------------------")
    segment_keys = {key for key, _ in segment_fields}
    for key, value in sorted(design.items()):
        if key not in segment_keys:
            print(f"{key}: {value}")


def show_summary() -> None:
    designs = load_design_db(); sims = load_simulation_db(); master, expanded, model_ready = rebuild_outputs()
    print(f"\nDesigns: {len(designs)} | Simulations: {len(sims)}")
    print(f"Audit master: {len(master)} rows x {len(master.columns)} columns")
    print(f"Expanded model-ready: {len(expanded)} rows x {len(expanded.columns)} columns")
    print(
        f"Compact model-ready: {len(model_ready)} rows x {len(model_ready.columns)} columns "
        f"({len(COMPACT_MODEL_FEATURES)} predictors)"
    )
    print(
        f"\n{DESIGN_DB_CSV}\n{SIMULATION_DB_CSV}\n{MASTER_CSV}"
        f"\n{EXPANDED_MODEL_READY_CSV}\n{MODEL_READY_CSV}"
    )
    if not model_ready.empty:
        unavailable = [
            feature
            for feature in COMPACT_MODEL_FEATURES
            if feature in model_ready and model_ready[feature].isna().all()
        ]
        if unavailable:
            source_map = compact_feature_dictionary().set_index("column_name")["collection_source"]
            print("\nCompact features not yet populated in any record:")
            for feature in unavailable:
                print(f"  - {feature}  [{source_map.get(feature, 'source unavailable')}]")
        else:
            print("\nEvery compact predictor is populated in at least one record.")


def _discover_electro_com_progids() -> list[str]:
    """Find the COM identifier already used by the local ELECTRO scripts."""
    candidates: list[str] = []

    configured = os.environ.get("ELECTRO_COM_PROGID", "").strip()
    if configured:
        candidates.append(configured)

    call_pattern = re.compile(
        r"(?:Dispatch|EnsureDispatch|GetActiveObject)\s*\(\s*[rubfRUBF]*"
        r"['\"]([^'\"]+)['\"]"
    )
    for key in ("simulation", "electro_geometry"):
        try:
            script = resolve_script(key)
            source = script.read_text(encoding="utf-8-sig", errors="ignore")
        except (FileNotFoundError, OSError):
            continue
        for progid in call_pattern.findall(source):
            # Ignore unrelated automation servers if a companion script uses
            # more than one COM application.
            if "solidworks" not in progid.lower() and progid not in candidates:
                candidates.append(progid)

    if DEFAULT_ELECTRO_COM_PROGID not in candidates:
        candidates.append(DEFAULT_ELECTRO_COM_PROGID)
    return candidates


def _attach_to_electro() -> tuple[Any, str]:
    """Attach to ELECTRO using the same COM identifier as companion scripts."""
    import win32com.client

    errors: list[str] = []
    progids = _discover_electro_com_progids()

    # First try the Running Object Table so this option does not accidentally
    # start a second application while an existing instance is available.
    for progid in progids:
        try:
            return win32com.client.GetActiveObject(progid), progid
        except Exception as exc:
            errors.append(f"GetActiveObject({progid!r}): {exc}")

    # Some ELECTRO COM servers are single-instance but are not registered in
    # the Running Object Table. Dispatch then returns their active instance.
    for progid in progids:
        try:
            return win32com.client.Dispatch(progid), progid
        except Exception as exc:
            errors.append(f"Dispatch({progid!r}): {exc}")

    details = "\n".join(f"  {line}" for line in errors)
    raise RuntimeError(
        "Could not connect to ELECTRO. Open ELECTRO and its model, then retry "
        "option 13. If your installation uses a different COM identifier, set "
        "ELECTRO_COM_PROGID before launching this application.\n" + details
    )


def stabilize_electro_after_simulation() -> None:
    """Optionally save ELECTRO, clear temporary displays, and release COM state."""
    print("\nELECTRO post-simulation stability cleanup")
    print("------------------------------------------")
    print("Run this only after the required E/Q graphs and result files are exported.")
    print("This option can save the model before removing temporary plots/streamlines.")
    print("It does NOT delete the calculated solution, mesh, geometry, or assignments.")

    if not confirm_yes_no("Have all required simulation results been exported?"):
        print("Cleanup cancelled. Export/process the required results first.")
        return

    save_before_cleanup = confirm_yes_no(
        "Save the active ELECTRO model before cleanup?",
        default=True,
    )

    pythoncom = None
    electro = None
    com_initialized = False
    completed: list[str] = []
    warnings: list[str] = []

    try:
        import pythoncom as _pythoncom

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        com_initialized = True

        electro, progid = _attach_to_electro()
        print(f"Connected to ELECTRO ({progid}).")

        try:
            model_path = electro.File_GetModelPath()
        except Exception as exc:
            model_path = None
            warnings.append(f"Could not read the active model path: {exc}")

        if save_before_cleanup:
            # When saving is requested, do not modify analysis displays unless
            # the current model is first protected successfully.
            try:
                electro.File_Save()
                completed.append("active model saved")
            except Exception as exc:
                raise RuntimeError(
                    "ELECTRO could not save the active model, so cleanup stopped "
                    f"before modifying any analysis displays: {exc}"
                ) from exc

            if model_path:
                print(f"Saved model: {model_path}")
        else:
            completed.append("model save skipped by user")
            print("WARNING: Continuing cleanup without saving the active model.")

        # These calls target analysis-view objects only. ELECTRO versions can
        # expose different signatures, so each cleanup is isolated and a
        # failure does not prevent the remaining safe steps.
        try:
            electro.Analysis_DeleteStreamlines_All()
            completed.append("temporary streamlines cleared")
        except Exception as exc:
            warnings.append(f"Streamline cleanup was unavailable: {exc}")

        try:
            electro.Analysis_DeletePlot()
            completed.append("active temporary plot cleared")
        except Exception as exc:
            warnings.append(
                "Plot cleanup was unavailable or required a plot identifier: "
                f"{exc}"
            )

        try:
            electro.Window_Refresh()
            completed.append("ELECTRO window refreshed")
        except Exception as exc:
            warnings.append(f"Window refresh was unavailable: {exc}")

    finally:
        # Dropping the automation application's references is deliberately
        # safer than calling the raw COM Release method directly.
        electro = None
        gc.collect()
        if pythoncom is not None:
            try:
                pythoncom.CoFreeUnusedLibraries()
                completed.append("unused Python COM libraries released")
            except Exception as exc:
                warnings.append(f"Python COM library cleanup was unavailable: {exc}")
            if com_initialized:
                pythoncom.CoUninitialize()

    print("\nOption 13 completed.")
    for item in completed:
        print(f"  - {item}")
    if warnings:
        print("\nNon-fatal warnings:")
        for warning in warnings:
            print(f"  - {warning}")
    print(
        "If ELECTRO continues becoming unstable, close and reopen it after "
        "this cleanup to fully recycle ELECTRO's internal process memory."
    )


def main() -> None:
    ensure_runtime_dependencies()
    print(f"Running with Python: {sys.executable}")
    state = load_state()
    while True:
        print("\n" + "="*78)
        print("ELECTRO INTERNAL DATA COLLECTION")
        print(f"Active design: {state.get('active_design_id') or 'NONE'}")
        print("="*78)
        print("1. Create new design profile")
        print("2. Select existing design profile")
        print("3. Collect SolidWorks geometry")
        print("4. Collect ELECTRO segment geometry")
        print("5. Extract PDF data or calculate 20-700-000 creepage")
        print("6. Manually review/edit design-level input variables")
        print("7. Launch ELECTRO setup / solver / extraction")
        print("8. Process ELECTRO export and append simulation")
        print("9. Rebuild audit, expanded, and compact model-ready vectors")
        print("10. Review active design")
        print("11. Show dataset summary")
        print("12. Manage/delete dataset records")
        print("13. Stabilize ELECTRO after simulation")
        print("0. Exit")
        choice = input("Choose an option: ").strip()
        try:
            if choice == "1": create_design(state)
            elif choice == "2": select_design(state)
            elif choice == "3": collect_solidworks_geometry(state)
            elif choice == "4": collect_electro_geometry(state)
            elif choice == "5": collect_creepage(state)
            elif choice == "6": edit_manual_design_features(state)
            elif choice == "7":
                design_id, _ = get_active_design(state)
                run_module_main(
                    "simulation",
                    design_folder(design_id),
                    main_kwargs={
                        "show_bil_guidance": True,
                        "use_legacy_static_workflow": False,
                    },
                )
            elif choice == "8": process_export_and_append(state)
            elif choice == "9": rebuild_outputs(); print("Outputs rebuilt.")
            elif choice == "10": show_active_design(state)
            elif choice == "11": show_summary()
            elif choice == "12": manage_dataset(state)
            elif choice == "13": stabilize_electro_after_simulation()
            elif choice == "0": break
            else: print("Invalid selection.")
        except KeyboardInterrupt:
            print("\nOperation cancelled.")
        except Exception as exc:
            print(f"\nERROR: {exc}")
            if os.environ.get("ELECTRO_APP_DEBUG") == "1": traceback.print_exc()
        state = load_state()


if __name__ == "__main__":
    main()
