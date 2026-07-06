"""
easy_feature_extraction.py

Purpose:
    Automates extraction of the easiest bushing-design variables/features from:
    1) ELECTRO API-accessible model metadata/material/source information, when available
    2) exported ELECTRO plot/data CSV files, especially max values over distance ranges

This file is intentionally separate from the main API_Automation pipeline.
Import it into the main pipeline after the ELECTRO project is opened and/or after CSV exports are created.

Expected output:
    One row of extracted features that can be appended to your ML input dataset.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Optional, Sequence
import re

import numpy as np
import pandas as pd


# -------------------------------------------------------------------------
# 1. Feature row definition
# -------------------------------------------------------------------------

@dataclass
class EasyBushingFeatures:
    bushing_id: str = ""
    conductor_diameter: Optional[float] = None
    conductor_material: str = ""
    shield_length: Optional[float] = None
    shield_material: str = ""
    shell_material: str = ""
    bil_voltage: Optional[float] = None

    # Optional output-derived values from exported line plots
    conductor_max_e: Optional[float] = None
    shield_max_e: Optional[float] = None
    shed_max_e: Optional[float] = None
    global_max_e: Optional[float] = None

    conductor_max_charge: Optional[float] = None
    shield_max_charge: Optional[float] = None
    shed_max_charge: Optional[float] = None
    global_max_charge: Optional[float] = None


# -------------------------------------------------------------------------
# 2. General helpers
# -------------------------------------------------------------------------

def infer_bushing_id(project_path: str | Path) -> str:
    """
    Infers bushing ID from the ELECTRO project file/folder name.

    Example:
        T:/.../Bushing_700.dsb -> "700"
    """
    name = Path(project_path).stem
    match = re.search(r"(\d{3,})", name)
    return match.group(1) if match else name


def safe_get(obj: Any, attr_names: Sequence[str], default=None):
    """
    Attempts several possible API attribute/method names.

    This is useful because exact ELECTRO API method names may differ.
    Replace the names below with confirmed ELECTRO API calls as you verify them.
    """
    for attr in attr_names:
        if hasattr(obj, attr):
            value = getattr(obj, attr)
            return value() if callable(value) else value
    return default


def write_feature_row_csv(features: EasyBushingFeatures, output_csv: str | Path) -> None:
    """
    Appends one feature row to a CSV file. Creates the file if it does not exist.
    """
    output_csv = Path(output_csv)
    output_csv.parent.mkdir(parents=True, exist_ok=True)

    row = pd.DataFrame([asdict(features)])

    if output_csv.exists():
        row.to_csv(output_csv, mode="a", index=False, header=False)
    else:
        row.to_csv(output_csv, index=False)


# -------------------------------------------------------------------------
# 3. ELECTRO/CAD API feature extraction placeholders
# -------------------------------------------------------------------------

def extract_easy_api_features(
    electro_project: Any,
    project_path: str | Path,
    conductor_region_name: str = "conductor",
    shield_region_name: str = "shield",
    shell_region_name: str = "shell",
    voltage_source_name: str = "BIL",
) -> EasyBushingFeatures:
    """
    Extracts the easiest variables from ELECTRO API-accessible model information.

    IMPORTANT:
        This function contains adapter-style placeholders because ELECTRO's exact
        Python API method names must be confirmed in your installation.

    You should replace the internals of:
        get_region(...)
        get_material(...)
        get_bounding_box(...)
        get_voltage_source_peak(...)
    with the actual ELECTRO API calls.
    """

    features = EasyBushingFeatures()
    features.bushing_id = infer_bushing_id(project_path)

    # -----------------------------
    # Replace these with real API calls
    # -----------------------------
    def get_region(region_name: str) -> Any:
        """
        Expected behavior:
            Return an ELECTRO region/body/object by name.
        """
        # Example candidates only:
        if hasattr(electro_project, "get_region"):
            return electro_project.get_region(region_name)
        if hasattr(electro_project, "GetRegion"):
            return electro_project.GetRegion(region_name)
        if hasattr(electro_project, "regions"):
            return electro_project.regions[region_name]
        return None

    def get_material(region: Any) -> str:
        """
        Expected behavior:
            Return material name assigned to a region.
        """
        if region is None:
            return ""
        return str(safe_get(region, ["material", "Material", "get_material", "GetMaterial"], ""))

    def get_bounding_box(region: Any):
        """
        Expected behavior:
            Return bounding box as:
                xmin, xmax, ymin, ymax, zmin, zmax

        If ELECTRO returns another format, convert it here.
        """
        if region is None:
            return None

        bbox = safe_get(region, ["bounding_box", "BoundingBox", "get_bounding_box", "GetBoundingBox"], None)
        if bbox is None:
            return None

        # Accept dictionary format
        if isinstance(bbox, dict):
            return (
                bbox.get("xmin"), bbox.get("xmax"),
                bbox.get("ymin"), bbox.get("ymax"),
                bbox.get("zmin"), bbox.get("zmax"),
            )

        # Accept list/tuple format
        if len(bbox) == 6:
            return tuple(bbox)

        return None

    def get_voltage_source_peak(source_name: str) -> Optional[float]:
        """
        Expected behavior:
            Return peak BIL impulse voltage from source definition.
        """
        source = None

        if hasattr(electro_project, "get_voltage_source"):
            source = electro_project.get_voltage_source(source_name)
        elif hasattr(electro_project, "GetVoltageSource"):
            source = electro_project.GetVoltageSource(source_name)
        elif hasattr(electro_project, "sources"):
            source = electro_project.sources.get(source_name)

        if source is None:
            return None

        value = safe_get(
            source,
            ["peak_value", "PeakValue", "amplitude", "Amplitude", "voltage", "Voltage", "get_peak_value", "GetPeakValue"],
            None,
        )

        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    # -----------------------------
    # Actual easy extraction logic
    # -----------------------------
    conductor = get_region(conductor_region_name)
    shield = get_region(shield_region_name)
    shell = get_region(shell_region_name)

    features.conductor_material = get_material(conductor)
    features.shield_material = get_material(shield)
    features.shell_material = get_material(shell)
    features.bil_voltage = get_voltage_source_peak(voltage_source_name)

    # Diameter/length via bounding box.
    # You may need to change the axis assumptions depending on how your model is oriented.
    conductor_bbox = get_bounding_box(conductor)
    if conductor_bbox is not None:
        xmin, xmax, ymin, ymax, zmin, zmax = conductor_bbox
        dx = abs(xmax - xmin)
        dy = abs(ymax - ymin)
        dz = abs(zmax - zmin)

        # For a long cylindrical conductor, the diameter is usually the smaller transverse size.
        features.conductor_diameter = min(dx, dy, dz)

    shield_bbox = get_bounding_box(shield)
    if shield_bbox is not None:
        xmin, xmax, ymin, ymax, zmin, zmax = shield_bbox
        dx = abs(xmax - xmin)
        dy = abs(ymax - ymin)
        dz = abs(zmax - zmin)

        # Shield length is usually the dominant dimension.
        features.shield_length = max(dx, dy, dz)

    return features


# -------------------------------------------------------------------------
# 4. Output CSV extraction: max E/charge over known distance ranges
# -------------------------------------------------------------------------

def read_electro_xy_csv(csv_path: str | Path) -> pd.DataFrame:
    """
    Reads an exported ELECTRO line-plot CSV.

    Expected data style:
        distance, value

    The function is intentionally tolerant of column names.
    It returns a DataFrame with standardized columns:
        distance
        value
    """
    csv_path = Path(csv_path)
    df = pd.read_csv(csv_path)

    # Try to identify distance column
    distance_candidates = [
        c for c in df.columns
        if "dist" in c.lower() or "x" == c.lower().strip() or "position" in c.lower()
    ]

    # Try to identify value column
    value_candidates = [
        c for c in df.columns
        if c not in distance_candidates and pd.api.types.is_numeric_dtype(df[c])
    ]

    if not distance_candidates:
        # Fall back to first numeric column
        numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        if len(numeric_cols) < 2:
            raise ValueError(f"Could not identify two numeric columns in {csv_path}")
        distance_col = numeric_cols[0]
        value_col = numeric_cols[1]
    else:
        distance_col = distance_candidates[0]
        if not value_candidates:
            numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
            value_candidates = [c for c in numeric_cols if c != distance_col]
        if not value_candidates:
            raise ValueError(f"Could not identify value column in {csv_path}")
        value_col = value_candidates[0]

    clean = df[[distance_col, value_col]].copy()
    clean.columns = ["distance", "value"]
    clean = clean.apply(pd.to_numeric, errors="coerce").dropna()

    return clean


def max_value_in_distance_range(
    df: pd.DataFrame,
    distance_min: float,
    distance_max: float,
    use_absolute: bool = True,
) -> Optional[float]:
    """
    Returns max value over a given distance interval.

    use_absolute=True is recommended for:
        E-field magnitude
        signed charge data when the magnitude matters
    """
    section = df[(df["distance"] >= distance_min) & (df["distance"] <= distance_max)]

    if section.empty:
        return None

    values = section["value"].to_numpy()
    values = np.abs(values) if use_absolute else values

    return float(np.nanmax(values))


def extract_max_outputs_from_distance_ranges(
    e_csv_path: str | Path | None = None,
    charge_csv_path: str | Path | None = None,
    conductor_range: tuple[float, float] | None = None,
    shield_range: tuple[float, float] | None = None,
    shed_range: tuple[float, float] | None = None,
) -> dict:
    """
    Extracts output-derived features from exported ELECTRO line plots.

    You provide the distance ranges corresponding to conductor, shield, and shed areas.

    Example:
        conductor_range = (0.0, 25.0)
        shield_range    = (120.0, 170.0)
        shed_range      = (250.0, 500.0)
    """

    results = {}

    if e_csv_path is not None:
        e_df = read_electro_xy_csv(e_csv_path)

        results["global_max_e"] = float(np.nanmax(np.abs(e_df["value"].to_numpy())))

        if conductor_range:
            results["conductor_max_e"] = max_value_in_distance_range(e_df, *conductor_range)
        if shield_range:
            results["shield_max_e"] = max_value_in_distance_range(e_df, *shield_range)
        if shed_range:
            results["shed_max_e"] = max_value_in_distance_range(e_df, *shed_range)

    if charge_csv_path is not None:
        q_df = read_electro_xy_csv(charge_csv_path)

        results["global_max_charge"] = float(np.nanmax(np.abs(q_df["value"].to_numpy())))

        if conductor_range:
            results["conductor_max_charge"] = max_value_in_distance_range(q_df, *conductor_range)
        if shield_range:
            results["shield_max_charge"] = max_value_in_distance_range(q_df, *shield_range)
        if shed_range:
            results["shed_max_charge"] = max_value_in_distance_range(q_df, *shed_range)

    return results


def merge_output_features(features: EasyBushingFeatures, output_features: dict) -> EasyBushingFeatures:
    """
    Adds output-derived max values into the dataclass feature row.
    """
    for key, value in output_features.items():
        if hasattr(features, key):
            setattr(features, key, value)
    return features


# -------------------------------------------------------------------------
# 5. Standalone test/demo mode
# -------------------------------------------------------------------------

if __name__ == "__main__":
    # This demo only tests CSV max extraction.
    # API extraction requires an opened ELECTRO project object from your main pipeline.

    example_e_csv = "exports/E_vs_distance.csv"
    example_charge_csv = "exports/charge_vs_distance.csv"

    if Path(example_e_csv).exists():
        output_features = extract_max_outputs_from_distance_ranges(
            e_csv_path=example_e_csv,
            charge_csv_path=example_charge_csv if Path(example_charge_csv).exists() else None,
            conductor_range=(0.0, 25.0),
            shield_range=(120.0, 170.0),
            shed_range=(250.0, 500.0),
        )

        features = EasyBushingFeatures(bushing_id="demo")
        features = merge_output_features(features, output_features)

        write_feature_row_csv(features, "outputs/extracted_feature_rows.csv")
        print("Extracted features:")
        print(asdict(features))
    else:
        print("Demo CSV not found. Import this module into API_Automation.py instead.")
