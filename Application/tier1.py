from pathlib import Path
import os
import re
import sys
from tkinter import Tk, filedialog

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Allow importing DSP helper scripts when this file is run from another working directory.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

try:
    from wavelet_analysis import extract_zoned_wavelet_features
except Exception as exc:
    extract_zoned_wavelet_features = None
    WAVELET_IMPORT_ERROR = exc
else:
    WAVELET_IMPORT_ERROR = None

try:
    from fft_analysis import extract_zoned_fft_features
except Exception as exc:
    extract_zoned_fft_features = None
    FFT_IMPORT_ERROR = exc
else:
    FFT_IMPORT_ERROR = None

# =========================
# NUMPY COMPATIBILITY
# =========================

try:
    integrate = np.trapezoid
except AttributeError:
    integrate = np.trapz

# =========================
# USER SETTINGS
# =========================

EXPORT_FILE: Path | None = None
OUTPUT_FEATURES_XLSX: Path | None = None
OUTPUT_FEATURES_CSV: Path | None = None


def resolve_export_file(
    export_file: str | Path | None = None,
) -> Path:
    """
    Resolve the ELECTRO export in this order:

    1. A path passed directly to main(export_file=...)
    2. ELECTRO_EXPORT_FILE set by automation_application.py
    3. A file picker when tier1.py is run by itself

    This prevents a second file-selection dialog when the main application
    already selected the ELECTRO export.
    """

    # Preferred option: direct argument from the main application.
    if export_file is not None:
        candidate = Path(export_file).expanduser().resolve()

        if not candidate.exists():
            raise FileNotFoundError(
                "The supplied ELECTRO export does not exist:\n"
                f"{candidate}"
            )

        return candidate

    # Second option: environment variable supplied by the main application.
    environment_path = os.environ.get(
        "ELECTRO_EXPORT_FILE",
        "",
    ).strip()

    if environment_path:
        candidate = Path(environment_path).expanduser().resolve()

        if not candidate.exists():
            raise FileNotFoundError(
                "ELECTRO_EXPORT_FILE points to a file that does not exist:\n"
                f"{candidate}"
            )

        print("\nUsing ELECTRO export supplied by the main application:")
        print(candidate)

        return candidate

    # Standalone fallback.
    root = Tk()
    root.withdraw()

    try:
        root.attributes("-topmost", True)
        root.update()

        selected = filedialog.askopenfilename(
            parent=root,
            title="Select ELECTRO export file",
            filetypes=[
                ("ELECTRO exports", "*.csv *.txt"),
                ("CSV files", "*.csv"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected:
        raise SystemExit("No ELECTRO export selected.")

    return Path(selected).expanduser().resolve()


def configure_input_and_outputs(
    export_file: str | Path | None = None,
    output_directory: str | Path | None = None,
) -> Path:
    """
    Configure the global paths used by the existing Tier 1 functions.

    By default, output files are written beside the selected ELECTRO export.
    The main application can optionally provide an active-design folder.
    """

    global EXPORT_FILE
    global OUTPUT_FEATURES_XLSX
    global OUTPUT_FEATURES_CSV

    EXPORT_FILE = resolve_export_file(export_file)

    if output_directory is None:
        output_dir = EXPORT_FILE.parent
    else:
        output_dir = Path(output_directory).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    OUTPUT_FEATURES_XLSX = output_dir / "tier1_input_vector.xlsx"
    OUTPUT_FEATURES_CSV = output_dir / "tier1_input_vector.csv"

    return EXPORT_FILE

N_ZONES = 4
ZONE_EDGES_PERCENT = [0, 35, 55, 85, 100]  # fallback/fixed only

# Tier 1 = E-field + surface-bound-Q post-simulation feature vector.
# Geometry is merged later. Zoned Wavelet/FFT DSP features currently describe
# the E-field spatial shape; Q receives signed and absolute-magnitude summaries.
TIER1_FIELD = "E"
ZONING_FIELD = "E"
PLOT_MEASUREMENT_MODE = "normalized"  # options: "normalized", "raw"
INCLUDE_DSP_FEATURES = True
INCLUDE_PER_CURVE_DSP_FEATURES = False  # Set True only if you want a much wider input vector.

VALID_LABELS = [
    "top_stress",
    "bottom_stress",
    "shield",
    "conductor",
    "ignore"
]

# =========================
# FIELD / PARSING HELPERS
# =========================

def infer_field_from_text(text: str):
    lower = str(text).lower()

    if "re{qpm}" in lower or re.search(r"curve\s*:\s*q\b", lower):
        return "Q", "C/m^2"
    if "re{em}" in lower or re.search(r"curve\s*:\s*e\b", lower):
        return "E", "kV/mm"
    if "re{pm}" in lower or re.search(r"curve\s*:\s*p\b", lower):
        return "P", "C/m^2"
    if "re{dm}" in lower or re.search(r"curve\s*:\s*d\b", lower):
        return "D", "C/m^2"

    return "unknown", ""


def read_electro_export(file_path: Path) -> pd.DataFrame:
    """
    Reads ELECTRO graph exports in long format.

    Expected ELECTRO block format:
        Curve : E - step = 2
          Point d(mm) x(mm) y(mm) z(mm) Re{Em}(kV/mm)
          1     ...   ...   ...   ...   value

    Output columns:
        curve_id, field, field_curve_id, unit, point, d, x, y, z, value
    """
    if not file_path.exists():
        raise FileNotFoundError(f"File does not exist:\n{file_path}")

    rows = []
    current_curve = 0
    current_field = "unknown"
    current_unit = ""
    field_counts = {}
    current_field_curve_id = 0

    number_pattern = re.compile(r"[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?")

    with open(file_path, "r", errors="ignore") as f:
        for line in f:
            stripped = line.strip()

            if not stripped:
                continue

            if stripped.startswith("Curve"):
                current_curve += 1
                field, unit = infer_field_from_text(stripped)
                current_field = field
                current_unit = unit
                field_counts[current_field] = field_counts.get(current_field, 0) + 1
                current_field_curve_id = field_counts[current_field]
                continue

            # Header row can refine field/unit if Curve line was ambiguous.
            if "Point" in stripped and "d(" in stripped and "x(" in stripped and "y(" in stripped and "z(" in stripped:
                field, unit = infer_field_from_text(stripped)
                if field != "unknown":
                    # If header changes the field from unknown, update count/index.
                    if current_field == "unknown":
                        current_field = field
                        current_unit = unit
                        field_counts[current_field] = field_counts.get(current_field, 0) + 1
                        current_field_curve_id = field_counts[current_field]
                    else:
                        current_field = field
                        current_unit = unit
                continue

            nums = number_pattern.findall(line)
            if len(nums) < 6:
                continue

            try:
                point = int(float(nums[0]))
                if point < 1:
                    continue

                d = float(nums[1])
                x = float(nums[2])
                y = float(nums[3])
                z = float(nums[4])
                value = float(nums[5])

                rows.append([
                    current_curve,
                    current_field,
                    current_field_curve_id,
                    current_unit,
                    point,
                    d,
                    x,
                    y,
                    z,
                    value
                ])

            except ValueError:
                continue

    if not rows:
        raise ValueError(f"No numeric ELECTRO data found in:\n{file_path}")

    df = pd.DataFrame(
        rows,
        columns=[
            "curve_id", "field", "field_curve_id", "unit",
            "point", "d", "x", "y", "z", "value"
        ]
    )

    print("\nSuccessfully parsed ELECTRO export.")
    print(f"Rows found: {len(df)}")
    print(f"Curves found: {df['curve_id'].nunique()}")
    print("\nDetected curves:")
    print(df.groupby(["curve_id", "field", "field_curve_id", "unit"]).size())

    return df


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    required = ["curve_id", "field", "field_curve_id", "unit", "point", "d", "x", "y", "z", "value"]
    missing = [col for col in required if col not in df.columns]

    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    numeric_cols = ["curve_id", "field_curve_id", "point", "d", "x", "y", "z", "value"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["field"] = df["field"].astype(str)
    df["unit"] = df["unit"].astype(str)

    df = df.dropna(subset=numeric_cols)
    df = df.sort_values(["curve_id", "d"]).reset_index(drop=True)

    if len(df) < 4:
        raise ValueError("Not enough usable rows after cleaning.")

    return df

# =========================
# ZONE LOGIC HELPERS
# =========================

def add_normalized_distance(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["d_norm"] = np.nan
    df["d_percent"] = np.nan

    for curve_id in sorted(df["curve_id"].unique()):
        mask = df["curve_id"] == curve_id
        d_min = df.loc[mask, "d"].min()
        d_max = df.loc[mask, "d"].max()

        if d_max == d_min:
            raise ValueError(f"Distance range is zero for curve {curve_id}.")

        df.loc[mask, "d_norm"] = (df.loc[mask, "d"] - d_min) / (d_max - d_min)
        df.loc[mask, "d_percent"] = df.loc[mask, "d_norm"] * 100

    return df


def robust_scale(y):
    y = np.asarray(y, dtype=float)
    med = np.nanmedian(y)
    mad = np.nanmedian(np.abs(y - med))

    if mad < 1e-12:
        std = np.nanstd(y)
        if std < 1e-12:
            return np.zeros_like(y)
        return (y - med) / std

    return (y - med) / (1.4826 * mad)


def rolling_std(y, window=7):
    return (
        pd.Series(y)
        .rolling(window=window, center=True, min_periods=1)
        .std()
        .fillna(0)
        .to_numpy()
    )


def moving_average(y, window=9):
    y = np.asarray(y, dtype=float)

    if window <= 1:
        return y

    window = int(window)
    if window % 2 == 0:
        window += 1

    pad = window // 2
    y_pad = np.pad(y, pad_width=pad, mode="edge")
    kernel = np.ones(window) / window
    return np.convolve(y_pad, kernel, mode="valid")


def apply_zone_edges_to_df(df: pd.DataFrame, zone_edges_percent, n_zones: int, method_name="unknown") -> pd.DataFrame:
    df = df.copy()
    df.attrs["zone_edges_percent"] = zone_edges_percent
    df.attrs["zoning_method"] = method_name

    edges = np.array(zone_edges_percent, dtype=float) / 100.0
    df["zone_id"] = np.nan

    for curve_id in sorted(df["curve_id"].unique()):
        mask = df["curve_id"] == curve_id
        zone_ids = np.digitize(df.loc[mask, "d_norm"], edges, right=False) - 1
        zone_ids = np.clip(zone_ids, 0, n_zones - 1)
        df.loc[mask, "zone_id"] = zone_ids + 1

    df["zone_id"] = df["zone_id"].astype(int)
    return df

# =========================
# DP / ADAPTIVE SEGMENTATION
# =========================

def segment_cost(prefix_sum, prefix_sq, start, end):
    n = end - start
    if n <= 0:
        return np.inf

    s = prefix_sum[end] - prefix_sum[start]
    ss = prefix_sq[end] - prefix_sq[start]
    return np.sum(ss - (s * s) / n)


def find_changepoint_edges(feature_matrix, n_zones=4, min_width=20):
    X = np.asarray(feature_matrix, dtype=float)
    n = len(X)

    prefix_sum = np.vstack([np.zeros(X.shape[1]), np.cumsum(X, axis=0)])
    prefix_sq = np.vstack([np.zeros(X.shape[1]), np.cumsum(X * X, axis=0)])

    dp = np.full((n_zones + 1, n + 1), np.inf)
    back = np.full((n_zones + 1, n + 1), -1, dtype=int)
    dp[0, 0] = 0

    for k in range(1, n_zones + 1):
        for end in range(k * min_width, n + 1):
            for start in range((k - 1) * min_width, end - min_width + 1):
                cost = dp[k - 1, start] + segment_cost(prefix_sum, prefix_sq, start, end)
                if cost < dp[k, end]:
                    dp[k, end] = cost
                    back[k, end] = start

    edges_idx = [n]
    end = n

    for k in range(n_zones, 0, -1):
        start = back[k, end]
        if start < 0:
            raise ValueError("Adaptive zoning failed. Try reducing min_width.")
        edges_idx.append(start)
        end = start

    edges_idx = sorted(edges_idx)
    edges_idx = [max(0, min(i, n - 1)) for i in edges_idx]
    edges_idx[0] = 0
    edges_idx[-1] = n - 1
    return edges_idx


def get_zoning_df(df: pd.DataFrame) -> pd.DataFrame:
    zoning_df = df[df["field"] == ZONING_FIELD].copy()

    if zoning_df.empty:
        available = sorted(df["field"].unique())
        raise ValueError(f"ZONING_FIELD='{ZONING_FIELD}' not found. Available fields: {available}")

    return zoning_df


def build_shape_feature_matrix(df, grid):
    feature_list = []
    zoning_df = get_zoning_df(df)

    for curve_id in sorted(zoning_df["curve_id"].unique()):
        curve_df = zoning_df[zoning_df["curve_id"] == curve_id].sort_values("d_percent")
        x = curve_df["d_percent"].to_numpy()
        y = curve_df["value"].to_numpy()
        y_interp = np.interp(grid, x, y)

        y_norm = robust_scale(y_interp)
        slope = robust_scale(np.gradient(y_norm, grid))
        curvature = robust_scale(np.gradient(slope, grid))
        local_variation = robust_scale(rolling_std(y_norm, window=9))

        feature_list.append(y_norm)
        feature_list.append(np.abs(slope))
        feature_list.append(np.abs(curvature))
        feature_list.append(local_variation)

    return np.vstack(feature_list).T


def compute_dp_zone_edges_percent(df, n_zones=4):
    grid = np.linspace(0, 100, 301)
    feature_matrix = build_shape_feature_matrix(df, grid)
    edges_idx = find_changepoint_edges(feature_matrix, n_zones=n_zones, min_width=25)

    edges_percent = [float(grid[min(i, len(grid) - 1)]) for i in edges_idx]
    edges_percent[0] = 0.0
    edges_percent[-1] = 100.0

    print("\nDP/adaptive zone edges:")
    for i in range(n_zones):
        print(f"  Zone {i + 1}: {edges_percent[i]:.2f}% to {edges_percent[i + 1]:.2f}%")

    return edges_percent

# =========================
# DERIVATIVE SEGMENTATION
# =========================

def choose_spaced_boundaries(candidate_positions, candidate_scores, n_needed, min_spacing_percent=12.0, preferred_windows=None):
    selected = []

    if preferred_windows is None:
        preferred_windows = [(15, 30), (50, 65), (75, 88)]

    candidate_positions = np.asarray(candidate_positions, dtype=float)
    candidate_scores = np.asarray(candidate_scores, dtype=float)

    for window_start, window_end in preferred_windows[:n_needed]:
        mask = (candidate_positions >= window_start) & (candidate_positions <= window_end)
        if np.any(mask):
            local_positions = candidate_positions[mask]
            local_scores = candidate_scores[mask]
            best_idx = np.argmax(local_scores)
            selected.append(float(local_positions[best_idx]))

    order = np.argsort(candidate_scores)[::-1]
    for idx in order:
        pos = float(candidate_positions[idx])
        if pos <= 5 or pos >= 95:
            continue
        if all(abs(pos - s) >= min_spacing_percent for s in selected):
            selected.append(pos)
        if len(selected) == n_needed:
            break

    return sorted(selected[:n_needed])


def compute_derivative_score(df, grid):
    combined_score = np.zeros_like(grid, dtype=float)
    zoning_df = get_zoning_df(df)

    for curve_id in sorted(zoning_df["curve_id"].unique()):
        curve_df = zoning_df[zoning_df["curve_id"] == curve_id].sort_values("d_percent")
        x = curve_df["d_percent"].to_numpy()
        y = curve_df["value"].to_numpy()

        y_interp = np.interp(grid, x, y)
        y_norm = robust_scale(y_interp)
        y_smooth = moving_average(y_norm, window=17)

        slope = np.gradient(y_smooth, grid)
        curvature = np.gradient(slope, grid)

        slope_score = np.abs(robust_scale(slope))
        curvature_score = np.abs(robust_scale(curvature))
        variation_score = np.abs(robust_scale(rolling_std(y_smooth, window=17)))

        curve_score = 0.55 * slope_score + 0.25 * curvature_score + 0.20 * variation_score
        combined_score += curve_score

    combined_score = moving_average(combined_score, window=15)
    ignore_mask = (grid < 8) | (grid > 95)
    combined_score[ignore_mask] = 0
    return combined_score


def get_derivative_candidates(df, grid):
    score = compute_derivative_score(df, grid)
    candidate_indices = []

    for i in range(1, len(score) - 1):
        if score[i] >= score[i - 1] and score[i] >= score[i + 1]:
            candidate_indices.append(i)

    if len(candidate_indices) == 0:
        return np.array([]), np.array([]), score

    return grid[candidate_indices], score[candidate_indices], score


def compute_derivative_zone_edges_percent(df, n_zones=4):
    grid = np.linspace(0, 100, 401)
    candidate_positions, candidate_scores, _ = get_derivative_candidates(df, grid)

    if len(candidate_positions) == 0:
        print("\nNo derivative candidates found. Falling back to fixed zones.")
        return ZONE_EDGES_PERCENT

    boundaries = choose_spaced_boundaries(
        candidate_positions,
        candidate_scores,
        n_needed=n_zones - 1,
        min_spacing_percent=12.0,
        preferred_windows=[(15, 30), (50, 65), (75, 88)]
    )

    if len(boundaries) < n_zones - 1:
        print("\nNot enough derivative boundaries found. Falling back to fixed zones.")
        return ZONE_EDGES_PERCENT

    edges_percent = [0.0] + boundaries + [100.0]

    print("\nDerivative-based zone edges:")
    for i in range(n_zones):
        print(f"  Zone {i + 1}: {edges_percent[i]:.2f}% to {edges_percent[i + 1]:.2f}%")

    return edges_percent

# =========================
# HYBRID SEGMENTATION
# =========================

def constrained_dp_edges(feature_matrix, grid, candidate_boundaries, n_zones=4, min_width_percent=8.0, window_percent=7.0):
    n = len(grid)
    min_width_idx = max(1, int(round(min_width_percent / 100.0 * (n - 1))))

    if len(candidate_boundaries) < n_zones - 1:
        return None

    candidate_lists = []
    for boundary in candidate_boundaries[:n_zones - 1]:
        low = boundary - window_percent
        high = boundary + window_percent
        idx = np.where((grid >= low) & (grid <= high))[0]
        idx = idx[(idx > 0) & (idx < n - 1)]
        if len(idx) == 0:
            nearest = int(np.argmin(np.abs(grid - boundary)))
            idx = np.array([nearest])
        candidate_lists.append(idx)

    prefix_sum = np.vstack([np.zeros(feature_matrix.shape[1]), np.cumsum(feature_matrix, axis=0)])
    prefix_sq = np.vstack([np.zeros(feature_matrix.shape[1]), np.cumsum(feature_matrix * feature_matrix, axis=0)])

    best_cost = np.inf
    best_edges = None

    def recurse(level, chosen):
        nonlocal best_cost, best_edges

        if level == len(candidate_lists):
            edges = [0] + chosen + [n - 1]
            if any(edges[i + 1] - edges[i] < min_width_idx for i in range(len(edges) - 1)):
                return

            seg_bounds = edges[:-1] + [n]
            cost = 0.0
            for i in range(len(seg_bounds) - 1):
                start = seg_bounds[i]
                end = seg_bounds[i + 1]
                cost += segment_cost(prefix_sum, prefix_sq, start, end)

            if cost < best_cost:
                best_cost = cost
                best_edges = edges
            return

        for idx in candidate_lists[level]:
            idx = int(idx)
            if chosen and idx <= chosen[-1]:
                continue
            recurse(level + 1, chosen + [idx])

    recurse(0, [])
    return best_edges


def compute_hybrid_zone_edges_percent(df, n_zones=4):
    grid = np.linspace(0, 100, 401)

    feature_matrix = build_shape_feature_matrix(df, grid)
    candidate_positions, candidate_scores, _ = get_derivative_candidates(df, grid)

    if len(candidate_positions) == 0:
        print("\nNo derivative candidates found for hybrid. Falling back to DP zones.")
        return compute_dp_zone_edges_percent(df, n_zones)

    derivative_boundaries = choose_spaced_boundaries(
        candidate_positions,
        candidate_scores,
        n_needed=n_zones - 1,
        min_spacing_percent=12.0,
        preferred_windows=[(15, 30), (50, 65), (75, 88)]
    )

    if len(derivative_boundaries) < n_zones - 1:
        print("\nNot enough derivative candidates for hybrid. Falling back to DP zones.")
        return compute_dp_zone_edges_percent(df, n_zones)

    hybrid_edges_idx = constrained_dp_edges(
        feature_matrix,
        grid,
        derivative_boundaries,
        n_zones=n_zones,
        min_width_percent=8.0,
        window_percent=7.0
    )

    if hybrid_edges_idx is None:
        print("\nHybrid constrained DP failed. Falling back to derivative zones.")
        return [0.0] + derivative_boundaries + [100.0]

    edges_percent = [float(grid[i]) for i in hybrid_edges_idx]
    edges_percent[0] = 0.0
    edges_percent[-1] = 100.0

    print("\nHybrid zone edges:")
    for i in range(n_zones):
        print(f"  Zone {i + 1}: {edges_percent[i]:.2f}% to {edges_percent[i + 1]:.2f}%")

    return edges_percent

# =========================
# PLOTTING / LABELING
# =========================

def create_zoned_versions(df: pd.DataFrame, n_zones: int):
    print("\nComputing DP/adaptive zoning...")
    dp_edges = compute_dp_zone_edges_percent(df, n_zones)
    df_dp = apply_zone_edges_to_df(df, dp_edges, n_zones, method_name="dp")

    print("\nComputing derivative-based zoning...")
    derivative_edges = compute_derivative_zone_edges_percent(df, n_zones)
    df_derivative = apply_zone_edges_to_df(df, derivative_edges, n_zones, method_name="derivative")

    print("\nComputing hybrid zoning...")
    hybrid_edges = compute_hybrid_zone_edges_percent(df, n_zones)
    df_hybrid = apply_zone_edges_to_df(df, hybrid_edges, n_zones, method_name="hybrid")

    return df_dp, df_derivative, df_hybrid


def plot_zones(df: pd.DataFrame, n_zones: int, title="ELECTRO Export - Zoned Field Measurements"):
    plt.figure(figsize=(13, 7))
    zone_edges_percent = df.attrs.get("zone_edges_percent", ZONE_EDGES_PERCENT)

    zone_colors = ["#ffd166", "#06d6a0", "#118ab2", "#ef476f"]

    for i in range(n_zones):
        start = zone_edges_percent[i]
        end = zone_edges_percent[i + 1]

        plt.axvspan(start, end, alpha=0.25, color=zone_colors[i % len(zone_colors)])
        plt.axvline(start, color="black", linewidth=1.1, alpha=0.65)
        plt.text(
            (start + end) / 2,
            0.98,
            f"Zone {i + 1}",
            transform=plt.gca().get_xaxis_transform(),
            ha="center",
            va="top",
            fontsize=10,
            fontweight="bold"
        )

    plt.axvline(100, color="black", linewidth=1.1, alpha=0.65)

    cmap = plt.get_cmap("tab10")

    for idx, curve_id in enumerate(sorted(df["curve_id"].unique())):
        curve_df = df[df["curve_id"] == curve_id].sort_values("d_percent")
        y = curve_df["value"].to_numpy(dtype=float)

        if np.all(np.isnan(y)):
            continue

        if PLOT_MEASUREMENT_MODE.lower() == "normalized":
            y_plot = robust_scale(y)
        elif PLOT_MEASUREMENT_MODE.lower() == "raw":
            y_plot = y
        else:
            raise ValueError("PLOT_MEASUREMENT_MODE must be 'normalized' or 'raw'.")

        field = curve_df["field"].iloc[0]
        field_curve_id = int(curve_df["field_curve_id"].iloc[0])
        unit = curve_df["unit"].iloc[0]
        label = f"Curve {int(curve_id)} - {field}{field_curve_id} ({unit})"

        plt.plot(
            curve_df["d_percent"],
            y_plot,
            label=label,
            color=cmap(idx % 10),
            linestyle="-",
            marker=".",
            linewidth=1.5,
            markersize=3
        )

    plt.xlabel("Normalized distance (%)")
    ylabel = "Robust-scaled measurement value" if PLOT_MEASUREMENT_MODE.lower() == "normalized" else "Measurement value (raw units)"
    plt.ylabel(ylabel)
    plt.title(title)
    plt.grid(True, alpha=0.35)
    plt.legend(fontsize=8, ncol=2)
    plt.tight_layout()
    plt.show()


def prompt_zone_labels(n_zones: int):
    """
    Prompt the user to assign a physical label to each temporary zone.

    Numeric shortcuts:
        1 -> bottom_stress
        2 -> conductor
        3 -> top_stress
        4 -> shield
        5 -> ignore

    The stored labels remain the full semantic strings, so the input vector
    still contains values such as ``bottom_stress`` and ``conductor`` rather
    than the numeric shortcuts.
    """
    label_choices = {
        "1": "bottom_stress",
        "2": "conductor",
        "3": "top_stress",
        "4": "shield",
        "5": "ignore",
    }

    # Also allow the full text labels for backward compatibility.
    text_aliases = {
        "bottom_stress": "bottom_stress",
        "bottom": "bottom_stress",
        "conductor": "conductor",
        "top_stress": "top_stress",
        "top": "top_stress",
        "shield": "shield",
        "ignore": "ignore",
    }

    print("\nZone label options:")
    print("  1 = bottom_stress")
    print("  2 = conductor")
    print("  3 = top_stress")
    print("  4 = shield")
    print("  5 = ignore")

    zone_labels = {}

    for zone in range(1, n_zones + 1):
        while True:
            user_entry = input(
                f"\nEnter label number for Zone {zone} "
                "(1=bottom, 2=conductor, 3=top, 4=shield, 5=ignore): "
            ).strip().lower()

            if user_entry in label_choices:
                resolved_label = label_choices[user_entry]
            elif user_entry in text_aliases:
                resolved_label = text_aliases[user_entry]
            else:
                print(
                    "Invalid label. Enter 1, 2, 3, 4, or 5 "
                    "(or type the full label name)."
                )
                continue

            zone_labels[zone] = resolved_label
            print(f"  Zone {zone} assigned: {resolved_label}")
            break

    return zone_labels


def prompt_pass_fail_label():
    """
    Prompt user for known physical/simulation outcome label.

    Stored as both:
        pass_fail_label: 'pass', 'fail', or 'n/a'
        pass_fail_code:  1 for pass, 0 for fail, NaN for n/a
    """
    valid = {
        "pass": ("pass", 1),
        "p": ("pass", 1),
        "1": ("pass", 1),
        "fail": ("fail", 0),
        "f": ("fail", 0),
        "0": ("fail", 0),
        "n/a": ("n/a", np.nan),
        "na": ("n/a", np.nan),
        "n": ("n/a", np.nan),
        "unknown": ("n/a", np.nan),
        "": ("n/a", np.nan),
    }

    print("\nOutcome label options:")
    print("  - Pass")
    print("  - Fail")
    print("  - N/A")

    while True:
        user_label = input("\nEnter Pass/Fail/N/A label for this simulation: ").strip().lower()
        if user_label in valid:
            label, code = valid[user_label]
            return label, code

        print("Invalid label. Enter Pass, Fail, or N/A.")

# =========================
# TIER 1 FEATURE EXTRACTION
# =========================

def safe_divide(num, den):
    if den is None or not np.isfinite(den) or abs(den) < 1e-30:
        return np.nan
    return num / den


def add_basic_field_stats(features: dict, prefix: str, data: pd.DataFrame):
    values = data["value"].to_numpy(dtype=float)
    d_norm = data["d_norm"].to_numpy(dtype=float)

    values = values[np.isfinite(values)]
    if len(values) == 0:
        return

    data_valid = data[np.isfinite(data["value"].to_numpy(dtype=float))]

    max_idx = data_valid["value"].idxmax()
    features[f"{prefix}_max"] = float(data_valid["value"].max())
    features[f"{prefix}_mean"] = float(data_valid["value"].mean())
    features[f"{prefix}_p95"] = float(np.percentile(data_valid["value"], 95))
    features[f"{prefix}_auc"] = float(integrate(data_valid["value"], data_valid["d_norm"]))
    features[f"{prefix}_peak_d_percent"] = float(data_valid.loc[max_idx, "d_percent"])
    features[f"{prefix}_peak_x"] = float(data_valid.loc[max_idx, "x"])
    features[f"{prefix}_peak_y"] = float(data_valid.loc[max_idx, "y"])
    features[f"{prefix}_peak_z"] = float(data_valid.loc[max_idx, "z"])
    features[f"{prefix}_peak_zone_id"] = int(data_valid.loc[max_idx, "zone_id"])


def _mean_curve_auc(data: pd.DataFrame, absolute: bool = False) -> float:
    """
    Integrate each physical curve independently, then average curve AUCs.

    This avoids joining the end of one curve to the beginning of another when
    several curves share the same field. Averaging also prevents AUC magnitude
    from changing merely because an export contains an additional Q curve.
    """
    auc_values = []
    for _, curve in data.groupby("field_curve_id", sort=True):
        curve = curve.sort_values("d_norm")
        x = curve["d_norm"].to_numpy(dtype=float)
        y = curve["value"].to_numpy(dtype=float)
        valid = np.isfinite(x) & np.isfinite(y)
        if valid.sum() < 2:
            continue
        if absolute:
            y = np.abs(y)
        auc_values.append(float(integrate(y[valid], x[valid])))
    return float(np.mean(auc_values)) if auc_values else np.nan


def add_surface_bound_q_stats(features: dict, prefix: str, data: pd.DataFrame):
    """Add polarity-preserving and severity-oriented surface-bound-Q features."""
    data_valid = data[np.isfinite(data["value"].to_numpy(dtype=float))].copy()
    if data_valid.empty:
        return

    values = data_valid["value"].to_numpy(dtype=float)
    abs_values = np.abs(values)
    abs_peak_position = int(np.argmax(abs_values))
    abs_peak_index = data_valid.index[abs_peak_position]

    # Signed values preserve polarity behavior; absolute values describe stress
    # severity without treating a large negative charge as a small response.
    features[f"{prefix}_max"] = float(np.max(values))
    features[f"{prefix}_min"] = float(np.min(values))
    features[f"{prefix}_mean"] = float(np.mean(values))
    features[f"{prefix}_p95"] = float(np.percentile(values, 95))
    features[f"{prefix}_auc"] = _mean_curve_auc(data_valid, absolute=False)
    features[f"{prefix}_abs_max"] = float(np.max(abs_values))
    features[f"{prefix}_mean_abs"] = float(np.mean(abs_values))
    features[f"{prefix}_p95_abs"] = float(np.percentile(abs_values, 95))
    features[f"{prefix}_auc_abs"] = _mean_curve_auc(data_valid, absolute=True)
    features[f"{prefix}_abs_peak_d_percent"] = float(data_valid.loc[abs_peak_index, "d_percent"])
    features[f"{prefix}_abs_peak_x_mm"] = float(data_valid.loc[abs_peak_index, "x"])
    features[f"{prefix}_abs_peak_y_mm"] = float(data_valid.loc[abs_peak_index, "y"])
    features[f"{prefix}_abs_peak_z_mm"] = float(data_valid.loc[abs_peak_index, "z"])
    features[f"{prefix}_abs_peak_zone_id"] = int(data_valid.loc[abs_peak_index, "zone_id"])


def extract_tier1_input_vector(df: pd.DataFrame, zone_labels: dict, selected_method: str, pass_fail_label: str, pass_fail_code) -> dict:
    """
    Tier 1 post-simulation input vector.

    Includes E-field and surface-bound-Q features. When enabled in main(),
    zoned E-field DSP features are appended from the selected zoning.
    Geometry features are still excluded here and can be merged later.
    """
    e_df = df[df["field"] == TIER1_FIELD].copy()
    if e_df.empty:
        available = sorted(df["field"].unique())
        raise ValueError(f"Tier 1 extraction requires field '{TIER1_FIELD}'. Available fields: {available}")

    q_df = df[df["field"] == "Q"].copy()

    features = {
        "source_file": EXPORT_FILE.name,
        "num_total_points": int(len(df)),
        "num_E_points": int(len(e_df)),
        "num_Q_points": int(len(q_df)),
        "surface_bound_Q_available": int(not q_df.empty),
        "tier": "tier1_E_Q_with_zoned_E_DSP_no_geometry",
        "zoning_method": selected_method,
        "zoning_field": ZONING_FIELD,
        "tier1_field": TIER1_FIELD,
        "n_zones": int(N_ZONES),
        "pass_fail_label": pass_fail_label,
        "pass_fail_code": pass_fail_code,
    }

    edges = df.attrs.get("zone_edges_percent", [])
    for i, edge in enumerate(edges):
        features[f"zone_edge_{i}_percent"] = float(edge)

    for zone_id in range(1, N_ZONES + 1):
        label = zone_labels.get(zone_id, "unknown")
        features[f"zone{zone_id}_label"] = label

    # Global E stats across all E curves.
    add_basic_field_stats(features, "global_E", e_df)

    # Global peak zone label.
    if "global_E_peak_zone_id" in features:
        peak_zone_id = int(features["global_E_peak_zone_id"])
        features["global_E_peak_zone_label"] = zone_labels.get(peak_zone_id, "unknown")

    global_E_max = features.get("global_E_max", np.nan)
    global_E_auc = features.get("global_E_auc", np.nan)

    # Per E-curve stats, usually inside/outside curves.
    for field_curve_id in sorted(e_df["field_curve_id"].unique()):
        curve_df = e_df[e_df["field_curve_id"] == field_curve_id]
        prefix = f"E_curve{int(field_curve_id)}"
        add_basic_field_stats(features, prefix, curve_df)

        if f"{prefix}_peak_zone_id" in features:
            peak_zone_id = int(features[f"{prefix}_peak_zone_id"])
            features[f"{prefix}_peak_zone_label"] = zone_labels.get(peak_zone_id, "unknown")

    # Labeled zone stats, combining E curves within each zone.
    for zone_id, label in zone_labels.items():
        if label == "ignore":
            continue

        zone_df = e_df[e_df["zone_id"] == zone_id]
        if zone_df.empty:
            continue

        prefix = f"zone{zone_id}_{label}_E"
        features[f"zone{zone_id}_{label}_start_percent"] = float(zone_df["d_percent"].min())
        features[f"zone{zone_id}_{label}_end_percent"] = float(zone_df["d_percent"].max())
        add_basic_field_stats(features, prefix, zone_df)

        # Relative zone severity values.
        features[f"{prefix}_max_over_global_E_max"] = safe_divide(features.get(f"{prefix}_max", np.nan), global_E_max)
        features[f"{prefix}_auc_over_global_E_auc"] = safe_divide(features.get(f"{prefix}_auc", np.nan), global_E_auc)

        # Zone + E curve stats.
        for field_curve_id in sorted(zone_df["field_curve_id"].unique()):
            curve_zone_df = zone_df[zone_df["field_curve_id"] == field_curve_id]
            if curve_zone_df.empty:
                continue

            cz_prefix = f"E_curve{int(field_curve_id)}_zone{zone_id}_{label}"
            add_basic_field_stats(features, cz_prefix, curve_zone_df)
            features[f"{cz_prefix}_max_over_global_E_max"] = safe_divide(features.get(f"{cz_prefix}_max", np.nan), global_E_max)
            features[f"{cz_prefix}_auc_over_global_E_auc"] = safe_divide(features.get(f"{cz_prefix}_auc", np.nan), global_E_auc)

    # Which labeled zone dominates E max/AUC?
    zone_max_candidates = {}
    zone_auc_candidates = {}
    for zone_id, label in zone_labels.items():
        if label == "ignore":
            continue
        prefix = f"zone{zone_id}_{label}_E"
        if f"{prefix}_max" in features:
            zone_max_candidates[label] = features[f"{prefix}_max"]
        if f"{prefix}_auc" in features:
            zone_auc_candidates[label] = features[f"{prefix}_auc"]

    if zone_max_candidates:
        features["dominant_E_max_zone_label"] = max(zone_max_candidates, key=zone_max_candidates.get)
    if zone_auc_candidates:
        features["dominant_E_auc_zone_label"] = max(zone_auc_candidates, key=zone_auc_candidates.get)

    # Surface-bound Q is optional for backward compatibility with E-only
    # exports. When present, use the same physical zones derived from E.
    if not q_df.empty:
        global_q_prefix = "global_surface_bound_Q"
        add_surface_bound_q_stats(features, global_q_prefix, q_df)
        global_q_abs_max = features.get(f"{global_q_prefix}_abs_max", np.nan)
        global_q_auc_abs = features.get(f"{global_q_prefix}_auc_abs", np.nan)

        if f"{global_q_prefix}_abs_peak_zone_id" in features:
            peak_zone_id = int(features[f"{global_q_prefix}_abs_peak_zone_id"])
            features[f"{global_q_prefix}_abs_peak_zone_label"] = zone_labels.get(
                peak_zone_id,
                "unknown",
            )

        # Retain curve-specific Q summaries in the expanded research vector.
        for field_curve_id in sorted(q_df["field_curve_id"].unique()):
            curve_df = q_df[q_df["field_curve_id"] == field_curve_id]
            add_surface_bound_q_stats(
                features,
                f"surface_bound_Q_curve{int(field_curve_id)}",
                curve_df,
            )

        q_zone_max_candidates = {}
        q_zone_auc_candidates = {}
        for zone_id, label in zone_labels.items():
            if label == "ignore":
                continue
            zone_q_df = q_df[q_df["zone_id"] == zone_id]
            if zone_q_df.empty:
                continue

            prefix = f"zone{zone_id}_{label}_surface_bound_Q"
            add_surface_bound_q_stats(features, prefix, zone_q_df)
            features[f"{prefix}_abs_max_over_global_surface_bound_Q_abs_max"] = safe_divide(
                features.get(f"{prefix}_abs_max", np.nan),
                global_q_abs_max,
            )
            features[f"{prefix}_auc_abs_over_global_surface_bound_Q_auc_abs"] = safe_divide(
                features.get(f"{prefix}_auc_abs", np.nan),
                global_q_auc_abs,
            )
            q_zone_max_candidates[label] = features.get(f"{prefix}_abs_max", np.nan)
            q_zone_auc_candidates[label] = features.get(f"{prefix}_auc_abs", np.nan)

        finite_q_max = {k: v for k, v in q_zone_max_candidates.items() if np.isfinite(v)}
        finite_q_auc = {k: v for k, v in q_zone_auc_candidates.items() if np.isfinite(v)}
        if finite_q_max:
            features["dominant_surface_bound_Q_abs_max_zone_label"] = max(
                finite_q_max,
                key=finite_q_max.get,
            )
        if finite_q_auc:
            features["dominant_surface_bound_Q_auc_abs_zone_label"] = max(
                finite_q_auc,
                key=finite_q_auc.get,
            )

    return features



def extract_zoned_dsp_input_features(df: pd.DataFrame, zone_labels: dict) -> dict:
    """
    Runs the imported WaveletAnalysis.py and FFT_Analysis.py feature functions on
    the same four zones selected for the Tier 1 vector.

    Important implementation choice:
        The ELECTRO export used here is E-field vs normalized distance, so both
        DSP methods are applied spatially per zone. The wavelet feature names use
        spatial frequency units (cycles/mm equivalent after normalization as
        cycles/percent because d_percent is the axis). The ratios are still valid
        comparative model inputs because every bushing is transformed the same way.
    """
    features = {}

    if not INCLUDE_DSP_FEATURES:
        features["dsp_features_enabled"] = 0
        return features

    features["dsp_features_enabled"] = 1

    if extract_zoned_wavelet_features is None:
        features["dsp_wavelet_available"] = 0
        features["dsp_wavelet_error"] = str(WAVELET_IMPORT_ERROR)
    else:
        try:
            wavelet_features = extract_zoned_wavelet_features(
                df,
                zone_labels=zone_labels,
                field=TIER1_FIELD,
                include_per_curve=INCLUDE_PER_CURVE_DSP_FEATURES,
            )
            features.update(wavelet_features)
        except Exception as exc:
            features["dsp_wavelet_available"] = 0
            features["dsp_wavelet_error"] = str(exc)

    if extract_zoned_fft_features is None:
        features["dsp_fft_available"] = 0
        features["dsp_fft_error"] = str(FFT_IMPORT_ERROR)
    else:
        try:
            fft_features = extract_zoned_fft_features(
                df,
                zone_labels=zone_labels,
                field=TIER1_FIELD,
                include_per_curve=INCLUDE_PER_CURVE_DSP_FEATURES,
            )
            features.update(fft_features)
        except Exception as exc:
            features["dsp_fft_available"] = 0
            features["dsp_fft_error"] = str(exc)

    return features


def append_features_to_outputs(features: dict, csv_path: Path, xlsx_path: Path):
    new_row = pd.DataFrame([features])

    if csv_path.exists():
        old = pd.read_csv(csv_path)
        combined = pd.concat([old, new_row], ignore_index=True, sort=False)
    else:
        combined = new_row

    combined.to_csv(csv_path, index=False)

    try:
        combined.to_excel(xlsx_path, index=False, sheet_name="Tier1_Input_Vector")
        print(f"Saved Excel output to:\n{xlsx_path}")
    except Exception as exc:
        print("\nWarning: Could not write .xlsx output.")
        print(f"Reason: {exc}")
        print("CSV output was still written successfully.")

    print(f"Saved CSV output to:\n{csv_path}")

# =========================
# MAIN
# =========================

def main(export_file=None, output_dir=None):
    """
    Run Tier 1 extraction.

    When called by automation_application.py, ``export_file`` and ``output_dir``
    are passed directly, so no additional file-selection dialog is opened.

    When run standalone, ``resolve_export_file`` retains the file-picker
    fallback. ``ELECTRO_EXPORT_FILE`` is also supported for compatibility.
    """
    selected_export = configure_input_and_outputs(
        export_file=export_file,
        output_directory=output_dir,
    )

    output_directory = OUTPUT_FEATURES_CSV.parent

    print(f"Reading export file:\n{selected_export}")
    print(f"Tier 1 outputs will be saved to:\n{output_directory}")

    df = read_electro_export(selected_export)
    df = standardize_columns(df)
    df = add_normalized_distance(df)

    df_dp, df_derivative, df_hybrid = create_zoned_versions(
        df,
        N_ZONES
    )

    plot_zones(
        df_dp,
        N_ZONES,
        title="ELECTRO Export - DP Zones - All Measurements"
    )

    plot_zones(
        df_derivative,
        N_ZONES,
        title="ELECTRO Export - Derivative Zones - All Measurements"
    )

    plot_zones(
        df_hybrid,
        N_ZONES,
        title="ELECTRO Export - Hybrid Zones - All Measurements"
    )

    while True:
        choice = input(
            "\nWhich zoning method do you want to use for Tier 1 "
            "data extraction? Enter 'dp', 'derivative', or 'hybrid': "
        ).strip().lower()

        if choice in {"dp", "adaptive"}:
            selected_df = df_dp
            selected_method = "dp"
            break

        if choice in {"derivative", "deriv"}:
            selected_df = df_derivative
            selected_method = "derivative"
            break

        if choice in {"hybrid", "h"}:
            selected_df = df_hybrid
            selected_method = "hybrid"
            break

        print(
            "Invalid choice. Enter 'dp', 'derivative', or 'hybrid'."
        )

    print(f"\nSelected zoning method: {selected_method}")

    zone_labels = prompt_zone_labels(N_ZONES)

    print("\nZone labels selected:")
    for zone, label in zone_labels.items():
        print(f"  Zone {zone}: {label}")

    pass_fail_label, pass_fail_code = prompt_pass_fail_label()

    print(f"\nSelected outcome label: {pass_fail_label}")

    features = extract_tier1_input_vector(
        selected_df,
        zone_labels,
        selected_method,
        pass_fail_label,
        pass_fail_code
    )

    print("\nRunning zoned Wavelet and FFT feature extraction...")

    dsp_features = extract_zoned_dsp_input_features(
        selected_df,
        zone_labels
    )

    features.update(dsp_features)

    append_features_to_outputs(
        features,
        OUTPUT_FEATURES_CSV,
        OUTPUT_FEATURES_XLSX
    )

    print("\nDone.")
    print(
        "Tier 1 output now includes zoned E-field features, signed/absolute "
        "surface-bound-Q features, and zoned E-field Wavelet/FFT DSP features."
    )
    print(
        "Geometry features are still excluded here and should be "
        "merged by the main application."
    )

    return features


if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    selected_file = filedialog.askopenfilename(
        title="Select ELECTRO export file",
        filetypes=[
            ("CSV and TXT files", "*.csv *.txt"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()

    if not selected_file:
        raise SystemExit("No file selected.")

    main(export_file=selected_file)
