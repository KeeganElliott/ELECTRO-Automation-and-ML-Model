from pathlib import Path
import re
from tkinter import Tk, filedialog

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


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

Tk().withdraw()

EXPORT_FILE = Path(filedialog.askopenfilename(
    title="Select ELECTRO export file",
    filetypes=[
        ("CSV and TXT files", "*.csv *.txt"),
        ("All files", "*.*")
    ]
))

if not EXPORT_FILE:
    raise SystemExit("No file selected.")

OUTPUT_FEATURES_CSV = EXPORT_FILE.parent / "labeled_post_sim_features.csv"

N_ZONES = 4
ZONE_EDGES_PERCENT = [0, 35, 55, 85, 100]  # fallback/fixed only

VALID_LABELS = [
    "top_stress",
    "bottom_stress",
    "shield",
    "conductor",
    "ignore"
]


# =========================
# FILE READING
# =========================

def read_electro_export(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File does not exist:\n{file_path}")

    rows = []
    current_curve = 0
    number_pattern = re.compile(r"[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?")

    with open(file_path, "r", errors="ignore") as f:
        for line in f:
            stripped = line.strip()

            if stripped.startswith("Curve"):
                current_curve += 1
                continue

            nums = number_pattern.findall(line)
            if len(nums) < 6:
                continue

            try:
                point = int(float(nums[0]))
                d = float(nums[1])
                x = float(nums[2])
                y = float(nums[3])
                z = float(nums[4])
                E = float(nums[5])

                if point < 1:
                    continue

                rows.append([current_curve, point, d, x, y, z, E])

            except ValueError:
                continue

    if not rows:
        raise ValueError(f"No numeric ELECTRO data found in:\n{file_path}")

    df = pd.DataFrame(rows, columns=["curve_id", "point", "d", "x", "y", "z", "E"])

    print("\nSuccessfully parsed ELECTRO export.")
    print(f"Rows found: {len(df)}")
    print(f"Curves found: {df['curve_id'].nunique()}")
    print(df.head())

    return df


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    required = ["curve_id", "point", "d", "x", "y", "z", "E"]
    missing = [col for col in required if col not in df.columns]

    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    for col in required:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=required)
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


def build_shape_feature_matrix(df, grid):
    feature_list = []

    for curve_id in sorted(df["curve_id"].unique()):
        curve_df = df[df["curve_id"] == curve_id].sort_values("d_percent")
        x = curve_df["d_percent"].to_numpy()
        y = curve_df["E"].to_numpy()
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
        preferred_windows = [
            (15, 30),
            (50, 65),
            (75, 88),
        ]

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

    for curve_id in sorted(df["curve_id"].unique()):
        curve_df = df[df["curve_id"] == curve_id].sort_values("d_percent")
        x = curve_df["d_percent"].to_numpy()
        y = curve_df["E"].to_numpy()

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
    """
    Hybrid method:
    - derivative finds likely transition neighborhoods
    - DP chooses the lowest-cost set of boundaries near those neighborhoods
    """
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

    # N_ZONES=4 expected here, but this recursive version works generally.
    def recurse(level, chosen):
        nonlocal best_cost, best_edges

        if level == len(candidate_lists):
            edges = [0] + chosen + [n - 1]
            if any(edges[i + 1] - edges[i] < min_width_idx for i in range(len(edges) - 1)):
                return

            # Use end-exclusive indices for segment_cost; add 1 to final edge.
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


def plot_zones(df: pd.DataFrame, n_zones: int, title="ELECTRO Export - Zoned E-Stress"):
    plt.figure(figsize=(12, 6))
    zone_edges_percent = df.attrs.get("zone_edges_percent", ZONE_EDGES_PERCENT)

    zone_colors = ["#ffd166", "#06d6a0", "#118ab2", "#ef476f"]

    for i in range(n_zones):
        start = zone_edges_percent[i]
        end = zone_edges_percent[i + 1]

        plt.axvspan(start, end, alpha=0.32, color=zone_colors[i % len(zone_colors)])
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

    curve_styles = {
        1: {"label": "Curve 1 - Inside", "color": "tab:blue", "linestyle": "-", "marker": "."},
        2: {"label": "Curve 2 - Outside", "color": "tab:orange", "linestyle": "--", "marker": "."},
    }

    for curve_id in sorted(df["curve_id"].unique()):
        curve_df = df[df["curve_id"] == curve_id].sort_values("d_percent")
        style = curve_styles.get(
            curve_id,
            {"label": f"Curve {curve_id}", "color": None, "linestyle": "-", "marker": "."}
        )

        plt.plot(
            curve_df["d_percent"],
            curve_df["E"],
            label=style["label"],
            color=style["color"],
            linestyle=style["linestyle"],
            marker=style["marker"],
            linewidth=1.8,
            markersize=4
        )

    plt.xlabel("Normalized distance (%)")
    plt.ylabel("Electric field E")
    plt.title(title)
    plt.grid(True, alpha=0.35)
    plt.legend()
    plt.tight_layout()
    plt.show()


def prompt_zone_labels(n_zones: int):
    print("\nAvailable labels:")
    for label in VALID_LABELS:
        print(f"  - {label}")

    zone_labels = {}
    for zone in range(1, n_zones + 1):
        while True:
            user_label = input(f"\nEnter label for Zone {zone}: ").strip().lower()
            if user_label in VALID_LABELS:
                zone_labels[zone] = user_label
                break
            print("Invalid label. Please choose from the available labels.")

    return zone_labels


# =========================
# FEATURE EXTRACTION
# =========================

def extract_zone_features(df: pd.DataFrame, zone_labels: dict) -> dict:
    features = {}
    features["source_file"] = EXPORT_FILE.name
    features["num_points"] = len(df)

    features["global_max_E"] = df["E"].max()
    features["global_mean_E"] = df["E"].mean()
    features["global_p95_E"] = np.percentile(df["E"], 95)
    features["global_auc_E"] = integrate(df["E"], df["d_norm"])

    global_peak_idx = df["E"].idxmax()
    features["global_peak_d_percent"] = df.loc[global_peak_idx, "d_percent"]
    features["global_peak_x"] = df.loc[global_peak_idx, "x"]
    features["global_peak_y"] = df.loc[global_peak_idx, "y"]
    features["global_peak_z"] = df.loc[global_peak_idx, "z"]

    for curve_id in sorted(df["curve_id"].unique()):
        curve_df = df[df["curve_id"] == curve_id]
        curve_prefix = f"curve{int(curve_id)}"

        features[f"{curve_prefix}_max_E"] = curve_df["E"].max()
        features[f"{curve_prefix}_mean_E"] = curve_df["E"].mean()
        features[f"{curve_prefix}_p95_E"] = np.percentile(curve_df["E"], 95)
        features[f"{curve_prefix}_auc_E"] = integrate(curve_df["E"], curve_df["d_norm"])

        curve_peak_idx = curve_df["E"].idxmax()
        features[f"{curve_prefix}_peak_d_percent"] = curve_df.loc[curve_peak_idx, "d_percent"]
        features[f"{curve_prefix}_peak_x"] = curve_df.loc[curve_peak_idx, "x"]
        features[f"{curve_prefix}_peak_y"] = curve_df.loc[curve_peak_idx, "y"]
        features[f"{curve_prefix}_peak_z"] = curve_df.loc[curve_peak_idx, "z"]

    for zone_id, label in zone_labels.items():
        zone_df = df[df["zone_id"] == zone_id]
        if zone_df.empty or label == "ignore":
            continue

        prefix = f"zone{zone_id}_{label}"
        peak_idx = zone_df["E"].idxmax()

        features[f"{prefix}_start_percent"] = zone_df["d_percent"].min()
        features[f"{prefix}_end_percent"] = zone_df["d_percent"].max()
        features[f"{prefix}_max_E"] = zone_df["E"].max()
        features[f"{prefix}_mean_E"] = zone_df["E"].mean()
        features[f"{prefix}_p95_E"] = np.percentile(zone_df["E"], 95)
        features[f"{prefix}_auc_E"] = integrate(zone_df["E"], zone_df["d_norm"])
        features[f"{prefix}_peak_d_percent"] = zone_df.loc[peak_idx, "d_percent"]
        features[f"{prefix}_peak_x"] = zone_df.loc[peak_idx, "x"]
        features[f"{prefix}_peak_y"] = zone_df.loc[peak_idx, "y"]
        features[f"{prefix}_peak_z"] = zone_df.loc[peak_idx, "z"]

        for curve_id in sorted(zone_df["curve_id"].unique()):
            curve_zone_df = zone_df[zone_df["curve_id"] == curve_id]
            if curve_zone_df.empty:
                continue

            curve_zone_prefix = f"curve{int(curve_id)}_zone{zone_id}_{label}"
            curve_zone_peak_idx = curve_zone_df["E"].idxmax()

            features[f"{curve_zone_prefix}_max_E"] = curve_zone_df["E"].max()
            features[f"{curve_zone_prefix}_mean_E"] = curve_zone_df["E"].mean()
            features[f"{curve_zone_prefix}_p95_E"] = np.percentile(curve_zone_df["E"], 95)
            features[f"{curve_zone_prefix}_auc_E"] = integrate(curve_zone_df["E"], curve_zone_df["d_norm"])
            features[f"{curve_zone_prefix}_peak_d_percent"] = curve_zone_df.loc[curve_zone_peak_idx, "d_percent"]
            features[f"{curve_zone_prefix}_peak_x"] = curve_zone_df.loc[curve_zone_peak_idx, "x"]
            features[f"{curve_zone_prefix}_peak_y"] = curve_zone_df.loc[curve_zone_peak_idx, "y"]
            features[f"{curve_zone_prefix}_peak_z"] = curve_zone_df.loc[curve_zone_peak_idx, "z"]

    return features


def append_features_to_csv(features: dict, output_csv: Path):
    new_row = pd.DataFrame([features])

    if output_csv.exists():
        old = pd.read_csv(output_csv)
        combined = pd.concat([old, new_row], ignore_index=True)
    else:
        combined = new_row

    combined.to_csv(output_csv, index=False)


# =========================
# MAIN
# =========================

def main():
    print(f"Reading export file:\n{EXPORT_FILE}")

    df = read_electro_export(EXPORT_FILE)
    df = standardize_columns(df)
    df = add_normalized_distance(df)

    df_dp, df_derivative, df_hybrid = create_zoned_versions(df, N_ZONES)

    plot_zones(df_dp, N_ZONES, title="ELECTRO Export - DP / Adaptive Shape-Based Zones")
    plot_zones(df_derivative, N_ZONES, title="ELECTRO Export - Derivative-Based Zones")
    plot_zones(df_hybrid, N_ZONES, title="ELECTRO Export - Hybrid Zones")

    while True:
        choice = input(
            "\nWhich zoning method do you want to use for data extraction? "
            "Enter 'dp', 'derivative', or 'hybrid': "
        ).strip().lower()

        if choice in ["dp", "adaptive"]:
            selected_df = df_dp
            selected_method = "dp"
            break
        elif choice in ["derivative", "deriv"]:
            selected_df = df_derivative
            selected_method = "derivative"
            break
        elif choice in ["hybrid", "h"]:
            selected_df = df_hybrid
            selected_method = "hybrid"
            break

        print("Invalid choice. Enter 'dp', 'derivative', or 'hybrid'.")

    print(f"\nSelected zoning method: {selected_method}")

    zone_labels = prompt_zone_labels(N_ZONES)

    print("\nZone labels selected:")
    for zone, label in zone_labels.items():
        print(f"  Zone {zone}: {label}")

    features = extract_zone_features(selected_df, zone_labels)
    features["zoning_method"] = selected_method

    edges = selected_df.attrs.get("zone_edges_percent", [])
    for i, edge in enumerate(edges):
        features[f"zone_edge_{i}_percent"] = edge

    append_features_to_csv(features, OUTPUT_FEATURES_CSV)

    print("\nDone.")
    print(f"Saved labeled post-simulation features to:\n{OUTPUT_FEATURES_CSV}")


if __name__ == "__main__":
    main()
