"""
Zoned Wavelet feature extraction for ELECTRO line-graph exports.

This file keeps the original standalone behavior, but adds callable functions that
can be imported by data_extraction_tier1_with_label.py. The callable path expects
an already-parsed/zoned pandas DataFrame with columns:
field, field_curve_id, zone_id, d, d_percent, d_norm, value.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, Optional

import numpy as np
import matplotlib.pyplot as plt

try:
    import pywt
except ImportError:  # handled cleanly by feature extractor
    pywt = None

from tkinter import Tk
from tkinter.filedialog import askopenfilename


RESULTS_DIR = "results"
WAVELET = "cmor1.5-1.0"
MIN_SPATIAL_FREQ = 1e-4      # cycles/mm; low enough for long bushing regions
MAX_SPATIAL_FREQ = None      # None => Nyquist from zone spacing
NUM_FREQS = 80
REMOVE_DC_OFFSET = True
NORMALIZE_SIGNAL = True
MIN_ZONE_POINTS = 8


def _safe_divide(num, den):
    if den is None or not np.isfinite(den) or abs(den) < 1e-30:
        return np.nan
    return float(num / den)


def _sanitize_label(label: str) -> str:
    label = str(label).strip().lower()
    label = re.sub(r"[^a-z0-9]+", "_", label).strip("_")
    return label or "unknown"


def select_file(title="Select ELECTRO Export"):
    root = Tk()
    root.withdraw()
    return askopenfilename(
        title=title,
        filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*")],
    )


def load_electro_graph(filename):
    time = []
    value = []
    number_pattern = r"[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?"
    with open(filename, "r", errors="ignore") as f:
        for line in f:
            nums = re.findall(number_pattern, line)
            if len(nums) >= 3:
                try:
                    point = float(nums[0])
                    t = float(nums[1])
                    y = float(nums[2])
                    if point >= 0:
                        time.append(t)
                        value.append(y)
                except ValueError:
                    pass
    return np.array(time), np.array(value)


def preprocess_signal(signal):
    signal = np.array(signal, dtype=float)
    if REMOVE_DC_OFFSET:
        signal = signal - np.nanmean(signal)
    if NORMALIZE_SIGNAL:
        std = np.nanstd(signal)
        if std > 0:
            signal = signal / std
    return signal


def _uniform_resample(x, y, min_points: int = MIN_ZONE_POINTS):
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    valid = np.isfinite(x) & np.isfinite(y)
    x = x[valid]
    y = y[valid]
    if len(x) < min_points:
        return None, None
    order = np.argsort(x)
    x = x[order]
    y = y[order]
    x_unique, idx = np.unique(x, return_index=True)
    y_unique = y[idx]
    if len(x_unique) < min_points or np.nanmax(x_unique) == np.nanmin(x_unique):
        return None, None
    x_uniform = np.linspace(float(x_unique[0]), float(x_unique[-1]), len(x_unique))
    y_uniform = np.interp(x_uniform, x_unique, y_unique)
    return x_uniform, y_uniform


def run_wavelet_analysis(axis, signal, min_freq=MIN_SPATIAL_FREQ, max_freq=MAX_SPATIAL_FREQ, num_freqs=NUM_FREQS):
    """Generic CWT over a uniformly sampled axis. For zoned ELECTRO line data, axis is mm."""
    if pywt is None:
        raise ImportError("PyWavelets is required. Install with: py -m pip install PyWavelets")

    axis = np.asarray(axis, dtype=float)
    signal = np.asarray(signal, dtype=float)
    step = float(np.mean(np.diff(axis)))
    sample_rate = 1.0 / step
    nyquist = sample_rate / 2.0

    max_usable = nyquist if max_freq is None else min(float(max_freq), nyquist)
    min_usable = min(float(min_freq), max_usable * 0.5)
    if min_usable <= 0 or max_usable <= min_usable:
        return np.array([]), np.empty((0, len(signal))), sample_rate, nyquist

    freqs = np.geomspace(min_usable, max_usable, int(num_freqs))
    central_freq = pywt.central_frequency(WAVELET)
    scales = central_freq / (freqs * step)
    coeffs, actual_freqs = pywt.cwt(signal, scales, WAVELET, sampling_period=step, method="fft")
    power = np.abs(coeffs) ** 2
    return actual_freqs, power, sample_rate, nyquist


def compute_wavelet_metrics(axis, raw_signal, freqs, power):
    if len(axis) < 2 or power.size == 0:
        return {}
    step = float(np.mean(np.diff(axis)))
    derivative = np.gradient(raw_signal, step)
    energy_vs_axis = np.sum(power, axis=0)
    total_energy = float(np.sum(power))
    max_power = float(np.max(power))

    max_flat = int(np.argmax(power))
    max_freq_idx, max_axis_idx = np.unravel_index(max_flat, power.shape)
    dominant_freq_vs_axis = freqs[np.argmax(power, axis=0)]

    high_freq_limit = 0.25 * float(np.max(freqs))
    high_mask = freqs >= high_freq_limit
    high_energy = float(np.sum(power[high_mask, :])) if np.any(high_mask) else np.nan

    first_quarter_limit = axis[0] + 0.25 * (axis[-1] - axis[0])
    first_quarter_mask = axis <= first_quarter_limit
    first_quarter_energy = float(np.sum(energy_vs_axis[first_quarter_mask])) if np.any(first_quarter_mask) else np.nan

    return {
        "wavelet_peak_magnitude": float(np.max(np.abs(raw_signal))),
        "wavelet_max_abs_dE_dd": float(np.max(np.abs(derivative))),
        "wavelet_total_energy": total_energy,
        "wavelet_max_power": max_power,
        "wavelet_energy_density": _safe_divide(total_energy, axis[-1] - axis[0]),
        "wavelet_high_freq_energy_ratio": _safe_divide(high_energy, total_energy),
        "wavelet_first_quarter_energy_ratio": _safe_divide(first_quarter_energy, total_energy),
        "wavelet_peak_power_spatial_freq_cyc_per_mm": float(freqs[max_freq_idx]),
        "wavelet_peak_power_d_percent": float(axis[max_axis_idx]),
        "wavelet_dominant_spatial_freq_median": float(np.nanmedian(dominant_freq_vs_axis)),
    }


def extract_zoned_wavelet_features(df, zone_labels: Optional[dict] = None, field: str = "E", include_per_curve: bool = False) -> Dict[str, float]:
    """Return wavelet metrics for each zone. Uses d_percent as the spatial axis for zoning consistency."""
    features: Dict[str, float] = {}
    if pywt is None:
        features["dsp_wavelet_available"] = 0
        return features

    e_df = df[df["field"] == field].copy()
    if e_df.empty:
        features["dsp_wavelet_available"] = 0
        return features
    features["dsp_wavelet_available"] = 1

    global_energy_by_zone = {}

    for zone_id in sorted(e_df["zone_id"].unique()):
        label = _sanitize_label(zone_labels.get(int(zone_id), f"zone{zone_id}") if zone_labels else f"zone{zone_id}")
        zone_df = e_df[e_df["zone_id"] == zone_id].sort_values("d_percent")
        prefix = f"zone{int(zone_id)}_{label}_E"

        x, y = _uniform_resample(zone_df["d_percent"].to_numpy(), zone_df["value"].to_numpy())
        if x is not None:
            y_proc = preprocess_signal(y)
            freqs, power, _, _ = run_wavelet_analysis(x, y_proc)
            metrics = compute_wavelet_metrics(x, y, freqs, power)
            for k, v in metrics.items():
                features[f"{prefix}_{k}"] = v
            global_energy_by_zone[int(zone_id)] = metrics.get("wavelet_total_energy", np.nan)

        if include_per_curve:
            for curve_id in sorted(zone_df["field_curve_id"].unique()):
                curve_zone = zone_df[zone_df["field_curve_id"] == curve_id].sort_values("d_percent")
                x, y = _uniform_resample(curve_zone["d_percent"].to_numpy(), curve_zone["value"].to_numpy())
                if x is None:
                    continue
                y_proc = preprocess_signal(y)
                freqs, power, _, _ = run_wavelet_analysis(x, y_proc)
                metrics = compute_wavelet_metrics(x, y, freqs, power)
                cprefix = f"E_curve{int(curve_id)}_zone{int(zone_id)}_{label}"
                for k, v in metrics.items():
                    features[f"{cprefix}_{k}"] = v

    total_global_energy = np.nansum(list(global_energy_by_zone.values())) if global_energy_by_zone else np.nan
    for zone_id, energy in global_energy_by_zone.items():
        label = _sanitize_label(zone_labels.get(int(zone_id), f"zone{zone_id}") if zone_labels else f"zone{zone_id}")
        features[f"zone{int(zone_id)}_{label}_E_wavelet_energy_over_all_zones"] = _safe_divide(energy, total_global_energy)

    return features


# Standalone plotting helpers retained from the original workflow.
def main():
    os.makedirs(RESULTS_DIR, exist_ok=True)
    filename = select_file()
    if not filename:
        print("No file selected.")
        return
    name = Path(filename).stem
    t, raw_signal = load_electro_graph(filename)
    if len(t) < 5:
        print("Error: Not enough numeric data was found.")
        return
    processed_signal = preprocess_signal(raw_signal)
    freqs, power, fs, nyquist = run_wavelet_analysis(t, processed_signal, min_freq=1e3, max_freq=5e5)
    metrics = compute_wavelet_metrics(t, raw_signal, freqs, power)
    print(f"\nSamples: {len(t)}")
    print(f"Sampling frequency: {fs:.6e}")
    print(f"Nyquist frequency: {nyquist:.6e}")
    for key, value in metrics.items():
        print(f"{key}: {value:.6e}")


if __name__ == "__main__":
    main()
