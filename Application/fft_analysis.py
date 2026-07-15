"""
Zoned spatial FFT/PSD feature extraction for ELECTRO line-graph exports.

This file keeps standalone FFT behavior while adding callable functions that can
be imported by data_extraction_tier1_with_label.py.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, Optional

import numpy as np
import matplotlib.pyplot as plt

from tkinter import Tk
from tkinter.filedialog import askopenfilename


RESULTS_DIR = "fft_results"
REMOVE_DC_OFFSET = True
NORMALIZE_SIGNAL = False
APPLY_HANN_WINDOW = True
MIN_ZONE_POINTS = 8


def _safe_divide(num, den):
    if den is None or not np.isfinite(den) or abs(den) < 1e-30:
        return np.nan
    return float(num / den)


def _sanitize_label(label: str) -> str:
    label = str(label).strip().lower()
    label = re.sub(r"[^a-z0-9]+", "_", label).strip("_")
    return label or "unknown"


def select_file(title="Select ELECTRO Spatial Export"):
    root = Tk()
    root.withdraw()
    return askopenfilename(
        title=title,
        filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*")],
    )


def load_electro_spatial_graph(filename):
    distance = []
    signal = []
    number_pattern = r"[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?"
    with open(filename, "r", errors="ignore") as f:
        for line in f:
            nums = re.findall(number_pattern, line)
            if len(nums) >= 6:
                try:
                    point = float(nums[0])
                    d_mm = float(nums[1])
                    signal_value = float(nums[-1])
                    if point >= 0:
                        distance.append(d_mm)
                        signal.append(signal_value)
                except ValueError:
                    pass
    return np.array(distance), np.array(signal)


def preprocess_signal(signal):
    signal = np.array(signal, dtype=float)
    if REMOVE_DC_OFFSET:
        signal = signal - np.nanmean(signal)
    if NORMALIZE_SIGNAL:
        std = np.nanstd(signal)
        if std > 0:
            signal = signal / std
    return signal


def resample_uniform_distance(distance, signal, min_points: int = MIN_ZONE_POINTS):
    distance = np.asarray(distance, dtype=float)
    signal = np.asarray(signal, dtype=float)
    valid = np.isfinite(distance) & np.isfinite(signal)
    distance = distance[valid]
    signal = signal[valid]
    if len(distance) < min_points:
        return None, None
    sort_idx = np.argsort(distance)
    distance = distance[sort_idx]
    signal = signal[sort_idx]
    unique_distance, unique_indices = np.unique(distance, return_index=True)
    unique_signal = signal[unique_indices]
    if len(unique_distance) < min_points or unique_distance[-1] == unique_distance[0]:
        return None, None
    uniform_distance = np.linspace(unique_distance[0], unique_distance[-1], len(unique_distance))
    uniform_signal = np.interp(uniform_distance, unique_distance, unique_signal)
    return uniform_distance, uniform_signal


def compute_spatial_derivative(distance_mm, signal):
    return np.gradient(signal, distance_mm)


def run_spatial_fft_and_psd(distance_mm, signal):
    dx_mm = float(np.mean(np.diff(distance_mm)))
    if not np.isfinite(dx_mm) or dx_mm <= 0:
        return np.array([]), np.array([]), np.array([]), np.array([]), np.nan

    signal_proc = preprocess_signal(signal)
    if APPLY_HANN_WINDOW and len(signal_proc) > 2:
        signal_proc = signal_proc * np.hanning(len(signal_proc))

    fft_values = np.fft.rfft(signal_proc)
    fft_magnitude = np.abs(fft_values)
    psd = (np.abs(fft_values) ** 2) / (len(signal_proc) * dx_mm)
    spatial_freq = np.fft.rfftfreq(len(signal_proc), d=dx_mm)
    wavelength_mm = np.full_like(spatial_freq, np.inf)
    nonzero = spatial_freq > 0
    wavelength_mm[nonzero] = 1 / spatial_freq[nonzero]
    return spatial_freq, wavelength_mm, fft_magnitude, psd, dx_mm


def compute_fft_metrics(distance_mm, signal, spatial_freq, wavelength_mm, fft_magnitude, psd, dx_mm):
    valid = spatial_freq > 0
    if len(distance_mm) < MIN_ZONE_POINTS or not np.any(valid):
        return {}

    valid_indices = np.where(valid)[0]
    true_index = valid_indices[int(np.argmax(fft_magnitude[valid]))]
    total_psd_energy = float(np.sum(psd[valid]))

    high_freq_limit = 0.25 * float(np.max(spatial_freq))
    high_mask = spatial_freq >= high_freq_limit
    high_energy = float(np.sum(psd[high_mask])) if np.any(high_mask) else np.nan

    low_mid_limit = 0.10 * float(np.max(spatial_freq))
    low_mid_mask = (spatial_freq > 0) & (spatial_freq <= low_mid_limit)
    low_mid_energy = float(np.sum(psd[low_mid_mask])) if np.any(low_mid_mask) else np.nan

    derivative = compute_spatial_derivative(distance_mm, signal)
    _, _, derivative_fft, derivative_psd, _ = run_spatial_fft_and_psd(distance_mm, derivative)
    derivative_total = float(np.sum(derivative_psd[valid])) if len(derivative_psd) == len(valid) else np.nan
    derivative_high = float(np.sum(derivative_psd[high_mask])) if len(derivative_psd) == len(high_mask) and np.any(high_mask) else np.nan

    return {
        "fft_peak_magnitude": float(np.max(np.abs(signal))),
        "fft_min_magnitude": float(np.min(signal)),
        "fft_max_abs_dE_dd": float(np.max(np.abs(derivative))),
        "fft_dominant_spatial_freq_cyc_per_mm": float(spatial_freq[true_index]),
        "fft_dominant_wavelength_percent": float(wavelength_mm[true_index]),
        "fft_max_magnitude": float(fft_magnitude[true_index]),
        "fft_total_psd_energy": total_psd_energy,
        "fft_psd_energy_density": _safe_divide(total_psd_energy, distance_mm[-1] - distance_mm[0]),
        "fft_high_freq_psd_ratio": _safe_divide(high_energy, total_psd_energy),
        "fft_low_mid_freq_psd_ratio": _safe_divide(low_mid_energy, total_psd_energy),
        "fft_derivative_total_psd_energy": derivative_total,
        "fft_derivative_high_freq_psd_ratio": _safe_divide(derivative_high, derivative_total),
        "fft_dx_percent": float(dx_mm),
    }


def extract_zoned_fft_features(df, zone_labels: Optional[dict] = None, field: str = "E", include_per_curve: bool = False) -> Dict[str, float]:
    """Return FFT/PSD metrics for each zone using d_percent as the spatial axis."""
    features: Dict[str, float] = {}
    e_df = df[df["field"] == field].copy()
    if e_df.empty:
        features["dsp_fft_available"] = 0
        return features
    features["dsp_fft_available"] = 1

    psd_energy_by_zone = {}

    for zone_id in sorted(e_df["zone_id"].unique()):
        label = _sanitize_label(zone_labels.get(int(zone_id), f"zone{zone_id}") if zone_labels else f"zone{zone_id}")
        zone_df = e_df[e_df["zone_id"] == zone_id].sort_values("d_percent")
        prefix = f"zone{int(zone_id)}_{label}_E"

        x, y = resample_uniform_distance(zone_df["d_percent"].to_numpy(), zone_df["value"].to_numpy())
        if x is not None:
            spatial_freq, wavelength, fft_mag, psd, dx = run_spatial_fft_and_psd(x, y)
            metrics = compute_fft_metrics(x, y, spatial_freq, wavelength, fft_mag, psd, dx)
            for k, v in metrics.items():
                features[f"{prefix}_{k}"] = v
            psd_energy_by_zone[int(zone_id)] = metrics.get("fft_total_psd_energy", np.nan)

        if include_per_curve:
            for curve_id in sorted(zone_df["field_curve_id"].unique()):
                curve_zone = zone_df[zone_df["field_curve_id"] == curve_id].sort_values("d_percent")
                x, y = resample_uniform_distance(curve_zone["d_percent"].to_numpy(), curve_zone["value"].to_numpy())
                if x is None:
                    continue
                spatial_freq, wavelength, fft_mag, psd, dx = run_spatial_fft_and_psd(x, y)
                metrics = compute_fft_metrics(x, y, spatial_freq, wavelength, fft_mag, psd, dx)
                cprefix = f"E_curve{int(curve_id)}_zone{int(zone_id)}_{label}"
                for k, v in metrics.items():
                    features[f"{cprefix}_{k}"] = v

    total_psd_energy = np.nansum(list(psd_energy_by_zone.values())) if psd_energy_by_zone else np.nan
    for zone_id, energy in psd_energy_by_zone.items():
        label = _sanitize_label(zone_labels.get(int(zone_id), f"zone{zone_id}") if zone_labels else f"zone{zone_id}")
        features[f"zone{int(zone_id)}_{label}_E_fft_psd_energy_over_all_zones"] = _safe_divide(energy, total_psd_energy)

    return features


# Standalone behavior retained from the original script.
def main():
    filename = select_file()
    if not filename:
        print("No file selected.")
        return
    distance_raw, signal_raw = load_electro_spatial_graph(filename)
    if len(distance_raw) < MIN_ZONE_POINTS:
        print("Error: Not enough spatial numeric data was found.")
        return
    distance_mm, signal_uniform = resample_uniform_distance(distance_raw, signal_raw)
    spatial_freq, wavelength, fft_mag, psd, dx = run_spatial_fft_and_psd(distance_mm, signal_uniform)
    metrics = compute_fft_metrics(distance_mm, signal_uniform, spatial_freq, wavelength, fft_mag, psd, dx)
    for key, value in metrics.items():
        print(f"{key}: {value:.6e}")


if __name__ == "__main__":
    main()
