import os
import re
import numpy as np
import matplotlib.pyplot as plt

from tkinter import Tk
from tkinter.filedialog import askopenfilename


# ==========================================
# Settings
# ==========================================

RESULTS_DIR = "fft_results"

REMOVE_DC_OFFSET = True
NORMALIZE_SIGNAL = False
APPLY_HANN_WINDOW = True


# ==========================================
# File Selection
# ==========================================

def select_file(title="Select ELECTRO Spatial Export"):
    root = Tk()
    root.withdraw()

    filename = askopenfilename(
        title=title,
        filetypes=[
            ("CSV Files", "*.csv"),
            ("Text Files", "*.txt"),
            ("All Files", "*.*")
        ]
    )

    return filename


# ==========================================
# Load ELECTRO Spatial Export
# Expected data rows:
# Point, d(mm), x(mm), y(mm), z(mm), Re{Em}(kV/mm)
# Uses the last numeric column as the signal value.
# ==========================================

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


# ==========================================
# Preprocess Spatial Signal
# ==========================================

def preprocess_signal(signal):
    signal = np.array(signal, dtype=float)

    if REMOVE_DC_OFFSET:
        signal = signal - np.mean(signal)

    if NORMALIZE_SIGNAL:
        std = np.std(signal)
        if std != 0:
            signal = signal / std

    return signal


# ==========================================
# Resample to Uniform Distance Spacing
# FFT requires uniform spacing
# ==========================================

def resample_uniform_distance(distance, signal):
    sort_idx = np.argsort(distance)
    distance = distance[sort_idx]
    signal = signal[sort_idx]

    unique_distance, unique_indices = np.unique(distance, return_index=True)
    unique_signal = signal[unique_indices]

    num_points = len(unique_distance)

    uniform_distance = np.linspace(
        unique_distance[0],
        unique_distance[-1],
        num_points
    )

    uniform_signal = np.interp(
        uniform_distance,
        unique_distance,
        unique_signal
    )

    return uniform_distance, uniform_signal


# ==========================================
# Spatial Derivative dE/dx
# ==========================================

def compute_spatial_derivative(distance_mm, signal):
    derivative = np.gradient(signal, distance_mm)
    return derivative


# ==========================================
# FFT + PSD Analysis
# ==========================================

def run_spatial_fft_and_psd(distance_mm, signal):
    dx_mm = np.mean(np.diff(distance_mm))

    signal_proc = preprocess_signal(signal)

    if APPLY_HANN_WINDOW:
        window = np.hanning(len(signal_proc))
        signal_proc = signal_proc * window

    fft_values = np.fft.rfft(signal_proc)
    fft_magnitude = np.abs(fft_values)

    # PSD estimate. Units are approximately signal^2 per cycles/mm.
    psd = (np.abs(fft_values) ** 2) / (len(signal_proc) * dx_mm)

    spatial_freq_cycles_per_mm = np.fft.rfftfreq(
        len(signal_proc),
        d=dx_mm
    )

    wavelength_mm = np.full_like(spatial_freq_cycles_per_mm, np.inf)
    nonzero = spatial_freq_cycles_per_mm > 0
    wavelength_mm[nonzero] = 1 / spatial_freq_cycles_per_mm[nonzero]

    return spatial_freq_cycles_per_mm, wavelength_mm, fft_magnitude, psd, dx_mm


# ==========================================
# Plot Spatial Signal
# ==========================================

def plot_spatial_signal(distance_mm, signal, name, ylabel, title_suffix, filename_suffix):
    plt.figure(figsize=(10, 5))
    plt.plot(distance_mm, signal)
    plt.xlabel("Distance from Origin (mm)")
    plt.ylabel(ylabel)
    plt.title(f"{title_suffix}: {name}")
    plt.grid(True)
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_{filename_suffix}.png")
    plt.savefig(save_path, dpi=300)
    plt.show()


# ==========================================
# Plot FFT Magnitude
# ==========================================

def plot_fft(spatial_freq, fft_magnitude, name, title_suffix, filename_suffix):
    plt.figure(figsize=(10, 5))
    plt.plot(spatial_freq, fft_magnitude)
    plt.xlabel("Spatial Frequency (cycles/mm)")
    plt.ylabel("FFT Magnitude")
    plt.title(f"{title_suffix}: {name}")
    plt.grid(True)
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_{filename_suffix}.png")
    plt.savefig(save_path, dpi=300)
    plt.show()


# ==========================================
# Plot PSD
# ==========================================

def plot_psd(spatial_freq, psd, name, title_suffix, filename_suffix):
    plt.figure(figsize=(10, 5))
    plt.semilogy(spatial_freq, psd)
    plt.xlabel("Spatial Frequency (cycles/mm)")
    plt.ylabel("PSD")
    plt.title(f"{title_suffix}: {name}")
    plt.grid(True)
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_{filename_suffix}.png")
    plt.savefig(save_path, dpi=300)
    plt.show()


# ==========================================
# Plot FFT vs Wavelength
# ==========================================

def plot_fft_wavelength(wavelength_mm, fft_magnitude, name, title_suffix, filename_suffix):
    valid = np.isfinite(wavelength_mm)

    plt.figure(figsize=(10, 5))
    plt.plot(wavelength_mm[valid], fft_magnitude[valid])
    plt.xlabel("Spatial Wavelength (mm)")
    plt.ylabel("FFT Magnitude")
    plt.title(f"{title_suffix}: {name}")
    plt.grid(True)
    plt.tight_layout()

    plt.gca().invert_xaxis()

    save_path = os.path.join(RESULTS_DIR, f"{name}_{filename_suffix}.png")
    plt.savefig(save_path, dpi=300)
    plt.show()


# ==========================================
# Print Metrics
# ==========================================

def print_fft_metrics(label, distance_mm, signal, spatial_freq, wavelength_mm, fft_magnitude, psd, dx_mm):
    valid = spatial_freq > 0

    dominant_index = np.argmax(fft_magnitude[valid])
    valid_indices = np.where(valid)[0]
    true_index = valid_indices[dominant_index]

    dominant_freq = spatial_freq[true_index]
    dominant_wavelength = wavelength_mm[true_index]

    total_psd_energy = np.sum(psd[valid])

    high_freq_limit = 0.25 * np.max(spatial_freq)
    high_freq_indices = spatial_freq >= high_freq_limit

    if np.any(high_freq_indices):
        high_freq_psd_energy = np.sum(psd[high_freq_indices])
        high_freq_psd_ratio = high_freq_psd_energy / np.sum(psd[valid])
    else:
        high_freq_psd_ratio = np.nan

    print(f"\n{label} Spatial Data Summary")
    print("-------------------")
    print(f"Samples: {len(distance_mm)}")
    print(f"Start distance: {distance_mm[0]:.6e} mm")
    print(f"End distance: {distance_mm[-1]:.6e} mm")
    print(f"Average spacing dx: {dx_mm:.6e} mm")
    print(f"Peak magnitude: {np.max(signal):.6e}")
    print(f"Minimum magnitude: {np.min(signal):.6e}")
    print(f"Maximum absolute value: {np.max(np.abs(signal)):.6e}")

    print(f"\n{label} FFT / PSD Metrics")
    print("-------------------")
    print(f"Dominant spatial frequency: {dominant_freq:.6e} cycles/mm")
    print(f"Dominant spatial wavelength: {dominant_wavelength:.6e} mm")
    print(f"Maximum FFT magnitude: {fft_magnitude[true_index]:.6e}")
    print(f"Total PSD energy: {total_psd_energy:.6e}")
    print(f"High-frequency PSD ratio: {high_freq_psd_ratio:.6e}")


# ==========================================
# Main
# ==========================================

def main():
    os.makedirs(RESULTS_DIR, exist_ok=True)

    print("Select ELECTRO spatial export...")

    filename = select_file()

    if not filename:
        print("No file selected.")
        return

    name = os.path.splitext(os.path.basename(filename))[0]

    print("\nLoaded File:")
    print(filename)

    distance_raw, signal_raw = load_electro_spatial_graph(filename)

    if len(distance_raw) < 5:
        print("Error: Not enough spatial numeric data was found.")
        print("Expected rows similar to:")
        print("1    0.00000E+00    x    y    z    1.23456E+00")
        return

    distance_mm, signal_uniform = resample_uniform_distance(distance_raw, signal_raw)

    derivative_uniform = compute_spatial_derivative(distance_mm, signal_uniform)

    spatial_freq, wavelength_mm, fft_magnitude, psd, dx_mm = run_spatial_fft_and_psd(
        distance_mm,
        signal_uniform
    )

    derivative_freq, derivative_wavelength, derivative_fft, derivative_psd, derivative_dx = run_spatial_fft_and_psd(
        distance_mm,
        derivative_uniform
    )

    plot_spatial_signal(
        distance_mm,
        signal_uniform,
        name,
        "Signal Magnitude",
        "Spatial Signal",
        "spatial_signal"
    )

    plot_fft(
        spatial_freq,
        fft_magnitude,
        name,
        "Spatial FFT",
        "spatial_fft"
    )

    plot_psd(
        spatial_freq,
        psd,
        name,
        "Spatial PSD",
        "spatial_psd"
    )

    plot_fft_wavelength(
        wavelength_mm,
        fft_magnitude,
        name,
        "Spatial Wavelength Spectrum",
        "spatial_wavelength_spectrum"
    )

    plot_spatial_signal(
        distance_mm,
        derivative_uniform,
        name,
        "d(signal)/dx",
        "Spatial Derivative",
        "spatial_derivative"
    )

    plot_fft(
        derivative_freq,
        derivative_fft,
        name,
        "FFT of Spatial Derivative",
        "derivative_fft"
    )

    plot_psd(
        derivative_freq,
        derivative_psd,
        name,
        "PSD of Spatial Derivative",
        "derivative_psd"
    )

    print_fft_metrics(
        "Original Signal",
        distance_mm,
        signal_uniform,
        spatial_freq,
        wavelength_mm,
        fft_magnitude,
        psd,
        dx_mm
    )

    print_fft_metrics(
        "Spatial Derivative",
        distance_mm,
        derivative_uniform,
        derivative_freq,
        derivative_wavelength,
        derivative_fft,
        derivative_psd,
        derivative_dx
    )

    print("\nSaved plots to:")
    print(os.path.abspath(RESULTS_DIR))


if __name__ == "__main__":
    main()