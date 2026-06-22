import os
import numpy as np
import matplotlib.pyplot as plt
import pywt

from tkinter import Tk
from tkinter.filedialog import askopenfilename


# ==========================================
# Settings
# ==========================================

RESULTS_DIR = "results"

WAVELET = "cmor1.5-1.0"

MIN_FREQ_HZ = 1e3
MAX_FREQ_HZ = 5e5
NUM_FREQS = 100

REMOVE_DC_OFFSET = True
NORMALIZE_SIGNAL = True


# ==========================================
# File Selection
# ==========================================

def select_file(title="Select ELECTRO Export"):
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
# Load ELECTRO Export
# ==========================================

def load_electro_graph(filename):
    import re
    import numpy as np

    time = []
    value = []

    number_pattern = r"[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?"

    with open(filename, "r", errors="ignore") as f:
        for line in f:
            nums = re.findall(number_pattern, line)

            # Need at least: point, time, value
            if len(nums) >= 3:
                try:
                    point = float(nums[0])
                    t = float(nums[1])
                    y = float(nums[2])

                    # Only accept normal data rows where first number is the point index
                    if point >= 0:
                        time.append(t)
                        value.append(y)

                except ValueError:
                    pass

    return np.array(time), np.array(value)


# ==========================================
# Preprocess Signal
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
# Wavelet Analysis
# ==========================================

def run_wavelet_analysis(t, signal):
    dt = np.mean(np.diff(t))
    fs = 1 / dt
    nyquist = fs / 2

    max_freq = min(MAX_FREQ_HZ, nyquist)

    freqs = np.geomspace(MIN_FREQ_HZ, max_freq, NUM_FREQS)

    central_freq = pywt.central_frequency(WAVELET)
    scales = central_freq / (freqs * dt)

    coeffs, actual_freqs = pywt.cwt(
        signal,
        scales,
        WAVELET,
        sampling_period=dt,
        method="fft"
    )

    power = np.abs(coeffs) ** 2

    return actual_freqs, power, fs, nyquist


# ==========================================
# Plot Raw Signal
# ==========================================

def plot_raw_signal(t, signal, name):
    plt.figure(figsize=(10, 5))
    plt.plot(t, signal)
    plt.xlabel("Time (s)")
    plt.ylabel("Magnitude")
    plt.title(f"Raw ELECTRO Signal: {name}")
    plt.grid(True)
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_raw_signal.png")
    plt.savefig(save_path, dpi=300)
    plt.show()


# ==========================================
# Plot Wavelet Scalogram
# ==========================================

def plot_scalogram(t, freqs, power, name):
    plt.figure(figsize=(10, 6))
    plt.pcolormesh(t, freqs, power, shading="auto")
    plt.yscale("log")
    plt.xlabel("Time (s)")
    plt.ylabel("Frequency (Hz)")
    plt.title(f"Wavelet Scalogram: {name}")
    plt.colorbar(label="Wavelet Power")
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_wavelet_scalogram.png")
    plt.savefig(save_path, dpi=300)
    plt.show()


# ==========================================
# Plot Wavelet Energy vs Time
# ==========================================

def plot_energy_vs_time(t, power, name):
    energy_vs_time = np.sum(power, axis=0)

    plt.figure(figsize=(10, 5))
    plt.plot(t, energy_vs_time)
    plt.xlabel("Time (s)")
    plt.ylabel("Total Wavelet Energy")
    plt.title(f"Wavelet Energy vs Time: {name}")
    plt.grid(True)
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_wavelet_energy_vs_time.png")
    plt.savefig(save_path, dpi=300)
    plt.show()

    return energy_vs_time


# ==========================================
# Plot Dominant Frequency vs Time
# ==========================================

def plot_dominant_frequency(t, freqs, power, name):
    max_indices = np.argmax(power, axis=0)
    dominant_freq = freqs[max_indices]

    plt.figure(figsize=(10, 5))
    plt.plot(t, dominant_freq)
    plt.yscale("log")
    plt.xlabel("Time (s)")
    plt.ylabel("Dominant Frequency (Hz)")
    plt.title(f"Dominant Frequency vs Time: {name}")
    plt.grid(True)
    plt.tight_layout()

    save_path = os.path.join(RESULTS_DIR, f"{name}_dominant_frequency.png")
    plt.savefig(save_path, dpi=300)
    plt.show()

    return dominant_freq


# ==========================================
# Flashover-Oriented Metrics
# ==========================================

def compute_metrics(t, raw_signal, processed_signal, freqs, power):
    dt = np.mean(np.diff(t))

    derivative = np.gradient(raw_signal, dt)

    total_wavelet_energy = np.sum(power)
    max_wavelet_power = np.max(power)

    energy_vs_time = np.sum(power, axis=0)

    early_time_limit = 5e-6
    early_indices = t <= early_time_limit

    if np.any(early_indices):
        early_wavelet_energy = np.sum(energy_vs_time[early_indices])
    else:
        early_wavelet_energy = np.nan

    high_freq_limit = 0.25 * np.max(freqs)
    high_freq_indices = freqs >= high_freq_limit

    if np.any(high_freq_indices):
        high_freq_energy = np.sum(power[high_freq_indices, :])
        high_freq_ratio = high_freq_energy / total_wavelet_energy
    else:
        high_freq_energy = np.nan
        high_freq_ratio = np.nan

    metrics = {
        "Peak magnitude": np.max(np.abs(raw_signal)),
        "Maximum d(signal)/dt": np.max(np.abs(derivative)),
        "Total wavelet energy": total_wavelet_energy,
        "Maximum wavelet power": max_wavelet_power,
        "Early-time wavelet energy, 0-5 us": early_wavelet_energy,
        "High-frequency energy ratio": high_freq_ratio
    }

    return metrics


# ==========================================
# Main
# ==========================================

def main():
    os.makedirs(RESULTS_DIR, exist_ok=True)

    print("Select ELECTRO export file...")

    filename = select_file()

    if not filename:
        print("No file selected.")
        return

    name = os.path.splitext(os.path.basename(filename))[0]

    print("\nLoaded File:")
    print(filename)

    t, raw_signal = load_electro_graph(filename)

    if len(t) < 5:
        print("Error: Not enough numeric data was found.")
        print("Make sure the ELECTRO file contains rows like:")
        print("1    0.00000E+00    0.00000E+00")
        return

    processed_signal = preprocess_signal(raw_signal)

    freqs, power, fs, nyquist = run_wavelet_analysis(t, processed_signal)

    plot_raw_signal(t, raw_signal, name)
    plot_scalogram(t, freqs, power, name)
    plot_energy_vs_time(t, power, name)
    plot_dominant_frequency(t, freqs, power, name)

    metrics = compute_metrics(t, raw_signal, processed_signal, freqs, power)

    print("\nData Summary")
    print("-------------------")
    print(f"Samples: {len(t)}")
    print(f"Start time: {t[0]:.6e} s")
    print(f"End time: {t[-1]:.6e} s")
    print(f"Time step: {np.mean(np.diff(t)):.6e} s")
    print(f"Sampling frequency: {fs:.6e} Hz")
    print(f"Nyquist frequency: {nyquist:.6e} Hz")

    print("\nFlashover-Oriented Metrics")
    print("-------------------")
    for key, value in metrics.items():
        print(f"{key}: {value:.6e}")

    print("\nSaved plots to:")
    print(os.path.abspath(RESULTS_DIR))


if __name__ == "__main__":
    main()