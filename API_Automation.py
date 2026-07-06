from pathlib import Path

import win32com.client
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

try:
    import pywt
    PYWT_AVAILABLE = True
except ImportError:
    PYWT_AVAILABLE = False


EPOXY_MATERIAL = "Standard molded Epoxy @60Hz"
EPOXY_OUTSIDE_OBJECT_NAME = "epoxy outside"


def check_return(name, result):
    print(f"{name} returned: {result}")
    if result != 0:
        print(f"WARNING: {name} returned nonzero code.")


def get_current_model_path(ies):
    current_model, err = ies.File_GetModelPath("", 0)

    print("Current model:", current_model)
    print("Error code:", err)

    if err != 0:
        raise RuntimeError("Could not determine current model path.")

    if current_model == "Untitled":
        raise RuntimeError("No saved model is currently open in ELECTRO.")

    return Path(current_model)


def apply_safety_settings(ies):
    print("\nApplying safety settings...")

    try:
        ies.Window_SetRefresh_OFF()
        ies.Window_SetUndo_OFF()
        ies.System_SetNumberThreadsUse(4, 0)
    except Exception as e:
        print("WARNING: Some safety settings failed.")
        print(e)


def restore_interface(ies):
    try:
        ies.Window_SetRefresh_ON()
        ies.Window_SetUndo_ON()
        ies.Window_Refresh()
    except Exception:
        pass


def geometry_validity_check(ies, epoxy_x, epoxy_y):
    print("\n--- Geometry Validity Check ---")

    region_id = ies.Geometry2D_GetRegion_FromPoint(epoxy_x, epoxy_y, 0)
    print("Epoxy shell region ID:", region_id)

    if region_id <= 0:
        print("FAILED: Invalid epoxy shell region.")
        return False

    test_voltage = ies.Physics_Set2DVoltage("conductor", 1.0, 0)
    print("Conductor voltage test returned:", test_voltage)

    if test_voltage != 0:
        print("FAILED: Conductor voltage assignment failed.")
        return False

    print("Geometry validity check passed.")
    return True


def pre_simulation_setup(ies):
    get_current_model_path(ies)

    x_shift = float(input("Enter horizontal movement amount for GND: "))
    voltage = float(input("Enter voltage to apply to conductor: "))

    epoxy_x = float(input("Enter x-coordinate inside epoxy shell: "))
    epoxy_y = float(input("Enter y-coordinate inside epoxy shell: "))

    apply_safety_settings(ies)

    try:
        move_result = ies.Geometry2D_Displace("GND", x_shift, 0.0, 0)
        check_return("Geometry2D_Displace", move_result)

        voltage_result = ies.Physics_Set2DVoltage("conductor", voltage, 0)
        check_return("Physics_Set2DVoltage", voltage_result)

        region_id = ies.Geometry2D_GetRegion_FromPoint(epoxy_x, epoxy_y, 0)
        print("Epoxy shell region ID:", region_id)

        create_result = ies.Object_Create("epoxy shell", 0, 0)
        check_return("Object_Create epoxy shell", create_result)

        add_result = ies.Object_AddRegion("epoxy shell", region_id, 0)
        check_return("Object_AddRegion epoxy shell", add_result)

        material_result = ies.Physics_SetMaterial(
            "epoxy shell",
            EPOXY_MATERIAL,
            0
        )
        check_return("Physics_SetMaterial epoxy shell", material_result)

        valid = geometry_validity_check(ies, epoxy_x, epoxy_y)

        if not valid:
            print("Geometry validity check failed. Review model before simulating.")
            return

        print("\nPre-simulation setup complete.")
        print("You may now run the simulation manually in ELECTRO.")

    finally:
        restore_interface(ies)


def extract_e_field_from_epoxy_outside_segment(ies, output_dir):
    print("\n--- Direct API E-field Extraction ---")
    print(f"Target object label: {EPOXY_OUTSIDE_OBJECT_NAME}")
    print("The API requires a segment ID, so enter a point on the epoxy outside segment.")

    x_on_seg = float(input("Enter x-coordinate ON epoxy outside segment: "))
    y_on_seg = float(input("Enter y-coordinate ON epoxy outside segment: "))

    side_choice = input(
        "Enter side to sample field on [1 = left, 2 = right, 3 = both]: "
    ).strip()

    if side_choice not in ["1", "2", "3"]:
        raise ValueError("Invalid side selection. Use 1, 2, or 3.")

    num_points = int(input("Enter number of sample points along segment, e.g. 500: "))

    seg_result = ies.Geometry2D_GetSegment_FromPoint(x_on_seg, y_on_seg, 0)
    print("Raw segment result:", seg_result)

    if isinstance(seg_result, tuple):
        seg_id = seg_result[0]
        seg_err = seg_result[-1]
    else:
        seg_id = seg_result
        seg_err = 0

    print("Detected epoxy outside segment ID:", seg_id)
    print("Segment detection error code:", seg_err)

    if seg_err != 0 or seg_id <= 0:
        raise RuntimeError("Could not detect a valid segment ID from the point entered.")

    def extract_side(side):
        result = ies.Analysis_Get2DElectricField_FromSegment(
            seg_id,
            side,
            num_points,
            [],
            [],
            [],
            [],
            [],
            [],
            0
        )

        print(f"Raw API return for side {side}:", result)

        if not isinstance(result, tuple) or len(result) < 7:
            raise RuntimeError(
                f"Unexpected return format from Analysis_Get2DElectricField_FromSegment for side {side}."
            )

        err = result[-1]
        print(f"Extraction error code for side {side}:", err)

        if err != 0:
            raise RuntimeError(
                f"E-field extraction failed for side {side}. "
                "Make sure the model is solved before running option 2."
            )

        df = pd.DataFrame({
            "x": np.array(result[0], dtype=float),
            "y": np.array(result[1], dtype=float),
            "distance": np.array(result[2], dtype=float),
            "E_tangential": np.array(result[3], dtype=float),
            "E_normal": np.array(result[4], dtype=float),
            "E_magnitude": np.array(result[5], dtype=float)
        })

        side_label = "left" if side == 1 else "right"
        output_csv = output_dir / f"epoxy_outside_direct_E_field_{side_label}.csv"
        df.to_csv(output_csv, index=False)

        print(f"Direct E-field data saved for {side_label} side:", output_csv)

        return output_csv

    if side_choice == "1":
        return [extract_side(1)]

    if side_choice == "2":
        return [extract_side(2)]

    return [extract_side(1), extract_side(2)]


def run_fft_and_wavelet(csv_path, output_dir, signal_column):
    df = pd.read_csv(csv_path)

    x = df["distance"].to_numpy(dtype=float)
    y = df[signal_column].to_numpy(dtype=float)

    x_coord = df["x"].to_numpy(dtype=float)
    y_coord = df["y"].to_numpy(dtype=float)

    y_centered = y - np.mean(y)
    dx = np.mean(np.diff(x))

    # -----------------------------
    # Spatial FFT
    # -----------------------------

    freq = np.fft.rfftfreq(len(y_centered), d=dx)
    mag = np.abs(np.fft.rfft(y_centered))

    signal_name = f"{csv_path.stem}_{signal_column}"

    fft_df = pd.DataFrame({
        "spatial_frequency": freq,
        "fft_magnitude": mag
    })

    fft_csv = output_dir / f"{signal_name}_spatial_fft.csv"
    fft_df.to_csv(fft_csv, index=False)

    # Ignore DC / zero-frequency component when finding dominant frequency
    if len(freq) > 1:
        nonzero_freq = freq[1:]
        nonzero_mag = mag[1:]

        dominant_index = np.argmax(nonzero_mag)

        dominant_spatial_frequency = nonzero_freq[dominant_index]
        dominant_fft_magnitude = nonzero_mag[dominant_index]

        if dominant_spatial_frequency != 0:
            dominant_wavelength = 1.0 / dominant_spatial_frequency
        else:
            dominant_wavelength = np.inf
    else:
        dominant_spatial_frequency = np.nan
        dominant_fft_magnitude = np.nan
        dominant_wavelength = np.nan

    summary_df = pd.DataFrame([{
        "signal": signal_name,
        "dominant_spatial_frequency": dominant_spatial_frequency,
        "dominant_wavelength": dominant_wavelength,
        "dominant_fft_magnitude": dominant_fft_magnitude,
        "distance_units": "same as ELECTRO model units",
        "frequency_units": "cycles per ELECTRO length unit"
    }])

    summary_csv = output_dir / f"{signal_name}_fft_summary.csv"
    summary_df.to_csv(summary_csv, index=False)

    print("\n--- Spatial FFT Summary ---")
    print("Signal:", signal_name)
    print("Dominant spatial frequency:", dominant_spatial_frequency)
    print("Dominant wavelength:", dominant_wavelength)
    print("Dominant FFT magnitude:", dominant_fft_magnitude)
    print("FFT summary saved to:", summary_csv)

    plt.figure()
    plt.plot(freq, mag)
    plt.xlabel("Spatial frequency")
    plt.ylabel("FFT magnitude")
    plt.title(f"{signal_name} Spatial FFT")
    plt.grid(True)

    if np.isfinite(dominant_spatial_frequency):
        plt.axvline(dominant_spatial_frequency, linestyle="--")
        plt.text(
            dominant_spatial_frequency,
            dominant_fft_magnitude,
            f"λ={dominant_wavelength:.3g}",
            rotation=90,
            verticalalignment="bottom"
        )

    plt.savefig(output_dir / f"{signal_name}_spatial_fft.png", dpi=300)
    plt.close()

    # -----------------------------
    # Wavelet Analysis + Hotspot Table
    # -----------------------------

        # -----------------------------
    # Wavelet Analysis + Hotspot Table
    # -----------------------------

    if PYWT_AVAILABLE:

        # Smaller max scale = more localized features.
        # Increase this if you want broader geometry trends.
        min_scale = 1
        max_scale = 40

        scales = np.arange(min_scale, max_scale + 1)

        coef, freqs = pywt.cwt(
            y_centered,
            scales,
            "morl",
            sampling_period=dx
        )

        wavelet_mag = np.abs(coef)

        plt.figure()
        plt.imshow(
            wavelet_mag,
            aspect="auto",
            origin="lower",
            extent=[x.min(), x.max(), scales.min(), scales.max()]
        )
        plt.xlabel("Distance")
        plt.ylabel("Wavelet scale")
        plt.title(f"{signal_name} Wavelet Magnitude")
        plt.colorbar(label="Magnitude")
        plt.savefig(output_dir / f"{signal_name}_wavelet.png", dpi=300)
        plt.close()

        # Maximum wavelet magnitude at each physical location,
        # restricted to the selected scale range
        max_over_scales = wavelet_mag.max(axis=0)

        # Scale where that maximum magnitude occurs
        scale_index_at_max = wavelet_mag.argmax(axis=0)
        scale_at_max = scales[scale_index_at_max]

        hotspot_df = pd.DataFrame({
            "distance": x,
            "x_coordinate": x_coord,
            "y_coordinate": y_coord,
            "signal_value": y,
            "max_wavelet_magnitude": max_over_scales,
            "wavelet_scale_at_max": scale_at_max
        })

        # Sort by greatest wavelet magnitude, NOT greatest scale
        hotspot_df = hotspot_df.sort_values(
            by="max_wavelet_magnitude",
            ascending=False
        )

        # Avoid selecting adjacent duplicate points from the same hotspot
        selected_rows = []
        min_distance_separation = 0.02 * (x.max() - x.min())

        for _, row in hotspot_df.iterrows():
            if len(selected_rows) >= 10:
                break

            candidate_distance = row["distance"]

            too_close = False
            for selected in selected_rows:
                if abs(candidate_distance - selected["distance"]) < min_distance_separation:
                    too_close = True
                    break

            if not too_close:
                selected_rows.append(row)

        top_hotspots_df = pd.DataFrame(selected_rows)

        hotspot_csv = output_dir / f"{signal_name}_wavelet_hotspots.csv"
        top_hotspots_df.to_csv(hotspot_csv, index=False)

        print("\n--- Wavelet Hotspot Table ---")
        print(top_hotspots_df)
        print("Wavelet hotspot table saved to:", hotspot_csv)

    else:
        print("PyWavelets not installed. Wavelet analysis skipped.")

    print(f"DSP complete for {signal_name}.")


def plot_combined_original_e_magnitude(e_csv_files, output_dir):
    plt.figure()

    for csv_path in e_csv_files:
        df = pd.read_csv(csv_path)

        distance = df["distance"].to_numpy(dtype=float)
        e_mag = df["E_magnitude"].to_numpy(dtype=float)

        label = csv_path.stem.replace("epoxy_outside_direct_E_field_", "")

        plt.plot(distance, e_mag, label=f"E magnitude {label}")

    plt.xlabel("Distance")
    plt.ylabel("E magnitude")
    plt.title("Original E Magnitude vs Distance")
    plt.grid(True)
    plt.legend()

    output_png = output_dir / "combined_original_E_magnitude_vs_distance.png"
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("Combined original E vs d plot saved to:", output_png)
    

def post_simulation_direct_api_dsp(ies):
    model_path = get_current_model_path(ies)

    run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = model_path.parent / "API_DSP_Output" / f"run_{run_stamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    e_csv_files = extract_e_field_from_epoxy_outside_segment(ies, output_dir)

    plot_combined_original_e_magnitude(e_csv_files, output_dir)

    for e_csv in e_csv_files:
        run_fft_and_wavelet(e_csv, output_dir, "E_magnitude")
        run_fft_and_wavelet(e_csv, output_dir, "E_normal")
        run_fft_and_wavelet(e_csv, output_dir, "E_tangential")

    print("\nPost-simulation direct API DSP analysis complete.")
    print("Results saved to:", output_dir)


def main():
    ies = win32com.client.Dispatch("IES.Document")
    print("Connected to ELECTRO.")

    print("\nSelect mode:")
    print("1. Pre-simulation setup")
    print("2. Post-simulation direct API DSP analysis")

    choice = input("Enter 1 or 2: ").strip()

    if choice == "1":
        pre_simulation_setup(ies)

    elif choice == "2":
        post_simulation_direct_api_dsp(ies)

    else:
        print("Invalid selection.")


if __name__ == "__main__":
    main()