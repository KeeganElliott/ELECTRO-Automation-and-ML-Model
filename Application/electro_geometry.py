# electro_geometry_with_physical_deltas.py
# Requires: pip install pywin32
# Run: py electro_geometry_with_physical_deltas.py

import csv
import math
from pathlib import Path

import pythoncom
import win32com.client

MODEL_PATH = r""
OUTPUT_CSV = "electro_simple_ref_features.csv"
MIN_ALIGNMENT_TOLERANCE = 5.0
SHIELD_LENGTH_TOLERANCE_FRACTION = 0.04
SHELL_DIAMETER_TOLERANCE_FRACTION = 0.08

# Diameter correction currently validated only for the 20-700-000 design.
DIAMETER_CORRECTION_REFERENCE_MM = 63.5
CONDUCTOR_1_5_IN_MM = 38.1
CONDUCTOR_3_5_IN_MM = 88.9
CONDUCTOR_MATCH_TOLERANCE_MM = 7.0


class ElectroAPI:
    def __init__(self):
        pythoncom.CoInitialize()
        self.ies = win32com.client.Dispatch("IES.Document")

    def call(self, method_name, *args):
        return getattr(self.ies, method_name)(*args)

    def open_model(self, path):
        if path:
            ret = self.call("File_Open", path, 0)
            print(f"DEBUG File_Open returned: {ret}")

    def get_line_points(self, seg_id):
        try:
            ret = self.call(
                "Geometry2D_GetLinePointCoordinates",
                int(seg_id), 0.0, 0.0, 0.0, 0.0, 0,
            )
            print(f"DEBUG Geometry2D_GetLinePointCoordinates({seg_id}) returned: {ret}")
            nums = [float(v) for v in ret if isinstance(v, (int, float))] if isinstance(ret, tuple) else []
            if len(nums) >= 5 and int(nums[-1]) == 0:
                return [(nums[0], nums[1]), (nums[2], nums[3])]
            return None
        except Exception as exc:
            print(f"DEBUG line-coordinate lookup failed for segment {seg_id}: {exc}")
            return None

    def get_arc_points(self, seg_id):
        try:
            ret = self.call(
                "Geometry2D_GetArcPointCoordinates",
                int(seg_id), 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0,
            )
            print(f"DEBUG Geometry2D_GetArcPointCoordinates({seg_id}) returned: {ret}")
            nums = [float(v) for v in ret if isinstance(v, (int, float))] if isinstance(ret, tuple) else []
            if len(nums) >= 7 and int(nums[-1]) == 0:
                return [(nums[0], nums[1]), (nums[2], nums[3]), (nums[4], nums[5])]
            return None
        except Exception as exc:
            print(f"DEBUG arc-coordinate lookup failed for segment {seg_id}: {exc}")
            return None

    def get_reference_geometry_from_segment(self, seg_id, ref_name):
        pts = self.get_line_points(seg_id)
        kind = "line"
        if not pts:
            pts = self.get_arc_points(seg_id)
            kind = "arc"
        if not pts:
            raise RuntimeError(f"Unable to read line or arc coordinates for {ref_name} segment ID {seg_id}.")

        x_vals = [p[0] for p in pts]
        y_vals = [p[1] for p in pts]
        x_min, x_max = min(x_vals), max(x_vals)
        y_min, y_max = min(y_vals), max(y_vals)

        return {
            "reference": ref_name,
            "segment_id": int(seg_id),
            "type": kind,
            "points": pts,
            "x": sum(x_vals) / len(x_vals),
            "y": sum(y_vals) / len(y_vals),
            "x_span": abs(x_max - x_min),
            "y_span": abs(y_max - y_min),
            "x_min": x_min,
            "x_max": x_max,
            "y_min": y_min,
            "y_max": y_max,
        }


def prompt_segment_id(label):
    while True:
        value = input(f"Enter the segment ID for {label}: ").strip()
        try:
            seg_id = int(value)
            if seg_id <= 0:
                print("Segment ID must be a positive integer.")
                continue
            return seg_id
        except ValueError:
            print("Please enter a valid integer segment ID.")


def point_at_y_extreme(geometry, extreme):
    points = geometry["points"]
    if extreme == "max":
        target_y = max(point[1] for point in points)
    elif extreme == "min":
        target_y = min(point[1] for point in points)
    else:
        raise ValueError("extreme must be 'max' or 'min'")

    tolerance = 1e-9
    tied_points = [point for point in points if abs(point[1] - target_y) <= tolerance]
    return max(tied_points, key=lambda point: point[0])


def calculate_alignment_tolerance(
    shield_length,
    shell_diameter,
):
    """
    Calculate a scale-aware tolerance for deciding whether the shield bulb
    and selected shed reference point are approximately aligned.

    The largest of the following is used:
        - a minimum absolute tolerance;
        - a fraction of shield length;
        - a fraction of shell diameter.
    """
    return max(
        MIN_ALIGNMENT_TOLERANCE,
        SHIELD_LENGTH_TOLERANCE_FRACTION * shield_length,
        SHELL_DIAMETER_TOLERANCE_FRACTION * shell_diameter,
    )


def classify_plane_relation(
    global_delta_y,
    tolerance,
):
    """
    global_delta_y = shell_y - shield_y

    relation_code:
         1 = shield is below the shed
         0 = shield and shed are approximately in line
        -1 = shield is above the shed
    """
    if abs(global_delta_y) <= tolerance:
        return 0

    if global_delta_y > tolerance:
        return 1

    return -1


def relation_label(relation_code):
    labels = {
        1: "Shield is below the shed",
        0: "Shield and shed are in-line with each other",
        -1: "Shield is above the shed",
    }

    return labels[relation_code]


def compute_physical_relationship(
    shield_point,
    shell_point,
    side,
    alignment_tolerance,
):
    """
    Compute both global and physically oriented deltas.

    Global Cartesian convention:
        global_delta_x = shell_x - shield_x
        global_delta_y = shell_y - shield_y

    Physical outward convention:
        Upper:
            positive outward_delta_y means the shed is above/outward
            from the upper shield bulb.

        Lower:
            positive outward_delta_y means the shed is below/outward
            from the lower shield bulb.

    The human-readable relation label is based on global vertical position:
        Shield is above the shed
        Shield and shed are in-line with each other
        Shield is below the shed
    """
    shield_x, shield_y = shield_point
    shell_x, shell_y = shell_point

    global_delta_x = shell_x - shield_x
    global_delta_y = shell_y - shield_y

    if side == "upper":
        outward_delta_y = shell_y - shield_y
    elif side == "lower":
        outward_delta_y = shield_y - shell_y
    else:
        raise ValueError("side must be 'upper' or 'lower'")

    outward_delta_x = shell_x - shield_x

    distance = math.hypot(
        global_delta_x,
        global_delta_y,
    )

    relation_code = classify_plane_relation(
        global_delta_y,
        alignment_tolerance,
    )

    return {
        "side": side,

        "shield_x": shield_x,
        "shield_y": shield_y,
        "shell_x": shell_x,
        "shell_y": shell_y,

        "global_delta_x": global_delta_x,
        "global_delta_y": global_delta_y,

        "outward_delta_x": outward_delta_x,
        "outward_delta_y": outward_delta_y,

        "distance": distance,
        "alignment_tolerance": alignment_tolerance,

        "relation_code": relation_code,
        "relation_label": relation_label(relation_code),
    }

def print_reference(data):
    return (
        f"ID={data['segment_id']} | {data['type']} | "
        f"midpoint=({data['x']:.6f}, {data['y']:.6f}) | "
        f"x_span={data['x_span']:.6f} | y_span={data['y_span']:.6f}"
    )


def print_relationship(label, relationship):
    print(f"\n{label}")
    print("-" * len(label))
    print(f"Shield point: ({relationship['shield_x']:.6f}, {relationship['shield_y']:.6f})")
    print(f"Shell point:  ({relationship['shell_x']:.6f}, {relationship['shell_y']:.6f})")
    print(f"global_delta_x: {relationship['global_delta_x']:.6f}")
    print(f"global_delta_y: {relationship['global_delta_y']:.6f}")
    print(f"outward_delta_x: {relationship['outward_delta_x']:.6f}")
    print(f"outward_delta_y: {relationship['outward_delta_y']:.6f}")
    print(f"distance: {relationship['distance']:.6f}")
    print(
        f"alignment_tolerance: "
        f"{relationship['alignment_tolerance']:.6f}"
    )
    print(f"relation_code: {relationship['relation_code']}")
    print(f"relation_label: {relationship['relation_label']}")


def calculate_20_700_000_corrected_diameters(
    conductor_diameter_mm,
    shield_diameter_collected_mm,
    shell_mean_diameter_collected_mm,
):
    """Return corrected shield and shell diameters for design 20-700-000."""
    conductor_diameter_mm = float(conductor_diameter_mm)
    shield_diameter_collected_mm = float(shield_diameter_collected_mm)
    shell_mean_diameter_collected_mm = float(shell_mean_diameter_collected_mm)

    conductor_cases = [
        (CONDUCTOR_1_5_IN_MM, "1.5 in", 1.0),
        (CONDUCTOR_3_5_IN_MM, "3.5 in", -1.0),
    ]
    nominal_mm, nominal_label, correction_sign = min(
        conductor_cases,
        key=lambda case: abs(conductor_diameter_mm - case[0]),
    )
    if abs(conductor_diameter_mm - nominal_mm) > CONDUCTOR_MATCH_TOLERANCE_MM:
        raise ValueError(
            "The 20-700-000 diameter correction requires a conductor diameter "
            f"within +/-{CONDUCTOR_MATCH_TOLERANCE_MM:g} mm of 38.1 mm "
            f"(1.5 in) or 88.9 mm (3.5 in); collected "
            f"{conductor_diameter_mm:.3f} mm. Verify the conductor and origin segments."
        )

    correction_mm = abs(
        DIAMETER_CORRECTION_REFERENCE_MM - conductor_diameter_mm
    )
    corrected_shield_mm = (
        shield_diameter_collected_mm + correction_sign * correction_mm
    )
    corrected_shell_mm = (
        shell_mean_diameter_collected_mm + correction_sign * correction_mm
    )
    if corrected_shield_mm <= 0 or corrected_shell_mm <= 0:
        raise ValueError(
            "The 20-700-000 correction produced a non-positive diameter. "
            "Verify the selected origin, conductor, shield, and shell segments."
        )

    return {
        "shield_diameter_mm": corrected_shield_mm,
        "shell_mean_diameter_mm": corrected_shell_mm,
        "shield_diameter_segment_raw_mm": shield_diameter_collected_mm,
        "shell_mean_diameter_segment_raw_mm": shell_mean_diameter_collected_mm,
        "diameter_correction_mm": correction_mm,
        "diameter_correction_sign": int(correction_sign),
        "diameter_correction_conductor_nominal_mm": nominal_mm,
        "diameter_correction_conductor_nominal_label": nominal_label,
        "diameter_calculation_method": "20-700-000 segment-diameter correction",
    }


def main(apply_20_700_diameter_correction=False):
    print("\nELECTRO geometry extraction by segment ID")
    print("-----------------------------------------")
    print("Enter six ELECTRO segment IDs:")
    print("  1. origin")
    print("  2. conductor")
    print("  3. shield")
    print("  4. largest-diameter shell")
    print("  5. top shell/shed")
    print("  6. bottom shell/shed")
    print()
    print("Physical delta convention:")
    print("  positive outward_delta_y = shed lies outward from shield bulb")
    print("  near-zero outward_delta_y = approximately in line")
    print("  negative outward_delta_y = shield bulb extends beyond shed plane")
    print()
    print("Relation-label convention:")
    print("  Shield is above the shed")
    print("  Shield and shed are in-line with each other")
    print("  Shield is below the shed")
    print("  The in-line tolerance is scaled to shield length and shell diameter.")
    print("  Shell diameter uses the dedicated largest-diameter shell segment.")
    if apply_20_700_diameter_correction:
        print()
        print("IMPORTANT: 20-700-000 diameter correction is enabled.")
        print("This correction is currently intended only for the 20-700-000 design.")
        print("  1.5 in conductor: add |63.5 mm - conductor diameter|")
        print("  3.5 in conductor: subtract |63.5 mm - conductor diameter|")
    print()

    api = ElectroAPI()
    api.open_model(MODEL_PATH)

    origin_id = prompt_segment_id("origin")
    conductor_id = prompt_segment_id("conductor")
    shield_id = prompt_segment_id("shield")
    largest_shell_id = prompt_segment_id("largest-diameter shell")
    upper_shell_id = prompt_segment_id("top shell/shed")
    lower_shell_id = prompt_segment_id("bottom shell/shed")

    origin = api.get_reference_geometry_from_segment(origin_id, "origin")
    conductor = api.get_reference_geometry_from_segment(conductor_id, "conductor")
    shield = api.get_reference_geometry_from_segment(shield_id, "shield")
    largest_shell = api.get_reference_geometry_from_segment(
        largest_shell_id,
        "largest_diameter_shell",
    )
    upper_shell = api.get_reference_geometry_from_segment(
        upper_shell_id,
        "upper_shell",
    )
    lower_shell = api.get_reference_geometry_from_segment(
        lower_shell_id,
        "lower_shell",
    )

    conductor_diameter = 2.0 * abs(conductor["x"] - origin["x"])
    conductor_length = origin["y_span"]
    shield_diameter_collected = 2.0 * abs(shield["x"] - origin["x"])
    shield_length = shield["y_span"]

    shell_outer_x = largest_shell["x_max"]
    shell_diameter_collected = 2.0 * abs(
        shell_outer_x - origin["x"]
    )

    diameter_correction_features = {}
    if apply_20_700_diameter_correction:
        diameter_correction_features = calculate_20_700_000_corrected_diameters(
            conductor_diameter,
            shield_diameter_collected,
            shell_diameter_collected,
        )
        shield_diameter = diameter_correction_features["shield_diameter_mm"]
        shell_diameter = diameter_correction_features["shell_mean_diameter_mm"]
    else:
        shield_diameter = shield_diameter_collected
        shell_diameter = shell_diameter_collected

    shield_upper_point = point_at_y_extreme(shield, "max")
    shield_lower_point = point_at_y_extreme(shield, "min")
    upper_shell_point = point_at_y_extreme(upper_shell, "max")
    lower_shell_point = point_at_y_extreme(lower_shell, "min")

    alignment_tolerance = calculate_alignment_tolerance(
        shield_length=shield_length,
        shell_diameter=shell_diameter,
    )

    upper_relationship = compute_physical_relationship(
        shield_upper_point,
        upper_shell_point,
        side="upper",
        alignment_tolerance=alignment_tolerance,
    )

    lower_relationship = compute_physical_relationship(
        shield_lower_point,
        lower_shell_point,
        side="lower",
        alignment_tolerance=alignment_tolerance,
    )

    # Canonical output names match the centralized compact-vector schema.
    # Every ELECTRO coordinate-derived distance is expressed in millimetres.
    features = {
        # Persist the manually selected ELECTRO segment IDs so the centralized
        # application can store them with the active design profile and display
        # them later under "Review active design".
        "origin_segment_id": origin_id,
        "conductor_segment_id": conductor_id,
        "shield_segment_id": shield_id,
        "largest_diameter_shell_segment_id": largest_shell_id,
        "top_shed_segment_id": upper_shell_id,
        "bottom_shed_segment_id": lower_shell_id,

        "conductor_diameter_mm": conductor_diameter,
        "conductor_length_mm": conductor_length,
        "shield_diameter_mm": shield_diameter,
        "shield_length_mm": shield_length,
        "shell_mean_diameter_mm": shell_diameter,
        "shield_y_max_mm": shield["y_max"],
        "shield_y_min_mm": shield["y_min"],
        "top_shed_y_max_mm": upper_shell["y_max"],
        "bottom_shed_y_min_mm": lower_shell["y_min"],
        "shell_outer_x_max_mm": shell_outer_x,
        "shield_shed_alignment_tolerance_mm": alignment_tolerance,

        # Top/upper shield bulb to selected nearest-shed reference.
        "top_shed_global_delta_x_mm": upper_relationship["global_delta_x"],
        "top_shed_global_delta_y_mm": upper_relationship["global_delta_y"],
        "top_shed_outward_delta_x_mm": upper_relationship["outward_delta_x"],
        "top_shed_outward_delta_y_mm": upper_relationship["outward_delta_y"],
        "top_bulb_distance_to_nearest_shed_mm": upper_relationship["distance"],
        "top_shield_to_shed_relation_code": upper_relationship["relation_code"],

        # Bottom/lower shield bulb to selected nearest-shed reference.
        "bottom_shed_global_delta_x_mm": lower_relationship["global_delta_x"],
        "bottom_shed_global_delta_y_mm": lower_relationship["global_delta_y"],
        "bottom_shed_outward_delta_x_mm": lower_relationship["outward_delta_x"],
        "bottom_shed_outward_delta_y_mm": lower_relationship["outward_delta_y"],
        "bottom_bulb_distance_to_nearest_shed_mm": lower_relationship["distance"],
        "bottom_shield_to_shed_relation_code": lower_relationship["relation_code"],
    }
    features.update(diameter_correction_features)

    references = {
        "origin": origin,
        "conductor": conductor,
        "shield": shield,
        "largest_diameter_shell": largest_shell,
        "upper_shell": upper_shell,
        "lower_shell": lower_shell,
    }

    print("\nExtracted features")
    print("------------------")
    for key, value in features.items():
        if isinstance(value, float):
            print(f"{key}: {value:.6f}")
        else:
            print(f"{key}: {value}")

    print_relationship("Upper shield-to-shed relationship", upper_relationship)
    print_relationship("Lower shield-to-shed relationship", lower_relationship)

    print("\nReference geometry")
    print("------------------")
    for key, data in references.items():
        print(f"{key}: {print_reference(data)}")

    output_path = Path(OUTPUT_CSV)
    with output_path.open("w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["feature", "value"])
        for key, value in features.items():
            writer.writerow([key, value])

        writer.writerow([])
        writer.writerow([
            "reference", "segment_id", "type", "midpoint_x_mm", "midpoint_y_mm",
            "x_span_mm", "y_span_mm", "x_min_mm", "x_max_mm", "y_min_mm", "y_max_mm",
        ])
        for key, data in references.items():
            writer.writerow([
                key, data["segment_id"], data["type"], data["x"], data["y"],
                data["x_span"], data["y_span"], data["x_min"], data["x_max"],
                data["y_min"], data["y_max"],
            ])

        writer.writerow([])
        writer.writerow([
            "relationship", "shield_x_mm", "shield_y_mm", "shell_x_mm", "shell_y_mm",
            "global_delta_x_mm", "global_delta_y_mm", "outward_delta_x_mm",
            "outward_delta_y_mm", "distance_mm", "alignment_tolerance_mm",
            "relation_code", "relation_label",
        ])

        for relationship_name, relationship in [
            ("upper_shield_to_shed", upper_relationship),
            ("lower_shield_to_shed", lower_relationship),
        ]:
            writer.writerow([
                relationship_name,
                relationship["shield_x"], relationship["shield_y"],
                relationship["shell_x"], relationship["shell_y"],
                relationship["global_delta_x"], relationship["global_delta_y"],
                relationship["outward_delta_x"], relationship["outward_delta_y"],
                relationship["distance"],
                relationship["alignment_tolerance"],
                relationship["relation_code"],
                relationship["relation_label"],
            ])

    print(f"\nSaved: {output_path.resolve()}")
    return features


if __name__ == "__main__":
    main()
