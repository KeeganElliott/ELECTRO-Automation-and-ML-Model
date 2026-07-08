# electro_simple_ref_geometry_extraction_FIXED.py
# Requires: pip install pywin32
# Run: python electro_simple_ref_geometry_extraction_FIXED.py

import csv
from pathlib import Path

import pythoncom
import win32com.client


MODEL_PATH = r""
OUTPUT_CSV = "electro_simple_ref_features.csv"

REF_NAMES = {
    "origin": ["origin", "Object 1"],
    "conductor": ["conductor", "Object 2"],
    "shield": ["shield", "Object 3"],
    "shell": ["shell", "Object 4"],
}


class ElectroAPI:
    def __init__(self):
        pythoncom.CoInitialize()
        self.ies = win32com.client.Dispatch("IES.Document")

    def call(self, method_name, *args):
        return getattr(self.ies, method_name)(*args)

    def open_model(self, path):
        if path:
            self.call("File_Open", path, 0)

    def get_segment_id_from_name(self, name):
        try:
            ret = self.call("Geometry_GetID_FromName", name, "Segment", 0, 0)
            print(f"DEBUG segment lookup {name!r}: {ret}")

            if isinstance(ret, tuple) and len(ret) >= 2:
                seg_id = int(ret[0])
                err = int(ret[-1])

                if err == 0 and seg_id > 0:
                    return seg_id

            return None

        except Exception as e:
            print(f"DEBUG segment lookup failed for {name!r}: {e}")
            return None

    def get_object_id_from_name(self, name):
        try:
            ret = self.call("Object_GetObjectID", name, 0, 0)
            print(f"DEBUG object lookup {name!r}: {ret}")

            if isinstance(ret, tuple) and len(ret) >= 2:
                obj_id = int(ret[0])
                err = int(ret[-1])

                if err == 0 and obj_id > 0:
                    return obj_id

            return None

        except Exception as e:
            print(f"DEBUG object lookup failed for {name!r}: {e}")
            return None

    def get_line_points(self, seg_id):
        try:
            ret = self.call(
                "Geometry2D_GetLinePointCoordinates",
                int(seg_id),
                0.0, 0.0, 0.0, 0.0, 0,
            )

            nums = [float(v) for v in ret if isinstance(v, (int, float))]

            if len(nums) >= 5 and int(nums[-1]) == 0:
                return [(nums[0], nums[1]), (nums[2], nums[3])]

            return None

        except Exception:
            return None

    def get_arc_points(self, seg_id):
        try:
            ret = self.call(
                "Geometry2D_GetArcPointCoordinates",
                int(seg_id),
                0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0,
            )

            nums = [float(v) for v in ret if isinstance(v, (int, float))]

            if len(nums) >= 7 and int(nums[-1]) == 0:
                return [(nums[0], nums[1]), (nums[2], nums[3]), (nums[4], nums[5])]

            return None

        except Exception:
            return None

    def get_reference_geometry(self, ref_key, possible_names, required=True):
        if isinstance(possible_names, str):
            possible_names = [possible_names]

        object_was_found = False

        for name in possible_names:
            seg_id = self.get_segment_id_from_name(name)

            if seg_id:
                pts = self.get_line_points(seg_id)
                kind = "line"

                if not pts:
                    pts = self.get_arc_points(seg_id)
                    kind = "arc"

                if not pts:
                    raise RuntimeError(
                        f"Found segment named {name!r} with ID {seg_id}, "
                        f"but could not read its line/arc coordinates."
                    )

                x_vals = [p[0] for p in pts]
                y_vals = [p[1] for p in pts]

                x_mid = sum(x_vals) / len(x_vals)
                y_mid = sum(y_vals) / len(y_vals)

                x_span = max(x_vals) - min(x_vals)
                y_span = max(y_vals) - min(y_vals)

                return {
                    "reference": ref_key,
                    "name_used": name,
                    "segment_id": seg_id,
                    "type": kind,
                    "x": x_mid,
                    "y": y_mid,
                    "x_span": abs(x_span),
                    "y_span": abs(y_span),
                    "x_min": min(x_vals),
                    "x_max": max(x_vals),
                    "y_min": min(y_vals),
                    "y_max": max(y_vals),
                }

            obj_id = self.get_object_id_from_name(name)
            if obj_id:
                object_was_found = True

        if object_was_found:
            msg = (
                f"\nFound an ELECTRO Object for {ref_key}, but not a Segment.\n"
                f"The API coordinate calls require a true Segment ID.\n"
                f"Fix: rename the actual reference segment geometry to one of: {possible_names}\n"
                f"Do not only label/create an Object containing the segment."
            )
        else:
            msg = f"\nNo Segment or Object found for {ref_key} using names: {possible_names}"

        if required:
            raise RuntimeError(msg)

        print("\nNo shield available to measure.")
        print(msg)
        return None


def print_ref(data):
    if data is None:
        return "N/A"

    return (
        f"name={data['name_used']} | ID={data['segment_id']} | {data['type']} | "
        f"midpoint=({data['x']:.6f}, {data['y']:.6f}) | "
        f"x_span={data['x_span']:.6f} | y_span={data['y_span']:.6f}"
    )


def main():
    print("\nELECTRO simple reference geometry extraction")
    print("-------------------------------------------")
    print("Accepted names:")
    print("  origin:    origin or Object 1")
    print("  conductor: conductor or Object 2")
    print("  shield:    shield or Object 3")
    print("  shell:     shell or Object 4")
    print("\nIMPORTANT:")
    print("These names must belong to actual reference SEGMENTS, not only ELECTRO Objects.")
    print("Diameters are measured using segment midpoint x-distance from origin.")
    print("Conductor length is measured using the origin segment's total y-span.")
    print("Shield length is measured using the shield segment's total y-span.\n")

    api = ElectroAPI()
    api.open_model(MODEL_PATH)

    origin = api.get_reference_geometry("origin", REF_NAMES["origin"], required=True)
    conductor = api.get_reference_geometry("conductor", REF_NAMES["conductor"], required=True)
    shell = api.get_reference_geometry("shell", REF_NAMES["shell"], required=True)
    shield = api.get_reference_geometry("shield", REF_NAMES["shield"], required=False)

    conductor_diameter = 2.0 * abs(conductor["x"] - origin["x"])
    conductor_length = origin["y_span"]

    shell_diameter = 2.0 * abs(shell["x"] - origin["x"])

    if shield is None:
        shield_diameter = "N/A"
        shield_length = "N/A"
    else:
        shield_diameter = 2.0 * abs(shield["x"] - origin["x"])
        shield_length = shield["y_span"]

    features = {
        "conductor_diameter": conductor_diameter,
        "conductor_length": conductor_length,
        "shield_diameter": shield_diameter,
        "shield_length": shield_length,
        "shell_diameter": shell_diameter,
    }

    refs = {
        "origin": origin,
        "conductor": conductor,
        "shield": shield,
        "shell": shell,
    }

    print("\nExtracted features")
    print("------------------")
    for key, value in features.items():
        if isinstance(value, float):
            print(f"{key}: {value:.6f}")
        else:
            print(f"{key}: {value}")

    print("\nReference geometry")
    print("------------------")
    for key, data in refs.items():
        print(f"{key}: {print_ref(data)}")

    out = Path(OUTPUT_CSV)
    with out.open("w", newline="") as f:
        writer = csv.writer(f)

        writer.writerow(["feature", "value"])
        for key, value in features.items():
            writer.writerow([key, value])

        writer.writerow([])
        writer.writerow([
            "reference", "name_used", "segment_id", "type",
            "midpoint_x", "midpoint_y",
            "x_span", "y_span",
            "x_min", "x_max", "y_min", "y_max",
        ])

        for key, data in refs.items():
            if data is None:
                writer.writerow([key, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"])
            else:
                writer.writerow([
                    key,
                    data["name_used"],
                    data["segment_id"],
                    data["type"],
                    data["x"],
                    data["y"],
                    data["x_span"],
                    data["y_span"],
                    data["x_min"],
                    data["x_max"],
                    data["y_min"],
                    data["y_max"],
                ])

    print(f"\nSaved: {out.resolve()}")


if __name__ == "__main__":
    main()