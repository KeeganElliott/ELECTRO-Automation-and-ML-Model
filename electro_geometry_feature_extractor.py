"""
electro_geometry_feature_extractor.py

Extracts:
    Conductor Diameter (mm)
    Shield Diameter (mm)
    Shield Length (mm)
    Epoxy Diameter (mm)

Also creates small crosshair markers at the measurement locations.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional
import csv
import win32com.client


PROJECT_PATH = None  # Uses currently open ELECTRO model

OUTPUT_CSV = (
    r"T:\Interns Team Share\Keegan Elliott - ElectricalEng\ELECTRO"
    r"\Python Script\extracted_geometry_features.csv"
)

UNIT_TO_MM = 1.0
DOUBLE_RADIAL_DIMENSIONS = True

CREATE_MEASUREMENT_MARKERS = True
MARKER_HALF_SIZE = 1.0  # likely mm


@dataclass
class GeometryFeatures:
    project_name: str
    conductor_diameter_mm: Optional[float] = None
    shield_diameter_mm: Optional[float] = None
    shield_length_mm: Optional[float] = None
    # epoxy_diameter_mm: Optional[float] = None


def unwrap_api_result(result):
    if isinstance(result, tuple):
        return result[0], result[-1]
    return result, 0


def ensure_ok(name: str, err: int):
    if err != 0:
        raise RuntimeError(f"{name} failed with ELECTRO error code {err}")


def as_list(value):
    if value is None:
        return []
    if isinstance(value, tuple):
        return list(value)
    if isinstance(value, list):
        return value
    return [value]


def prompt_region_points():
    print("\nEnter one point inside each target region.")
    print("Avoid selecting a point exactly on a boundary.\n")

    region_points = {}

    for region in ["conductor", "shield"]:
        display_name = region.replace("_", " ").title()

        print(display_name)
        x = float(input("    X coordinate: "))
        y = float(input("    Y coordinate: "))

        region_points[region] = (x, y)
        print()

    return region_points


class ElectroGeometryAdapter:
    def __init__(
        self,
        project_path: Optional[str],
        region_points: dict[str, tuple[float, float]],
    ):
        self.project_path = Path(project_path) if project_path else None
        self.region_points = region_points
        self.project = None
        self.current_model_path = None

    def open_project(self) -> None:
        self.project = win32com.client.Dispatch("IES.Document")

        if self.project_path is not None:
            result = self.project.File_Open(str(self.project_path), 0)
            _, err = unwrap_api_result(result)
            ensure_ok("File_Open", err)
            self.current_model_path = self.project_path
        else:
            current_model, err = self.project.File_GetModelPath("", 0)
            ensure_ok("File_GetModelPath", err)

            if current_model == "Untitled":
                raise RuntimeError("No saved ELECTRO model is currently open.")

            self.current_model_path = Path(current_model)

    def get_region_ids_from_region_point(self, region_key: str) -> list[int]:
        x, y = self.region_points[region_key]

        result = self.project.Geometry2D_GetRegion_FromPoint(x, y, 0)
        region_id, err = unwrap_api_result(result)

        if err != 0 or region_id <= 0:
            raise RuntimeError(
                f"Could not find valid region for '{region_key}' "
                f"using point ({x}, {y})."
            )

        print(f"{region_key}: region ID = {region_id}")
        return [int(region_id)]

    def get_all_2d_segment_ids(self) -> list[int]:
        result = self.project.Geometry_GetIDsOfSegments([], 0)
        segment_ids, err = unwrap_api_result(result)
        ensure_ok("Geometry_GetIDsOfSegments", err)
        return [int(s) for s in as_list(segment_ids)]

    def get_region_on_left(self, seg_id: int) -> int:
        result = self.project.Geometry2D_GetRegionOnLeft_FromSegment(seg_id, 0)
        value, err = unwrap_api_result(result)
        if err != 0:
            return -1
        return int(value)

    def get_region_on_right(self, seg_id: int) -> int:
        result = self.project.Geometry2D_GetRegionOnRight_FromSegment(seg_id, 0)
        value, err = unwrap_api_result(result)
        if err != 0:
            return -1
        return int(value)

    def segment_touches_any_region(self, seg_id: int, region_ids: list[int]) -> bool:
        left_region = self.get_region_on_left(seg_id)
        right_region = self.get_region_on_right(seg_id)
        return left_region in region_ids or right_region in region_ids

    def get_line_points(self, seg_id: int):
        try:
            result = self.project.Geometry2D_GetLinePointCoordinates(
                seg_id, 0.0, 0.0, 0.0, 0.0, 0
            )

            if isinstance(result, tuple) and len(result) >= 5:
                x1, y1, x2, y2, err = result[:5]
                if err == 0:
                    return [(float(x1), float(y1)), (float(x2), float(y2))]
        except Exception:
            pass

        return None

    def get_arc_points(self, seg_id: int):
        try:
            result = self.project.Geometry2D_GetArcPointCoordinates(
                seg_id, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0
            )

            if isinstance(result, tuple) and len(result) >= 7:
                x1, y1, x2, y2, x3, y3, err = result[:7]
                if err == 0:
                    return [
                        (float(x1), float(y1)),
                        (float(x2), float(y2)),
                        (float(x3), float(y3)),
                    ]
        except Exception:
            pass

        return None

    def get_segment_points(self, seg_id: int) -> list[tuple[float, float]]:
        line_points = self.get_line_points(seg_id)
        if line_points is not None:
            return line_points

        arc_points = self.get_arc_points(seg_id)
        if arc_points is not None:
            return arc_points

        raise RuntimeError(f"Could not extract coordinates for segment ID {seg_id}")

    def get_bounding_box(self, region_key: str) -> tuple[float, float, float, float]:
        region_ids = self.get_region_ids_from_region_point(region_key)
        all_segments = self.get_all_2d_segment_ids()

        boundary_segments = [
            seg_id
            for seg_id in all_segments
            if self.segment_touches_any_region(seg_id, region_ids)
        ]

        if not boundary_segments:
            raise RuntimeError(f"No boundary segments found for region '{region_key}'.")

        points = []

        for seg_id in boundary_segments:
            points.extend(self.get_segment_points(seg_id))

        xs = [p[0] for p in points]
        ys = [p[1] for p in points]

        bbox = min(xs), max(xs), min(ys), max(ys)
        print(f"{region_key} bounding box:", bbox)

        return bbox

    def create_cross_marker(self, x: float, y: float, name: str) -> None:
        s = MARKER_HALF_SIZE

        try:
            self.project.Geometry2D_CreateLine(x - s, y, x + s, y, 0)
            self.project.Geometry2D_CreateLine(x, y - s, x, y + s, 0)
            print(f"Created measurement marker: {name} at ({x}, {y})")
        except Exception as e:
            print(f"WARNING: Could not create marker '{name}'.")
            print(e)


def bbox_height(bbox: tuple[float, float, float, float]) -> float:
    _, _, ymin, ymax = bbox
    return abs(ymax - ymin)


def bbox_max_x_from_origin(bbox: tuple[float, float, float, float]) -> float:
    xmin, xmax, _, _ = bbox
    return max(abs(xmin), abs(xmax))


def radial_diameter_from_quarter_section(
    bbox: tuple[float, float, float, float],
    unit_to_mm: float = UNIT_TO_MM,
    double_radial: bool = DOUBLE_RADIAL_DIMENSIONS,
) -> float:
    radial_extent = bbox_max_x_from_origin(bbox)

    if double_radial:
        return 2.0 * radial_extent * unit_to_mm

    return radial_extent * unit_to_mm


def axial_length_from_bbox(
    bbox: tuple[float, float, float, float],
    unit_to_mm: float = UNIT_TO_MM,
) -> float:
    return bbox_height(bbox) * unit_to_mm


def midpoint(a: float, b: float) -> float:
    return 0.5 * (a + b)


def add_measurement_markers(
    adapter: ElectroGeometryAdapter,
    conductor_bbox,
    shield_bbox,
) -> None:
    cxmin, cxmax, cymin, cymax = conductor_bbox
    sxmin, sxmax, symin, symax = shield_bbox

    adapter.create_cross_marker(
        cxmax,
        midpoint(cymin, cymax),
        "conductor diameter radial max",
    )

    adapter.create_cross_marker(
        sxmax,
        midpoint(symin, symax),
        "shield diameter radial max",
    )

    adapter.create_cross_marker(
        midpoint(sxmin, sxmax),
        symin,
        "shield length bottom",
    )

    adapter.create_cross_marker(
        midpoint(sxmin, sxmax),
        symax,
        "shield length top",
    )


def extract_geometry_features(project_path: Optional[str]) -> GeometryFeatures:
    region_points = prompt_region_points()

    adapter = ElectroGeometryAdapter(project_path, region_points)
    adapter.open_project()

    project_name = adapter.current_model_path.stem

    conductor_bbox = adapter.get_bounding_box("conductor")
    shield_bbox = adapter.get_bounding_box("shield")

    if CREATE_MEASUREMENT_MARKERS:
        add_measurement_markers(
            adapter,
            conductor_bbox,
            shield_bbox,
        )

    return GeometryFeatures(
    project_name=project_name,
    conductor_diameter_mm=radial_diameter_from_quarter_section(conductor_bbox),
    shield_diameter_mm=radial_diameter_from_quarter_section(shield_bbox),
    shield_length_mm=axial_length_from_bbox(shield_bbox),
)


def append_features_to_csv(features: GeometryFeatures, output_csv: str) -> None:
    output_path = Path(output_csv)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    row = asdict(features)
    write_header = not output_path.exists()

    with output_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=row.keys())

        if write_header:
            writer.writeheader()

        writer.writerow(row)


def print_feature_summary(features: GeometryFeatures) -> None:
    print("\nExtracted ELECTRO Geometry Features")
    print("----------------------------------")
    print(f"Project Name:              {features.project_name}")
    print(f"Conductor Diameter (mm):   {features.conductor_diameter_mm}")
    print(f"Shield Diameter (mm):      {features.shield_diameter_mm}")
    print(f"Shield Length (mm):        {features.shield_length_mm}")
    # print(f"Epoxy Diameter (mm):       {features.epoxy_diameter_mm}")


def main() -> None:
    features = extract_geometry_features(PROJECT_PATH)
    print_feature_summary(features)
    append_features_to_csv(features, OUTPUT_CSV)

    print(f"\nFeature row appended to:\n{OUTPUT_CSV}")


if __name__ == "__main__":
    main()