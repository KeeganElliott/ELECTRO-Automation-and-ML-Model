"""
API_Automation_with_easy_feature_integration.py

This is not meant to replace your full original API_Automation.py file.
Instead, copy the marked section into your existing pipeline after:

    1) ELECTRO project is opened
    2) simulation/export steps are complete
    3) E-field/charge CSV files exist

Required companion file:
    easy_feature_extraction.py
"""

from pathlib import Path

from easy_feature_extraction import (
    extract_easy_api_features,
    extract_max_outputs_from_distance_ranges,
    merge_output_features,
    write_feature_row_csv,
)


def run_api_automation_pipeline():
    # ---------------------------------------------------------------------
    # Existing pipeline section
    # ---------------------------------------------------------------------
    PROJECT_PATH = Path(r"T:\Path\To\Your\Bushing_700.dsb")
    EXPORT_FOLDER = Path(r"T:\Path\To\Your\exports")
    OUTPUT_FEATURE_CSV = Path(r"T:\Path\To\Your\outputs\extracted_feature_rows.csv")

    # Replace with your actual ELECTRO API project-open logic
    electro_project = None
    # electro_project = electro.open_project(PROJECT_PATH)

    # Replace with your existing simulation/export logic
    # run_transient_solution(electro_project)
    # export_line_plot(electro_project, EXPORT_FOLDER / "E_vs_distance.csv")
    # export_line_plot(electro_project, EXPORT_FOLDER / "charge_vs_distance.csv")

    e_csv_path = EXPORT_FOLDER / "E_vs_distance.csv"
    charge_csv_path = EXPORT_FOLDER / "charge_vs_distance.csv"

    # ---------------------------------------------------------------------
    # NEW IMPLEMENTATION SECTION
    # ---------------------------------------------------------------------

    # These ranges should come from your line-plot distance coordinate system.
    # Replace the example values with the distances you provide from ELECTRO.
    conductor_range = (0.0, 25.0)
    shield_range = (120.0, 170.0)
    shed_range = (250.0, 500.0)

    # Step 1: Extract easiest model/API variables
    features = extract_easy_api_features(
        electro_project=electro_project,
        project_path=PROJECT_PATH,
        conductor_region_name="conductor",
        shield_region_name="shield",
        shell_region_name="shell",
        voltage_source_name="BIL",
    )

    # Step 2: Extract output max values from exported CSV files
    output_features = extract_max_outputs_from_distance_ranges(
        e_csv_path=e_csv_path if e_csv_path.exists() else None,
        charge_csv_path=charge_csv_path if charge_csv_path.exists() else None,
        conductor_range=conductor_range,
        shield_range=shield_range,
        shed_range=shed_range,
    )

    # Step 3: Combine model/API features + output-derived max features
    features = merge_output_features(features, output_features)

    # Step 4: Append one row to the ML input/output dataset
    write_feature_row_csv(features, OUTPUT_FEATURE_CSV)

    print(f"Feature row written to: {OUTPUT_FEATURE_CSV}")


if __name__ == "__main__":
    run_api_automation_pipeline()
