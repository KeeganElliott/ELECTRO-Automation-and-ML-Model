"""
ELECTRO Boundary E-field Extraction + DSP Automation
Version: 2026-07-13-electro-only-transient-static-setup-v13

Purpose
-------
Transient setup + post-simulation automation script for INTEGRATED ELECTRO via the COM API.

Primary automated workflow:
    - use fixed ELECTRO object names directly,
    - assign default bushing materials by object name,
    - assign named voltage definitions to ELECTRO geometry,
    - point the HV voltage definition to an existing transient source/waveform,
    - optionally run the solver,
    - optionally continue into the existing post-simulation DSP workflow.

No transient setup CSV is required. The only required user input is the name of
the existing transient BIL source/waveform in the open ELECTRO model.

Post-simulation workflow:
    - extract boundary-following E-field data from segment IDs or seed points,
    - save raw E-field CSV files,
    - compute E-field summary metrics,
    - run spatial FFT and optional wavelet analysis.

Important API limitation
------------------------
The ELECTRO API function Analysis_Get2DElectricField_FromSegment requires a
segment ID. If your boundary is named as an object such as "measure", this
script can attempt a named-object workflow only if the API exposes the needed
segment lookup on your installation. Otherwise, use the config CSV or seed-point
workflow.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import win32com.client

try:
    import pywt
    PYWT_AVAILABLE = True
except ImportError:
    PYWT_AVAILABLE = False


# -----------------------------------------------------------------------------
# Default project constants
# -----------------------------------------------------------------------------

DEFAULT_MEASURE_OBJECT = "measure"
DEFAULT_NUM_POINTS = 500
DEFAULT_OUTPUT_ROOT_NAME = "API_DSP_Output"
EPOXY_OUTSIDE_OBJECT_NAME = "epoxy outside"

# Standard material mapping used when each model keeps these ELECTRO object names.
# These names must match the ELECTRO object/model tree and material library spelling exactly.
DEFAULT_MATERIAL_MAP = {
    "conductor": "Copper",
    "gnd": "Copper",
    "shell": "Standard molded Epoxy @60Hz",
}

# Default voltage setup. These ELECTRO object names must exist in the open model.
DEFAULT_HV_OBJECT_NAME = "conductor"
DEFAULT_GROUND_OBJECT_NAME = "gnd"
DEFAULT_HV_VOLTAGE_NAME = "source"
DEFAULT_GROUND_VOLTAGE_NAME = "gnd"
DEFAULT_TRANSIENT_SOURCE_NAME = "impulse"
DEFAULT_TRANSIENT_GROUND_SOURCE_NAME = "Ground"
DEFAULT_STATIC_HV_VOLTAGE = 350000.0


@dataclass
class BoundaryMeasurement:
    label: str
    side: int
    num_points: int
    segment_id: Optional[int] = None
    seed_x: Optional[float] = None
    seed_y: Optional[float] = None


# -----------------------------------------------------------------------------
# General API utilities
# -----------------------------------------------------------------------------


def call_api(ies: Any, method_name: str, *args: Any, required: bool = False) -> Any:
    """Call an ELECTRO API method and print the returned value.

    COM return signatures vary between API versions and between Python COM
    wrappers. This helper does not assume a fixed return shape.
    """
    try:
        method = getattr(ies, method_name)
    except AttributeError as exc:
        msg = f"API method not available: {method_name}"
        if required:
            raise RuntimeError(msg) from exc
        print(f"WARNING: {msg}")
        return None

    try:
        result = method(*args)
    except Exception as exc:
        msg = f"API call failed: {method_name}({args}) -> {exc}"
        if required:
            raise RuntimeError(msg) from exc
        print(f"WARNING: {msg}")
        return None

    print(f"{method_name} returned: {result}")
    return result


def extract_error_code(result: Any) -> int:
    """Best-effort extraction of an ELECTRO iErr code from a COM result."""
    if result is None:
        return -999
    if isinstance(result, tuple) and len(result) > 0:
        last = result[-1]
        try:
            return int(last)
        except Exception:
            return 0
    try:
        return int(result)
    except Exception:
        return 0


def yes_no(prompt: str, default: bool = False) -> bool:
    default_text = "Y/n" if default else "y/N"
    response = input(f"{prompt} [{default_text}]: ").strip().lower()
    if not response:
        return default
    return response in {"y", "yes"}


def prompt_float(prompt: str, default: Optional[float] = None) -> float:
    if default is None:
        return float(input(f"{prompt}: ").strip())
    value = input(f"{prompt} [{default}]: ").strip()
    return float(value) if value else float(default)


def prompt_int(prompt: str, default: Optional[int] = None) -> int:
    if default is None:
        return int(input(f"{prompt}: ").strip())
    value = input(f"{prompt} [{default}]: ").strip()
    return int(value) if value else int(default)


def prompt_text(prompt: str, default: str) -> str:
    value = input(f"{prompt} [{default}]: ").strip()
    return value if value else default


def get_current_model_path(ies: Any) -> Path:
    result = call_api(ies, "File_GetModelPath", "", 0, required=True)

    if isinstance(result, tuple):
        current_model = result[0]
        err = extract_error_code(result)
    else:
        current_model = str(result)
        err = 0

    print("Current model:", current_model)
    print("Error code:", err)

    if err != 0:
        raise RuntimeError("Could not determine current model path.")
    if str(current_model).strip().lower() == "untitled":
        raise RuntimeError("No saved model is currently open in ELECTRO.")

    return Path(str(current_model))


def make_output_dir(model_path: Path, phase_label: str = "postsim_extract_dsp") -> Path:
    run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = model_path.parent / DEFAULT_OUTPUT_ROOT_NAME / f"{phase_label}_{run_stamp}"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


# -----------------------------------------------------------------------------
# Transient pre-simulation setup
# -----------------------------------------------------------------------------

@dataclass
class PhysicsAssignment:
    """One row from the transient setup manifest.

    Preferred usage is object_name + geometry_type.  If an object cannot be
    resolved reliably in your ELECTRO version, provide geometry_id directly.

    voltage_mode options:
        transient  -> PhysicsVoltage_SetTransientValue(voltage_name, source_name)
        static     -> PhysicsVoltage_SetStaticValue(voltage_name, static_value)
        none/blank -> no voltage assignment
    """
    role: str
    object_name: str = ""
    geometry_type: str = "Object"
    geometry_id: Optional[int] = None
    material: str = ""
    voltage_name: str = ""
    voltage_mode: str = ""
    source_name: str = ""
    static_value: Optional[float] = None
    color_r: Optional[int] = None
    color_g: Optional[int] = None
    color_b: Optional[int] = None
    active: int = 1
    assign_connected: int = 1


def create_default_transient_manifest(path: Path) -> None:
    """Create a manifest template for object-name based transient setup."""
    df = pd.DataFrame([
        {
            "role": "conductor_hv",
            "object_name": "conductor",
            "geometry_type": "Object",
            "geometry_id": "",
            "material": "Copper",
            "voltage_name": "source",
            "voltage_mode": "transient",
            "source_name": "impulse",
            "static_value": "",
            "color_r": 255,
            "color_g": 0,
            "color_b": 0,
            "active": 1,
            "assign_connected": 1,
            "notes": "Rename object/source/material to match the ELECTRO model tree and material library.",
        },
        {
            "role": "ground_shield",
            "object_name": "shield",
            "geometry_type": "Object",
            "geometry_id": "",
            "material": "Copper",
            "voltage_name": "gnd",
            "voltage_mode": "static",
            "source_name": "",
            "static_value": 0.0,
            "color_r": 0,
            "color_g": 0,
            "color_b": 255,
            "active": 1,
            "assign_connected": 1,
            "notes": "For grounded conductor or shield. Use your exact object name.",
        },
        {
            "role": "epoxy_shell",
            "object_name": "shell",
            "geometry_type": "Object",
            "geometry_id": "",
            "material": "Standard molded Epoxy @60Hz",
            "voltage_name": "",
            "voltage_mode": "none",
            "source_name": "",
            "static_value": "",
            "color_r": "",
            "color_g": "",
            "color_b": "",
            "active": 1,
            "assign_connected": 1,
            "notes": "Dielectric material only; no voltage assignment. Object name must be shell.",
        },
    ])
    df.to_csv(path, index=False)


def load_transient_manifest(path: Path) -> list[PhysicsAssignment]:
    df = pd.read_csv(path)
    required_cols = {"role"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Transient manifest missing columns: {sorted(missing)}")

    def text(row: pd.Series, name: str, default: str = "") -> str:
        value = row.get(name, default)
        if pd.isna(value):
            return default
        return str(value).strip()

    def opt_int(row: pd.Series, name: str) -> Optional[int]:
        value = row.get(name, None)
        if pd.isna(value) or str(value).strip() == "":
            return None
        return int(float(value))

    def opt_float(row: pd.Series, name: str) -> Optional[float]:
        value = row.get(name, None)
        if pd.isna(value) or str(value).strip() == "":
            return None
        return float(value)

    assignments: list[PhysicsAssignment] = []
    for _, row in df.iterrows():
        role = text(row, "role")
        if not role:
            continue
        assignments.append(PhysicsAssignment(
            role=role,
            object_name=text(row, "object_name"),
            geometry_type=text(row, "geometry_type", "Object") or "Object",
            geometry_id=opt_int(row, "geometry_id"),
            material=text(row, "material"),
            voltage_name=text(row, "voltage_name"),
            voltage_mode=text(row, "voltage_mode").lower(),
            source_name=text(row, "source_name"),
            static_value=opt_float(row, "static_value"),
            color_r=opt_int(row, "color_r"),
            color_g=opt_int(row, "color_g"),
            color_b=opt_int(row, "color_b"),
            active=opt_int(row, "active") if opt_int(row, "active") is not None else 1,
            assign_connected=opt_int(row, "assign_connected") if opt_int(row, "assign_connected") is not None else 1,
        ))
    if not assignments:
        raise ValueError("No usable rows found in transient manifest.")
    return assignments


def parse_id_and_error(result: Any) -> tuple[Optional[int], int]:
    """Best-effort extraction of an ID plus iErr from variable COM returns."""
    if result is None:
        return None, -999
    if isinstance(result, tuple):
        err = extract_error_code(result)
        for item in result:
            if isinstance(item, (int, float)) and int(item) > 0:
                return int(item), err
        return None, err
    try:
        value = int(result)
        return (value if value > 0 else None), 0
    except Exception:
        return None, 0


def resolve_geometry_id(ies: Any, assignment: PhysicsAssignment) -> Optional[int]:
    if assignment.geometry_id is not None and assignment.geometry_id > 0:
        return assignment.geometry_id

    if not assignment.object_name:
        return None

    # Preferred for model-tree objects created from named CAD/ELECTRO groups.
    result = call_api(ies, "Object_GetObjectID", assignment.object_name, 0, required=False)
    object_id, err = parse_id_and_error(result)
    if object_id and err == 0:
        return object_id

    # Backup for named geometry objects if object-tree lookup fails.
    result = call_api(
        ies,
        "Geometry_GetID_FromName",
        assignment.object_name,
        assignment.geometry_type,
        0,
        0,
        required=False,
    )
    geom_id, err = parse_id_and_error(result)
    if geom_id and err == 0:
        return geom_id

    return None


def assign_material(ies: Any, assignment: PhysicsAssignment, geometry_id: Optional[int]) -> None:
    if not assignment.material:
        return

    print(f"\n--- Material assignment: {assignment.role} -> {assignment.material} ---")

    success = False
    if assignment.object_name:
        # Most convenient old-format call: assigns by named object where supported.
        result = call_api(ies, "Physics_SetMaterial", assignment.object_name, assignment.material, 0, required=False)
        success = extract_error_code(result) == 0 if result is not None else False

    if not success and geometry_id and assignment.geometry_type.lower() == "volume":
        result = call_api(ies, "Physics_SetMaterial_ByVolume", int(geometry_id), assignment.material, 0, required=False)
        success = extract_error_code(result) == 0 if result is not None else False

    if not success:
        print(
            f"WARNING: Material assignment may not have succeeded for role '{assignment.role}'. "
            "Check object_name/material spelling or provide geometry_type=Volume with geometry_id."
        )


def assign_voltage(ies: Any, assignment: PhysicsAssignment, geometry_id: Optional[int]) -> None:
    """Update an existing ELECTRO voltage definition by name.

    This v8 workflow is intentionally model-tree based because your ELECTRO
    template already contains voltage definitions named `source` and `gnd`, and
    those definitions are already attached to the correct objects in the tree.

    Therefore this function does NOT delete/create/clear/reassign voltage
    geometry. It only sets the value/source on the existing voltage definition:

        source -> transient source named impulse
        gnd    -> static value 0 V

    This avoids the geometry-ID mismatch seen when attempting
    PhysicsVoltage_AssignGeometry with object IDs.
    """
    if not assignment.voltage_name or assignment.voltage_mode in {"", "none", "nan"}:
        return

    print(f"\n--- Voltage value update: {assignment.role} -> {assignment.voltage_name} ---")
    print(
        "Using existing ELECTRO model-tree voltage definition. "
        "No create/delete/clear/geometry reassignment is attempted."
    )

    call_api(ies, "PhysicsVoltage_SetActive", assignment.voltage_name, int(assignment.active), 0, required=False)

    if assignment.color_r is not None and assignment.color_g is not None and assignment.color_b is not None:
        call_api(
            ies,
            "PhysicsVoltage_SetColor",
            assignment.voltage_name,
            int(assignment.color_r),
            int(assignment.color_g),
            int(assignment.color_b),
            0,
            required=False,
        )

    if assignment.voltage_mode.startswith("trans"):
        if not assignment.source_name:
            print(f"WARNING: Transient voltage '{assignment.voltage_name}' needs a source_name.")
            return
        result = call_api(
            ies,
            "PhysicsVoltage_SetTransientValue",
            assignment.voltage_name,
            assignment.source_name,
            0,
            required=False,
        )
        if extract_error_code(result) != 0:
            print(
                f"WARNING: Could not set transient source '{assignment.source_name}' on voltage "
                f"'{assignment.voltage_name}'. Check that the model tree contains Voltage > "
                f"{assignment.voltage_name} and Sources > {assignment.source_name}, spelled exactly."
            )
        else:
            print(f"Assigned transient source '{assignment.source_name}' to voltage '{assignment.voltage_name}'.")

    elif assignment.voltage_mode.startswith("stat"):
        value = 0.0 if assignment.static_value is None else float(assignment.static_value)
        result = call_api(ies, "PhysicsVoltage_SetStaticValue", assignment.voltage_name, value, 0, required=False)
        if extract_error_code(result) != 0:
            print(
                f"WARNING: Could not set static value {value} on voltage '{assignment.voltage_name}'. "
                "Check that the model tree contains the exact voltage definition name."
            )
        else:
            print(f"Assigned static voltage {value} V to voltage '{assignment.voltage_name}'.")
    else:
        print(f"WARNING: Unsupported voltage_mode '{assignment.voltage_mode}' for role '{assignment.role}'.")




def _call_api_variants(ies: Any, method_name: str, variants: list[tuple[Any, ...]]) -> Any:
    """Try several COM signatures and return the first successful API result."""
    last_result = None
    for args in variants:
        result = call_api(ies, method_name, *args, required=False)
        last_result = result
        if result is not None and extract_error_code(result) == 0:
            return result
    return last_result


def configure_common_electric_settings(ies: Any) -> None:
    """Apply settings required by both transient and static bushing runs."""
    print("\n--- Common ELECTRO settings ---")

    # Material mode: Permittivity only.
    result = _call_api_variants(
        ies,
        "PhysicsElectricSettings_SetMaterialMode_AsPermittivity",
        [(), (0,)],
    )
    if result is None or extract_error_code(result) != 0:
        print("WARNING: Could not confirm Material Mode = Permittivity.")

    # Grounded Voltage Sources unchecked. The API guide describes SetUnbalanced
    # as checking Grounded Voltage Sources, so SetBalanced is the inverse.
    result = _call_api_variants(ies, "Physics_SetBalanced", [(), (0,)])
    if result is None or extract_error_code(result) != 0:
        print("WARNING: Could not confirm Grounded Voltage Sources is unselected.")

    # Geometry model type: Axisymmetric - Y-axis.
    result = _call_api_variants(
        ies,
        "Physics_Set2DModelType_AsYRotationalSymmetric",
        [(0,), ()],
    )
    if result is None or extract_error_code(result) != 0:
        result = _call_api_variants(ies, "Physics_Set2DModelType", [(2, 0), (2,)])
    if result is None or extract_error_code(result) != 0:
        print("WARNING: Could not confirm Model Type = Axisymmetric - Y-axis.")


def configure_transient_operation_mode(ies: Any) -> None:
    """Handle ELECTRO transient operation mode without using an invalid API call.

    The documented Model_SetAnalysisMode method selects the analysis family
    (Electric/Magnetic/Trajectory), not Static/Transient operation mode. The
    available ELECTRO COM method list does not expose a documented
    Physics_SetTransientMode equivalent to Physics_SetStaticMode. Therefore,
    transient mode must already be selected in the open template/model.
    """
    print("\n--- Transient operation settings ---")
    print(
        "ELECTRO's documented API does not expose a confirmed command for changing "
        "Operation Mode from Static to Transient. Model_SetAnalysisMode('Transient') "
        "is not valid for this purpose; it controls Electric/Magnetic/Trajectory mode."
    )
    print(
        "REQUIRED: Set Operation Mode = Transient in the ELECTRO template/model before "
        "running transient setup. The script will still assign source->impulse and "
        "gnd->Ground, materials, material mode, grounding option, and model type."
    )

    if not yes_no("Confirm ELECTRO currently shows Operation Mode = Transient", default=False):
        raise RuntimeError(
            "Transient setup stopped because Operation Mode was not confirmed as Transient. "
            "Select Transient in ELECTRO, then rerun option 1 or 2."
        )


def configure_static_operation_and_solver(ies: Any) -> None:
    """Set Static operation mode and Boundary Element solution method."""
    print("\n--- Static operation and solver settings ---")

    result = _call_api_variants(ies, "Physics_SetStaticMode", [(), (0,)])
    if result is None or extract_error_code(result) != 0:
        print("WARNING: Could not confirm Operation Mode = Static.")

    result = _call_api_variants(ies, "Solution_SetSolutionMethod_AsBEM", [(0,), ()])
    if result is None or extract_error_code(result) != 0:
        print("WARNING: Could not confirm Method of Solution = Boundary Element.")


def apply_default_materials_by_name(ies: Any) -> None:
    """Assign the standard bushing materials using fixed ELECTRO object names.

    Required object/material names:
        conductor -> Copper
        gnd       -> Copper
        shell     -> Standard molded Epoxy @60Hz

    This is the lowest-manual-action path when every design uses the same object
    names. It does not assign voltages; use the transient manifest workflow for
    voltage/source setup.
    """
    call_api(ies, "Window_SetRefresh_OFF", 0, required=False)
    call_api(ies, "Window_SetUndo_OFF", 0, required=False)
    configure_common_electric_settings(ies)

    try:
        for object_name, material_name in DEFAULT_MATERIAL_MAP.items():
            assignment = PhysicsAssignment(
                role=f"default_material_{object_name}",
                object_name=object_name,
                geometry_type="Object",
                material=material_name,
                voltage_mode="none",
            )
            geometry_id = resolve_geometry_id(ies, assignment)
            print(
                f"\nDefault material row: object='{object_name}', material='{material_name}', "
                f"resolved geometry_id={geometry_id}"
            )
            assign_material(ies, assignment, geometry_id)
    finally:
        call_api(ies, "Window_SetUndo_ON", 0, required=False)
        call_api(ies, "Window_SetRefresh_ON", 0, required=False)
        call_api(ies, "Window_Refresh", required=False)

    print("\nDefault material assignment complete.")

def apply_transient_manifest(ies: Any, manifest_path: Path, run_solver: bool = False) -> None:
    assignments = load_transient_manifest(manifest_path)
    print(f"Loaded {len(assignments)} transient setup rows from: {manifest_path}")

    call_api(ies, "Window_SetRefresh_OFF", 0, required=False)
    call_api(ies, "Window_SetUndo_OFF", 0, required=False)
    configure_common_electric_settings(ies)

    try:
        for assignment in assignments:
            geometry_id = resolve_geometry_id(ies, assignment)
            print(
                f"\nRole '{assignment.role}': object='{assignment.object_name}', "
                f"geometry_type='{assignment.geometry_type}', resolved geometry_id={geometry_id}"
            )
            assign_material(ies, assignment, geometry_id)
            assign_voltage(ies, assignment, geometry_id)

        if run_solver:
            if yes_no("Delete existing solution before solving", default=True):
                call_api(ies, "Solution_DeleteSolution", 0, required=False)
            result = call_api(ies, "Solution_RunSolver", 0, required=True)
            if extract_error_code(result) != 0:
                raise RuntimeError("Solver returned a nonzero error code.")
    finally:
        call_api(ies, "Window_SetUndo_ON", 0, required=False)
        call_api(ies, "Window_SetRefresh_ON", 0, required=False)
        call_api(ies, "Window_Refresh", required=False)

    print("\nTransient manifest setup complete.")


def build_default_electro_only_assignments(source_name: str) -> list[PhysicsAssignment]:
    """Return the fixed bushing setup rows without reading any CSV.

    Required ELECTRO object names:
        conductor -> Copper + transient HV BIL voltage
        gnd       -> Copper + transient ground source named Ground
        shell     -> Standard molded Epoxy @60Hz, no voltage

    The transient source/waveform must already exist in the open ELECTRO model.
    """
    return [
        PhysicsAssignment(
            role="conductor_hv",
            object_name=DEFAULT_HV_OBJECT_NAME,
            geometry_type="Object",
            material=DEFAULT_MATERIAL_MAP["conductor"],
            voltage_name=DEFAULT_HV_VOLTAGE_NAME,
            voltage_mode="transient",
            source_name=source_name,
            color_r=255,
            color_g=0,
            color_b=0,
            active=1,
        ),
        PhysicsAssignment(
            role="ground_shield",
            object_name=DEFAULT_GROUND_OBJECT_NAME,
            geometry_type="Object",
            material=DEFAULT_MATERIAL_MAP["gnd"],
            voltage_name=DEFAULT_GROUND_VOLTAGE_NAME,
            voltage_mode="transient",
            source_name=DEFAULT_TRANSIENT_GROUND_SOURCE_NAME,
            color_r=0,
            color_g=0,
            color_b=255,
            active=1,
        ),
        PhysicsAssignment(
            role="epoxy_shell",
            object_name="shell",
            geometry_type="Object",
            material=DEFAULT_MATERIAL_MAP["shell"],
            voltage_mode="none",
            active=1,
        ),
    ]


def apply_physics_assignments(ies: Any, assignments: list[PhysicsAssignment], run_solver: bool = False) -> None:
    """Apply material and voltage assignments directly to the open ELECTRO model."""
    call_api(ies, "Window_SetRefresh_OFF", 0, required=False)
    call_api(ies, "Window_SetUndo_OFF", 0, required=False)
    configure_common_electric_settings(ies)

    try:
        for assignment in assignments:
            geometry_id = resolve_geometry_id(ies, assignment)
            print(
                f"\nRole '{assignment.role}': object='{assignment.object_name}', "
                f"geometry_type='{assignment.geometry_type}', resolved geometry_id={geometry_id}"
            )
            assign_material(ies, assignment, geometry_id)
            assign_voltage(ies, assignment, geometry_id)

        if run_solver:
            if yes_no("Delete existing solution before solving", default=True):
                call_api(ies, "Solution_DeleteSolution", 0, required=False)
            result = call_api(ies, "Solution_RunSolver", 0, required=True)
            if extract_error_code(result) != 0:
                raise RuntimeError("Solver returned a nonzero error code.")
    finally:
        call_api(ies, "Window_SetUndo_ON", 0, required=False)
        call_api(ies, "Window_SetRefresh_ON", 0, required=False)
        call_api(ies, "Window_Refresh", required=False)

    print("\nELECTRO-only transient setup complete.")


def apply_default_electro_only_transient_setup(ies: Any, run_solver: bool = False) -> None:
    """Apply transient source names to the existing source and gnd voltage definitions."""
    print("\nELECTRO-only pre-simulation setup")
    print("Materials:")
    for object_name, material_name in DEFAULT_MATERIAL_MAP.items():
        print(f"  {object_name} -> {material_name}")
    print("Voltages:")
    print(f"  {DEFAULT_HV_OBJECT_NAME} -> {DEFAULT_HV_VOLTAGE_NAME} -> transient source")
    print(f"  {DEFAULT_GROUND_OBJECT_NAME} -> {DEFAULT_GROUND_VOLTAGE_NAME} -> transient source {DEFAULT_TRANSIENT_GROUND_SOURCE_NAME}")

    configure_common_electric_settings(ies)
    configure_transient_operation_mode(ies)

    source_name = prompt_text("Existing Sources-tree impulse/source name in ELECTRO", DEFAULT_TRANSIENT_SOURCE_NAME)
    assignments = build_default_electro_only_assignments(source_name)
    apply_physics_assignments(ies, assignments, run_solver=run_solver)


def build_default_electro_only_static_assignments(hv_voltage: float) -> list[PhysicsAssignment]:
    """Return the fixed bushing setup for a static electric simulation.

    Existing ELECTRO voltage definitions are addressed by their model-tree names:
        source -> static HV value entered by the user
        gnd    -> static 0 V

    The voltage definitions must already be attached to conductor and gnd,
    respectively. No geometry coordinates or segment IDs are used.
    """
    return [
        PhysicsAssignment(
            role="conductor_static_hv",
            object_name=DEFAULT_HV_OBJECT_NAME,
            geometry_type="Object",
            material=DEFAULT_MATERIAL_MAP["conductor"],
            voltage_name=DEFAULT_HV_VOLTAGE_NAME,
            voltage_mode="static",
            static_value=float(hv_voltage),
            color_r=255,
            color_g=0,
            color_b=0,
            active=1,
        ),
        PhysicsAssignment(
            role="ground_shield_static",
            object_name=DEFAULT_GROUND_OBJECT_NAME,
            geometry_type="Object",
            material=DEFAULT_MATERIAL_MAP["gnd"],
            voltage_name=DEFAULT_GROUND_VOLTAGE_NAME,
            voltage_mode="static",
            static_value=0.0,
            color_r=0,
            color_g=0,
            color_b=255,
            active=1,
        ),
        PhysicsAssignment(
            role="epoxy_shell_static",
            object_name="shell",
            geometry_type="Object",
            material=DEFAULT_MATERIAL_MAP["shell"],
            voltage_mode="none",
            active=1,
        ),
    ]


def apply_default_electro_only_static_setup(ies: Any, run_solver: bool = False) -> None:
    """Apply materials and static named-voltage values to the open model."""
    print("\nELECTRO-only static pre-simulation setup")
    print("Materials:")
    for object_name, material_name in DEFAULT_MATERIAL_MAP.items():
        print(f"  {object_name} -> {material_name}")
    print("Static voltages:")
    print(f"  {DEFAULT_HV_OBJECT_NAME} -> Voltage > {DEFAULT_HV_VOLTAGE_NAME} -> user-entered static voltage")
    print(f"  {DEFAULT_GROUND_OBJECT_NAME} -> Voltage > {DEFAULT_GROUND_VOLTAGE_NAME} -> 0 V static")

    hv_voltage = prompt_float(
        "Static voltage assigned to Voltage > source (kV)",
        DEFAULT_STATIC_HV_VOLTAGE,
    )

    configure_common_electric_settings(ies)
    configure_static_operation_and_solver(ies)

    assignments = build_default_electro_only_static_assignments(hv_voltage)
    apply_physics_assignments(ies, assignments, run_solver=run_solver)
    print("\nELECTRO-only static setup complete.")


def run_static_solver_only(ies: Any) -> None:
    """Apply required static/global settings and run the solver without changing assignments."""
    configure_common_electric_settings(ies)
    configure_static_operation_and_solver(ies)
    if yes_no("Delete existing solution before solving", default=True):
        call_api(ies, "Solution_DeleteSolution", 0, required=False)
    result = call_api(ies, "Solution_RunSolver", 0, required=True)
    if extract_error_code(result) != 0:
        raise RuntimeError("Static solver returned a nonzero error code.")
    print("\nStatic solver run complete.")


def run_transient_solver_only(ies: Any) -> None:
    """Apply required transient/global settings and run the solver."""
    configure_common_electric_settings(ies)
    configure_transient_operation_mode(ies)
    if yes_no("Delete existing solution before solving", default=True):
        call_api(ies, "Solution_DeleteSolution", 0, required=False)
    result = call_api(ies, "Solution_RunSolver", 0, required=True)
    if extract_error_code(result) != 0:
        raise RuntimeError("Solver returned a nonzero error code.")
    print("\nSolver run complete.")


# -----------------------------------------------------------------------------
# Post-simulation extraction and DSP analysis
# -----------------------------------------------------------------------------


def parse_segment_result(result: Any) -> tuple[int, int]:
    if isinstance(result, tuple):
        seg_id = int(result[0])
        err = extract_error_code(result)
    else:
        seg_id = int(result)
        err = 0
    return seg_id, err


def get_segment_id_from_seed(ies: Any, x: float, y: float) -> int:
    result = call_api(ies, "Geometry2D_GetSegment_FromPoint", x, y, 0, required=True)
    seg_id, err = parse_segment_result(result)
    if err != 0 or seg_id <= 0:
        raise RuntimeError(f"Could not detect valid segment from seed point ({x}, {y}).")
    return seg_id


def extract_e_field_from_segment(
    ies: Any,
    measurement: BoundaryMeasurement,
    output_dir: Path,
) -> Path:
    if measurement.segment_id is None:
        if measurement.seed_x is None or measurement.seed_y is None:
            raise ValueError(f"Measurement {measurement.label} has no segment_id or seed point.")
        measurement.segment_id = get_segment_id_from_seed(ies, measurement.seed_x, measurement.seed_y)

    print("\n--- Extracting boundary E-field ---")
    print(measurement)

    result = call_api(
        ies,
        "Analysis_Get2DElectricField_FromSegment",
        int(measurement.segment_id),
        int(measurement.side),
        int(measurement.num_points),
        [], [], [], [], [], [],
        0,
        required=True,
    )

    if not isinstance(result, tuple) or len(result) < 7:
        raise RuntimeError("Unexpected return format from Analysis_Get2DElectricField_FromSegment.")

    err = extract_error_code(result)
    if err != 0:
        raise RuntimeError(
            f"E-field extraction failed for {measurement.label}. "
            "Make sure the model has been solved before running extraction."
        )

    df = pd.DataFrame({
        "label": measurement.label,
        "segment_id": int(measurement.segment_id),
        "side": int(measurement.side),
        "x": np.array(result[0], dtype=float),
        "y": np.array(result[1], dtype=float),
        "distance": np.array(result[2], dtype=float),
        "E_tangential": np.array(result[3], dtype=float),
        "E_normal": np.array(result[4], dtype=float),
        "E_magnitude": np.array(result[5], dtype=float),
    })

    safe_label = "".join(c if c.isalnum() or c in "_-" else "_" for c in measurement.label)
    output_csv = output_dir / f"{safe_label}_seg{measurement.segment_id}_side{measurement.side}_E_field.csv"
    df.to_csv(output_csv, index=False)
    print("E-field data saved:", output_csv)
    return output_csv


def create_default_measurement_config(path: Path) -> None:
    df = pd.DataFrame([
        {
            "label": "measure_boundary_1",
            "segment_id": "",
            "seed_x": "",
            "seed_y": "",
            "side": 1,
            "num_points": DEFAULT_NUM_POINTS,
            "notes": "Fill segment_id OR seed_x/seed_y. side: 1=left, 2=right.",
        },
        {
            "label": "measure_boundary_2",
            "segment_id": "",
            "seed_x": "",
            "seed_y": "",
            "side": 2,
            "num_points": DEFAULT_NUM_POINTS,
            "notes": "Add rows for conductor/shield/epoxy/shed boundaries.",
        },
    ])
    df.to_csv(path, index=False)


def load_measurement_config(path: Path) -> list[BoundaryMeasurement]:
    df = pd.read_csv(path)
    required_cols = {"label", "side", "num_points"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Measurement config missing columns: {sorted(missing)}")

    measurements: list[BoundaryMeasurement] = []
    for _, row in df.iterrows():
        label = str(row["label"]).strip()
        if not label or label.lower() == "nan":
            continue

        def opt_int(value: Any) -> Optional[int]:
            if pd.isna(value) or str(value).strip() == "":
                return None
            return int(float(value))

        def opt_float(value: Any) -> Optional[float]:
            if pd.isna(value) or str(value).strip() == "":
                return None
            return float(value)

        measurements.append(BoundaryMeasurement(
            label=label,
            segment_id=opt_int(row.get("segment_id", None)),
            seed_x=opt_float(row.get("seed_x", None)),
            seed_y=opt_float(row.get("seed_y", None)),
            side=int(row["side"]),
            num_points=int(row["num_points"]),
        ))

    if not measurements:
        raise ValueError("No valid measurements found in config CSV.")
    return measurements


def interactive_measurement_entry() -> list[BoundaryMeasurement]:
    measurements: list[BoundaryMeasurement] = []
    count = prompt_int("How many boundary measurements do you want to extract", 1)

    for i in range(count):
        print(f"\nMeasurement {i + 1} of {count}")
        label = prompt_text("Label", f"boundary_{i + 1}")
        method = prompt_text("Selection method: segment_id or seed", "seed").lower()
        side_choice = prompt_text("Side [1=left, 2=right, 3=both]", "1")
        num_points = prompt_int("Number of sample points", DEFAULT_NUM_POINTS)

        sides = [1, 2] if side_choice == "3" else [int(side_choice)]

        if method == "segment_id":
            seg_id = prompt_int("Segment ID")
            for side in sides:
                measurements.append(BoundaryMeasurement(label=f"{label}_side{side}", segment_id=seg_id, side=side, num_points=num_points))
        else:
            seed_x = prompt_float("Seed x-coordinate near/on boundary")
            seed_y = prompt_float("Seed y-coordinate near/on boundary")
            for side in sides:
                measurements.append(BoundaryMeasurement(label=f"{label}_side{side}", seed_x=seed_x, seed_y=seed_y, side=side, num_points=num_points))

    return measurements


def legacy_epoxy_outside_measurement_prompt() -> list[BoundaryMeasurement]:
    print("\n--- Legacy epoxy outside extraction setup ---")
    print(f"Target object label retained from original script: {EPOXY_OUTSIDE_OBJECT_NAME}")
    x_on_seg = prompt_float("Enter x-coordinate ON epoxy outside segment")
    y_on_seg = prompt_float("Enter y-coordinate ON epoxy outside segment")
    side_choice = prompt_text("Enter side [1=left, 2=right, 3=both]", "3")
    num_points = prompt_int("Number of sample points", DEFAULT_NUM_POINTS)

    sides = [1, 2] if side_choice == "3" else [int(side_choice)]
    return [
        BoundaryMeasurement(
            label=f"{EPOXY_OUTSIDE_OBJECT_NAME.replace(' ', '_')}_side{side}",
            seed_x=x_on_seg,
            seed_y=y_on_seg,
            side=side,
            num_points=num_points,
        )
        for side in sides
    ]


def named_measure_object_prompt(ies: Any) -> Optional[list[BoundaryMeasurement]]:
    """Attempt to use an object named 'measure'.

    The documented extraction call still needs a segment ID. This function tries
    a few likely API calls, but if your ELECTRO API does not expose reverse
    object-to-segment enumeration, it will tell you to use config/seed instead.
    """
    measure_name = prompt_text("Named measurement object", DEFAULT_MEASURE_OBJECT)
    side_choice = prompt_text("Side [1=left, 2=right, 3=both]", "1")
    num_points = prompt_int("Number of sample points", DEFAULT_NUM_POINTS)
    sides = [1, 2] if side_choice == "3" else [int(side_choice)]

    # Validate object existence if possible.
    call_api(ies, "Object_GetObjectID", measure_name, 0, required=False)

    candidate_methods = [
        "Object_GetIDsOfSegments",
        "Object_GetSegments",
        "Object_GetSegmentIDs",
    ]

    for method in candidate_methods:
        result = call_api(ies, method, measure_name, [], 0, required=False)
        if result is None:
            continue

        # Best-effort segment parsing.
        possible_array = None
        if isinstance(result, tuple):
            for item in result:
                if isinstance(item, (list, tuple)):
                    possible_array = item
                    break
        elif isinstance(result, (list, tuple)):
            possible_array = result

        if possible_array:
            seg_ids = [int(s) for s in possible_array if int(s) > 0]
            if seg_ids:
                measurements: list[BoundaryMeasurement] = []
                for seg_id in seg_ids:
                    for side in sides:
                        measurements.append(BoundaryMeasurement(
                            label=f"{measure_name}_seg{seg_id}_side{side}",
                            segment_id=seg_id,
                            side=side,
                            num_points=num_points,
                        ))
                return measurements

    print("\nCould not extract segment IDs directly from the named 'measure' object.")
    print("Use option 1, 3, or 4 instead: config CSV, interactive seed entry, or legacy seed extraction.")
    return None


def compute_spatial_metrics(csv_path: Path, output_dir: Path) -> Path:
    df = pd.read_csv(csv_path)
    signal = df["E_magnitude"].to_numpy(dtype=float)
    distance = df["distance"].to_numpy(dtype=float)

    idx_max = int(np.nanargmax(signal))
    emax = float(signal[idx_max])
    eavg = float(np.nanmean(signal))
    emin = float(np.nanmin(signal))

    threshold = 0.90 * emax
    above = signal >= threshold
    high_field_length = 0.0
    if len(distance) > 1:
        increments = np.diff(distance)
        high_field_length = float(np.sum(increments[above[:-1]]))

    summary = pd.DataFrame([{
        "source_csv": str(csv_path),
        "label": str(df["label"].iloc[0]) if "label" in df else csv_path.stem,
        "segment_id": int(df["segment_id"].iloc[0]) if "segment_id" in df else np.nan,
        "side": int(df["side"].iloc[0]) if "side" in df else np.nan,
        "E_min": emin,
        "E_max": emax,
        "E_avg": eavg,
        "E_max_distance": float(distance[idx_max]),
        "E_max_x": float(df["x"].iloc[idx_max]),
        "E_max_y": float(df["y"].iloc[idx_max]),
        "high_field_threshold_90pct_Emax": threshold,
        "high_field_length_approx": high_field_length,
    }])

    out = output_dir / f"{csv_path.stem}_field_summary.csv"
    summary.to_csv(out, index=False)
    print("Field summary saved:", out)
    return out


def run_fft_and_wavelet(csv_path: Path, output_dir: Path, signal_column: str) -> None:
    df = pd.read_csv(csv_path)

    x = df["distance"].to_numpy(dtype=float)
    y = df[signal_column].to_numpy(dtype=float)
    x_coord = df["x"].to_numpy(dtype=float)
    y_coord = df["y"].to_numpy(dtype=float)

    if len(x) < 4:
        print(f"Skipping DSP for {csv_path.name}: not enough points.")
        return

    dx = float(np.mean(np.diff(x)))
    if not np.isfinite(dx) or dx == 0:
        print(f"Skipping DSP for {csv_path.name}: invalid distance spacing.")
        return

    y_centered = y - np.nanmean(y)
    signal_name = f"{csv_path.stem}_{signal_column}"

    # Spatial FFT
    freq = np.fft.rfftfreq(len(y_centered), d=dx)
    mag = np.abs(np.fft.rfft(y_centered))

    fft_df = pd.DataFrame({"spatial_frequency": freq, "fft_magnitude": mag})
    fft_csv = output_dir / f"{signal_name}_spatial_fft.csv"
    fft_df.to_csv(fft_csv, index=False)

    if len(freq) > 1:
        dominant_index = int(np.argmax(mag[1:]) + 1)
        dominant_spatial_frequency = float(freq[dominant_index])
        dominant_fft_magnitude = float(mag[dominant_index])
        dominant_wavelength = float(1.0 / dominant_spatial_frequency) if dominant_spatial_frequency != 0 else np.inf
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
        "frequency_units": "cycles per ELECTRO length unit",
    }])
    summary_df.to_csv(output_dir / f"{signal_name}_fft_summary.csv", index=False)

    plt.figure()
    plt.plot(freq, mag)
    plt.xlabel("Spatial frequency")
    plt.ylabel("FFT magnitude")
    plt.title(f"{signal_name} Spatial FFT")
    plt.grid(True)
    if np.isfinite(dominant_spatial_frequency):
        plt.axvline(dominant_spatial_frequency, linestyle="--")
        plt.text(dominant_spatial_frequency, dominant_fft_magnitude, f"λ={dominant_wavelength:.3g}", rotation=90, verticalalignment="bottom")
    plt.savefig(output_dir / f"{signal_name}_spatial_fft.png", dpi=300)
    plt.close()

    # Wavelet analysis
    if PYWT_AVAILABLE:
        scales = np.arange(1, 41)
        coef, _ = pywt.cwt(y_centered, scales, "morl", sampling_period=dx)
        wavelet_mag = np.abs(coef)

        plt.figure()
        plt.imshow(wavelet_mag, aspect="auto", origin="lower", extent=[x.min(), x.max(), scales.min(), scales.max()])
        plt.xlabel("Distance")
        plt.ylabel("Wavelet scale")
        plt.title(f"{signal_name} Wavelet Magnitude")
        plt.colorbar(label="Magnitude")
        plt.savefig(output_dir / f"{signal_name}_wavelet.png", dpi=300)
        plt.close()

        max_over_scales = wavelet_mag.max(axis=0)
        scale_at_max = scales[wavelet_mag.argmax(axis=0)]
        hotspot_df = pd.DataFrame({
            "distance": x,
            "x_coordinate": x_coord,
            "y_coordinate": y_coord,
            "signal_value": y,
            "max_wavelet_magnitude": max_over_scales,
            "wavelet_scale_at_max": scale_at_max,
        }).sort_values(by="max_wavelet_magnitude", ascending=False)

        selected_rows = []
        min_sep = 0.02 * (x.max() - x.min())
        for _, row in hotspot_df.iterrows():
            if len(selected_rows) >= 10:
                break
            if all(abs(row["distance"] - selected["distance"]) >= min_sep for selected in selected_rows):
                selected_rows.append(row)
        pd.DataFrame(selected_rows).to_csv(output_dir / f"{signal_name}_wavelet_hotspots.csv", index=False)
    else:
        print("PyWavelets not installed. Wavelet analysis skipped.")

    print(f"DSP complete for {signal_name}.")


def plot_combined_original_e_magnitude(e_csv_files: Iterable[Path], output_dir: Path) -> None:
    plt.figure()
    for csv_path in e_csv_files:
        df = pd.read_csv(csv_path)
        plt.plot(df["distance"].to_numpy(dtype=float), df["E_magnitude"].to_numpy(dtype=float), label=csv_path.stem[:55])
    plt.xlabel("Distance")
    plt.ylabel("E magnitude")
    plt.title("Original E Magnitude vs Distance")
    plt.grid(True)
    plt.legend(fontsize=7)
    output_png = output_dir / "combined_original_E_magnitude_vs_distance.png"
    plt.savefig(output_png, dpi=300, bbox_inches="tight")
    plt.close()
    print("Combined E magnitude plot saved:", output_png)


def post_simulation_extract_and_analyze(ies: Any) -> None:
    model_path = get_current_model_path(ies)
    output_dir = make_output_dir(model_path, "postsim_extract_dsp")

    print("\nPost-simulation measurement selection:")
    print("1. Load measurement config CSV")
    print("2. Create blank measurement config CSV and stop")
    print("3. Interactive segment/seed entry")
    print("4. Legacy single epoxy outside seed-point extraction")
    print("5. Attempt named 'measure' object extraction")

    choice = input("Enter 1, 2, 3, 4, or 5: ").strip()

    if choice == "1":
        path = Path(input("Enter full path to measurement config CSV: ").strip().strip('"'))
        measurements = load_measurement_config(path)
    elif choice == "2":
        config_path = output_dir / "measurement_config_template.csv"
        create_default_measurement_config(config_path)
        print("Template created:", config_path)
        print("Fill this file, then rerun option 1.")
        return
    elif choice == "3":
        measurements = interactive_measurement_entry()
    elif choice == "4":
        measurements = legacy_epoxy_outside_measurement_prompt()
    elif choice == "5":
        maybe_measurements = named_measure_object_prompt(ies)
        if not maybe_measurements:
            return
        measurements = maybe_measurements
    else:
        print("Invalid selection.")
        return

    e_csv_files: list[Path] = []
    for measurement in measurements:
        e_csv_files.append(extract_e_field_from_segment(ies, measurement, output_dir))

    summary_files = [compute_spatial_metrics(csv_path, output_dir) for csv_path in e_csv_files]
    if summary_files:
        combined_summary = pd.concat([pd.read_csv(p) for p in summary_files], ignore_index=True)
        combined_summary.to_csv(output_dir / "combined_field_summary.csv", index=False)

    plot_combined_original_e_magnitude(e_csv_files, output_dir)

    signal_columns = ["E_magnitude", "E_normal", "E_tangential"]
    for e_csv in e_csv_files:
        for col in signal_columns:
            run_fft_and_wavelet(e_csv, output_dir, col)

    print("\nPost-simulation extraction and DSP analysis complete.")
    print("Results saved to:", output_dir)


# -----------------------------------------------------------------------------
# Main menu: separated pre-simulation / solver / post-simulation phases
# -----------------------------------------------------------------------------


def main() -> None:
    ies = win32com.client.Dispatch("IES.Document")
    print("Connected to ELECTRO.")
    print("Use an open ELECTRO template model with objects named conductor, gnd, and shell.")
    print("For transient setup, Sources > impulse and Sources > Ground must exist in the open ELECTRO model.")
    print("For static setup, Voltage > source and Voltage > gnd must already be attached to conductor and gnd.")

    while True:
        print("\nSelect mode:")
        print("1. Transient pre-simulation setup: materials + source->impulse + gnd->Ground")
        print("2. Full transient run: setup, solve, then optionally post-process")
        print("3. Run transient solver only")
        print("4. Static pre-simulation setup: materials + named static voltages")
        print("5. Full static run: setup and solve")
        print("6. Run static solver only")
        print("7. Post-simulation boundary E-field extraction + DSP analysis")
        print("8. Create measurement config template only")
        print("9. Apply default materials only: conductor/gnd/shell by object name")
        print("0. Exit")

        choice = input("Enter choice: ").strip()

        if choice == "1":
            apply_default_electro_only_transient_setup(ies, run_solver=False)
            print("\nTransient pre-simulation setup complete. Solver was NOT run.")
        elif choice == "2":
            apply_default_electro_only_transient_setup(ies, run_solver=True)
            if yes_no("Continue into post-simulation extraction + DSP", default=True):
                post_simulation_extract_and_analyze(ies)
        elif choice == "3":
            run_transient_solver_only(ies)
        elif choice == "4":
            apply_default_electro_only_static_setup(ies, run_solver=False)
            print("\nStatic pre-simulation setup complete. Solver was NOT run.")
        elif choice == "5":
            apply_default_electro_only_static_setup(ies, run_solver=True)
        elif choice == "6":
            run_static_solver_only(ies)
        elif choice == "7":
            post_simulation_extract_and_analyze(ies)
        elif choice == "8":
            model_path = get_current_model_path(ies)
            out = make_output_dir(model_path, "config_template") / "measurement_config_template.csv"
            create_default_measurement_config(out)
            print("Template created:", out)
        elif choice == "9":
            apply_default_materials_by_name(ies)
        elif choice == "0":
            print("Exiting.")
            break
        else:
            print("Invalid selection.")


if __name__ == "__main__":
    main()
