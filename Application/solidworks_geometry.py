import csv
import os
import win32com.client

SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2
MM_PER_M = 1000.0

OUTPUT_CSV = "solidworks_extracted_features.csv"


def get_com(obj, attr):
    value = getattr(obj, attr)
    return value() if callable(value) else value


def call_com(obj, attr, *args):
    value = getattr(obj, attr)
    return value(*args) if callable(value) else value


def get_doc_type_from_path(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".sldprt":
        return SW_DOC_PART
    if ext == ".sldasm":
        return SW_DOC_ASSEMBLY
    return None


def open_solidworks_doc(sw_app, path):
    doc_type = get_doc_type_from_path(path)

    if doc_type is None or not path:
        return None

    try:
        errors = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0
        )
        warnings = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0
        )

        return sw_app.OpenDoc6(path, doc_type, 1, "", errors, warnings)

    except Exception:
        return None


def get_box_dimensions_mm(box):
    xmin, ymin, zmin, xmax, ymax, zmax = box

    dx = abs(xmax - xmin) * MM_PER_M
    dy = abs(ymax - ymin) * MM_PER_M
    dz = abs(zmax - zmin) * MM_PER_M

    dims = sorted([dx, dy, dz])

    # Shield-only softened diameter estimate.
    # This reduces the influence of inserts/protrusions extending beyond the shield body.
    x_avg_span = (abs(xmin) + abs(xmax)) * MM_PER_M
    y_avg_span = (abs(ymin) + abs(ymax)) * MM_PER_M
    z_avg_span = (abs(zmin) + abs(zmax)) * MM_PER_M

    avg_spans = sorted([x_avg_span, y_avg_span, z_avg_span])

    return {
        "dx": dx,
        "dy": dy,
        "dz": dz,
        "small_dim_mm": dims[0],
        "mid_dim_mm": dims[1],
        "large_dim_mm": dims[2],
        "diameter_avg_mm": (dims[0] + dims[1]) / 2.0,
        "diameter_max_cross_mm": dims[1],
        "shield_diameter_avg_mm": avg_spans[1],
        "length_like_mm": dims[2],
        "aspect_ratio": dims[2] / max(((dims[0] + dims[1]) / 2.0), 1e-9),
    }


def highlight_component(model, comp):
    if comp is None:
        return

    comp_name = get_com(comp, "Name2")

    try:
        call_com(model, "ClearSelection2", True)
    except Exception:
        pass

    selected = False

    try:
        selected = comp.Select(False)
    except Exception:
        selected = False

    if not selected:
        try:
            selected = comp.Select2(False, 0)
        except Exception:
            selected = False

    if selected:
        print(f"Selected marker component in SolidWorks: {comp_name}")
        try:
            model.ViewZoomtoSelection2()
        except Exception:
            pass
    else:
        print(f"Could not highlight component: {comp_name}")


def scan_components(sw_app, parent_doc, components, parent_label="Top Assembly", marker_component=None):
    candidates = []

    for comp in components:
        try:
            name = get_com(comp, "Name2")
            path = get_com(comp, "GetPathName")
            box = comp.GetBox(False, False)

            current_marker = marker_component if marker_component is not None else comp

            if box is not None:
                dims = get_box_dimensions_mm(box)

                candidates.append({
                    "component": comp,
                    "marker_component": current_marker,
                    "owner_doc": parent_doc,
                    "name": name,
                    "path": path,
                    "source": parent_label,
                    **dims
                })

            if path and path.lower().endswith(".sldasm"):
                sub_doc = open_solidworks_doc(sw_app, path)

                if sub_doc is not None:
                    sub_components = sub_doc.GetComponents(True)

                    if sub_components:
                        candidates.extend(
                            scan_components(
                                sw_app,
                                sub_doc,
                                sub_components,
                                parent_label=f"Inside {name}",
                                marker_component=comp
                            )
                        )

        except Exception as e:
            print(f"Skipped component during scan: {e}")

    return candidates


def prompt_user_choice(title, options, value_key):
    print(f"\n{title}")
    print("-" * 125)

    for i, c in enumerate(options, start=1):
        print(
            f"{i:<3} "
            f"{c['name'][:34]:<36} "
            f"value={c[value_key]:>10.3f} mm | "
            f"dx={c['dx']:>8.3f}, dy={c['dy']:>8.3f}, dz={c['dz']:>8.3f}, "
            f"aspect={c['aspect_ratio']:>6.2f} | "
            f"{c['source']}"
        )

    while True:
        choice = input(f"\nPick option number for {title}: ")

        try:
            index = int(choice)

            if 1 <= index <= len(options):
                return options[index - 1]

            print("Invalid number. Try again.")

        except ValueError:
            print("Enter a number only.")


def make_no_shield_option(model):
    return {
        "component": None,
        "marker_component": None,
        "owner_doc": model,
        "name": "NO SHIELD",
        "path": "",
        "source": "Manual option",
        "dx": 0,
        "dy": 0,
        "dz": 0,
        "small_dim_mm": 0,
        "mid_dim_mm": 0,
        "large_dim_mm": 0,
        "diameter_avg_mm": 0,
        "diameter_max_cross_mm": 0,
        "shield_diameter_avg_mm": 0,
        "length_like_mm": 0,
        "aspect_ratio": 0,
    }


def main():
    sw_app = win32com.client.Dispatch("SldWorks.Application")
    sw_app.Visible = True

    model = sw_app.ActiveDoc

    if model is None:
        print("No SolidWorks document is open.")
        return

    if get_com(model, "GetType") != SW_DOC_ASSEMBLY:
        print("Open an assembly first.")
        return

    assembly_name = get_com(model, "GetTitle")
    print(f"\nAssembly: {assembly_name}")

    top_components = model.GetComponents(True)

    candidates = scan_components(
        sw_app,
        model,
        top_components,
        parent_label="Top Assembly"
    )

    if not candidates:
        print("No measurable components found.")
        return

    print("\nAll detected component geometry:")
    print("-" * 125)
    print(f"{'#':<4}{'Component':<36}{'dx':>10}{'dy':>10}{'dz':>10}{'Aspect':>10}   Source")
    print("-" * 125)

    for i, c in enumerate(candidates, start=1):
        print(
            f"{i:<4}"
            f"{c['name'][:34]:<36}"
            f"{c['dx']:>10.3f}"
            f"{c['dy']:>10.3f}"
            f"{c['dz']:>10.3f}"
            f"{c['aspect_ratio']:>10.2f}   "
            f"{c['source']}"
        )

    conductor_options = [
        c for c in candidates
        if c["aspect_ratio"] > 10
        and c["length_like_mm"] > 500
        and c["diameter_avg_mm"] > 10
    ]

    conductor_options.sort(key=lambda c: c["diameter_avg_mm"])

    if not conductor_options:
        print("No conductor options found.")
        return

    conductor = prompt_user_choice(
        "Conductor Diameter Options",
        conductor_options,
        "diameter_avg_mm"
    )

    highlight_component(model, conductor["marker_component"])

    shield_diameter_options = [
        c for c in candidates
        if c["shield_diameter_avg_mm"] > 75
        and c["length_like_mm"] > 25
        and c["length_like_mm"] < 900
        and c["aspect_ratio"] < 8
    ]

    shield_diameter_options.sort(
        key=lambda c: (
            c["shield_diameter_avg_mm"],
            c["length_like_mm"]
        ),
        reverse=True
    )

    no_shield_option = make_no_shield_option(model)
    shield_diameter_options.insert(0, no_shield_option)

    shield_diameter_candidate = prompt_user_choice(
        "Shield Diameter Options",
        shield_diameter_options,
        "shield_diameter_avg_mm"
    )

    if shield_diameter_candidate["name"] == "NO SHIELD":
        shield_length_candidate = no_shield_option
    else:
        highlight_component(model, shield_diameter_candidate["marker_component"])

        shield_length_options = [
            c for c in candidates
            if c["shield_diameter_avg_mm"] > 75
            and c["length_like_mm"] > 25
            and c["length_like_mm"] < 900
            and c["aspect_ratio"] < 8
        ]

        shield_length_options.insert(0, no_shield_option)
        shield_length_options.sort(key=lambda c: c["length_like_mm"], reverse=True)

        shield_length_candidate = prompt_user_choice(
            "Shield Length Options",
            shield_length_options,
            "length_like_mm"
        )

        if shield_length_candidate["name"] != "NO SHIELD":
            highlight_component(model, shield_length_candidate["marker_component"])

    conductor_diameter = conductor["diameter_avg_mm"]
    shield_diameter = shield_diameter_candidate["shield_diameter_avg_mm"]
    shield_length = shield_length_candidate["length_like_mm"]

    print("\nFinal selected extracted features:")
    print("-" * 60)
    print(f"Conductor component: {conductor['name']}")
    print(f"Conductor diameter:  {conductor_diameter:.3f} mm")
    print()
    print(f"Shield diameter component: {shield_diameter_candidate['name']}")
    print(f"Shield diameter:           {shield_diameter:.3f} mm")
    print()
    print(f"Shield length component:   {shield_length_candidate['name']}")
    print(f"Shield length:             {shield_length:.3f} mm")

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        writer.writerow([
            "Assembly",
            "Conductor Component",
            "Conductor Diameter mm",
            "Shield Diameter Component",
            "Shield Diameter mm",
            "Shield Length Component",
            "Shield Length mm"
        ])

        writer.writerow([
            assembly_name,
            conductor["name"],
            round(conductor_diameter, 6),
            shield_diameter_candidate["name"],
            round(shield_diameter, 6),
            shield_length_candidate["name"],
            round(shield_length, 6)
        ])

    print(f"\nSaved results to: {OUTPUT_CSV}")


if __name__ == "__main__":
    main()