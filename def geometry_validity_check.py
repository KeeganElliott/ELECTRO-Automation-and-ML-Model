def geometry_validity_check(ies, epoxy_x, epoxy_y):
    print("\n--- Geometry Validity Check ---")

    passed = True

    # 1. Check epoxy region exists
    try:
        region_id = ies.Geometry2D_GetRegion_FromPoint(epoxy_x, epoxy_y, 0)
        print("Epoxy region ID:", region_id)

        if region_id <= 0:
            print("FAIL: Invalid epoxy region ID.")
            passed = False

    except Exception as e:
        print("FAIL: Could not find epoxy region.")
        print(e)
        return False

    # 2. Check epoxy area
    try:
        area_result = ies.Geometry2D_GetRegionArea(region_id, 0.0, 0)
        print("Epoxy region area result:", area_result)

        # Depending on COM return style, area_result may be a tuple or number
        if isinstance(area_result, tuple):
            area = area_result[0]
            err = area_result[-1]
        else:
            area = area_result
            err = 0

        print("Epoxy region area:", area)

        if err != 0 or area <= 0:
            print("FAIL: Epoxy region area invalid.")
            passed = False

    except Exception as e:
        print("WARNING: Could not check epoxy region area.")
        print(e)

    # 3. Check conductor voltage assignment still works with test value
    try:
        test_result = ies.Physics_Set2DVoltage("conductor", 1.0, 0)
        print("Conductor test voltage returned:", test_result)

        if test_result != 0:
            print("FAIL: Conductor object/segment voltage assignment failed.")
            passed = False

    except Exception as e:
        print("FAIL: Could not apply test voltage to conductor.")
        print(e)
        passed = False

    print("--- Validity Check Complete ---")

    return passed