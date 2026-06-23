import win32com.client

ies = win32com.client.Dispatch("IES.Document")

print("Connected to ELECTRO.")

# -----------------------------
# User inputs
# -----------------------------
x_shift = float(input("Enter horizontal movement amount for GND: "))
voltage = float(input("Enter voltage to apply to conductor: "))

epoxy_x = float(input("Enter x-coordinate inside epoxy shell: "))
epoxy_y = float(input("Enter y-coordinate inside epoxy shell: "))

epoxy_material = "Standard molded Epoxy @60Hz"

# -----------------------------
# 1. Move GND horizontally
# -----------------------------
move_result = ies.Geometry2D_Displace("GND", x_shift, 0.0, 0)
print("Geometry2D_Displace returned:", move_result)

# -----------------------------
# 2. Apply voltage to conductor
# -----------------------------
voltage_result = ies.Physics_Set2DVoltage("conductor", voltage, 0)
print("Physics_Set2DVoltage returned:", voltage_result)

# -----------------------------
# 3. Assign epoxy shell object
# -----------------------------
region_id = ies.Geometry2D_GetRegion_FromPoint(epoxy_x, epoxy_y, 0)
print("Epoxy shell region ID:", region_id)

create_result = ies.Object_Create("epoxy shell", 0, 0)
print("Object_Create returned:", create_result)

add_result = ies.Object_AddRegion("epoxy shell", region_id, 0)
print("Object_AddRegion returned:", add_result)

# -----------------------------
# 4. Assign epoxy material
# -----------------------------
material_result = ies.Physics_SetMaterial("epoxy shell", epoxy_material, 0)
print("Physics_SetMaterial returned:", material_result)

ies.Window_Refresh()

print("Finished.")