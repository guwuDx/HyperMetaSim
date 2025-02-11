from utils.macors_canva import Canvas

def set_prj_wavelength(cst_handler, wavelength_min, wavelength_max):
    print(f"[INFO] Project wavelength would be set to {wavelength_min} - {wavelength_max}")

    cst_handler.crr_prj_properties["wavelegnth_min"] = wavelength_min
    cst_handler.crr_prj_properties["wavelegnth_max"] = wavelength_max

    canvas = Canvas()
    canvas.add_code("Solver", "WavelengthRange", wavelength_min, wavelength_max)
    
    print("[INFO] vba code to be executed:\n")
    canvas.preview()
    res = canvas.send(cst_handler)

    if res:
        print("[ OK ] Project wavelength set successfully")
    else:
        print("[ERRO] Failed to set project wavelength")
        raise RuntimeError("Failed to set project wavelength")


def define_material(cst_handler, materials_path, material_name):
    print(f"[INFO] Defining material: {material_name}")

    canvas = Canvas()
    canvas.add_code("Material", "Reset")
    canvas.add_code("Material", "Name", material_name)
    canvas.add_code("Material", "Folder", "")
    canvas.add_code("Material", "Rho", "0.0")
    canvas.add_code("Material", "ThermalType", "Normal")
    canvas.add_code("Material", "ThermalConductivity", "0")
    canvas.add_code("Material", "SpecificHeat", "0", "J/K/kg")
    canvas.add_code("Material", "DynamicViscosity", "0")
    canvas.add_code("Material", "Emissivity", "0")
    canvas.add_code("Material", "MetabolicRate", "0.0")
    canvas.add_code("Material", "VoxelConvection", "0.0")
    canvas.add_code("Material", "BloodFlow", "0")
    canvas.add_code("Material", "MechanicsType", "Unused")
    canvas.add_code("Material", "IntrinsicCarrierDensity", "0")
    canvas.add_code("Material", "FrqType", "all")
    canvas.add_code("Material", "Type", "Normal")
    canvas.add_code("Material", "MaterialUnit", "Frequency", "THz")
    canvas.add_code("Material", "MaterialUnit", "Geometry", "um")
    canvas.add_code("Material", "MaterialUnit", "Time", "ns")
    canvas.add_code("Material", "MaterialUnit", "Temperature", "Kelvin")
    canvas.add_code("Material", "Epsilon", "1")
    canvas.add_code("Material", "Mu", "1")
    canvas.add_code("Material", "Sigma", "0")
    canvas.add_code("Material", "TanD", "0.0")
    canvas.add_code("Material", "TanDFreq", "0.0")
    canvas.add_code("Material", "TanDGiven", "False")
    canvas.add_code("Material", "TanDModel", "ConstTanD")
    canvas.add_code("Material", "SetConstTanDStrategyEps", "AutomaticOrder")
    canvas.add_code("Material", "ConstTanDModelOrderEps", "3")
    canvas.add_code("Material", "DjordjevicSarkarUpperFreqEps", "0")
    canvas.add_code("Material", "SetElParametricConductivity", "False")
    canvas.add_code("Material", "ReferenceCoordSystem", "Global")
    canvas.add_code("Material", "CoordSystemType", "Cartesian")
    canvas.add_code("Material", "SigmaM", "0")
    canvas.add_code("Material", "TanDM", "0.0")
    canvas.add_code("Material", "TanDMFreq", "0.0")
    canvas.add_code("Material", "TanDMGiven", "False")
    canvas.add_code("Material", "TanDMModel", "ConstTanD")
    canvas.add_code("Material", "SetConstTanDStrategyMu", "AutomaticOrder")
    canvas.add_code("Material", "ConstTanDModelOrderMu", "3")
    canvas.add_code("Material", "DjordjevicSarkarUpperFreqMu", "0")
    canvas.add_code("Material", "SetMagParametricConductivity", "False")
    canvas.add_code("Material", "DispModelEps", "None")
    canvas.add_code("Material", "DispModelMu", "None")
    canvas.add_code("Material", "DispersiveFittingSchemeEps", "Nth Order")
    canvas.add_code("Material", "MaximalOrderNthModelFitEps", "10")
    canvas.add_code("Material", "ErrorLimitNthModelFitEps", "0.1")
    canvas.add_code("Material", "DispersiveFittingSchemeMu", "Nth Order")
    canvas.add_code("Material", "MaximalOrderNthModelFitMu", "10")
    canvas.add_code("Material", "ErrorLimitNthModelFitMu", "0.1")
    canvas.add_code("Material", "DispersiveFittingFormatEps", "Real_Imag")

    with open(f"{materials_path}/{material_name}.csv", "r") as f:
        lines = f.readlines()
        for line in lines:
            if line.startswith("#"): continue

            freq, re, im = line.replace(" ", "").replace("\n", "").split(",")
            canvas.add_code("Material", "AddDispersionFittingValueEps", freq, re, im, "1.0")

    canvas.add_code("Material", "UseGeneralDispersionEps", "True")
    canvas.add_code("Material", "UseGeneralDispersionMu", "False")
    canvas.add_code("Material", "NLAnisotropy", "False")
    canvas.add_code("Material", "NLAStackingFactor", "1")
    canvas.add_code("Material", "NLADirectionX", "1")
    canvas.add_code("Material", "NLADirectionY", "0")
    canvas.add_code("Material", "NLADirectionZ", "0")
    canvas.add_code("Material", "LatticeScattering", "Electron", "0.1", "0.")
    canvas.add_code("Material", "LatticeScattering", "Hole", "0.1", "0.")
    canvas.add_code("Material", "EffectiveMassForConductivity", "Electron", "0.25")
    canvas.add_code("Material", "EffectiveMassForConductivity", "Hole", "0.35")
    canvas.add_code("Material", "Colour", "0.701961", "0.301961", "0.301961")
    canvas.add_code("Material", "Wireframe", "False")
    canvas.add_code("Material", "Reflection", "False")
    canvas.add_code("Material", "Allowoutline", "True")
    canvas.add_code("Material", "Transparentoutline", "False")
    canvas.add_code("Material", "Transparency", "0")
    canvas.add_code("Material", "Create")

    # print("[INFO] vba code to be executed:\n")
    # canvas.preview()
    res = canvas.send(cst_handler)
    if res:
        print(f"[ OK ] Material {material_name} defined successfully")
    else:
        print(f"[ERRO] Failed to define material {material_name}")
        raise RuntimeError(f"Failed to define material {material_name}")


def update_param(cst_handler, force=False):
    print("[INFO] Updating parameters ...")

    canvas = Canvas()
    vba_code = f"RebuildOnParametricChange \"{force}\", \"True\""
    res = canvas.write_send(cst_handler, vba_code, None)
    return res


def modify_param(cst_handler, param_name: str, value: int):
    print(f"[INFO] Modifying parameter {param_name} to {value}")

    canvas = Canvas()
    vba_code = f"StoreParameter \"{param_name}\", \"{value}\""
    res = canvas.write_send(cst_handler, vba_code, None)
    if res:
        print(f"[ OK ] Parameter {param_name} modified successfully")
    else:
        print(f"[ERRO] Failed to modify parameter {param_name}, please check whether the parameter exists")
        raise RuntimeError(f"Failed to modify parameter {param_name}, please check whether the parameter exists")
    update_param(cst_handler)
