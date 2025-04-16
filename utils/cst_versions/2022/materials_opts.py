from utils.macors_canva import Canvas
from tabulate import tabulate



def define_material(csth, material_name, materials_path="materials"):
    # check if material_name is in csth.materials
    if material_name in csth.materials:
        redefine = True
        print(f"[INFO] redefining material: {material_name}")
    else:
        redefine = False
        print(f"[INFO] Defining material: {material_name}")

    csth.materials[material_name] = {
        "freq": [],
        "re": [],
        "im": []
    }

    wl_min = csth.crr_prj_properties["wavelength_min"]
    wl_max = csth.crr_prj_properties["wavelength_max"]
    freq_min = 300 / wl_max # in THz
    freq_max = 300 / wl_min # in THz

    # Add tolerance
    tolerance = 0.02
    freq_min *= 1 - tolerance
    freq_max *= 1 + tolerance

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
            if float(freq) < freq_min or float(freq) > freq_max:
                continue
            canvas.add_code("Material", "AddDispersionFittingValueEps", freq, re, im, "1.0")
            csth.materials[material_name]["freq"].append(freq)
            csth.materials[material_name]["re"].append(re)
            csth.materials[material_name]["im"].append(im)

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
    res = canvas.send(csth, cmt=f"Define material {material_name}")
    if res:
        print(f"[ OK ] Material {material_name} defined successfully")
    else:
        print(f"[ERRO] Failed to define material {material_name}")
        raise RuntimeError(f"Failed to define material {material_name}")

    print(f"[INFO] A New material {material_name} has been added to the project:")
    print(tabulate(csth.materials[material_name], headers="keys", tablefmt="pretty"))
    print("\n")


def change_substrate(csth, material_name, materials_path="materials", component="component1", force_define=False):
    # check if material_name is in csth.materials
    if material_name not in csth.materials:
        print(f"[WARN] Material {material_name} not found in materials list")
        print(f"[INFO] Program will try to define the material {material_name} first")
        define_material(csth, material_name, materials_path)
    elif force_define:
        print(f"[WARN] Material {material_name} already defined, but force_define is set to True")
        print(f"[INFO] Program will try to redefine the material {material_name} first")
        define_material(csth, material_name, materials_path)
    else:
        print(f"[INFO] Material {material_name} already defined")

    canvas = Canvas()
    # Solid.ChangeMaterial "component1:substrate", "Material"
    res = canvas.write_send(csth, 
                            f"Solid.ChangeMaterial \"{component}:substrate\", \"{material_name}\"",
                            "Change substrate material"
                            )
    if res:
        print(f"[INFO] Substrate material changed to {material_name}")
    else:
        print(f"[ERRO] Failed to change substrate material to {material_name}")
        raise RuntimeError("Failed to change substrate material")

    csth.crr_prj_properties["substrate_material"] = material_name
    # update crr_prj_properties to prjs list (csth.prjs = pd.DataFrame(columns=["project_instance", "project_properties"]))
    csth.update_project_properties()


def change_pillar(csth, material_name, materials_path="materials", component="component1", force_define=False):
    # check if material_name is in csth.materials
    if material_name not in csth.materials:
        print(f"[WARN] Material {material_name} not found in materials list")
        print(f"[INFO] Program will try to define the material {material_name} first")
        define_material(csth, material_name, materials_path)
    elif force_define:
        print(f"[WARN] Material {material_name} already defined, but force_define is set to True")
        print(f"[INFO] Program will try to redefine the material {material_name} first")
        define_material(csth, material_name, materials_path)
    else:
        print(f"[INFO] Material {material_name} already defined")

    canvas = Canvas()
    # Solid.ChangeMaterial "component1:pillar", "Material"
    res = canvas.write_send(csth, 
                                    f"Solid.ChangeMaterial \"{component}:pillar\", \"{material_name}\"",
                                    "Change pillar material"
                                    )
    if res:
        print(f"[INFO] Pillar material changed to {material_name}")
    else:
        print(f"[ERRO] Failed to change pillar material to {material_name}")
        raise RuntimeError("Failed to change pillar material")

    csth.crr_prj_properties["pillar_material"] = material_name
    # update crr_prj_properties to prjs list (csth.prjs = pd.DataFrame(columns=["project_instance", "project_properties"]))
    csth.update_project_properties()