from utils.macors_canva import Canvas

import threading



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
    wl_min = cst_handler.crr_prj_properties["wavelegnth_min"]
    wl_max = cst_handler.crr_prj_properties["wavelegnth_max"]
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
    res = canvas.send(cst_handler, cmt=f"Define material {material_name}")
    if res:
        print(f"[ OK ] Material {material_name} defined successfully")
    else:
        print(f"[ERRO] Failed to define material {material_name}")
        raise RuntimeError(f"Failed to define material {material_name}")


def update_params(cst_handler, force=False):
    print("[INFO] Updating parameters ...")

    canvas = Canvas()
    vba_code = f"RebuildOnParametricChange \"{force}\", \"True\""
    res = canvas.write_send(cst_handler, vba_code, None)

    if res:
        print("[ OK ] Parameters updated successfully")
    else:
        print("[ERRO] Failed to update parameters")
        raise RuntimeError("Failed to update parameters")
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
    update_params(cst_handler)


def set_acc_dc(cst_handler, solver="FDSolver"):
    print(f"[INFO] Setting solver to {solver}")

    canvas = Canvas()

    canvas.add_code(solver, "MPIParallelization", "False")
    canvas.add_code(solver, "UseDistributedComputing", "False")
    canvas.add_code(solver, "NetworkComputingStrategy", "RunRemote")
    canvas.add_code(solver, "NetworkComputingJobCount", "99")
    canvas.add_code(solver, "UseParallelization", "True")
    canvas.add_code(solver, "MaxCPUs", cst_handler._ACC_DC.max_threads)
    canvas.add_code(solver, "MaximumNumberOfCPUDevices", cst_handler._ACC_DC.max_num_of_cpu_devs)

    canvas.add_code("MeshSettings", "SetMeshType", "Unstr")
    canvas.add_code("MeshSettings", "Set", "UseDC", cst_handler._ACC_DC.remote_mesh)

    if cst_handler._ACC_DC.max_params_parallel <= 0:
        canvas.write("UseDistributedComputingForParameters \"False\"")
        canvas.write("MaxNumberOfDistributedComputingParameters \"1\"")
        canvas.write("ParameterSweep.UseDistributedComputing \"False\"")
    else:
        canvas.write("UseDistributedComputingForParameters \"True\"")
        canvas.write(f"MaxNumberOfDistributedComputingParameters \"{cst_handler._ACC_DC.max_params_parallel}\"")
        canvas.write("ParameterSweep.UseDistributedComputing \"True\"")
    canvas.write(f"UseDistributedComputingMemorySetting \"{cst_handler._ACC_DC.use_dc_mem_setting}\"")
    canvas.write(f"MinDistributedComputingMemoryLimit \"{cst_handler._ACC_DC.min_dc_mem_limit}\"")
    canvas.write(f"UseDistributedComputingSharedDirectory \"{cst_handler._ACC_DC.use_shared_dir}\"")
    canvas.write(f"OnlyConsider0D1DResultsForDC \"{cst_handler._ACC_DC.only_0D1D}\"")

    canvas.preview()
    canvas.send(cst_handler, "Set accerlation and distributed computing")


def set_FDSolver_source(cst_handler, source_port="Zmin", mode="TM(0,0)"):
    print("[INFO] Setting up FDSolver ...")
    canvas = Canvas()

    # vbac = Canvas.vba_template.get_mode_num_by_name(source_port, mode)
    # canvas.write(vbac)

    canvas.write(f"FDSolver.Stimulation \"{source_port}\", \"{mode}\"")

    canvas.preview()
    res = canvas.send(cst_handler, "Set FDSolver")
    if res:
        print("[ OK ] FDSolver set successfully")
    else:
        print("[ERRO] Failed to set FDSolver, please check whether the mode/port exists")
        raise RuntimeError("Failed to set FDSolver, please check whether the mode/port exists")


def cpu_monitor(interval:   int = 1,    # additonal check interval
                threshold:  int = 6,    # threshold for the CPU occupancy rate
                confidence: int = 10    # confidence level
                ):


def exec_paramSweep(cst_handler):
    print("[INFO] Executing parameter sweep ...")
    canvas = Canvas()
    canvas.write("ParameterSweep.Start")
    canvas.preview(0)
    res = canvas.send(cst_handler, add_to_history=False)
    if res:
        print("[INFO] Parameter sweep finished, please check the results")
    else:
        print("[ERRO] Failed to start parameter sweep")
        raise RuntimeError("Failed to start parameter sweep")