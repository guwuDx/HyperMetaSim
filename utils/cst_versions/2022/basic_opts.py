from utils.macors_canva import Canvas
from utils import cst_handler
from utils import misc

from tabulate import tabulate
from multiprocessing import Process
import os
import sys
import time
import psutil
import platform
import contextlib



def set_prj_wavelength(csth, wavelength_min, wavelength_max):
    print(f"[INFO] Project wavelength would be set to {wavelength_min} - {wavelength_max}")

    csth.crr_prj_properties["wavelegnth_min"] = wavelength_min
    csth.crr_prj_properties["wavelegnth_max"] = wavelength_max

    canvas = Canvas()
    canvas.add_code("Solver", "WavelengthRange", wavelength_min, wavelength_max)
    
    print("[INFO] vba code to be executed:\n")
    canvas.preview()
    res = canvas.send(csth)

    if res:
        print("[ OK ] Project wavelength set successfully")
    else:
        print("[ERRO] Failed to set project wavelength")
        raise RuntimeError("Failed to set project wavelength")


def define_material(csth, materials_path, material_name):
    csth.materials[material_name] = {
        "freq": [],
        "re": [],
        "im": []
    }

    print(f"[INFO] Defining material: {material_name}")
    wl_min = csth.crr_prj_properties["wavelegnth_min"]
    wl_max = csth.crr_prj_properties["wavelegnth_max"]
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


def update_params(csth, force=False):
    print("[INFO] Updating parameters ...")

    canvas = Canvas()
    vba_code = f"RebuildOnParametricChange \"{force}\", \"True\""
    res = canvas.write_send(csth, vba_code, add_to_history=False)

    if res:
        print("[ OK ] Parameters updated successfully")
    else:
        print("[ERRO] Failed to update parameters")
        raise RuntimeError("Failed to update parameters")
    return res


def modify_param(csth, param_name: str, value: int):
    print(f"[INFO] Modifying parameter {param_name} to {value}")

    canvas = Canvas()
    vba_code = f"StoreParameter \"{param_name}\", \"{value}\""
    res = canvas.write_send(csth, vba_code, add_to_history=False)
    if res:
        print(f"[ OK ] Parameter {param_name} modified successfully")
    else:
        print(f"[ERRO] Failed to modify parameter {param_name}, please check whether the parameter exists")
        raise RuntimeError(f"Failed to modify parameter {param_name}, please check whether the parameter exists")
    update_params(csth)


def set_acc_dc(csth, solver="FDSolver"):
    print(f"[INFO] Setting solver to {solver}")

    canvas = Canvas()

    canvas.add_code(solver, "MPIParallelization", "False")
    canvas.add_code(solver, "UseDistributedComputing", "False")
    canvas.add_code(solver, "NetworkComputingStrategy", "RunRemote")
    canvas.add_code(solver, "NetworkComputingJobCount", "99")
    canvas.add_code(solver, "UseParallelization", "True")
    canvas.add_code(solver, "MaxCPUs", csth._ACC_DC.max_threads)
    canvas.add_code(solver, "MaximumNumberOfCPUDevices", csth._ACC_DC.max_num_of_cpu_devs)

    canvas.add_code("MeshSettings", "SetMeshType", "Unstr")
    canvas.add_code("MeshSettings", "Set", "UseDC", csth._ACC_DC.remote_mesh)

    if csth._ACC_DC.max_params_parallel <= 0:
        canvas.write("UseDistributedComputingForParameters \"False\"")
        canvas.write("MaxNumberOfDistributedComputingParameters \"1\"")
        canvas.write("ParameterSweep.UseDistributedComputing \"False\"")
    else:
        canvas.write("UseDistributedComputingForParameters \"True\"")
        canvas.write(f"MaxNumberOfDistributedComputingParameters \"{csth._ACC_DC.max_params_parallel}\"")
        canvas.write("ParameterSweep.UseDistributedComputing \"True\"")
    canvas.write(f"UseDistributedComputingMemorySetting \"{csth._ACC_DC.use_dc_mem_setting}\"")
    canvas.write(f"MinDistributedComputingMemoryLimit \"{csth._ACC_DC.min_dc_mem_limit}\"")
    canvas.write(f"UseDistributedComputingSharedDirectory \"{csth._ACC_DC.use_shared_dir}\"")
    canvas.write(f"OnlyConsider0D1DResultsForDC \"{csth._ACC_DC.only_0D1D}\"")

    canvas.preview()
    canvas.send(csth, "Set accerlation and distributed computing")


def set_FDSolver_source(csth, source_port="Zmin", mode="TM(0,0)"):
    print("[INFO] Setting up FDSolver ...")
    canvas = Canvas()

    # vbac = Canvas.vba_template.get_mode_num_by_name(source_port, mode)
    # canvas.write(vbac)

    canvas.write(f"FDSolver.Stimulation \"{source_port}\", \"{mode}\"")

    canvas.preview()
    res = canvas.send(csth, "Set FDSolver")
    if res:
        print("[ OK ] FDSolver set successfully")
    else:
        print("[ERRO] Failed to set FDSolver, please check whether the mode/port exists")
        raise RuntimeError("Failed to set FDSolver, please check whether the mode/port exists")


def beep_alert():
    """ make a beep sound """
    if platform.system() == "Windows":
        import winsound
        winsound.Beep(1000, 500)  # 1000Hz, 0.5s
    else:
        os.system('echo -e "\\a"')  # TTYS BELL (macOS/Linux)


def sweep_monitor(crr_pid:        int   = None, # current PID
                  crr_prj_path:   str   = None, # current project path
                  interval:       float = 0.9,  # additonal check interval
                  threshold:      int   = 6,    # threshold for the CPU occupancy rate
                  monitor_secs:   int   = 60,   # monitor time
                  confidence:     int   = 20    # confidence level
                  ):
    '''
    Monitor the CPU occupancy rate and trigger an alarm if the rate is less than the threshold
    :param interval: additional check interval
    :param threshold: threshold for the CPU occupancy rate
    :param monitor_secs: monitor time
    :param confidence: confidence level
    '''
    if not crr_prj_path:
        print("[ERRO] Project path is not specified")
        raise RuntimeError("Project path is not specified")

    print("[INFO] Monitoring CPU occupancy rate ...")
    # seq = 0
    cnt = 0
    cpu_usage_list = []
    cpu_usage_tmp = [0] * 10
    with contextlib.redirect_stdout(None):
        csth = cst_handler.CSTHandler(crr_pid)
    csth.crr_prj = csth.de.get_open_project(crr_prj_path)
    # sample_size = monitor_secs
    print(f"[ OK ] Monitor started successfully, connected to PID: {crr_pid}")

    while True:
        for i in range(10):
            cpu_usage_tmp[i] = psutil.cpu_percent(interval=0.01)
        cpu_usage = sum(cpu_usage_tmp) / 10
        cpu_usage_list.append(cpu_usage)

        # keep the CPU record of the last 60 seconds
        if len(cpu_usage_list) > monitor_secs:
            cpu_usage_list.pop(0)

        # calculate the average CPU occupancy rate
        avg_cpu = sum(cpu_usage_list) / monitor_secs
        print(f"\r> Currrent CPU occupancy rate: {cpu_usage:.2f}%, \taverage in the past {monitor_secs} seconds: {avg_cpu:.2f}%", end='', flush=True)
        print("\033[K", end='', flush=True)
        # seq += 1; print(f"seq: {seq}")

        # trigger alarm
        if len(cpu_usage_list) == monitor_secs and avg_cpu < threshold:
            print(f"\n[ALER] CPU occupancy rate is less than {threshold}%, the situation may be abnormal!")
            beep_alert()
            cnt += 1
            if cnt >= confidence:
                print("[WARN] The solver runs abnormally, aborting ...")
                csth.crr_prj.modeler.abort_solver()
                print("[ OK ] Solver aborted successfully")
                sys.exit(1)
        else:
            cnt = 0

        time.sleep(interval)  # additional second interval to prevent frequent CPU monitoring


def exec_paramSweep(csth, project=None):
    if project:
        csth.crr_prj = project
        csth.crr_prj.activate()
    print("[INFO] Executing parameter sweep ...")
    canvas = Canvas()
    canvas.write("ParameterSweep.Start")
    canvas.preview(0)
    res = canvas.send(csth, add_to_history=False)
    if res:
        print("[INFO] Parameter sweep finished, please check the results")
    else:
        print("[ERRO] Failed to start parameter sweep")
        raise RuntimeError("Failed to start parameter sweep")


def exec_parallel_sweep(crr_pid, crr_prj_path):
    print("[INFO] Executing parallel parameter sweep in background ...")
    with contextlib.redirect_stdout(None):
        csth = cst_handler.CSTHandler(crr_pid)
    csth.crr_prj = csth.de.get_open_project(crr_prj_path)
    exec_paramSweep(csth)
    return


def exec_paramSweep_safe(csth, 
                         interval:       float = 0.9,  # additonal check interval
                         threshold:      int = 6,    # threshold for the CPU occupancy rate
                         monitor_secs:   int = 60,   # monitor time
                         confidence:     int = 10    # confidence level):
                         ):
    print("[INFO] Executing parameter sweep in a safe way ...")
    canvas = Canvas()
    canvas.write("ParameterSweep.Start")
    canvas.preview(0)

    cnt = 0
    crr_pid = csth.pid
    crr_prj_path = csth.crr_prj.filename()

    while True:
        # Create a new monitor thread to monitor the CPU occupancy rate
        monitor_proc = Process(target=sweep_monitor,
                            args=(crr_pid,
                                    crr_prj_path,
                                    interval,
                                    threshold,
                                    monitor_secs,
                                    confidence),
                            daemon=True)
        # time.sleep(10)
        monitor_proc.start()
        res = canvas.send(csth, add_to_history=False)

        while True:
            time.sleep(2)
            if csth.crr_prj.modeler.is_solver_running():
                print(".", end="")
            else:
                print()
                print("[INFO] Solver finished")
                break

        if not res:
            print("[ERRO] Failed to start parameter sweep")
            raise RuntimeError("Failed to start parameter sweep")

        if monitor_proc.is_alive():
            if cnt:
                print(f"[INFO] The parameter sweep done after {cnt} retries")
            print("[INFO] Parameter sweep finished, please check the results")
            break
        else:
            monitor_proc.join()
            if monitor_proc.exitcode == 1:
                print("[WARN] The solver runs abnormally, retrying ...")
                cnt += 1


def set_basic_params(csth,
                     p:       int = 0,
                     theta:   int = 0,
                     phi:     int = 0,
                     ):
    
    # get the maximum frequency and search the closest frequency in the substrate material
    wavelength_min = csth.crr_prj_properties["wavelength_min"]
    wavelength_max = csth.crr_prj_properties["wavelength_max"]
    freq_max = 300 / wavelength_min

    try:
        substrate_material = csth.crr_prj_properties["substrate_material"]
    except KeyError:
        print("[ERRO] Please specify the substrate material first")
        raise RuntimeError("Please specify the substrate material first")
    freq_list = csth.materials[substrate_material]["freq"]

    # find the closest frequency in the substrate material
    idx = misc.find_closest_idx(freq_list, freq_max)
    re = csth.materials[substrate_material]["re"][idx]
    im = csth.materials[substrate_material]["im"][idx]

    if not (p or csth.crr_prj_properties["period"]):
        print("[ERRO] Please specify the period p first")
        raise RuntimeError("Please specify the period p first")

    if not (p or csth.crr_prj_properties["farfield"]):
        print("[ERRO] Please specify the farfield distance first")
        raise RuntimeError("Please specify the farfield distance first")

    print("[INFO] Setting basic parameters ...")
    if p:
        modify_param(csth, "p", p)
        csth.crr_prj_properties["period"] = p
        print("[ OK ] p is updated to ", p)

        csth.crr_prj_properties["farfield"] = wavelength_max # reduce calculate amount
        # csth.crr_prj_properties["farfield"] = misc.farfield_evaluator(wavelength_min, "standard", p, 1)
        farfield = csth.crr_prj_properties["farfield"]
        print("[INFO] farfield will be set to ", farfield)

        # update crr_prj_properties to prjs list (csth.prjs = pd.DataFrame(columns=["project_instance", "project_properties"]))
        csth.prjs.loc[csth.prjs["project_instance"] == csth.crr_prj, "project_properties"] = csth.crr_prj_properties
    else:
        p = csth.crr_prj_properties["period"]
        print("[INFO] Current period is ", p)

        farfield = csth.crr_prj_properties["farfield"]
        print("[INFO] Current farfield distance is ", farfield)

        # update crr_prj_properties to prjs list (csth.prjs = pd.DataFrame(columns=["project_instance", "project_properties"]))
        csth.prjs.loc[csth.prjs["project_instance"] == csth.crr_prj, "project_properties"] = csth.crr_prj_properties

    if theta:
        modify_param(csth, "theta", theta)
        csth.crr_prj_properties["theta"] = theta
        print("[ OK ] theta is updated to ", theta)

    if phi:
        modify_param(csth, "phi", phi)
        csth.crr_prj_properties["phi"] = phi
        print("[ OK ] phi is updated to ", phi)

    # calculate refractive index
    n_re, _, _, _ = misc.dielectric2refractive(float(re), float(im))

    # calculate floquet boundaries and enable modes
    _, enable_modes = misc.floquet_evaluator(n_re, freq_max, p, theta, phi)

    # update background and its material
    canvas = Canvas()
    vbas = Canvas.vba_template.set_background(farfield_distance=farfield)
    canvas.write(vbas)
    vbas = Canvas.vba_template.set_background_normal_material()
    canvas.write(vbas)
    # canvas.preview()
    canvas.send(csth, "Set background")

    # update the number of modes
    vbas = Canvas.vba_template.set_floquet_port_boundaries(enable_modes, farfield_distance=farfield)
    canvas.write(vbas)
    canvas.preview(0)
    canvas.send(csth, "Set Floquet port boundaries")

    # update boundaries
    vbas = Canvas.vba_template.set_boundaries()
    canvas.write(vbas)
    # canvas.preview()
    canvas.send(csth, "Set boundaries")

    print("[ OK ] Basic parameters set successfully")