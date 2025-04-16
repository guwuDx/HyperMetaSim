import logging
logger = logging.getLogger(__name__)

import sys
import json
import numpy as np
import re

from tabulate import tabulate


def read_config(file_path: str, key: str = None):
    with open(file_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    if key:
        return config[key]
    return config


def add_cst_lib_path():
    lib_path = read_config("./config/service.json", "cst")["cst_py_lib_path"]
    sys.path.append(lib_path)


def ranges(start, end, step):
    if start==end:
        return np.array([start])
    arr = np.arange(start, end, step)
    if arr[-1] != end:
        arr = np.append(arr, end)
    return np.around(arr, decimals=3)


def floquet_evaluator(n, freq_thz, p_um, theta_deg, phi_deg, orderx_max=5, ordery_max=5):
    """
    n: refractive index
    freq_hz: frequency in Hz
    p: period in micrometer (for x, y same)
    theta_deg, phi_deg: incidence angles in degrees
    m_range, n_range: integer range for enumerating Floquet order 
                      (e.g. m_range = range(-5,6), n_range = range(-5,6))
    Return: a list of (m, n, beta, alpha, is_propagating)
    """
    m_range = range(0, orderx_max)
    n_range = range(0, ordery_max)

    c0 = 3e8
    enable_modes = 0

    # convert units
    PI = np.pi
    p = p_um * 1e-6
    freq_hz = freq_thz * 1e12
    freq_rad = 2.0 * PI * freq_hz
    k0 = freq_rad / c0   # wave number in vacuum
    # Convert angles to rad
    theta = np.deg2rad(theta_deg)
    phi = np.deg2rad(phi_deg)

    kx0 = k0*n*np.sin(theta)*np.cos(phi)
    ky0 = k0*n*np.sin(theta)*np.sin(phi)

    results = []

    for mm in m_range:
        for nn in n_range:
            kx = kx0 + 2*PI*mm/p
            ky = ky0 + 2*PI*nn/p
            kz_sq = (k0*n)**2 - (kx**2 + ky**2)

            if kz_sq >= 0:
                # real propagation
                beta = np.sqrt(kz_sq)
                alpha = 0.0
                is_propagating = True

                if not (mm or nn):
                    enable_modes += 2 # TE and TM, X=0 and Y'=0
                    for EMW in ["TE", "TM"]:
                        results.append({
                            "EMW": EMW,
                            "X": mm,
                            "Y'": nn,
                            "beta": beta,
                            "alpha": alpha,
                            "is_propagating": is_propagating
                        })

                elif (not mm) and nn:
                    enable_modes += 4 # TE and TM, X=0 and Y'!=0
                    for sign_y in [1, -1]:
                        for EMW in ["TE", "TM"]:
                            results.append({
                                "EMW": EMW,
                                "X": mm,
                                "Y'": sign_y*nn,
                                "beta": beta,
                                "alpha": alpha,
                                "is_propagating": is_propagating
                            })

                elif mm and (not nn):
                    enable_modes += 4 # TE and TM, X!=0 and Y'=0
                    for sign_x in [1, -1]:
                        for EMW in ["TE", "TM"]:
                            results.append({
                                "EMW": EMW,
                                "X": sign_x*mm,
                                "Y'": nn,
                                "beta": beta,
                                "alpha": alpha,
                                "is_propagating": is_propagating
                            })

                else:
                    enable_modes += 8 # TE and TM, X!=0 and Y'!=0
                    for sign_x in [1, -1]:
                        for sign_y in [1, -1]:
                            for EMW in ["TE", "TM"]:
                                results.append({
                                    "EMW": EMW,
                                    "X": sign_x*mm,
                                    "Y'": sign_y*nn,
                                    "beta": beta,
                                    "alpha": alpha,
                                    "is_propagating": is_propagating
                                })

            else:
                beta = 0.0
                alpha = np.sqrt(-kz_sq)
                is_propagating = False

                if not (mm or nn):
                    for EMW in ["TE", "TM"]:
                        results.append({
                            "EMW": EMW,
                            "X": mm,
                            "Y'": nn,
                            "beta": beta,
                            "alpha": alpha,
                            "is_propagating": is_propagating
                        })

                elif (not mm) and nn:
                    for sign_y in [1, -1]:
                        for EMW in ["TE", "TM"]:
                            results.append({
                                "EMW": EMW,
                                "X": mm,
                                "Y'": sign_y*nn,
                                "beta": beta,
                                "alpha": alpha,
                                "is_propagating": is_propagating
                            })

                elif mm and (not nn):
                    for sign_x in [1, -1]:
                        for EMW in ["TE", "TM"]:
                            results.append({
                                "EMW": EMW,
                                "X": sign_x*mm,
                                "Y'": nn,
                                "beta": beta,
                                "alpha": alpha,
                                "is_propagating": is_propagating
                            })

                else:
                    for sign_x in [1, -1]:
                        for sign_y in [1, -1]:
                            for EMW in ["TE", "TM"]:
                                results.append({
                                    "EMW": EMW,
                                    "X": sign_x*mm,
                                    "Y'": sign_y*nn,
                                    "beta": beta,
                                    "alpha": alpha,
                                    "is_propagating": is_propagating
                                })

            # results.append((mm, nn, beta, alpha, is_propagating))

    sorted_results = sorted(results, key=lambda x: -x["beta"])
    for i, row in enumerate(sorted_results, start=1):
        row["Index"] = i

    print(f"[INFO] {enable_modes} Floquet modes will propagate in metastructure cell.")
    print( "[INFO] All Floquet modes:")
    print(tabulate(sorted_results,
                   headers="keys",
                   tablefmt="pretty",
                   floatfmt=(".0f",
                             ".0f",
                             ".5e",
                             ".5e",
                             ""
                             )
                  )
         )

    if enable_modes == 0:
        print("[WARN] No propagating Floquet modes found.")
        print("[WARN] Any EM wave will be reflected or absorbed in the structure.")

    if enable_modes == len(results):
        print("[WARN] Too many propagating Floquet modes found.")
        print("[WARN] Consider increasing the orderx_max and ordery_max to discover more modes.")

    return sorted_results, enable_modes


def farfield_evaluator(lambda_um:   int,
                       solve_mode:  str = None,
                       p_um:        int = None,
                       n:           int = 1
                       ):
    """
    lambda_um: wavelength in micrometer"
    solve_mode: "loose", "standard" or "strict"
                strict mode may means more accurate but slower
    p_um: period in micrometer
    n: refractive index
    Return: farfield resolution in um
    """
    if solve_mode is None or solve_mode == "loose":
        print("[INFO] work in loose mode")
        # loose mode
        return 0.5 * lambda_um
    elif solve_mode == "standard":
        print("[INFO] work in standard mode")
        # standard mode
        if p_um is None:
            print("[ERROR] p_um is required for standard mode")
            raise ValueError("p_um is required for standard mode")
        return (2 * n * p_um) / lambda_um
    elif solve_mode == "strict":
        print("[INFO] work in strict mode")
        # strict mode
        if p_um is None:
            print("[ERROR] p_um is required for strict mode")
            raise ValueError("p_um is required for strict mode")
        return (20 * n * p_um * 1.41) / lambda_um


def dielectric2refractive(re, im):
    """
    re: real part of dielectric constant
    im: imaginary part of dielectric constant
    Return: im & re of refractive index and refractive index
    """
    n = np.sqrt(re + 1j*im)
    return n.real, n.imag, n, np.abs(n)


def load_material(materials_path: str,
                  material_name: str,
                  wavelegnth_min: float = None,
                  wavelegnth_max: float = None,
                  is_freq=False,
                  unit="um/THz",
                  calculate_refractive=False):
    TOLERANCE = 0.001  # tolerance for frequency range
    res = []

    if not (wavelegnth_max and wavelegnth_min):
        logging.warning("wavelegnth_min and wavelegnth_max are not provided, using default values")
        logging.warning("This may cause performance losses")
        wavelegnth_min = 0
        wavelegnth_max = float("inf")
    elif is_freq:
        if unit == "um/THz":
            freq_min = 299 / wavelegnth_max
            freq_max = 299 / wavelegnth_min

    freq_min = freq_min * (1 - TOLERANCE)
    freq_max = freq_max * (1 + TOLERANCE)
    with open(f"{materials_path}/{material_name}.csv", "r") as f:
        lines = f.readlines()
        for line in lines:
            if line.startswith("#"): continue

            freq, re, im = line.replace(" ", "").replace("\n", "").split(",")
            if float(freq) < freq_min or float(freq) > freq_max:
                continue

            if calculate_refractive:
                n_re, n_im, n_complex, n_abs = dielectric2refractive(float(re), float(im))
                res.append({
                    "freq": float(freq),
                    "re": float(re),
                    "im": float(im),
                    "n_re": n_re,
                    "n_im": n_im,
                    "complex": n_complex,
                    "abs": n_abs
                })
            else:
                res.append({
                    "freq": float(freq),
                    "re": float(re),
                    "im": float(im)
                })

    return res


def find_closest_idx(lst, target):
    """
    lst: list of numbers
    target: number
    Return: closest number in arr to target
    """
    lst = list(map(float, lst))
    return min(range(len(lst)), key=lambda i: abs(lst[i] - target))


def compensate_substrate(data: np.ndarray,
                         compensation_length: float,  # in um
                         compensation_material: str,
                         wavelength_min: float = None,
                         wavelength_max: float = None):
    """
    compensate_substrate

    data: shape(N,3), each row = [freq(THz), real_part, imag_part]
    compensation_length: length of substrate (um)
    compensation_material: material name in local folder
    wavelength_min, wavelength_max: optional wave range

    Returns data (in-place) with updated real_part, imag_part.
    """
    PI = np.pi
    c0 = 2.998e8  # m/s (speed of light in vacuum)

    # 1) load material data
    material = load_material(
        "./materials",
        compensation_material,
        wavelegnth_min=wavelength_min,
        wavelegnth_max=wavelength_max,
        is_freq=True,
        calculate_refractive=True
    )

    # 2) convert length to meters
    length_m = compensation_length * 1e-6

    # extract frequency list from material data
    freq_list = [entry['freq'] for entry in material]

    # 3) process data
    for i in range(data.shape[0]):
        # freq(THz)、real、imag
        freq_thz = data[i, 0]
        real_val = data[i, 1]
        imag_val = data[i, 2]

        # find closest frequency index in material data
        idx = find_closest_idx(freq_list, freq_thz)
        n_re = material[idx]["n_re"]

        # frequency in Hz and wave number
        freq_hz = freq_thz * 1e12
        k0 = 2.0 * PI * freq_hz / c0

        # phase shift
        phase_shift = np.exp(1j * k0 * n_re * length_m)

        # complex number multiplication
        old_complex = real_val + 1j * imag_val
        new_complex = old_complex * phase_shift

        # update data
        data[i, 1] = new_complex.real
        data[i, 2] = new_complex.imag

    return data


def sparam_id(sparam_name: str, sparam_id: int):

    sparam_list = [
        "TE(0,0)",  "TM(0,0)",
        "TE(0,1)",  "TM(0,1)",
        "TE(0,-1)", "TM(0,-1)",
        "TE(1,0)",  "TM(1,0)",
        "TE(-1,0)", "TM(-1,0)",
        "TE(1,1)",  "TM(1,1)",
        "TE(1,-1)", "TM(1,-1)",
        "TE(-1,1)", "TM(-1,1)",
        "TE(-1,-1)","TM(-1,-1)",
        "TE(0,2)",  "TM(0,2)",
        "TE(0,-2)", "TM(0,-2)",
        "TE(2,0)",  "TM(2,0)",
        "TE(-2,0)", "TM(-2,0)",
        "TE(1,2)",  "TM(1,2)",
        "TE(1,-2)", "TM(1,-2)",
        "TE(-1,2)", "TM(-1,2)",
        "TE(-1,-2)","TM(-1,-2)",
        "TE(2,1)",  "TM(2,1)",
        "TE(2,-1)", "TM(2,-1)",
        "TE(-2,1)", "TM(-2,1)",
        "TE(-2,-1)","TM(-2,-1)",
        "TE(2,2)",  "TM(2,2)",
        "TE(2,-2)", "TM(2,-2)",
        "TE(-2,2)", "TM(-2,2)",
        "TE(-2,-2)","TM(-2,-2)",
    ]

    if sparam_name:
        res = sparam_list.index(sparam_name)
        return res
    elif sparam_id:
        return sparam_list[sparam_id]
    else:
        print("[ERROR] sparam_name or sparam_id is required")
        return None
    

def identify_param(params: dict, param_info: dict):
    """
    eg. params: {"p": 1.0, "h": 2.0, "t": 3.0, "e_theta": 0.0, "e_phi": 0.0, "parameter1": 1.0, "parameter2": 2.0}
    + param_info: [{
                    "ID": 1,
                    "name": "height",
                    "name_zn": "单元结构高度",
                    "keywords": [
                        "h",
                        "height"
                    ],
                    "belongs_to": "generic_parameters"
               }, { ...
               }, ...]
    ↓
    . param_map: dict
    |
    ├── generic_parameters
    │   ├── period
    │   ├── height
    │   ├── thickness
    │   ├── e_theta
    │   └── e_phi
    └── shape__parameters
        ├── parameter1
        ├── parameter2
        ···
        └── parameterN
    """
    param_columns = {}
    generic_parameters = {}
    shape_parameters = {}
    for param, value in params.items():
        for info in param_info:
            if param in info["keywords"]:
                if info["belongs_to"] == "generic_parameters":
                    generic_parameters[info["name"]] = value
                else:
                    shape_parameters[info["name"]] = value
                break
        else:
            logging.exception(f"[ERRO] Parameter {param} not found in param_info")
            raise ValueError(f"Parameter {param} not found in param_info")

    param_columns["generic_parameters"] = generic_parameters
    param_columns["shape_parameters"] = shape_parameters
    return param_columns


def get_material_name(material_name: str):
    """
    material_name: "[Si_crystal]_freq-r-i_0.0310-310um_ByFranta-300K_2017"
    return: "Si_crystal" 
    """
    pattern = r'^\=(.*?)\='
    match = re.search(pattern, material_name)
    if match:
        material_name = match.group(1)
    else:
        logging.error(f"[ERRO] Invalid material name: {material_name}")
        raise ValueError(f"[ERRO] Invalid material name: {material_name}")
    return material_name


def get_project_properties(project_name: str):

    # project_properties.json is in the same directory as the project file
    project_properties = project_name.replace(".cst", "/project_properties.json")
    print(f"[INFO] Reading project properties from {project_properties}")

    try:
        project_properties = read_config(project_properties)
    except:
        print("[ERRO] project_properties.json not found in the project directory")
        print("[ERRO] make sure the project file was generated by the HyperMetaSim(R) tool")
        raise FileNotFoundError()

    return project_properties


def print_logo():
    print(r"""
   __    _  _  _  _  _  _  _  _
  / /__       __  __            \ _  _  _  _  _  _  _
 / // /      / / / /__  __ ____   ___   _____
/ // /      / /_/ // / / // __ \ / _ \ / ___/             __
\ \\ \     / __  // /_/ // /_/ //  __// /               __\_\
 \ \\ \   /_/ /_/ \__, // .___/ \___//_/ \   \ BY:      \_\\ \
  \ \\_\         /____//_/    ===========/    \ guwudx   \ \\ \
   \_\\_\      __  __       _          ___  _  \ GPL 3.0  \ \\ \
    \_\       |  \/  | ___ | |_  __ _ / __|(_) _ __        \ \\ \
              | |\/| |/ -_)|  _|/ _` |\__ \| || '  \       / // /
              |_|  |_|\___| \__|\__,_||___/|_||_|_|_|     /_// /
            - - - - - - - - - - - - - - - - - - - - - - -   /_/
    """)
    print("<<<<<<<<<<<<<<<<<<<<<< CST Automation >>>>>>>>>>>>>>>>>>>>>>")