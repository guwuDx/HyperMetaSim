import sys
import toml
import numpy as np

from tabulate import tabulate


def read_toml(file_path, chunk_name):
    with open(file_path, 'r') as f:
        config = toml.load(f)
    return config[chunk_name]


def add_cst_lib_path():
    lib_path = read_toml("./config/service.toml", "cst")["cst_py_lib_path"]
    sys.path.append(lib_path)


def configure_drc():
    drc_config = read_toml("./config/service.toml", "drc")
    return drc_config


def configure_acc_and_dc():
    acc_dc_config = read_toml("./config/service.toml", "acc_dc")
    return acc_dc_config

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


def find_closest_idx(lst, target):
    """
    lst: list of numbers
    target: number
    Return: closest number in arr to target
    """
    lst = list(map(float, lst))
    return min(range(len(lst)), key=lambda i: abs(lst[i] - target))


def print_logo():
    print(r"""
   __ __
  / // /      __  __    . . . . . . . . . . . . . .
 / // /      / / / /__  __ ____   ___   _____      
/ // /      / /_/ // / / // __ \ / _ \ / ___/          __ __
\ \\ \     / __  // /_/ // /_/ //  __// /              \_\\_\
 \ \\ \   /_/ /_/ \__, // .___/ \___//_/ \    \BY:      \ \\ \
  \ \\ \         /____//_/    ===========/     \guwudx   \ \\ \
   \_\\_\      __  __       _          ___  _   \GUN 3.0  \ \\ \
    \_\\_\    |  \/  | ___ | |_  __ _ / __|(_) _ __        \ \\ \
              | |\/| |/ -_)|  _|/ _` |\__ \| || '  \       / // /
              |_|  |_|\___| \__|\__,_||___/|_||_|_|_|     / // /
            - - - - - - - - - - - - - - - - - - - - - -  /_//_/
    """)
    print("<<<<<<<<<<<<<<<<<<<<<< CST Automation >>>>>>>>>>>>>>>>>>>>>>")