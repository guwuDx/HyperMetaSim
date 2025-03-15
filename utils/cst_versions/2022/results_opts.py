import cst.results

import numpy as np
from tqdm import tqdm
from typing import List


def fetch_sparams(project_name: str, 
                  sparam_names: List[str], 
                  project_type: str = None, 
                  plural=True):
    project = cst.results.ProjectFile(project_name)
    if project:
        print(f"[INFO] Fetching S-parameters from {project.filename}")
    else:
        print(f"[ERRO] Project {project_name} open failed")
        raise FileNotFoundError()

    project3d = project.get_3d()
    cells = []

    for sparam_name in sparam_names:
        try:
            runids = project3d.get_run_ids(f"1D Results\\S-Parameters\\{sparam_name}", True)
            if not runids:
                print(f"[WARN] No S-parameters found for {sparam_name}, the project may not have been simulated")
                return
        except:
            print(f"[WARN] tree path not found: 1D Results\\S-Parameters\\{sparam_name}")
            continue

        for runid in tqdm(runids, desc=f"Fetching {sparam_name}"):
            param = project3d.get_parameter_combination(runid)
            data = project3d.get_result_item(f"1D Results\\S-Parameters\\{sparam_name}", runid).get_data()

            if not plural:
                data = np.array(data, dtype=complex)
                data_complex = np.array(data[:, 1])

                mag = np.abs(data_complex)
                phase = np.angle(data_complex, deg=True)

                data = np.column_stack((data[:, 0], mag, phase))

            cell = {
                "runid": runid,
                "sparam": sparam_name,
                "param": param,
                "data": data
            }

            cells.append(cell)
            # print(f"[INFO] Fetched S-parameters for {sparam_name} with runid {runid}")

        print(f"[ OK ] Fetched S-parameters for {sparam_name} successfully")

    res = {
        "type": project_type,
        ""
        "cells": cells
    }
    return res
    # sparam = project.get_3d().get_result_item(f"1D Results\\S-Parameters\\{sparam_name}", 2)