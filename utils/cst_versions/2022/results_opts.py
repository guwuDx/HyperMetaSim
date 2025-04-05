import logging
logger = logging.getLogger(__name__)

import utils.misc as misc
import utils.mcbs as mcbs
import cst.results

import numpy as np
import threading
from tqdm import tqdm
from typing import List
from concurrent.futures import ThreadPoolExecutor



def fetch_runids(project3d, runids, sparam_name, plural=True):
    cells = []
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
        # print(f"Fetched S-parameters for {sparam_name} with runid {runid}")

    print(f"[ OK ] Fetched S-parameters for {sparam_name} successfully")
    return cells


def get_sparam_names(project3d):
    tree = project3d.get_tree_items()
    sparam_names = []

    tree = project3d.get_tree_items()
    for item in tree:
        if "1D Results\\S-Parameters" in item:
            sparam_names.append(item.split("\\")[-1])
    return sparam_names


def fetch_sparams(project_name: str, 
                  sparam_names: List[str],
                  plural=True):
    project = cst.results.ProjectFile(project_name)
    if project:
        print(f"Fetching S-parameters from {project.filename}")
    else:
        print(f"Project {project_name} open failed")
        raise FileNotFoundError()

    project_properties = misc.get_project_properties(project.filename)
    project_type = project_properties["type"]
    substrate_material = project_properties["substrate_material"]
    pillar_material = project_properties["pillar_material"]

    project3d = project.get_3d()
    cells = []

    if sparam_names is None:
        sparam_names = get_sparam_names(project3d)

    for sparam_name in sparam_names:
        try:
            runids = project3d.get_run_ids(f"1D Results\\S-Parameters\\{sparam_name}", True)
            if not runids:
                print(f"No S-parameters found for {sparam_name}, the project may not have been simulated")
                return
        except:
            print(f"tree path not found: 1D Results\\S-Parameters\\{sparam_name}")
            continue

        cells += fetch_runids(project3d, runids, sparam_name, plural)

    res = {
        "type": project_type,
        "geometry": "um",
        "substrate_material": substrate_material,
        "pillar_material": pillar_material,
        "cells": cells
    }
    return res
    # sparam = project.get_3d().get_result_item(f"1D Results\\S-Parameters\\{sparam_name}", 2)


def get_split_points(project3d):
    """
    Returns the split points of the Frequency: N_M_idx, M_F_idx
    Splited by near-infrared, mid-infrared, far-infrared
    """
    N_M_SPLIT_POINT = 119.9170 # THz
    M_F_SPLIT_POINT = 59.9585 # THz
    N_M_idx = 0
    M_F_idx = 0

    samp_s_param = get_sparam_names(project3d)[0]
    if not samp_s_param:
        logging.info(f"No S-parameters found, the project may not have been simulated")
        return None

    samp_runid = project3d.get_run_ids(f"1D Results\\S-Parameters\\{samp_s_param}", True)[0]
    data = project3d.get_result_item(f"1D Results\\S-Parameters\\{samp_s_param}", samp_runid).get_data()

    freqs = [x[0] for x in data]
    M_F_idx = np.searchsorted(freqs, M_F_SPLIT_POINT, side="left")
    N_M_idx = np.searchsorted(freqs, N_M_SPLIT_POINT, side="right")
    logging.info(f"Split points: N_M_idx={N_M_idx}, M_F_idx={M_F_idx}")

    return N_M_idx, M_F_idx


def process_import_single_runid(project3d, 
                                project_properties:dict, 
                                param_info:dict,
                                runid:int, 
                                substrate_material_id,
                                pillar_material_id,
                                sparam_name:str,
                                N_M_idx:int,
                                M_F_idx:int,
                                compensation=0,
                                force=False):
    logging.info(f"Processing runid {runid} for {sparam_name}")

    sqlh = mcbs.SQLHandler()
    param = project3d.get_parameter_combination(runid)

    compensation_material = project_properties["substrate_material"]

    param_columns = misc.identify_param(param, param_info)

    # complete the rest parameters
    param_columns["generic_parameters"]['substrate_material_id'] = substrate_material_id
    param_columns["generic_parameters"]['pillar_material_id'] = pillar_material_id
    param_columns["generic_parameters"]['s_param_id'] = sqlh.get_sparam_id(sparam_name)

    # check and import generic parameters
    gp_id, is_gen_duplicate = sqlh.check_and_insert_param("generic_parameters", param_columns["generic_parameters"], return_duplicate_status=True)
    sqlh.commit()
    if gp_id is None:
        logging.info(f"No generic parameters found for {sparam_name} with runid {runid}")
        return

    # check and import shape parameters
    shape = project_properties["type"]
    shape_parameters_table = shape + "_parameters"
    # param_columns["shape_parameters"]["gp_id"] = gp_id
    shp_id, is_shp_duplicate = sqlh.check_and_insert_param(shape_parameters_table, param_columns["shape_parameters"], return_duplicate_status=True)
    sqlh.commit()
    if shp_id is None:
        logging.info(f"No shape parameters found for {sparam_name} with runid {runid}")
        return

    # process data
    data = project3d.get_result_item(f"1D Results\\S-Parameters\\{sparam_name}", runid).get_data()
    if len(data[0]) > 2:
        logging.warning(f"Data length is greater than 2, deleting extra columns (probably including Ref.Imp.)")
        data = [row[:2] for row in data]
    data = np.array([
        [freq, z.real, z.imag] for freq, z in data
    ])

    wavelength_min = project_properties["wavelength_min"]
    wavelength_max = project_properties["wavelength_max"]
    if compensation:
        logging.info(f"Compensating substrate for {sparam_name} with runid {runid}")
        data = misc.compensate_substrate(data, compensation, compensation_material,
                                         wavelength_min, wavelength_max)

    far_infrared_data = data[:M_F_idx, :]
    mid_infrared_data = data[M_F_idx:N_M_idx, :]
    near_infrared_data = data[N_M_idx:, :]

    # insert data into database
    if far_infrared_data.size:
        table = shape + "_freq_resp_FIR"
        sqlh.insert_freq_response(table, far_infrared_data, gp_id, shp_id, force=force)
        sqlh.commit()
    if mid_infrared_data.size:
        table = shape + "_freq_resp_MIR"
        sqlh.insert_freq_response(table, mid_infrared_data, gp_id, shp_id, force=force)
        sqlh.commit()
    if near_infrared_data.size:
        table = shape + "_freq_resp_NIR"
        sqlh.insert_freq_response(table, near_infrared_data, gp_id, shp_id, force=force)
        sqlh.commit()
    # logging.info(f"Imported S-parameters for {sparam_name} with runid {runid}")

def cst2mysql(project_name, 
              force=False, 
              compensation_length=0):

    project = cst.results.ProjectFile(project_name)
    if project:
        logging.info(f"Fetching S-parameters from {project.filename}")
    else:
        logging.error(f"Project {project_name} open failed")
        raise FileNotFoundError()

    sqlh = mcbs.SQLHandler()

    project_properties = misc.get_project_properties(project.filename)
    shape = project_properties["type"]

    # get the material id
    substrate_material = misc.get_material_name(project_properties["substrate_material"])
    pillar_material = misc.get_material_name(project_properties["pillar_material"])
    substrate_material_id = sqlh.get_material_id(substrate_material)
    pillar_material_id = sqlh.get_material_id(pillar_material)

    # sqlh = mcbs.SQLHandler()
    param_info = sqlh.get_param_info(shape)

    project3d = project.get_3d()
    sparam_names = get_sparam_names(project3d)

    N_M_idx, M_F_idx = get_split_points(project3d)

    if sparam_names is None:
        logging.info(f"No S-parameters found, the project may not have been simulated")
        return

    cnt = 0
    for sparam_name in sparam_names:
        runids = project3d.get_run_ids(f"1D Results\\S-Parameters\\{sparam_name}", True)
        for runid in runids:
            cnt += 1
            print(f"--------------------- {cnt}")
            process_import_single_runid(project3d,
                                        project_properties,
                                        param_info,
                                        runid,
                                        substrate_material_id,
                                        pillar_material_id,
                                        sparam_name,
                                        N_M_idx,
                                        M_F_idx,
                                        compensation=compensation_length,
                                        force=force)

    return