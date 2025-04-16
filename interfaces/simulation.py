from .models import SingleSimRequest
from . import state

from utils import basic_opts
from utils import materials_opts
from utils import param_opts

from fastapi import APIRouter

router = APIRouter(prefix="/simulation", tags=["simulation"])

import logging
logger = logging.getLogger(__name__)

timeout = 60 * 5



def create_new_project_full(csth, shape_type, wavelength_min, wavelength_max,
                            substrate_material, pillar_material, 
                            port, mode, period, e_theta, e_phi):
    logging.info(f"Creating a new project with type {shape_type}.")
    csth.build_project_vba(
        shape_type,
        shape_type + "_http_simulation",
        wavelength_min=wavelength_min,
        wavelength_max=wavelength_max,
        save=False)

    basic_opts.set_acc_dc(csth)
    basic_opts.set_FDSolver_source(csth, port, mode)

    materials_opts.change_substrate(csth, substrate_material, force_define=True)
    materials_opts.change_pillar(csth, pillar_material, force_define=True)

    basic_opts.set_basic_params(csth, period, e_theta, e_phi)



@router.post("/single")
def single_simulation(request: SingleSimRequest):
    """
    Perform a single simulation based on the provided parameters.
    """
    logging.info(f"Received request: {request.json()}")
    csth = state.csth
    if not csth:
        logging.error("CST handler is not initialized.")
        return {"error": "CST handler is not initialized."}

    # Extract parameters from the request
    shape_type = request.shape_type
    wavelength_min = request.wavelength_min
    wavelength_max = request.wavelength_max
    substrate_material = request.material.substrate_material
    pillar_material = request.material.pillar_material
    port = request.exciting_source.port
    mode = request.exciting_source.mode
    generic_params = request.params.generic_parameters
    shape_params = request.params.shape_parameters
    db_overwrite = request.db_overwrite
    redirect_to_db = request.redirect_to_db

    period = generic_params.period
    e_theta = generic_params.e_theta
    e_phi = generic_params.e_phi

    prj_list = csth.find_project_by_type(shape_type)
    if prj_list:
        full_name = prj_list[0]
        logging.info(f"Project with type {shape_type} already exists.")
        csth.switch_to_project(full_name)

        period_old = csth.crr_prj_properties.get("period")
        e_theta_old = csth.crr_prj_properties.get("e_theta", 0)
        e_phi_old = csth.crr_prj_properties.get("e_phi", 0)

        # delete old project if basic parameters have changed
        if period != period_old or e_theta != e_theta_old or e_phi != e_phi_old:
            # rebuild the project with new basic parameters
            logging.info("Basic parameters have changed. Deleting old project.")
            csth.delete_project(full_name)
            create_new_project_full(csth, shape_type, wavelength_min, wavelength_max,
                                    substrate_material, pillar_material, port, mode,
                                    period, e_theta, e_phi)
    else:
        # Create a new project with the specified shape type
        create_new_project_full(csth, shape_type, wavelength_min, wavelength_max,
                                substrate_material, pillar_material, port, mode,
                                period, e_theta, e_phi)

    height = generic_params.height

    if shape_type == "SquarePillar":
        side_length = shape_params.get("side_length", 0.2)
        param_opts.SquarePillar(csth).set_params(h=height, l=side_length)
        csth.run_solver(timeout=timeout)

    pass