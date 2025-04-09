import logging
import utils.log_conf

from utils import misc
from utils import cst_handler
from utils import basic_opts
from utils import materials_opts
from utils import param_opts
from utils import results_opts

import uvicorn
from interfaces.app import app

def main():

    misc.print_logo()
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True, log_level="info")
    # prj = r"C:/Users/27950/OneDrive/Desktop/SquarePillar__surface_cstl2/SquarePillar__for_res_test.cst"
    # results_opts.cst2mysql(prj, compensation_length=1, force=False)
    # # return
    # sparam_name = ["SZmax(2),Zmin(2)", "SZmax(1),Zmax(2)"]
    # data = results_opts.fetch_sparams(prj, sparam_name, plural=True)
    # return
    cst = cst_handler.CSTHandler()
    cst.open_template("SquarePillar")
    cst.instantiate_template("SquarePillar_Fine_surface_cstl3", wavelength_min=8, wavelength_max=14)

    basic_opts.set_acc_dc(cst)
    basic_opts.set_FDSolver_source(cst, "Zmin", "TE(0,0)")

    basic_opts.define_material(cst, "materials", "=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")
    materials_opts.SquarePillar(cst).change_substrate("=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")
    materials_opts.SquarePillar(cst).change_pillar("=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")

    param_opts.SquarePillar(cst).set_period_parallel_sweep(p_start=3.2, p_end=3.8, p_step=0.2,
                                                           h_step=0.25, l_step=0.04,
                                                           h_start=4, l_start=0.2,
                                                           start_now = False)
    # exec_parallel_sweep_from_list(cst, sweep_list, 3)

    # cst.close()
    print("Done")
    return


if __name__ == '__main__':
    main()