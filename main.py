from utils import misc
from utils import cst_general
from utils import basic_operations
from utils import materials_operations
from utils import param_operations
from utils.macors_canva import Canvas

import numpy as np

def main():

    print("<<<<<<<<<<<<<<< CST Automation >>>>>>>>>>>>>>>")
    cst = cst_general.CSTHandler()
    cst.open_template("SquarePillar")
    cst.instantiate_template("SquarePillar_inst", 8, 14)
    basic_operations.define_material(cst, "materials", "freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    materials_operations.SquarePillar(cst).change_substrate("freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    materials_operations.SquarePillar(cst).change_pillar("freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    # basic_operations.modify_param(cst_app=cst, param_name="p", value=10)
    basic_operations.set_acc_dc(cst)
    basic_operations.set_FDSolver_source(cst, "Zmin", "TM(0,0)")
    canvas = Canvas()

    sp_params = param_operations.SquarePillar(cst)
    # sp_params.generate_sweep_squence(1, 1, 0.1)
    # sp_params.set_sweep_from_list(start_now=False)
    sp_params.set_sweep_from_range(5.5, 8, 0.5, 0.25, 0.1, start_now=False)
    # cst.save_crr_prj()
    basic_operations.exec_paramSweep(cst)
    # for i in np.arange(0.5, 4.5, 0.5):
    #     sp_params.simulate_param_combination(5, 3, i, 0, 0, True, None)

    # cst.close()
    print("Done")


if __name__ == '__main__':
    main()