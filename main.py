from utils import misc
from utils import cst_handler
from utils import basic_opts
from utils import materials_opts
from utils import param_opts
from utils import results_opts


def main():

    misc.print_logo()
    # prj = r"C:\Users\27950\OneDrive\Desktop\SquarePillar__surface_cstl2\SquarePillar__surface_cstl2.cst"
    # sparam_name = ["SZmax(1),Zmin(2)", "SZmax(2),Zmin(2)", "SZmin(1),Zmin(2)", "SZmax(1),Zmax(2)"]
    # data = results_opts.fetch_sparams(prj, sparam_name, plural=True)
    # return
    cst = cst_handler.CSTHandler()
    cst.open_template("SquarePillar")
    cst.instantiate_template("SquarePillar__surface_inst", wavelength_min=8, wavelength_max=14)
    basic_opts.define_material(cst, "materials", "freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    materials_opts.SquarePillar(cst).change_substrate("freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    materials_opts.SquarePillar(cst).change_pillar("freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")

    basic_opts.set_acc_dc(cst)
    basic_opts.set_FDSolver_source(cst, "Zmin", "TM(0,0)")

    param_opts.SquarePillar(cst).set_period_parallel_sweep(p_start=1.6, p_end=2.2, p_step=0.2,
                                                           h_step=0.25, l_step=0.02,
                                                           h_start=4, l_start=0.1,
                                                           start_now = False)
    # exec_parallel_sweep_from_list(cst, sweep_list, 3)

    # cst.close()
    print("Done")


if __name__ == '__main__':
    main()