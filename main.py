from utils import misc
from utils import cst_handler
from utils import basic_operations
from utils import materials_operations
from utils import param_operations

# import concurrent.futures

# def exec_parallel_sweep_from_list(csth, sweep_list, max_workers=3):
#     with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
#         future_to_prjs = {
#             executor.submit(basic_operations.exec_parallel_sweep, csth.pid, prj.filename()): prj
#             for prj in sweep_list
#         }

#         for future in concurrent.futures.as_completed(future_to_prjs):
#             prj = future_to_prjs[future]
#             try:
#                 res = future.result()
#             except Exception as exc:
#                 print(f"[ERRO] {prj} generated an exception: {exc}")
#             else:
#                 if res:
#                     print(f"[INFO] {prj} completed")
#                 else:
#                     print(f"[ERRO] {prj} failed")


def main():

    misc.print_logo()
    cst = cst_handler.CSTHandler()
    cst.open_template("SquarePillar")
    cst.instantiate_template("SquarePillar__surface_inst", wavelength_min=8, wavelength_max=14)
    basic_operations.define_material(cst, "materials", "freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    materials_operations.SquarePillar(cst).change_substrate("freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")
    materials_operations.SquarePillar(cst).change_pillar("freq-r-i_Si_crystal_0.0310-310um_ByFranta-300K_2017")

    basic_operations.set_acc_dc(cst)
    basic_operations.set_FDSolver_source(cst, "Zmin", "TM(0,0)")

    sweep_list = param_operations.SquarePillar(cst).set_period_parallel_sweep(p_start=1.6, p_end=2, p_step=0.2,
                                                                              h_step=0.25, l_step=0.02,
                                                                              start_now = True)
    # exec_parallel_sweep_from_list(cst, sweep_list, 3)

    # cst.close()
    print("Done")


if __name__ == '__main__':
    main()