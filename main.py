import logging
import utils.log_conf

from utils import misc
from utils import cst_handler
from utils import basic_opts
from utils import materials_opts
from utils import param_opts
from utils import results_opts
from utils.cli_parser import parser_constructor

import os
import uvicorn
import argparse
from interfaces.app import app



def debug():
    # prj = r"C:/Users/27950/OneDrive/Desktop/SquarePillar__surface_cstl2"
    # results_opts.cst2mysql(prj, compensation_length=1, max_workers=4, parallel_num=-1, force=False)
    # return
    # sparam_name = ["SZmax(2),Zmin(2)", "SZmax(1),Zmax(2)"]
    # data = results_opts.fetch_sparams(prj, sparam_name, plural=True)
    # return
    cst = cst_handler.CSTHandler()
    # cst.open_template("SquarePillar")
    # cst.instantiate_template("SquarePillar_Fine_surface_cstl3", wavelength_min=8, wavelength_max=14)
    cst.build_project_vba("SquarePillar", "SquarePillar_vba_build", 8, 15, True)

    basic_opts.set_acc_dc(cst)
    basic_opts.set_FDSolver_source(cst, "Zmin", "TM(0,0)")

    materials_opts.change_substrate(cst, "=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")
    materials_opts.change_pillar(cst, "=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")

    param_opts.SquarePillar(cst).set_period_parallel_sweep(p_start=3.0, p_end=3.4, p_step=0.2,
                                                           h_step=0.5, l_step=0.05,
                                                           h_start=4.0, l_start=0.2,
                                                           h_end=5.0, l_end=0.3,
                                                           start_now = False)
    # exec_parallel_sweep_from_list(cst, sweep_list, 3)

    # cst.close()
    print("Done")
    return



def main():
    misc.print_logo()

    # Parse command line arguments
    args = parser_constructor()

    # Handle command line arguments
    if args.mode == "http":
        logging.info("Starting HyperMetaSim in HTTP mode...")
        uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")


    elif args.mode == "cli":
        logging.info("Starting HyperMetaSim in CLI mode...")
        pass


    elif args.mode == "playbook":
        logging.info("Starting HyperMetaSim in Playbook mode...")
        pass


    elif args.mode == "sql":
        logging.info("Starting HyperMetaSim in SQL mode...")

        # import data from .cst project file
        if args.action == "cst2mysql": 
            if not args.path:
                logging.error("Path is required for cst2mysql action.")
                return
            else: # check the existence of the path
                path = args.path
                if not os.path.exists(path):
                    logging.error(f"Path does not exist: {path}")
                    return

            # get args
            if args.sparam_names: 
                sparam_names = args.sparam_names[0] if type(args.sparam_names[0]) == list else args.sparam_names
            else:
                sparam_names = []
            compensation_length = args.compensation
            max_workers         = args.thread
            parallel_num        = args.file_parallel
            force               = args.force

            # execute import
            logging.info(f"Importing S-parameters from {path}...")
            results_opts.cst2mysql(path, sparam_names, compensation_length, parallel_num, max_workers, force)

        elif args.action == "dup-check":
            logging.info("Performing duplicate check...")
            results_opts.dup_check()

        else:
            logging.error(f"Unknown SQL action: {args.action}. Please use 'cst2mysql' or 'dup-check'.")
            return


    elif args.mode == "debug":
        logging.info("Starting HyperMetaSim in Debug mode...")
        debug()


    else:
        logging.error(f"Unknown mode: {args.mode}. Please use 'http', 'cli', 'playbook', 'sql', or 'debug'.")
        return


if __name__ == '__main__':
    main()