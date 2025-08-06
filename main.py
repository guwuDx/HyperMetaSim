import logging
import utils.log_conf

from utils import misc
from utils import cst_handler
from utils import basic_opts
from utils import materials_opts
from utils import param_opts
from utils import results_opts

import uvicorn
import argparse
from interfaces.app import app



def debug():
    prj = r"C:/Users/27950/OneDrive/Desktop/SquarePillar__surface_cstl2"
    results_opts.cst2mysql(prj, compensation_length=1, max_workers=4, parallel_num=-1, force=False)
    return
    # sparam_name = ["SZmax(2),Zmin(2)", "SZmax(1),Zmax(2)"]
    # data = results_opts.fetch_sparams(prj, sparam_name, plural=True)
    # return
    cst = cst_handler.CSTHandler()
    # cst.open_template("SquarePillar")
    # cst.instantiate_template("SquarePillar_Fine_surface_cstl3", wavelength_min=8, wavelength_max=14)
    cst.build_project_vba("SquarePillar", "SquarePillar_vba_build", 8, 15, True)

    basic_opts.set_acc_dc(cst)
    basic_opts.set_FDSolver_source(cst, "Zmin", "TE(0,0)")

    materials_opts.change_substrate(cst, "=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")
    materials_opts.change_pillar(cst, "=Si_crystal=_freq-r-i_0.0310-310um_ByFranta-300K_2017")

    param_opts.SquarePillar(cst).set_period_parallel_sweep(p_start=3.2, p_end=3.8, p_step=0.2,
                                                           h_step=0.25, l_step=0.04,
                                                           h_start=4, l_start=0.2,
                                                           start_now = False)
    # exec_parallel_sweep_from_list(cst, sweep_list, 3)

    # cst.close()
    print("Done")
    return



def main():
    misc.print_logo()
    parser = argparse.ArgumentParser(description="HyperMetaSim - A CST-based Automatic simulation tool for metamaterials.")
    parser.add_argument("--mode", type=str, default="http", 
                        help="Mode of operation: 'http' for web interface, \n" \
                            "                    'cli' for command line interface. \n" \
                            "                    'playbook' for running a predefined sequence of operations. \n" \
                            "                    'sql' for database operations. \n" \
                            "                    'debug' for debugging mode. Do not use in production. \n" \
                            "                     Default is 'http'.") 
    parser.add_argument("--action", type=str, 
                        help="SQL action type: 'import:[Sparameters]' (import data), 'dup-check' (check for duplicates).")
    parser.add_argument("--path", type=str, 
                        help="Path to data file/directory for project being processed.")

    args = parser.parse_args()


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
        if args.action == "import":
            if not args.path:
                logging.error("Path is required for import action.")
                return
            logging.info(f"Importing S-parameters from {args.path}...")
            results_opts.cst2mysql(args.path, sparam_names, compensation_length, max_workers, parallel_num, force)

        elif args.action == "dup-check":
            logging.info("Performing duplicate check...")
            results_opts.dup_check()

        else:
            logging.error(f"Unknown SQL action: {args.action}. Please use 'import' or 'dup-check'.")
            return


    elif args.mode == "debug":
        logging.info("Starting HyperMetaSim in Debug mode...")
        debug()


    else:
        logging.error(f"Unknown mode: {args.mode}. Please use 'http', 'cli', 'playbook', 'sql', or 'debug'.")
        return


if __name__ == '__main__':
    main()