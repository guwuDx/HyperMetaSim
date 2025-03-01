from utils import misc

import pandas as pd
import os

misc.add_cst_lib_path()
from cst.interface import DesignEnvironment, Project

# project_name = "square_pillar"
_template_path = "/templates/"
_instance_path = "/instances/"
_projects_path = misc.read_toml("./config/service.toml", "cst")["projects_path"]

class CSTHandler:
    class _DRC:
        padding = None
        h_l_ratio_upper_bound = None


    class _ACC_DC:
        max_num_of_cpu_devs = None
        max_threads = None
        max_params_parallel = None
        only_0D1D = None
        use_shared_dir = None
        use_dc_mem_setting = None
        min_dc_mem_limit = None
        remote_mesh = None


    class _TimeOut:
        solver = None


    def __init__(self, debug=False):
        self.de = None
        self.pid = None
        self._projects_path = None
        self._template_path = None
        self._instance_path = None
        self.crr_prj = None
        self.crr_prj_properties = {
            "type": None, 
            "wavelegnth_min": None,
            "wavelegnth_max": None,
        }
        self.prjs = pd.DataFrame(columns=["project_instance", "project_properties"])
        self._get_cnf()

        if debug: self._conn_de()
        else: self._new_de()


    def _new_de(self):
        # cst_app = misc.cst_conn()
        # print(csti.DesignEnvironment)

        print("[INFO] Starting a brand new CST Design Environment ...")
        de = DesignEnvironment()
        if de.is_connected():
            print('[ OK ] Connected to CST Design Environment successfully')
            print('[INFO] CST version: ', de.version())
            print('[INFO] pid: ', de.pid())
            prjs = de.list_open_projects()
            if prjs:
                print('[INFO] current projects: ')
                for prj in prjs:
                    print(prj)
            else:
                print('[INFO] No projects are opened')
            print('[INFO] Init OK')
        else:
            print('[ERRO] Failed to connect to CST Design Environment, please check the installation')
            raise RuntimeError("Failed to connect to CST Design Environment")
        self.de = de
        self.pid = de.pid()


    def _conn_de(self):
        self.de = DesignEnvironment.connect_to_any_or_new()


    def open_template(self, metastructure_type):
        projects_path = misc.read_toml("./config/service.toml", "cst")["projects_path"]

        source_project_path = f"{projects_path}{self._template_path}{metastructure_type}.cst"
        print("[INFO] Opening project: " + source_project_path)
        prj = self.de.open_project(source_project_path)
        prj.activate()
        print("[ OK ] Project \"" + prj.filename() + "\" opened successfully")
        # prj_type = prj.project_type()
        
        print("Accessing to Modeler ...")
        if prj.modeler.is_solver_running():
            print("[WARN] Solver is running, Aborting ...")
            prj.modeler.abort_solver()
            print("[ OK ] Solver aborted successfully")
        else:
            print("[INFO] Solver is not running, continuing ...")
        self.crr_prj = prj
        self.crr_prj_properties["type"] = metastructure_type


    def instantiate_template(self,
                             project_name,
                             wavelength_min,
                             wavelength_max
                             ):
        from utils import basic_operations

        crr_prj_type = self.crr_prj_properties["type"]

        print("[INFO] Instantiating Project ...")
        instance_project_path = f"{self._projects_path}{self._instance_path}/{crr_prj_type}/{project_name}.cst"
        self.crr_prj.save(path=instance_project_path, include_results=False)
        print("[ OK ] Project instantiated successfully")
        print("[INFO] current project is: ", self.crr_prj.filename())

        basic_operations.set_prj_wavelength(self, wavelength_min, wavelength_max)
        self.crr_prj_properties["type"] = crr_prj_type
        prj_dict = [{
            "project_instance": self.crr_prj, 
            "project_properties": self.crr_prj_properties
        }]
        self.prjs = pd.concat([self.prjs, pd.DataFrame(prj_dict)], ignore_index=True)
        # crr_proj = de.get_open_projects()
        # print(crr_proj[0], '\n---')


    def send_vba(self,
                 vba_code: str = None,
                 timeout: int = None
                 ):

        if not vba_code:
            print("[WARN] No VBA code is provided")
            return
        return self.crr_prj.schematic.execute_vba_code(vba_code, timeout=timeout)


    def run_solver(self,
                   prj=None,
                   blocked: bool = True,
                   timeout: int = None
                   ):

        if not prj:
            prj = self.crr_prj

        if blocked:
            print("[INFO] Running Solver in Foreground ...")
            prj.modeler.run_solver(timeout=timeout)
            print("[ OK ] Solver finished successfully")
        else:
            print("[INFO] Running Solver in Background ...")
            prj.modeler.start_solver(timeout=timeout)
            print("[INFO] the simulation is submitted to the solver")


    def _get_cnf(self):
        print("[INFO] Reading CST configurations ...")
        cnf = misc.read_toml("./config/service.toml", "cst")
        self._projects_path = cnf["projects_path"]
        self._template_path = cnf.get("template_path", _template_path)
        self._instance_path = cnf.get("instance_path", _instance_path)
        
        drc = misc.configure_drc()
        self._DRC.h_l_ratio_upper_bound = drc["h_l_ratio_upper_bound"]
        units = drc.get("geometric_units", "um")
        if units == "um":
            self._DRC.padding = drc["padding"]
        elif units == "mm":
            self._DRC.padding = drc["padding"] * 1000
        elif units == "nm":
            self._DRC.padding = drc["padding"] / 1000
        else:
            print("[ERRO] Unsupported geometric units")
            raise ValueError("Unsupported geometric units, could only be um, mm or nm")

        acc_dc = misc.configure_acc_and_dc()
        self._ACC_DC.max_num_of_cpu_devs = acc_dc.get("max_num_of_cpu_devs", 1)
        self._ACC_DC.max_threads = acc_dc.get("max_threads", 1024)

        self._ACC_DC.max_params_parallel = acc_dc.get("max_params_parallel", 99)
        self._ACC_DC.only_0D1D = acc_dc.get("only_0D1D", True)
        self._ACC_DC.use_shared_dir = acc_dc.get("use_shared_dir", True)
        self._ACC_DC.use_dc_mem_setting = acc_dc.get("use_dc_mem_setting", False)
        self._ACC_DC.min_dc_mem_limit = acc_dc.get("min_dc_mem_limit", 0)
        self._ACC_DC.remote_mesh = acc_dc.get("remote_mesh", False)

        self._TimeOut.solver = cnf.get("solver_timeout", 300)

        print("[ OK ] CST configurations read successfully")


    def save_crr_prj(self):
        print("[INFO] Saving current project ...")
        self.crr_prj.save()
        print("[ OK ] Project saved successfully")


    def close(self, force=False):
        if force:
            print("[WARN] Closing CST Design Environment compulsorily ...")
            os.kill(self.pid, 9)
            print("[ OK ] CST Design Environment closed successfully")
        else:
            print("[INFO] Closing CST Design Environment ...")
            self.de.close()
            print("[ OK ] CST Design Environment closed successfully")
