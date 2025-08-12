from utils import misc

import pandas as pd
import os
import time
import json

from typing import Dict

misc.add_cst_lib_path()
from cst.interface import DesignEnvironment

# project_name = "square_pillar"
_template_path = "/templates/"
_instance_path = "/instances/"
_projects_path = misc.read_config("./config/service.json", "cst")["projects_path"]



class CSTProjectWrapper:
    def __init__(self, project_instance, project_properties:dict):
        self.project_instance = project_instance
        self.project_properties = project_properties



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


    def __init__(self, pid=None):
        self.de = None
        self.pid = None
        self._projects_path = None
        self._template_path = None
        self._instance_path = None
        self.crr_prj = None
        self.materials = {}
        self.crr_prj_properties = {
            "type": None, 
            "period": None,
            "theta": 0,
            "phi": 0,
            "wavelength_min": None,
            "wavelength_max": None,
            "substrate_material": "",
            "pillar_material": "",
            "farfield": None
        }
        # self.prjs = pd.DataFrame(columns=["project_instance", "project_properties"])
        self.projects: Dict[str, CSTProjectWrapper] = {}
        self._get_cnf()

        if pid: self._conn_de(pid)
        else: self._new_de()
        print(self.de.version())


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


    def _conn_de(self, pid):
        self.de = DesignEnvironment.connect(pid)


    def open_template(self, metastructure_type):
        projects_path = misc.read_config("./config/service.json", "cst")["projects_path"]

        # construct the path to the template project
        source_project_path = f"{projects_path}{self._template_path}{metastructure_type}.cst"

        # check if the template project exists and open it
        if not os.path.exists(source_project_path):
            print("[ERRO] Template project does not exist: " + source_project_path)
            raise FileNotFoundError(f"Template project {metastructure_type} not found in {self._template_path}")
        print("[INFO] Opening project: " + source_project_path)
        prj = self.de.open_project(source_project_path)
        prj.activate()
        print("[ OK ] Project \"" + prj.filename() + "\" opened successfully")

        # check the solver status
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
        from utils import basic_opts
        crr_prj_type = self.crr_prj_properties["type"]

        # save as a new project instance
        print("[INFO] Instantiating Project ...")
        instance_project_path = f"{self._projects_path}{self._instance_path}/{crr_prj_type}/{project_name}.cst"
        self.crr_prj.save(path=instance_project_path, include_results=False)
        print("[ OK ] Project instantiated successfully")
        full_name = self.crr_prj.filename()
        file_name = os.path.basename(full_name)[__file__:-4]
        print("[INFO] Project name is: ", file_name)
        print("[INFO] current project is: ", full_name)

        # set basic project properties
        basic_opts.set_prj_wavelength(self, wavelength_min, wavelength_max)
        self.crr_prj_properties["type"] = crr_prj_type
        self.crr_prj_properties["wavelength_min"] = wavelength_min
        self.crr_prj_properties["wavelength_max"] = wavelength_max

        self.crr_prj_properties["farfield"] = wavelength_max
        # self.crr_prj_properties["farfield"] = misc.farfield_evaluator(wavelength_min)

        self.projects[full_name] = CSTProjectWrapper(
            project_instance=self.crr_prj,
            project_properties=self.crr_prj_properties
        )

        # crr_proj = de.get_open_projects()
        # print(crr_proj[0], '\n---')


    def build_project_vba(self, shape_type, prj_name, wavelength_min, wavelength_max, save=False):
        # self.restore_properties()
        from .build_model import SquarePillar

        self.crr_prj = self.de.new_mws()
        temp_file_name = self.crr_prj.filename()

        self.crr_prj_properties["type"] = shape_type
        self.crr_prj_properties["wavelength_min"] = wavelength_min
        self.crr_prj_properties["wavelength_max"] = wavelength_max

        self.crr_prj_properties["farfield"] = wavelength_max
        # self.crr_prj_properties["farfield"] = misc.farfield_evaluator(wavelength_min)

        if shape_type == "SquarePillar":
            SquarePillar(self)

        else:
            print("[ERRO] Unsupported shape type")
            raise ValueError("Unsupported shape type")

        self.crr_prj.activate()
        # restore current project properties
        self.restore_properties()

        if save:
            print("[INFO] Saving project ...")
            self.crr_prj.save(path=f"{self._projects_path}{self._instance_path}/{shape_type}/{prj_name}.cst",
                              include_results=False)
            self.restore_properties()
            self.close_project(temp_file_name)
            print("[ OK ] Project saved successfully")


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
                   blocked:     bool = True,
                   timeout:     int  = None,
                   safe_mode:   bool = False
                   ):

        if not prj:
            prj = self.crr_prj

        if blocked:
            if safe_mode and timeout:
                print("[INFO] Running Solver in safe mode with timeout ...")

                watch_dog = 0
                prj.modeler.start_solver(timeout=None)
                time.sleep(5)
                while True:
                    if prj.modeler.is_solver_running():
                        print(".", end="")
                        time.sleep(1)
                        watch_dog += 1
                        if watch_dog >= timeout:
                            print()
                            print("[WARN] Timeout reached, Aborting ...")
                            prj.modeler.abort_solver()
                            print("[ OK ] Solver aborted successfully")
                            break
                    else:
                        print()
                        print("[ OK ] Solver finished")
                        break
            else:
                print("[INFO] Running Solver in Foreground ...")
                prj.modeler.run_solver(timeout=timeout)
                print("[ OK ] Solver finished")
        else:
            print("[INFO] Running Solver in Background ...")
            prj.modeler.start_solver(timeout=timeout)
            print("[INFO] the simulation is submitted to the solver")


    def _get_cnf(self):
        print("[INFO] Reading CST configurations ...")
        cnf = misc.read_config("./config/service.json", "cst")
        self._projects_path = cnf["projects_path"]
        # self._permanent_path = cnf["permanent_path"]
        self._template_path = cnf.get("template_path", _template_path)
        self._instance_path = cnf.get("instance_path", _instance_path)
        self._permanent_path = cnf.get("permanent_path")

        drc = misc.read_config("./config/service.json", "drc")
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

        acc_dc = misc.read_config("./config/service.json", "acc_dc")
        self._ACC_DC.max_num_of_cpu_devs    = acc_dc.get("max_num_of_cpu_devs", 1)
        self._ACC_DC.max_threads            = acc_dc.get("max_threads", 1024)

        self._ACC_DC.max_params_parallel    = acc_dc.get("max_params_parallel", 99)
        self._ACC_DC.only_0D1D              = acc_dc.get("only_0D1D", True)
        self._ACC_DC.use_shared_dir         = acc_dc.get("use_shared_dir", True)
        self._ACC_DC.use_dc_mem_setting     = acc_dc.get("use_dc_mem_setting", False)
        self._ACC_DC.min_dc_mem_limit       = acc_dc.get("min_dc_mem_limit", 0)
        self._ACC_DC.remote_mesh            = acc_dc.get("remote_mesh", False)

        self._TimeOut.solver = cnf.get("solver_timeout", 300)

        print("[ OK ] CST configurations read successfully")


    def save_crr_prj(self):
        print("[INFO] Saving current project ...")
        self.crr_prj.save(self.crr_prj.filename())
        self.update_project_properties()
        self.write_properties()
        print("[ OK ] Project saved successfully")


    def duplicate_prj(self, new_prj_name):
        print("[INFO] Duplicating project ...")
        self.save_crr_prj()

        # get the current project name and path
        full_name = self.crr_prj.filename()
        prj_path = os.path.dirname(full_name)

        # duplicate the project and its properties
        # ppts = self.crr_prj_properties
        self.crr_prj.save(path=f"{prj_path}/{new_prj_name}.cst", include_results=False)
        print("[INFO] Duplicated project is: ", self.crr_prj.filename())

        # restore the new project and its properties
        self.restore_properties()
        new_prj = self.crr_prj

        # recover the duplicated project
        self.crr_prj = self.de.open_project(full_name)
        self.update_project_instance()
        self.restore_properties()
        self.crr_prj.activate()
        print("[ OK ] Project duplicated successfully")

        return new_prj


    def settle_prj(self, prj=None):
        if not prj:
            prj = self.crr_prj
        else:
            prj.activate()
        print("[INFO] Settling project ...")
        permant_path = self._permanent_path
        prj.save(path=permant_path, include_results=False)


    def close_project(self, full_name:str):
        """
        Close the specified or current project and remove its properties from the list.
        Args:
            full_name (str): Key to the project in self.projects. If None, close the current project.
        """
        print("[INFO] Closing project and removing its properties from the list ...")
        if not full_name: # close the current project
            # get the current project name and path
            full_name = self.crr_prj.filename()
            prj = self.crr_prj

            # delete the project from self.projects
            if full_name in self.projects:
                del self.projects[full_name]
            else:
                print("[WARN] Project not found in the list")
                return

            # close the project
            prj.close()

        else: # close a specific project
            # delete the project from self.projects
            if full_name in self.projects:
                prj = self.projects[full_name].project_instance
                del self.projects[full_name]

                if not (self.crr_prj == prj): # prevent closing the current project unintentionally
                    prj.close()
                else:
                    print("[WARN] Specified project possess the same pointer as the current project, refusing to close it")

        print("[ OK ] Project closed successfully")


    def close(self, force=False):
        if force:
            print("[WARN] Closing CST Design Environment compulsorily ...")
            os.kill(self.pid, 9)
            print("[ OK ] CST Design Environment closed successfully")
        else:
            print("[INFO] Closing CST Design Environment ...")
            self.de.close()
            print("[ OK ] CST Design Environment closed successfully")


    def write_properties(self):
        print("[INFO] Writing project properties ...")

        # get the current project directory and the path to the project properties
        prj_dir = self.crr_prj.filename()[:-4]
        prj_prop_path = prj_dir + "/project_properties.json"

        # save current project properties as json
        with open(prj_prop_path, "w") as f:
            json.dump(self.crr_prj_properties, f, indent=4, ensure_ascii=False)
        print("[ OK ] Project properties written successfully")


    def restore_properties(self):
        """
        Restore current project properties to the CSTProjectWrapper object.
        This method is called when a new project is opened or when the project is duplicated.
        """
        print("[INFO] Restoring project properties ...")
        # get the current project directory and the path to the project properties
        full_name = self.crr_prj.filename()
        self.projects[full_name] = CSTProjectWrapper(
            project_instance=self.crr_prj,
            project_properties=self.crr_prj_properties.copy()
        )
        print("[ OK ] Project properties restored successfully")


    def find_project_by_type(self, prj_type, return_form="full_name"):
        """
        Find all projects with a specific type.
        Returns a list of CSTProjectWrapper objects.
        """
        if return_form == "full_name":
            return [
                full_name for full_name, wrapper in self.projects.items()
                if wrapper.project_properties.get("type") == prj_type
            ]
        elif return_form == "instance":
            return [
                wrapper for wrapper in self.projects.values()
                if wrapper.project_properties.get("type") == prj_type
            ]


    def update_project_properties(self):
        """
        Update the current project's stored properties in the projects dict.
        """
        full_name = self.crr_prj.filename()
        if full_name in self.projects:
            self.projects[full_name].project_properties = self.crr_prj_properties.copy()
        else:
            print("[WARN] Project not found in the list")


    def update_project_instance(self):
        """
        Update the current project's instance in the projects dict.
        """
        full_name = self.crr_prj.filename()
        if full_name in self.projects:
            self.projects[full_name].project_instance = self.crr_prj
        else:
            print("[WARN] Project not found in the list")


    def switch_to_project(self, full_name:str):
        """
        Switch to the specified project using its key.

        Args:
            full_name (str): Key to the project in self.projects
        """
        # Save current properties before switching
        self.update_project_properties()

        # Switch project
        wrapper = self.projects[full_name]
        self.crr_prj = wrapper.project_instance
        self.crr_prj.activate()
        self.crr_prj_properties = wrapper.project_properties

        print(f"[ OK ] Switched to project: {self.crr_prj.filename()}")