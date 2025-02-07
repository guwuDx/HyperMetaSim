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

    def __init__(self, debug=False):
        self.de = None
        self.pid = None
        self.crr_prj = None
        self.crr_prj_type = None
        self._projects_path = None
        self._template_path = None
        self._instance_path = None
        self.prjs = pd.DataFrame(columns=["project_instance", "project_type"])
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
        self.crr_prj_type = metastructure_type


    def instantiate_template(self, project_name):
        print("[INFO] Instantiating Project ...")
        instance_project_path = f"{self._projects_path}{self._instance_path}/{self.crr_prj_type}/{project_name}.cst"
        self.crr_prj.save(path=instance_project_path, include_results=False)
        print("[ OK ] Project instantiated successfully")
        print("[INFO] current project is: ", self.crr_prj.filename())
        prj_dict = [{
            "project_instance": self.crr_prj, 
            "project_type": self.crr_prj_type
        }]
        self.prjs = pd.concat([self.prjs, pd.DataFrame(prj_dict)], ignore_index=True)
        # crr_proj = de.get_open_projects()
        # print(crr_proj[0], '\n---')


    def send_vba(self, vba_code, timeout=None):
        return self.crr_prj.schematic.execute_vba_code(vba_code, timeout=timeout)


    def _get_cnf(self):
        print("[INFO] Reading CST configurations ...")
        cnf = misc.read_toml("./config/service.toml", "cst")
        self._projects_path = cnf["projects_path"]
        self._template_path = cnf.get("template_path", _template_path)
        self._instance_path = cnf.get("instance_path", _instance_path)


    def close(self, force=False):
        if force:
            print("[WARN] Closing CST Design Environment compulsorily ...")
            os.kill(self.pid, 9)
            print("[ OK ] CST Design Environment closed successfully")
        else:
            print("[INFO] Closing CST Design Environment ...")
            self.de.close()
            print("[ OK ] CST Design Environment closed successfully")
