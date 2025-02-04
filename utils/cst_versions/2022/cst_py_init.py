from utils import misc

misc.add_cst_lib_path()
from cst.interface import DesignEnvironment, Project


def cst_py_init():
    # cst_app = misc.cst_conn()
    # print(csti.DesignEnvironment)

    project_name = "square_pillar"
    template_path = "/templates/"
    instance_path = "/instances/"
    projects_path = misc.read_toml("./config/service.toml", "cst")["projects_path"]

    print("[INFO] Connecting to CST Design Environment ...")
    de = DesignEnvironment()
    if de.is_connected():
        print('[ OK ] Connected to CST Design Environment successfully')
        print('[INFO] CST version: ', de.version())
        print('[INFO] pid: ', de.pid())
        print('[INFO] current project: ')
        de.list_open_projects()
        print('[INFO] Init OK')
    projects_path = misc.read_toml("./config/service.toml", "cst")["projects_path"]

    source_project_path = f"{projects_path}{template_path}{project_name}.cst"
    print("[INFO] Opening project: " + source_project_path)
    prj = de.open_project(source_project_path)
    prj.activate()
    print("[ OK ] Project " + prj.filename() + " opened successfully")
    # prj_type = prj.project_type()
    
    print("Accessing to Modeler ...")
    if prj.modeler.is_solver_running():
        print("[WARN] Solver is running, Aborting ...")
        prj.modeler.abort_solver()
        print("[ OK ] Solver aborted successfully")
    else:
        print("[INFO] Solver is not running, continuing ...")

    print("[INFO] Instantiating Project ...")
    instance_project_path = f"{projects_path}{instance_path}{project_name}_inst.cst"
    prj.save(path=instance_project_path, include_results=False)
    print("[ OK ] Project instantiated successfully")
    print("[INFO] current project is: ", prj.filename())
    # crr_proj = de.get_open_projects()
    # print(crr_proj[0], '\n---')

    return de, prj