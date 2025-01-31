from utils import misc

misc.add_cst_lib_path()
from cst.interface import DesignEnvironment
import cst.results


def main():
    # cst_app = misc.cst_conn()
    # print(csti.DesignEnvironment)

    project_name = "square_pillar"
    template_path = "/templates/"
    instance_path = "/instances/"
    projects_path = misc.read_toml("./config/service.toml", "cst")["projects_path"]

    de = DesignEnvironment()
    if de.is_connected():
        print('Connected to CST Design Environment successfully')
        print('CST version: ', de.version())
        print('pid: ', de.pid())
        print('current project: ')
        de.list_open_projects()
        print('Init OK')
    projects_path = misc.read_toml("./config/service.toml", "cst")["projects_path"]

    source_project_path = f"{projects_path}{template_path}{project_name}.cst"
    print("Opening project: " + source_project_path)
    prj = de.open_project(source_project_path)
    prj.activate()
    print("Project " + prj.filename() + " opened successfully")
    # prj_type = prj.project_type()

    print("Instantiating Project ...")
    instance_project_path = f"{projects_path}{instance_path}{project_name}_inst.cst"
    prj.save(path=instance_project_path, include_results=False)
    print("Project instantiated successfully")
    print("current project is: ", prj.filename())
    # crr_proj = de.get_open_projects()
    # print(crr_proj[0], '\n---')

    print("Accessing to Modeler ...")
    mdl = prj.get_modeler()

    de.close()


if __name__ == '__main__':
    main()