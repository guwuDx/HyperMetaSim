import utils.misc as misc

misc.add_cst_lib_path()
from cst.interface import DesignEnvironment


def main():
    # cst_app = misc.cst_conn()
    # print(csti.DesignEnvironment)
    de = DesignEnvironment.connect_to_any_or_new()
    crr_proj = de.get_open_projects()
    print(crr_proj[0], '\n---')
    # de.close()


if __name__ == '__main__':
    main()