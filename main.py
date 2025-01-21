import utils.misc as misc

misc.add_cst_lib_path()
from cst.interface import DesignEnvironment


def main():
    # cst_app = misc.cst_conn()
    # print(csti.DesignEnvironment)
    de = DesignEnvironment.new()
    
    de.close()


if __name__ == '__main__':
    main()