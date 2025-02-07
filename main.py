from utils import misc
from utils import cst_general
import utils.macors_canva as canvas

import pythoncom
pythoncom.CoInitialize()

def main():

    cst = cst_general.CSTHandler(1)
    cst.open_template("square_pillar")
    cst.instantiate_template("square_pillar_inst")
    # cst.close()
    print("Done")


if __name__ == '__main__':
    main()