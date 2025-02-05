from utils import misc
from utils import cst_gen
from utils import macros_canva

import pythoncom
pythoncom.CoInitialize()

def main():

    cst = cst_gen.CSTHandler()
    cst.open_template("square_pillar")
    cst.instantiate_template("square_pillar_inst")
    macros_canva.Canvas.GetApplicationName(cst)
    cst.close()
    print("Done")


if __name__ == '__main__':
    main()