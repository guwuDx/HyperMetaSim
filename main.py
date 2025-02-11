from utils import misc
from utils import cst_general
from utils import basic_operations
from utils import materials_operations
from utils import param_operations
from utils.macors_canva import Canvas

import pythoncom
pythoncom.CoInitialize()

def main():

    print("<<<<<<<<<<<<<<< CST Automation >>>>>>>>>>>>>>>")
    cst = cst_general.CSTHandler(debug=True)
    cst.open_template("SquarePillar")
    cst.instantiate_template("SquarePillar_inst", 10, 10.5)
    basic_operations.define_material(cst, "materials", "freq-r-i_TiO2_ThinFilm_0.211-1.69um_ByZhukovsky_2015")
    materials_operations.SquarePillar(cst).change_substrate("freq-r-i_TiO2_ThinFilm_0.211-1.69um_ByZhukovsky_2015")
    materials_operations.SquarePillar(cst).change_pillar("freq-r-i_TiO2_ThinFilm_0.211-1.69um_ByZhukovsky_2015")
    # basic_operations.modify_param(cst_app=cst, param_name="p", value=10)
    ls = param_operations.SquarePillar(cst).generate_sweep_squence(0.5, 0.25, 0.01)
    canvas = Canvas()

    # cst.close()
    print("Done")


if __name__ == '__main__':
    main()