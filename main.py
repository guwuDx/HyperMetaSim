import utils.misc as misc

misc.add_cst_lib_path()
import cst

print(cst.__file__)

cst_app = misc.cst_conn()