import sys
import toml
import win32com.client as win32


def read_toml(file_path, chunk_name):
    with open(file_path, 'r') as f:
        config = toml.load(f)
    return config[chunk_name]


def add_cst_lib_path():
    lib_path = read_toml("./config/service.toml", "cst")["install_path"]
    sys.path.append(lib_path)


def cst_conn(name="CST.Application"):
    cst_app = win32.Dispatch(name)
    print("CST version: ", cst_app.GetVersion())
    return cst_app


def new_project(cst_app):
    return cst_app.CreateNewProject()