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

