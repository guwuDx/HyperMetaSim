import sys
import toml


def read_toml(file_path, chunk_name):
    with open(file_path, 'r') as f:
        config = toml.load(f)
    return config[chunk_name]


def add_cst_lib_path():
    lib_path = read_toml("./config/service.toml", "cst")["cst_py_lib_path"]
    sys.path.append(lib_path)


def configure_drc():
    drc_config = read_toml("./config/service.toml", "drc")
    return drc_config