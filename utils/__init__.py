import importlib
from utils.misc import read_toml

version = read_toml("./config/service.toml", "cst")["version"]
print(f"[INFO] configued CST version: {version}")

modules_path = f"utils.cst_versions.{version}.cst_gen"

try:
    cst_py_init = importlib.import_module(modules_path)
except ModuleNotFoundError:
    print(f"[ERRO] CST version {version} is not supported")
    raise ImportError()