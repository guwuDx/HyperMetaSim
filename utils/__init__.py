import importlib
from utils.misc import read_toml

version = read_toml("./config/service.toml", "cst")["version"]
print(f"[INFO] configued CST version: {version}")
print(f"[INFO] importing libraries for CST version: {version}")

try:
    cst_handler             = importlib.import_module(f"utils.cst_versions.{version}.cst_handler")
    param_operations        = importlib.import_module(f"utils.cst_versions.{version}.param_operations")
    basic_operations        = importlib.import_module(f"utils.cst_versions.{version}.basic_operations")
    materials_operations    = importlib.import_module(f"utils.cst_versions.{version}.materials_operations")

except ModuleNotFoundError:
    print(f"[ERRO] CST version {version} is not supported")
    raise ImportError()