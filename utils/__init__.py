import importlib
from utils.misc import read_toml

version = read_toml("./config/service.toml", "cst")["version"]
print(f"[INFO] configued CST version: {version}")
print(f"[INFO] importing libraries for CST version: {version}")

try:
    cst_handler             = importlib.import_module(f"utils.cst_versions.{version}.cst_handler")

except ModuleNotFoundError:
    print(f"[ERRO] CST version {version} is not supported")
    raise ImportError()


basic_opts        = importlib.import_module(f"utils.cst_versions.{version}.basic_opts")
param_opts        = importlib.import_module(f"utils.cst_versions.{version}.param_opts")
materials_opts    = importlib.import_module(f"utils.cst_versions.{version}.materials_opts")
results_opts      = importlib.import_module(f"utils.cst_versions.{version}.results_opts")