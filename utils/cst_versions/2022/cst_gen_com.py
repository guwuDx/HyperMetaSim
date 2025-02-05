import win32com.client
from win32com.client import gencache
import pythoncom
import pandas as pd

class CstInit:

    def __init__(self, name, cst_version=2022):
        self.name = name
        self.cst_version = cst_version
        self.app = None
        self.cst_module = gencache.EnsureModule('{EDD3E836-F537-4C6F-BE7D-6014C155CC7B}', 0, 1, 0)
        self.env_init()


    def env_check(self):
        try:
            cst_app = win32com.client.GetActiveObject("CSTStudio.Application")
        except:
            print("[ OK ] CST is not running, environment is Clean")
            return

        if cst_app:
            print("[ERRO] A CST instance is already running, please close it and try again")
            raise RuntimeError("CST is already running")
        else:
            print("[ OK ] CST is not running, Environment is Clean")


    def env_up(self, interactive=False):
        print("[INFO] Starting CST ...")
        cst_com = win32com.client.Dispatch("CSTStudio.Application")
        if cst_com:
            if isinstance(cst_com, int):
                print(f"[ERRO] CST startup failed and returned error code: {cst_com}")
                raise RuntimeError(f"[ERRO] CST startup failed and returned error code: {cst_com}")
            if not hasattr(cst_com, "GetFileMainVersion"):
                print("[ERRO] failed to get CST MAIN VERSION, please check the installation")
                raise RuntimeError("[ERRO] failed to get CST MAIN VERSION, please check the installation")
            if not hasattr(cst_com, "GetFilePatchVersion"):
                print("[ERRO] failed to get CST PATCH VERSION, please check the installation")
                raise RuntimeError("[ERRO] failed to get CST PATCH VERSION, please check the installation")
            if not interactive:
                cst_com.SetQuietMode(interactive)
                print("[INFO] CST was configured to run in interactive mode, user input is blocked")
            else:
                print("[WARN] CST was configured to run in interactive mode")
                print("[WARN] DO NOT interfere with the GUI unless you know what you are doing")
            self.app = cst_com
            return None
        else:
            print("[ERRO] CST startup failed, please check the installation")
            raise SystemExit(1)


    def env_init(self):
        pythoncom.CoInitialize()
        self.env_check()
        self.env_up()
        print("[ OK ] CST is up and running")


    def env_down(self):
        print("[INFO] Shutting down CST ...")
        cst_app = win32com.client.GetActiveObject("CSTStudio.Application")
        if cst_app:
            cst_app.Quit()
            print("[ OK ] CST shutdown successfully")
        else:
            print("[INFO] CST is not running")
        return None


    def new_mws_prj(self, prj_full_path):
        print(f"[INFO] Creating new CST project")
        mws = self.app.NewMWS()
        print(f"[ OK ] New project created successfully")
        print(f"[INFO] Saving new project as: {prj_full_path}")
        mws.SaveAs(prj_full_path, False)
        print(f"[ OK ] Project saved successfully")
        mws.Quit()
        return mws


    @classmethod
    def get_env_inst(self):
        return self.app

if __name__ == "__main__":
    cst = CstInit("CST")
    cst.new_mws_prj("C:/Users/pc/Documents/CST Projects/com_test.cst")
    cst.env_down()