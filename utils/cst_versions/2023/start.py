import win32com.client

def main():
    # 1. 获取 CST 应用程序的 COM 对象
    cst_app = win32com.client.Dispatch("CSTStudio.Application")
    
    # 2. 新建一个 MWS (Microwave Studio) 工程
    #    注意: 如果使用不同模块，如Circuit，或CST有其他命名，也可能是 .NewMWS() / .NewProject() 等
    #    不同版本可能略有差异，可以结合 Macro Recorder 查看
    mws = cst_app.NewMWS()
    
    # 3. 通过“模板向导”方式，新建项目时指定Microwave&RF/Optical -> Periodic Structures -> Metamaterial – Full Structure
    #    如果你想直接跳过GUI模板向导，把必要的设置都在脚本里写好也是可以的。
    #    这里为了示例，先演示设置一下工程名称:
    project_name = "python_demo"
    #    一般可以在后续保存时指定工程路径即可
    
    # 4. 选择求解器为“Frequency Domain”
    #    宏录制时通常会看到类似： Solver.SelectSolver "Frequency Domain"
    mws.Invoke("SelectSolver", "Frequency Domain Solver")
    
    # 5. 设置单位
    #    宏录制下会类似： Units.SetLengthUnit "um"; Units.SetFreqUnit "THz" 等
    #    这里以 Invoke 的方式为例，也可以针对 "Units" 对象直接调用相关方法
    mws.Invoke("Units", "SetLengthUnit", "um")
    mws.Invoke("Units", "SetFrequencyUnit", "THz")
    mws.Invoke("Units", "SetTimeUnit", "ns")
    mws.Invoke("Units", "SetTemperatureUnit", "Kelvin")

    # 6. 设置波长范围(或者频率范围)
    #    在 Frequency Domain 下，如果你想用波长(um)来定义扫描范围，
    #    需要存储参数，或者切换到合适的频率设置方法。
    #    下面仅演示存储一个“参数”的方式，然后可以把它再用在 SolverSettings 中。
    #    比如把 120um, 220um 写成 CST 内的参数:
    mws.StoreParameter("wavelength_min", "120")
    mws.StoreParameter("wavelength_max", "220")
    
    #    也可以直接用：
    #    mws.Invoke("Solver", "SetRange", 120, 220)  # 视实际脚本接口定
    
    # 7. （可选）如果不需要添加任何监视器，这里就不做操作
    
    # 8. 最后保存工程到指定路径
    save_path = r"C:\Users\Public\Documents\CST_Projects\python_demo.cst"
    mws.Invoke("FileSaveAs", save_path)
    
    print(f"工程已保存至: {save_path}")

if __name__ == "__main__":
    main()
