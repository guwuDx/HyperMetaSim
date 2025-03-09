import copy
import time
import hashlib


class Canvas:
    def __init__(self):
        self.BR = "\" & vbLf & _\n\""
        self.vbac = list()
        self.objs = {}
        self.clear()


    def write(self, code, adapt=True):
        if adapt: # change the \n to BR, change the \" to \"\"
            code = code.replace("\"", "\"\"")
            code = code.replace("\n", self.BR)
        self.vbac.append(code)


    def add_code(self, obj_name, key, *values, adapt=True):
        if obj_name not in self.objs:
            self.objs[obj_name] = []

        if values:
            if adapt:
                value_str = "\"\", \"\"".join(list(map(str, values))) # convert to string
                code = f"  .{key} \"\"{value_str}\"\""
            else:
                value_str = "\", \"".join(list(map(str, values)))
                code = f"  .{key} \"{value_str}\""
        else:
            code = f"  .{key}"

        self.objs[obj_name].append(code)


    def del_obj(self, obj_name):
        if obj_name in self.objs:
            del self.objs[obj_name]
        else:
            print(f"[WARN] object {obj_name} does not exist")


    def _clr_obj(self):
        self.objs = {}


    def _write_obj(self, obj_name=None, adapt=True):
        if adapt:
            if obj_name is None:
                for obj in self.objs:
                    obj_code = f"{self.BR}".join(self.objs[obj])
                    obj_code = f"With {obj}{self.BR}{obj_code}{self.BR}End With"
                    self.vbac.append(obj_code)
            else:
                if obj_name in self.objs:
                    obj_code = self.BR.join(self.objs[obj_name])
                    self.vbac.append(obj_code)
                else:
                    print(f"[WARN] object {obj_name} does not exist")
        else:
            if obj_name is None:
                for obj in self.objs:
                    obj_code = "\n".join(self.objs[obj])
                    obj_code = f"With {obj}\n{obj_code}\nEnd With"
                    self.vbac.append(obj_code)
            else:
                if obj_name in self.objs:
                    obj_code = "\n".join(self.objs[obj_name])
                    self.vbac.append(obj_code)
                else:
                    print(f"[WARN] object {obj_name} does not exist")


    def clear(self):
        self._clr_obj()
        self.vbac = []


    def send(self, cst_app, cmt=None, add_to_history=True, timeout=None, clear=True):
        """
        Send the VBA code to the CST application, 
        default to the currently active CST project.

        Args:
            cst_app (object): CST application handler object
            cmt: (str): comment for the VBA code
            add_to_history (bool): whether to add the VBA code to the history
            timeout (int): timeout for the CST application to execute the VBA code
            clear (bool): whether to clear the VBA code after sending
        """
        if add_to_history:
            self._write_obj()
            vba_code = self.BR.join(self.vbac)
            vba_code = self.vba_template.add_send_frame(vba_code, cmt)
            res = cst_app.send_vba(vba_code, timeout)
        else:
            self._write_obj(adapt=False)
            vba_code = "\n".join(self.vbac)
            vba_code = f"Sub Main()\n{vba_code}\nEnd Sub"
            res = cst_app.send_vba(vba_code, timeout)

        if clear: self.clear()
        return res # bool\


    def write_send(self, cst_app, code, cmt=None, add_to_history=True, timeout=None):
        self.write(code, adapt=add_to_history)
        return self.send(cst_app, cmt, add_to_history, timeout, clear=True)


    def preview(self, add_to_history=True):
        print("[INFO] The VBA code to be executed:\n")
        preview_instance = copy.deepcopy(self)
        if add_to_history:
            preview_instance._write_obj()
            vba_code = self.BR.join(preview_instance.vbac)
            vba_code = self.vba_template.add_send_frame(vba_code)
            print(vba_code)
        else:
            preview_instance._write_obj(adapt=False)
            vba_code = "\n".join(preview_instance.vbac)
            vba_code = f"Sub Main()\n{vba_code}\nEnd Sub"
            print(vba_code)
        return vba_code


    def write_to_file(self, file_path):
        self.vbac.append("End Sub\n")
        vba_code = "\n".join(self.vbac)
        with open(file_path, 'w') as f:
            f.write(vba_code)


    class vba_template:

        @staticmethod
        def get_mode_num_by_name(port, mode_name):
            return f"""
Dim Num As Long
With FloquetPort
    .Port (\"{port}\")
    success = .GetModeNumberByName (Num, \"{mode_name}\")
End With

If Not success Then
    Err.Raise 1000, , \"Failed to get mode number by {mode_name}\"
End If
"""


        @staticmethod
        def add_send_frame(vba_code, cmt=None):
            if cmt:
                pass
            else:
                ts = time.time()
                hs = hashlib.md5(str(ts).encode()).hexdigest()
                cmt = f"Python CST Automation {ts}{hs}"
            return f"""
Sub Main()
AddToHistory \"{cmt}\", \"{vba_code}\"
End Sub
"""


        @staticmethod
        def set_background(farfield_distance):
            return f"""
With Background 
     .ResetBackground 
     .XminSpace "0.0" 
     .XmaxSpace "0.0" 
     .YminSpace "0.0" 
     .YmaxSpace "0.0" 
     .ZminSpace "0.0" 
     .ZmaxSpace "{farfield_distance}" 
     .ApplyInAllDirections "False" 
End With 
"""


        @staticmethod
        def set_background_normal_material():
            return """
With Material 
     .Reset 
     .Rho "1.204"
     .ThermalType "Normal"
     .ThermalConductivity "0.026"
     .SpecificHeat "1005", "J/K/kg"
     .DynamicViscosity "0"
     .Emissivity "0"
     .MetabolicRate "0.0"
     .VoxelConvection "0.0"
     .BloodFlow "0"
     .MechanicsType "Unused"
     .IntrinsicCarrierDensity "0"
     .FrqType "all"
     .Type "Normal"
     .MaterialUnit "Frequency", "Hz"
     .MaterialUnit "Geometry", "m"
     .MaterialUnit "Time", "s"
     .MaterialUnit "Temperature", "Kelvin"
     .Epsilon "1.0"
     .Mu "1.0"
     .Sigma "0"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstSigma"
     .SetConstTanDStrategyEps "AutomaticOrder"
     .ConstTanDModelOrderEps "3"
     .DjordjevicSarkarUpperFreqEps "0"
     .SetElParametricConductivity "False"
     .ReferenceCoordSystem "Global"
     .CoordSystemType "Cartesian"
     .SigmaM "0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstSigma"
     .SetConstTanDStrategyMu "AutomaticOrder"
     .ConstTanDModelOrderMu "3"
     .DjordjevicSarkarUpperFreqMu "0"
     .SetMagParametricConductivity "False"
     .DispModelEps  "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .MaximalOrderNthModelFitEps "10"
     .ErrorLimitNthModelFitEps "0.1"
     .UseOnlyDataInSimFreqRangeNthModelEps "False"
     .DispersiveFittingSchemeMu "Nth Order"
     .MaximalOrderNthModelFitMu "10"
     .ErrorLimitNthModelFitMu "0.1"
     .UseOnlyDataInSimFreqRangeNthModelMu "False"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMu "False"
     .NLAnisotropy "False"
     .NLAStackingFactor "1"
     .NLADirectionX "1"
     .NLADirectionY "0"
     .NLADirectionZ "0"
     .Colour "0.6", "0.6", "0.6" 
     .Wireframe "False" 
     .Reflection "False" 
     .Allowoutline "True" 
     .Transparentoutline "False" 
     .Transparency "0" 
     .ChangeBackgroundMaterial
End With
"""


        @staticmethod
        def set_floquet_port_boundaries(enable_modes, farfield_distance):
            return f"""
With FloquetPort
     .Reset
     .SetDialogFrequency "1" 
     .SetDialogMediaFactor "1" 
     .SetDialogTheta "theta" 
     .SetDialogPhi "phi" 
     .SetPolarizationIndependentOfScanAnglePhi "0.0", "False"  
     .SetSortCode "+beta/pw" 
     .SetCustomizedListFlag "False" 
     .Port "Zmin" 
     .SetNumberOfModesConsidered "{enable_modes}" 
     .SetDistanceToReferencePlane "0.0" 
     .SetUseCircularPolarization "False" 
     .Port "Zmax" 
     .SetNumberOfModesConsidered "{enable_modes}" 
     .SetDistanceToReferencePlane "-{farfield_distance}" 
     .SetUseCircularPolarization "False" 
End With
"""


        @staticmethod
        def set_boundaries():
            return """
With Boundary
     .Xmin "unit cell"
     .Xmax "unit cell"
     .Ymin "unit cell"
     .Ymax "unit cell"
     .Zmin "open"
     .Zmax "open"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
     .ApplyInAllDirections "False"
     .XPeriodicShift "0.0"
     .YPeriodicShift "0.0"
     .ZPeriodicShift "0.0"
     .PeriodicUseConstantAngles "False"
     .SetPeriodicBoundaryAngles "theta", "phi"
     .SetPeriodicBoundaryAnglesDirection "outward"
     .UnitCellFitToBoundingBox "True"
     .UnitCellDs1 "0.0"
     .UnitCellDs2 "0.0"
     .UnitCellAngle "90.0"
End With
"""

__all__ = ["Canvas"]