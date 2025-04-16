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
        def set_floquet_port_boundaries(zmax_enable_modes, zmin_enable_modes, farfield_distance):
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
     .SetNumberOfModesConsidered "{zmin_enable_modes}" 
     .SetDistanceToReferencePlane "-h1" 
     .SetUseCircularPolarization "False" 
     .Port "Zmax" 
     .SetNumberOfModesConsidered "{zmax_enable_modes}" 
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

        @staticmethod
        def set_mws_baisc(wavelength_min:float, wavelength_max:float):
            return f"""
'set the units
With Units
    .Geometry "um"
    .Frequency "THz"
    .Voltage "V"
    .Resistance "Ohm"
    .Inductance "NanoH"
    .TemperatureUnit  "Kelvin"
    .Time "ns"
    .Current "A"
    .Conductance "Siemens"
    .Capacitance "PikoF"
End With

'----------------------------------------------------------------------------

'set the wavelength range
Solver.WavelengthRange "{wavelength_min}", "{wavelength_max}"

'----------------------------------------------------------------------------

Plot.DrawBox True

With Background
     .Type "Normal"
     .Epsilon "1.0"
     .Mu "1.0"
     .Rho "1.204"
     .ThermalType "Normal"
     .ThermalConductivity "0.026"
      .SpecificHeat "1005", "J/K/kg"
     .XminSpace "0.0"
     .XmaxSpace "0.0"
     .YminSpace "0.0"
     .YmaxSpace "0.0"
     .ZminSpace "0.0"
     .ZmaxSpace "0.0"
End With

With Boundary
     .Xmin "expanded open"
     .Xmax "expanded open"
     .Ymin "expanded open"
     .Ymax "expanded open"
     .Zmin "expanded open"
     .Zmax "expanded open"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With

' optimize mesh settings for planar structures

With Mesh
     .MergeThinPECLayerFixpoints "True"
     .RatioLimit "20"
     .AutomeshRefineAtPecLines "True", "6"
     .FPBAAvoidNonRegUnite "True"
     .ConsiderSpaceForLowerMeshLimit "False"
     .MinimumStepNumber "5"
     .AnisotropicCurvatureRefinement "True"
     .AnisotropicCurvatureRefinementFSM "True"
End With

With MeshSettings
     .SetMeshType "Hex"
     .Set "RatioLimitGeometry", "20"
     .Set "EdgeRefinementOn", "1"
     .Set "EdgeRefinementRatio", "6"
End With

With MeshSettings
     .SetMeshType "Tet"
     .Set "VolMeshGradation", "1.5"
     .Set "SrfMeshGradation", "1.5"
End With

With MeshSettings
     .SetMeshType "HexTLM"
     .Set "RatioLimitGeometry", "20"
End With

' change mesh adaption scheme to energy
' 		(planar structures tend to store high energy
'     	 locally at edges rather than globally in volume)

MeshAdaption3D.SetAdaptionStrategy "Energy"

' switch on FD-TET setting for accurate farfields

FDSolver.ExtrudeOpenBC "True"

'----------------------------------------------------------------------------

'change problem type
ChangeProblemType "Optical"

'----------------------------------------------------------------------------

With MeshSettings
     .SetMeshType "Tet"
     .Set "Version", 1%
End With

With Mesh
     .MeshType "Tetrahedral"
End With

'set the solver type
ChangeSolverType("HF Frequency Domain")

'----------------------------------------------------------------------------

"""


        @staticmethod
        def create_component():
            return """
Component.New "component1"
"""


        @staticmethod
        def create_substrate():
            return """
With Brick
     .Reset 
     .Name "substrate" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-p/2", "p/2" 
     .Yrange "-p/2", "p/2" 
     .Zrange "-h1", "0" 
     .Create
End With
"""


        @staticmethod
        def create_pillar(model:str, param:str, pillar_name="pillar"):
            return f"""
With {model}
     .Reset 
     .Name "{pillar_name}" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-{param}/2", "{param}/2" 
     .Yrange "-{param}/2", "{param}/2" 
     .Zrange "0", "h" 
     .Create
End With
"""


        @staticmethod
        def set_solver_basic():
            return """
Mesh.SetCreator "High Frequency" 

With FDSolver
     .Reset 
     .SetMethod "Tetrahedral", "General purpose" 
     .OrderTet "Second" 
     .OrderSrf "First" 
     .Stimulation "Zmin", "TM(0,0)" 
     .ResetExcitationList 
     .AutoNormImpedance "False" 
     .NormingImpedance "50" 
     .ModesOnly "False" 
     .ConsiderPortLossesTet "False" 
     .SetShieldAllPorts "False" 
     .AccuracyHex "1e-6" 
     .AccuracyTet "1e-4" 
     .AccuracySrf "1e-3" 
     .LimitIterations "False" 
     .MaxIterations "0" 
     .SetCalcBlockExcitationsInParallel "True", "True", "" 
     .StoreAllResults "False" 
     .StoreResultsInCache "False" 
     .UseHelmholtzEquation "True" 
     .LowFrequencyStabilization "True" 
     .Type "Auto" 
     .MeshAdaptionHex "False" 
     .MeshAdaptionTet "True" 
     .AcceleratedRestart "True" 
     .FreqDistAdaptMode "Distributed" 
     .NewIterativeSolver "True" 
     .TDCompatibleMaterials "False" 
     .ExtrudeOpenBC "True" 
     .SetOpenBCTypeHex "Default" 
     .SetOpenBCTypeTet "Default" 
     .AddMonitorSamples "True" 
     .CalcPowerLoss "True" 
     .CalcPowerLossPerComponent "False" 
     .StoreSolutionCoefficients "True" 
     .UseDoublePrecision "False" 
     .UseDoublePrecision_ML "True" 
     .MixedOrderSrf "False" 
     .MixedOrderTet "False" 
     .PreconditionerAccuracyIntEq "0.15" 
     .MLFMMAccuracy "Default" 
     .MinMLFMMBoxSize "0.3" 
     .UseCFIEForCPECIntEq "True" 
     .UseEnhancedCFIE2 "True" 
     .UseFastRCSSweepIntEq "true" 
     .UseSensitivityAnalysis "False" 
     .UseEnhancedNFSImprint "False" 
     .RemoveAllStopCriteria "Hex"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
     .RemoveAllStopCriteria "Tet"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
     .RemoveAllStopCriteria "Srf"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
     .SweepMinimumSamples "3" 
     .SetNumberOfResultDataSamples "1001" 
     .SetResultDataSamplingMode "Automatic" 
     .SweepWeightEvanescent "1.0" 
     .AccuracyROM "1e-4" 
     .AddSampleInterval "", "", "1", "Automatic", "True" 
     .AddSampleInterval "", "", "", "Automatic", "False" 
     .ConsiderPortLossesTet "False" 
     .MPIParallelization "False"
     .UseDistributedComputing "False"
     .NetworkComputingStrategy "RunRemote"
     .NetworkComputingJobCount "3"
     .UseParallelization "True"
     .MaxCPUs "1024"
     .MaximumNumberOfCPUDevices "2"
End With

With IESolver
     .Reset 
     .UseFastFrequencySweep "True" 
     .UseIEGroundPlane "False" 
     .SetRealGroundMaterialName "" 
     .CalcFarFieldInRealGround "False" 
     .RealGroundModelType "Auto" 
     .PreconditionerType "Auto" 
     .ExtendThinWireModelByWireNubs "False" 
     .ExtraPreconditioning "False" 
End With

With IESolver
     .SetFMMFFCalcStopLevel "0" 
     .SetFMMFFCalcNumInterpPoints "6" 
     .UseFMMFarfieldCalc "True" 
     .SetCFIEAlpha "0.500000" 
     .LowFrequencyStabilization "False" 
     .LowFrequencyStabilizationML "True" 
     .Multilayer "False" 
     .SetiMoMACC_I "0.0001" 
     .SetiMoMACC_M "0.0001" 
     .DeembedExternalPorts "True" 
     .SetOpenBC_XY "True" 
     .OldRCSSweepDefintion "False" 
     .SetRCSOptimizationProperties "True", "100", "0.00001" 
     .SetAccuracySetting "Custom" 
     .CalculateSParaforFieldsources "True" 
     .ModeTrackingCMA "True" 
     .NumberOfModesCMA "3" 
     .StartFrequencyCMA "-1.0" 
     .SetAccuracySettingCMA "Default" 
     .FrequencySamplesCMA "0" 
     .SetMemSettingCMA "Auto" 
     .CalculateModalWeightingCoefficientsCMA "True" 
     .DetectThinDielectrics "True" 
End With

"""

__all__ = ["Canvas"]