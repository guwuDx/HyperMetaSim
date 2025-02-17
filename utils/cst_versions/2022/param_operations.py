import numpy as np
from math import ceil, floor
from utils.macors_canva import Canvas



class SquarePillar:
    def __init__(self, cst_handler):
        self.cst_handler = cst_handler
        self.padding = cst_handler._DRC.padding
        self.h_l_ratio_upper_bound = cst_handler._DRC.h_l_ratio_upper_bound
        self.wavelength_min = cst_handler.crr_prj_properties["wavelegnth_min"]
        self.wavelength_max = cst_handler.crr_prj_properties["wavelegnth_max"]
        self.canvas = Canvas()
        self.sweep_list = None


    def generate_sweep_squence(self, p_step, h_step, l_step, output_file_path=None):
        '''
        Generate a list of square pillar parameters for ParaSweep simulation.
        Output the list of parameters in the following format and save it to a file:
        l < p - 2 * padding;
        l > h / h_l_ratio_upper_bound;

        input:
        p_step(h_step, l_step): float, the step size of the parameters.
        output_file_path:       string, the path of the output file.

        output:
        if output_file_path is provided, save the list of parameters to the file.
        '''
        wavelength_min = self.wavelength_min
        wavelength_max = self.wavelength_max
        padding = self.padding
        h_l_ratio_upper_bound = self.h_l_ratio_upper_bound
        print(f"[INFO] Generating Parameter Sweep List ...")

        # solve for the range of parameters according to the wavelength range
        p_min = floor((wavelength_min / 4) / p_step) * p_step
        p_max = ceil((wavelength_max / 2) / p_step) * p_step

        # generate a list of parameters
        parameters = []
        for p in np.arange(p_min, p_max, p_step):
            for l in np.arange(l_step, p - 2 * padding, l_step):
                for h in np.arange(max(h_step, padding), l * h_l_ratio_upper_bound, h_step):
                    parameters.append([p, h, l])
        parameters = np.around(parameters, decimals=3)
        print(f"[INFO] {len(parameters)} parameters combinations were generated.")

        # save the list of parameters to a file if output_file_path is provided
        if output_file_path:
            with open(output_file_path, 'w') as f:
                f.write('p, h, l\n')
                for parameter in parameters:
                    f.write(', '.join(map(str, parameter)) + '\n')
        self.sweep_list = parameters
        return parameters


    def set_sweep_from_list(self, start_now=True):
        if self.sweep_list is None:
            print("[WARN] No sweep list is generated.")
            return

        print("[INFO] Setting up Parameter Sweep ...")
        obj = "ParameterSweep"

        self.canvas.add_code(obj, "DeleteAllSequences")
        for i, parameter in enumerate(self.sweep_list):
            seq = "seq" + str(i)
            p, h, l = parameter
            self.canvas.add_code(obj, "AddSequence", seq)
            self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "p", str(p))
            self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "h", str(h))
            self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "l", str(l))
        self.canvas.add_code(obj, "SetSimulationType", "Frequency")

        if start_now:
            print("[INFO] Setting completed, starting simulation ...")
            self.canvas.add_code(obj, "Start")
            res = self.canvas.send(self.cst_handler, add_to_history=False)
        else:
            print("[INFO] Setting completed, please start the simulation manually.")
            res = self.canvas.send(self.cst_handler, cmt="Set up paramSweep", add_to_history=False)

        # print("[INFO] vba code to be executed:\n")
        # self.canvas.preview()
        return res
    

    def set_params(self, 
                   p:       int = None, # the Arrangement period of the metastructure
                   h:       int = None, # the height of the pillar
                   l:       int = None, # the length of the pillar
                   phi:     int = None, # the azimuthal angle of the EM wave
                   theta:   int = None  # the incident angle of the EM wave
                   ):
        print("[INFO] Setting parameters ...")
        obj = "StoreParameter"

        if p:
            self.canvas.write(f"{obj} \"p\", \"{p}\"")
        if h:
            self.canvas.write(f"{obj} \"h\", \"{h}\"")
        if l:
            self.canvas.write(f"{obj} \"l\", \"{l}\"")
        if phi:
            self.canvas.write(f"{obj} \"phi\", \"{phi}\"")
        if theta:
            self.canvas.write(f"{obj} \"theta\", \"{theta}\"")
        res = self.canvas.send(self.cst_handler, cmt="Set parameters")
        if res:
            print("[ OK ] Parameters set successfully")
        else:
            print("[ERRO] Failed to set parameters, please check whether the parameters exist")
            raise RuntimeError("Failed to set parameters, please check whether the parameters exist")