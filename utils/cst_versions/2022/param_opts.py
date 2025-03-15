import numpy as np
from math import ceil, floor
from tqdm import tqdm

from utils.macors_canva import Canvas
from utils.misc import ranges

# from utils import basic_operations



class SquarePillar:

    def __init__(self, csth):
        self.csth = csth
        self.padding = csth._DRC.padding
        self.h_l_ratio_upper_bound = csth._DRC.h_l_ratio_upper_bound
        self.wavelength_min = csth.crr_prj_properties["wavelegnth_min"]
        self.wavelength_max = csth.crr_prj_properties["wavelegnth_max"]
        self.canvas = Canvas()
        self.sweep_list = None
        # self.sweep_element = {
        #     "theta_start": 0, "theta_end": 0, "theta_step": 0,
        #     "phi_start": 0,   "phi_end": 0,   "phi_step": 0,
        #     "p_start": 0,     "p_end": 0,     "p_step": 0,
        #     "h_start": 0,     "h_end": 0,     "h_step": 0,
        #     "l_start": 0,     "l_end": 0,     "l_step": 0,
        # }


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


    def set_sweep_from_list(self,
                            start_now=True
                            ):
        if self.sweep_list is None:
            print("[WARN] No sweep list is generated.")
            return

        print("[INFO] Setting up Parameter Sweep ...")
        obj = "ParameterSweep"

        self.canvas.add_code(obj, "DeleteAllSequences", adapt=False)
        for i, parameter in enumerate(self.sweep_list):
            seq = "seq" + str(i)
            p, h, l = parameter
            self.canvas.add_code(obj, "AddSequence", seq                                , adapt=False)
            self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "p", str(p)  , adapt=False)
            self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "h", str(h)  , adapt=False)
            self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "l", str(l)  , adapt=False)
        self.canvas.add_code(obj, "SetSimulationType", "Frequency", adapt=False)

        if start_now:
            self.csth.save_crr_prj()
            print("[INFO] Setting completed, starting simulation ...")
            self.canvas.add_code(obj, "Start", adapt=False)
            # self.canvas.preview(0)
            res = self.canvas.send(self.csth, add_to_history=False)
            if res:
                print("[INFO] Parameter sweep finished, please check the results.")
            else:
                print("[ERRO] Failed to start the simulation, please check the parameters.")
                raise RuntimeError("Failed to start the simulation, please check the parameters.")
        else:
            print("[INFO] Setting completed, please start the simulation manually.")
            # self.canvas.preview(0)
            res = self.canvas.send(self.csth, cmt="Set up paramSweep", add_to_history=False)
            if res:
                print("[INFO] Parameter sweep set successfully")
            else:
                print("[ERRO] Failed to set up the parameter sweep, please check the parameters.")
                raise RuntimeError("Failed to set up the parameter sweep, please check the parameters.")
            self.csth.save_crr_prj()

        # print("[INFO] vba code to be executed:\n")
        # self.canvas.preview()
        return res


    def set_sweep_from_period(self,
                             h_step:    float  = None,
                             l_step:    float  = None,
                             h_start:   float  = None,
                             h_end:     float  = None,
                             l_start:   float  = None,
                             l_end:     float  = None,
                             start_now: bool = False):

        if not h_step or not l_step:
            print("[ERRO] Step size of the parameters is not specified.")
            raise ValueError("Step size of the parameters is not specified.")

        try:
            p = self.csth.crr_prj_properties["period"]
        except KeyError:
            print("[ERRO] Period is not specified.")
            raise KeyError("Period is not specified.")

        padding = self.padding
        h_l_ratio_upper_bound = self.h_l_ratio_upper_bound
        print(f"[INFO] Generating Parameter Sweep List ...")
        obj = "ParameterSweep"
        self.canvas.add_code(obj, "DeleteAllSequences", adapt=False)

        h_start_drc = False
        h_end_drc = False
        l_start_drc = False
        l_end_drc = False
        if not h_start:
            print("[INFO] Minimum height is not specified, using default DRC configuration.")
            h_start_drc = True
        if not h_end:
            print("[INFO] Maximum height is not specified, using default DRC configuration.")
            h_end_drc = True
        if not l_start:
            print("[INFO] Minimum length is not specified, using default DRC configuration.")
            l_start = l_step
            l_start_drc = True
        if not l_end:
            print("[INFO] Maximum length is not specified, using default DRC configuration.")
            l_end = p - 2 * padding
            l_end_drc = True

        i = 0
        cnt = 0
        lenghts = ranges(l_start, l_end, l_step)

        if h_start_drc or h_end_drc:
            for l in lenghts:
                if h_start_drc:
                    h_start = max(h_step, padding)
                if h_end_drc:
                    h_end = np.around((min(l * h_l_ratio_upper_bound, (p-l) * h_l_ratio_upper_bound)), decimals=3)

                seq = f"seq{cnt}_l{l}"
                self.canvas.add_code(obj, "AddSequence", seq                                        , adapt=False)
                self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "l", str(l)          , adapt=False)
                if h_start == h_end:
                    self.canvas.add_code(obj, "AddParameter_ArbitraryPoints", seq, "h", str(h_start), adapt=False)
                    i += 1
                else:
                    self.canvas.add_code(obj, "AddParameter_Stepwidth", seq, "h", h_start, h_end, h_step, adapt=False)
                    i += ceil((h_end - h_start + 1) / h_step)

                cnt += 1
        else:
            if not (l_start_drc or l_end_drc):
                print("[INFO] ALL parameters are specified.")

            seq = "full_specifed"
            self.canvas.add_code(obj, "AddSequence", seq,                                           adapt=False)
            self.canvas.add_code(obj, "AddParameter_Stepwidth", seq, "h", h_start, h_end, h_step,   adapt=False)
            self.canvas.add_code(obj, "AddParameter_Stepwidth", seq, "l", l_start, l_end, l_step,   adapt=False)
            i += ceil((h_end - h_start + 1) / h_step) * ceil((l_end - l_start + 1) / l_step)

        print(f"[INFO] {i} parameters combinations were generated.")
        self.canvas.add_code(obj, "SetSimulationType", "Frequency", adapt=False)

        if start_now:
            self.csth.save_crr_prj()
            print("[INFO] Setting completed, starting simulation ...")
            self.canvas.add_code(obj, "Start", adapt=False)
            # self.canvas.preview(0)
            res = self.canvas.send(self.csth, add_to_history=False)
            if res:
                print("[INFO] Parameter sweep finished, please check the results.")
            else:
                print("[ERRO] Failed to start the simulation, please check the parameters.")
                raise RuntimeError("Failed to start the simulation, please check the parameters.")
        else:
            print("[INFO] Setting completed, please start the simulation manually.")
            # self.canvas.preview(0)
            res = self.canvas.send(self.csth, cmt="Set up paramSweep", add_to_history=False)
            if res:
                print("[INFO] Parameter sweep set successfully")
            else:
                print("[ERRO] Failed to set up the parameter sweep, please check the parameters.")
                raise RuntimeError("Failed to set up the parameter sweep, please check the parameters.")
            self.csth.save_crr_prj()


    def set_params(self, 
                   p:       float = None, # the Arrangement period of the metastructure
                   h:       float = None, # the height of the pillar
                   l:       float = None, # the length of the pillar
                   phi:     float = None, # the azimuthal angle of the EM wave
                   theta:   float = None  # the incident angle of the EM wave
                   ):
        import basic_opts
        print("[INFO] Setting parameters ...")
        obj = "StoreParameter"

        if h:
            self.canvas.write(f"{obj} \"h\", \"{h}\"", adapt=False)
        if l:
            self.canvas.write(f"{obj} \"l\", \"{l}\"", adapt=False)
        res = self.canvas.send(self.csth, cmt="Set parameters")
        
        if res:
            print("[ OK ] Parameters set successfully")
            # basic_operations.update_params(self.csth)
        else:
            print("[ERRO] Failed to set parameters, please check whether the parameters exist")
            raise RuntimeError("Failed to set parameters, please check whether the parameters exist")

        if p or theta or phi:
            basic_opts.set_basic_params(self.csth, p, theta, phi)


    def simulate_param_combination(self, 
                   p:       float  = None,    # the Arrangement period of the metastructure
                   h:       float  = None,    # the height of the pillar
                   l:       float  = None,    # the length of the pillar
                   phi:     float  = None,    # the azimuthal angle of the EM wave
                   theta:   float  = None,    # the incident angle of the EM wave
                   blocked: bool   = True,    # whether to block the simulation
                   timeout: int    = None     # the timeout of the simulation
                   ):
        print("[INFO] Simulating parameters ...")
        self.set_params(p, h, l, phi, theta)
        self.csth.run_solver(blocked=blocked, timeout=timeout)


    def calculate_combination_num(self,
                                  p_start: float = None,
                                  p_end:   float = None,
                                  p_step:  float = None,
                                  h_start: float = None,
                                  h_end:   float = None,
                                  h_step:  float = None,
                                  l_start: float = None,
                                  l_end:   float = None,
                                  l_step:  float = None):
        
        if not p_step or not h_step or not l_step:
            print("[WARN] Step size of the parameters is not specified.")
            return
        
        padding = self.padding
        h_l_ratio_upper_bound = self.h_l_ratio_upper_bound

        if p_start is None or p_end is None:
            print("[WARN] Arrangement period range is not specified, all periods will be considered.")
            print("[WARN] This method is not recommended for large/long wavelength range.")
            
            wavelength_min = self.wavelength_min
            wavelength_max = self.wavelength_max
            p_start = floor((wavelength_min / 4) / p_step) * p_step
            p_end = ceil((wavelength_max / 2) / p_step) * p_step

        i = 0
        periods = ranges(p_start, p_end, p_step)
        for p in periods:
            if not l_start:
                l_start = l_step
            if not l_end:
                l_end = p - 2 * padding
            lenghts = ranges(l_start, l_end, l_step)

            for l in lenghts:
                if not h_start:
                    h_start = max(h_step, padding)
                if not h_end:
                    h_end = np.around((min(l * h_l_ratio_upper_bound, (p-l) * h_l_ratio_upper_bound)), decimals=3)

                i += ceil((h_end - h_start + 1) / h_step)

        print(f"[INFO] {i} parameters combinations will be simulated.")
        return i


    def py_sweep_from_period(self,
                            p:       float = None,
                            h_start: float = None,
                            h_end:   float = None,
                            h_step:  float = None,
                            l_start: float = None,
                            l_end:   float = None,
                            l_step:  float = None,
                            timeout_once: int = None):
        if not p or not h_step or not l_step:
            print("[WARN] Step size of the parameters is not specified.")
            return
        
        padding = self.padding
        h_l_ratio_upper_bound = self.h_l_ratio_upper_bound
        print(f"[INFO] Generating Parameter Sweep List ...")
        total = self.calculate_combination_num(p, h_start, h_end, h_step, l_start, l_end, l_step)

        if not l_start:
            l_start = l_step
        if not l_end:
            l_end = p - 2 * padding
        lenghts = ranges(l_start, l_end, l_step)

        for l in lenghts:
            if not h_start:
                h_start = max(h_step, padding)
            if not h_end:
                h_end = np.around((min(l * h_l_ratio_upper_bound, (p-l) * h_l_ratio_upper_bound)), decimals=3)

            h_range = ranges(h_start, h_end, h_step)
            for h in h_range:
                self.simulate_param_combination(p, h, l, 0, 0, True, timeout_once)

        print(f"[INFO] All {total} parameters combinations have been simulated.")


    def set_period_parallel_sweep(self, 
                                p_start=0.0, p_end=0.0, p_step=0.0,
                                h_start=0.0, h_end=0.0, h_step=0.0,
                                l_start=0.0, l_end=0.0, l_step=0.0, 
                                start_now=False):
        from . import basic_opts

        csth = self.csth
        # original_prj = copy.copy(csth.crr_prj)
        p_list = ranges(p_start, p_end, p_step)
        sweep_prjs_list = []

        for p in p_list:
            # duplicate project
            print(f"[INFO] Duplicating project for period {p}")
            new_prj = csth.duplicate_prj(f"SquarePillar__surface_inst_period{p}")

            # set basic sweep
            original_prj = csth.crr_prj
            csth.crr_prj = new_prj
            csth.crr_prj.activate()
            print(f"[INFO] Setting primary sweep for period {p}")
            basic_opts.set_basic_params(csth, p=p)
            self.set_sweep_from_period(h_start=h_start,  l_start=l_start,
                                       h_step=h_step,    l_step=l_step,
                                       h_end=h_end,      l_end=l_end)
            sweep_prjs_list.append(csth.crr_prj)

            # start simulation
            if start_now:
                print(f"[INFO] Starting simulation for period {p}")
                csth.save_crr_prj()
                self.canvas.add_code("ParameterSweep", "Start", adapt=False)
                res = self.canvas.send(csth, add_to_history=False)
                if res:
                    print(f"[ OK ] Simulation started for period {p}")
                    csth.save_crr_prj()
                    csth.close_prj()
                else:
                    print(f"[ERRO] Failed to start simulation for period {p}")
                    raise RuntimeError(f"Failed to start simulation for period {p}")

            print(f"[ OK ] Done for period {p}")
            csth.crr_prj = original_prj
            csth.crr_prj.activate()

        return sweep_prjs_list
