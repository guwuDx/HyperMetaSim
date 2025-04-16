from utils.macors_canva import Canvas
from utils import basic_opts

def SquarePillar(csth):
    """SquarePillar class for CST simulation.

    Args:
        csth (CSTHandler): CST handler instance.
    """
    wavelength_min = csth.crr_prj_properties["wavelength_min"]
    wavelength_max = csth.crr_prj_properties["wavelength_max"]

    # Initialize the Canvas object with the CST handler
    canvas = Canvas()

    # Define the initial parameters for the square pillar
    initial_theta = 0
    initial_phi = 0
    initial_h1 = 1
    initial_h = 2 * wavelength_max / 3
    initial_p = 1 * wavelength_max / 2
    initial_l = 2 * wavelength_max / 5
    obj = "StoreParameter"
    canvas.write(f"{obj} \"theta\", \"{initial_theta}\"",   adapt=False)
    canvas.write(f"{obj} \"phi\", \"{initial_phi}\"",       adapt=False)
    canvas.write(f"{obj} \"h1\", \"{initial_h1}\"",         adapt=False)
    canvas.write(f"{obj} \"h\", \"{initial_h}\"",           adapt=False)
    canvas.write(f"{obj} \"p\", \"{initial_p}\"",           adapt=False)
    canvas.write(f"{obj} \"l\", \"{initial_l}\"",           adapt=False)
    canvas.send(csth, add_to_history=False)

    # initialize the project
    vbac = Canvas.vba_template.set_mws_baisc(wavelength_min, wavelength_max)
    canvas.write_send(csth, vbac, "init_mws", add_to_history=True)

    # Create new component
    vbac = Canvas.vba_template.create_component()
    canvas.write_send(csth, vbac, "create_component1", add_to_history=True)

    # Create the substrate based on the component
    vbac = Canvas.vba_template.create_substrate()
    canvas.write_send(csth, vbac, "create_substrate", add_to_history=True)

    # Create the square pillar based on the component
    vbac = Canvas.vba_template.create_pillar("Brick", "l")
    canvas.write_send(csth, vbac, "create_square_pillar", add_to_history=True)

    # Set the simulation solver 
    vbac = Canvas.vba_template.set_solver_basic()
    canvas.write_send(csth, vbac, "set_solver_basic", add_to_history=True)