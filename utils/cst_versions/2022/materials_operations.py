from utils.macors_canva import Canvas


class SquarePillar:
    def __init__(self, cst_handler, component="component1"):
        self.cst_handler = cst_handler
        self.canvas = Canvas()
        self.component = component


    def change_substrate(self, material_name):
        # Solid.ChangeMaterial "component1:substrate", "Material"
        res = self.canvas.write_send(self.cst_handler, 
                                     f"Solid.ChangeMaterial \"{self.component}:substrate\", \"{material_name}\"",
                                     "Change substrate material"
                                     )
        if res:
            print(f"[INFO] Substrate material changed to {material_name}")
        else:
            print(f"[ERRO] Failed to change substrate material to {material_name}")
            raise RuntimeError("Failed to change substrate material")


    def change_pillar(self, material_name):
        # Solid.ChangeMaterial "component1:pillar", "Material"
        res = self.canvas.write_send(self.cst_handler, 
                                     f"Solid.ChangeMaterial \"{self.component}:pillar\", \"{material_name}\"",
                                     "Change pillar material"
                                     )
        if res:
            print(f"[INFO] Pillar material changed to {material_name}")
        else:
            print(f"[ERRO] Failed to change pillar material to {material_name}")
            raise RuntimeError("Failed to change pillar material")