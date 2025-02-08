from utils.macors_canva import Canvas


class SquarePillar:
    def __init__(self, cst, component="component1"):
        self.cst = cst
        self.canvas = Canvas()
        self.component = component


    def change_substrate(self, material_name):
        # Solid.ChangeMaterial "component1:substrate", "Material"
        res = self.canvas.write_send(self.cst, 
                                     f"Solid.ChangeMaterial \"{self.component}:substrate\", \"{material_name}\"")
        if res:
            print(f"[INFO] Substrate material changed to {material_name}")
        else:
            print(f"[ERRO] Failed to change substrate material to {material_name}")
            raise RuntimeError("Failed to change substrate material")


    def change_pillar(self, material_name):
        # Solid.ChangeMaterial "component1:pillar", "Material"
        res = self.canvas.write_send(self.cst, 
                                     f"Solid.ChangeMaterial \"{self.component}:pillar\", \"{material_name}\"")
        if res:
            print(f"[INFO] Pillar material changed to {material_name}")
        else:
            print(f"[ERRO] Failed to change pillar material to {material_name}")
            raise RuntimeError("Failed to change pillar material")