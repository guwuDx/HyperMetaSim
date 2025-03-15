from utils.macors_canva import Canvas


class SquarePillar:
    def __init__(self, csth, component="component1"):
        self.csth = csth
        self.canvas = Canvas()
        self.component = component


    def change_substrate(self, material_name):
        csth = self.csth
        # Solid.ChangeMaterial "component1:substrate", "Material"
        res = self.canvas.write_send(csth, 
                                     f"Solid.ChangeMaterial \"{self.component}:substrate\", \"{material_name}\"",
                                     "Change substrate material"
                                     )
        if res:
            print(f"[INFO] Substrate material changed to {material_name}")
        else:
            print(f"[ERRO] Failed to change substrate material to {material_name}")
            raise RuntimeError("Failed to change substrate material")
        
        self.csth.crr_prj_properties["substrate_material"] = material_name
        # update crr_prj_properties to prjs list (csth.prjs = pd.DataFrame(columns=["project_instance", "project_properties"]))
        csth.prjs.loc[csth.prjs["project_instance"] == csth.crr_prj, "project_properties"] = csth.crr_prj_properties


    def change_pillar(self, material_name):
        csth = self.csth
        # Solid.ChangeMaterial "component1:pillar", "Material"
        res = self.canvas.write_send(csth, 
                                     f"Solid.ChangeMaterial \"{self.component}:pillar\", \"{material_name}\"",
                                     "Change pillar material"
                                     )
        if res:
            print(f"[INFO] Pillar material changed to {material_name}")
        else:
            print(f"[ERRO] Failed to change pillar material to {material_name}")
            raise RuntimeError("Failed to change pillar material")

        self.csth.crr_prj_properties["pillar_material"] = material_name
        # update crr_prj_properties to prjs list (csth.prjs = pd.DataFrame(columns=["project_instance", "project_properties"]))
        csth.prjs.loc[csth.prjs["project_instance"] == csth.crr_prj, "project_properties"] = csth.crr_prj_properties