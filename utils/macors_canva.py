import copy


class Canvas:
    def __init__(self):
        self.vbac = list()
        self.objs = {}
        self.clear()


    def write(self, code):
        self.vbac.append(code)


    def add_code(self, obj_name, key, *values):
        if obj_name not in self.objs:
            self.objs[obj_name] = []

        if values:
            value_str = "\", \"".join(list(map(str, values))) # convert to string
            code = f".{key} \"{value_str}\""
        else:
            code = f".{key}"

        self.objs[obj_name].append(code)


    def del_obj(self, obj_name):
        if obj_name in self.objs:
            del self.objs[obj_name]
        else:
            print(f"[WARN] object {obj_name} does not exist")


    def _clr_obj(self):
        self.objs = {}


    def _write_obj(self, obj_name=None):
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
        self.vbac.append("Sub Main()")


    def send(self, cst_app, timeout=None, clear=True):
        """
        Send the VBA code to the CST application, 
        default to the currently active CST project.

        Args:
            cst_app (object): CST application handler object
        """
        self._write_obj()
        self.vbac.append("End Sub\n")
        vba_code = "\n".join(self.vbac)
        res = cst_app.send_vba(vba_code, timeout)
        if clear: self.clear()
        return res


    def write_send(self, cst_app, code, timeout=None):
        self.write(code)
        return self.send(cst_app, timeout, clear=True)


    def preview(self):
        preview_instance = copy.deepcopy(self)
        preview_instance._write_obj()
        preview_instance.vbac.append("End Sub\n")
        vba_code = "\n".join(preview_instance.vbac)
        print(vba_code)
        return vba_code


    def write_to_file(self, file_path):
        self.vbac.append("End Sub")
        vba_code = "\n".join(self.vbac)
        with open(file_path, 'w') as f:
            f.write(vba_code)


__all__ = ["Canvas"]