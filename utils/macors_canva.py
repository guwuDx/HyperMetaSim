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


    def add_code(self, obj_name, key, *values):
        if obj_name not in self.objs:
            self.objs[obj_name] = []

        if values:
            value_str = "\"\", \"\"".join(list(map(str, values))) # convert to string
            code = f"  .{key} \"\"{value_str}\"\""
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
        self.write(code)
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
            preview_instance._write_obj(add_to_history)
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

__all__ = ["Canvas"]