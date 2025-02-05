from utils import cst_gen


class Canvas:
    def __init__(self):
        self.vbac = list()


    def clear(self):
        self.vbac = list()
        self.vbac.append("Sub Main()\n")


    def write(self, code):
        self.vbac.append(code)


    def send(self, cst_app, timeout=None):
        """
        Send the VBA code to the CST application, 
        default to the currently active CST project.

        Args:
            cst_app (object): CST application handler object
        """
        self.vbac.append("End Sub\n")
        vba_code = "\n".join(self.vbac)
        return cst_app.send_vba(vba_code, timeout)


    def write_send(self, code):
        self.write(code)
        self.send()


    def write_to_file(self, file_path):
        self.vbac.append("End Sub\n")
        vba_code = "\n".join(self.vbac)
        with open(file_path, 'w') as f:
            f.write(vba_code)


    # test the class
    @classmethod
    def GetApplicationName(cls, cst_app):
        cls.clear()
        cls.write("GetApplicationName")
        app = cls.send(cst_app)
        cls.clear()
        print(app)