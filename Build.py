import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages" : ["os", "tkinter", "ttkthemes", "tkcalendar", "docx", "datetime", "docxtpl", "os", "subprocess", "time", "sys"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"
    pass
setup(  name = "Assigment Creator",
       version = "0.1",
       description = "Create Assignments Really Quick",
       options = {"build_exe": build_exe_options},
       executables = [Executable("WordDocumentCreator.py", base=base)])
