import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages" : ["os", "tkinter", "tkcalendar", "docx", "datetime", "docxtpl", "os", "subprocess", "time", "sys"], "include_files" : ["StandardTemplate.docx", "config.ini"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"
    pass
setup(  name = "School Docs Assist",
       version = "0.1",
       description = "A tool that assists in the creation of word documents for school projects, using python.",
       options = {"build_exe": build_exe_options},
       executables = [Executable("WordDocumentCreator.py", base=base)])
