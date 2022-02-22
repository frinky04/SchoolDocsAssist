from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
from tkcalendar import Calendar
import datetime
from docxtpl import DocxTemplate
import os
import subprocess
import time
import sys
import tkinter
import configparser

#Get Date
now = datetime.datetime.now()
#Define Template
doc = DocxTemplate("StandardTemplate.docx")
#Config
config_obj = configparser.ConfigParser()
config_obj.read("config.ini")
userinfo = config_obj["user_info"]


subjects = []
docsPath = os.path.expanduser(userinfo["directory"])
schoolWorkPath = docsPath + '/' + userinfo["directoryname"]

os.makedirs(schoolWorkPath, exist_ok=True)


#File Browser
FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

#Create Window
window = Tk()
window.title('School Docs Assist')
#Create Calender
cal = Calendar(window, selectmode = 'day',
               year = now.year, month = now.month,
               day = now.day, date_pattern='dd/MM/yyyy')

#Title
titleAssign = ttk.Entry(window)
subjectAssign = ttk.Entry(window)

#Subject Selection
subjectSelection = ttk.Combobox()
subjectSelection['state'] = 'readonly'


def UpdateDirectory():
    global schoolWorkPath
    global docsPath
    docsPath = os.path.expanduser(userinfo["directory"])
    schoolWorkPath = docsPath + '/' + userinfo["directoryname"]
    print(schoolWorkPath)
    os.makedirs(schoolWorkPath, exist_ok=True)


#Create Doc Event
def CreateDocument():
    print(titleAssign.get())
    context = { 'assignname' : titleAssign.get(), 'due' : cal.get_date(), 'name' : userinfo["name"] }
    doc.render(context)
    filename = schoolWorkPath + "/" + subjectSelection.get()+'/'+titleAssign.get().replace(' ','_')+'.docx'
    print(filename)
    doc.save(filename)
    os.system("start " + filename)
    subprocess.run([FILEBROWSER_PATH, '/select,', os.path.normpath(filename)])
    #subprocess.check_call(['open', filename])
    quit()

def OpenDir():
    subprocess.run([FILEBROWSER_PATH, os.path.normpath(schoolWorkPath)])
    

def Refresh():
    subjectSelection["values"] = []
    subjects = []
    for subdir, dirs, files in os.walk(schoolWorkPath):
        for dirs in dirs:
            subjects.append(dirs)
    if len(subjects) > 0:
        subjectSelection.set(subjects[0])
        subjectSelection["values"] = subjects

def CreateSubject():
    if len(subjectAssign.get()) > 0:
        os.makedirs(schoolWorkPath + "/" + subjectAssign.get(), exist_ok=True)
        Refresh()
        subjectSelection.set(subjectAssign.get())


def SetUsersSettings():
    usersettingswindow = Tk()
    usersettingswindow.title('User Settings')
    usText1 = ttk.Label(usersettingswindow, text="User Name:").grid(row=0, column=0)
    usText2 = ttk.Label(usersettingswindow, text="Directory:").grid(row=1, column=0)
    usText3 = ttk.Label(usersettingswindow, text="Directory Name :").grid(row=2, column=0)
    usText2 = ttk.Label(usersettingswindow, text="~ represents your base user directory").grid(row=4, column=1)
    usText2 = ttk.Label(usersettingswindow, text="Directory name must NOT contain spaces!").grid(row=5, column=1)
    test1 = ttk.Entry(usersettingswindow)
    test2 = ttk.Entry(usersettingswindow)
    test3 = ttk.Entry(usersettingswindow)
    test1.insert(0, userinfo["directory"])
    test2.insert(0, userinfo["name"])
    test3.insert(0, userinfo["directoryname"])
    def SaveConfig():
        config_obj.set('user_info', 'name', test2.get())
        config_obj.set('user_info', 'directory', test1.get())
        config_obj.set('user_info', 'directoryname', test3.get())
        with open(r"config.ini", 'w') as configfile:
            config_obj.write(configfile)
            UpdateDirectory()
    test1.grid(row=1, column=2)
    test2.grid(row=0, column=2)
    test3.grid(row=2, column=2)
    usButton1 = ttk.Button(usersettingswindow, text="Save", command=SaveConfig).grid(row=3, column=1)


Refresh()

#Pack And Create Labels
Text = ttk.Label(text="Assignment Creator", font=("Helvetica", 14)).pack()
sep1 = ttk.Separator().pack(expand=True, fill="x", pady=5)
titleText = ttk.Label(text="Name of assignment:").pack()
titleAssign.pack(expand=True)
calendertext = ttk.Label(text="Pick A Date:").pack()
cal.pack(pady=5, padx=5)
subjectselectText = ttk.Label(text="Select Subject:").pack()
subjectSelection.pack()
createButton = ttk.Button(text = "Create Assignment", command=CreateDocument).pack(pady=15)
sep2 = ttk.Separator().pack(expand=True, fill="x", pady=5)
subjectCreateText = ttk.Label(text="Create Subject:").pack(pady=5)
subjectAssign.pack(pady=5)
CreateSubjectButton = ttk.Button(text = "Create Subject", command=CreateSubject).pack(pady=5)
RefreshButton = ttk.Button(text = "Refresh Subjects", command=Refresh).pack(pady=5)
openDir = ttk.Button(text = "Open Directory", command=OpenDir).pack(pady=5)
settingsButton = ttk.Button(text = "Set User Settings", command=SetUsersSettings).pack(pady=5)


#Loop


window.mainloop()   



