"""python version of Microsoft Office Activation Assistant
This requires for Office to be installed. Meant to be used when activation through GUI is not enough
Developed on Windows 2012r2 Office 2013 with Python 3.5.1
first release, shows outputs dstatus
TODO: setkms, setport and the rest of the functions
http://github.com/audricd
"""

from tkinter import filedialog
from tkinter import messagebox
import subprocess
import os
import sys

if sys.version_info[0] < 3:
    import Tkinter as Tk
else:
    import tkinter as Tk


def browse_file():
    fname = filedialog.askopenfilename(filetypes=(("OSPP.VBS", "OSPP.VBS"), ("OSPP.VBS", "OSPP.VBS")))
    return fname

root = Tk.Tk()
root.wm_title("Browser")
broButton = Tk.Button(master=root, text='Search for OSPP.vbs', width=15, command=browse_file)
broButton.pack(side=Tk.LEFT, padx=2, pady=2)

messagebox.showinfo("Welcome to Microsoft Office Activation Assistant v0.1",
                    "In the next step, you will have to browse OSPP.VBS, "
                    "built VBScript in the Office suite that handles activation"
                    "\n It is mandatory to be running this from an administrator account")

ruta = browse_file()


def dstatus():
    print("Launching OSPP.vbs status check from " + ruta)
    subprocess.check_call(['cscript', 'ospp.vbs', '/dstatus'], cwd=ruta[:-8])


def act():
    print("Launching OSPP.vbs activation attempt from " + ruta)
    subprocess.check_call(['cscript', 'ospp.vbs', '/act'], cwd=ruta[:-8])

def inpkey():
    keytoinstall = input("Type in your product key:")
    subprocess.check_call(['cscript', 'ospp.vbs', '/inpkey:' + keytoinstall], cwd=ruta[:-8])

def unpkey():
    keytoremove = input("Type in the last segment of the key you want to remove:")
    subprocess.check_call(['cscript', 'ospp.vbs', '/unpkey:' + keytoremove], cwd=ruta[:-8])

print("Python Microsoft Windows Office Activation Assistant Menu:"
      "\n 1. check current activation status"
      "\n 2. launch activation status"
      "\n 3. install key"
      "\n 4. remove key")

menuchoice = input("Type in the menu option:")
if menuchoice == "1":
    dstatus(),
elif menuchoice == "2":
    act(),
elif menuchoice == "3":
    inpkey(),
elif menuchoice == "4":
    unpkey(),
else:
    print("invalid choice")

os.system("pause")

