"""python version of Microsoft Office Activation Assistant
This requires for Office to be installed. Meant to be used when activation through GUI is not enough
Developed on Windows 2012r2 Office 2013 with Python 3.5.1
first release, shows outputs dstatus
TODO: inpkey, unpkey, setkms, and the rest of the functions
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
print("Launching OSPP.vbs status check from " + ruta)
subprocess.check_call(['cscript', 'ospp.vbs', '/dstatus'], cwd=ruta[:-8])

os.system("pause")

