# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 15:47:07 2018

©Copyright 2020 University of Florida Research Foundation, Inc. All Rights Reserved.
    William Watson and Mark Orazem

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

#---Import necessary libraries/modules---
import ctypes
import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import inputFrame
import helpFrame
import modelFrame
import errorFileFrame
import errorFrame
import settingsFrame
import formulaFrame
import configparser
import os
import re
import multiprocessing
from PIL import ImageTk, Image
import webbrowser

class CreateToolTip(object):
    """
        Create a tooltip for a given widget
        Code from: https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter with answer by crxguy52
    """
    def __init__(self, widget, text='widget info', color=None, resetColor=None):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        
        if not color is None and not resetColor is None:
            self.widget.bind("<Enter>", lambda e: self.enterColor(e, color))
            self.widget.bind("<Leave>", lambda e: self.leaveColor(e, resetColor))
        else:
            self.widget.bind("<Enter>", self.enter)
            self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None
    
    def enterColor(self, e, c):
        self.widget.configure(fg=c)
        self.schedule()
    
    def leaveColor(self, e, c):
        self.widget.configure(fg=c)
        self.unschedule()
        self.hidetip()
    
    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left', background="#ffffff", relief='solid', borderwidth=1, wraplength = self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()

class myGUI:
    
    def __init__(self, master, args):
        
        def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
        
            return os.path.join(base_path, relative_path)
        
        """" Windows taskbar code from: https://www.neowin.net/forum/topic/716968-using-apis-in-a-objectaction-format/#comment-590434472,
            https://stackoverflow.com/questions/1736394/using-windows-7-taskbar-features-in-pyqt/1744503#1744503,
            and https://github.com/tqdm/tqdm/issues/442 """
        
        self.master = master
        #---Appearance defaults---
        self.theme = "light"    #Default theme
        self.backgroundColor = '#333333'    #Default bar color (gray20)
        self.highlightColor = '#999999'     #Default color on mouseover (lighter gray)
        self.activeColor = '#737373'        #Default color when tab is active (light gray)
                                            #6B6B6B - gray for text entries with dark theme

        #---Application defaults for other classes; used in getters/setters at the end of this file---
        self.detectCommentsDefault = 1      #Whether or not to detect comments
        self.detectDelimiterDefault = 1     #Whether or not to detect delimiters
        self.commentsDefault = 0            #The default number of comment lines
        self.delimiterDefault = "Tab"       #The default delimiter; options are "Space", "Tab", ";", "|", ":", and ","
        self.MCDefault = 1000               #The default number of Monte Carlo simulations ; a positive integer
        self.fitDefault = "Complex"         #The default fit type; options are "Complex", "Real", and "Imaginary"
        self.weightDefault = "Modulus"      #The default weighting type; options are "None", "Modulus", "Proportional", and "Error Model"
        self.alphaDefault = 1               #Default noise
        self.errorAlphaDefault = 0          #Default error structure choices
        self.errorBetaDefault = 0
        self.errorReDefault = 0
        self.errorGammaDefault = 1
        self.errorDeltaDefault = 1
        self.errorWeightingDefault = "Variance"     #Default weighting for error structure fitting; options are "Variance" and "None"
        self.errorMADefault = "5 point"             #Default moving average for variance weighting; options are "5 point", "3 point", and "None"
        self.detrendDefault = "Off"                 #Default detrending; options are "Off" and "On"
        self.ellipseColor = "#FF0000"               #Default ellipse color
        self.currentDirectory = "C:\\"              #The current directory; used in various other classes, but not a default setting
        self.defaultDirectory = "C:\\"              #The default directory
        self.defaultFormulaDirectory = "C:\\"       #The default directory for .mmformula files
        self.inputExitAlert = 0                     #Whether or not to alert before closing the input tab
        self.customFormulaExitAlert = 0             #Whether or not to alert before closing the custom fitting tab
        self.defaultTab = 0                         #The default tab when the program is opened
        self.freqLoad = 0                           #Whether or not frequency range stays the same when a new file is loaded (mm tab)
        self.freqUndo = 0                           #Whether or not frequency range is undone with undo button
        self.freqLoadCustom = 0                     #Whether or not frequency rnge stays the same when a new file is loaded (custom tab)
        self.defaultScroll = 1                      #Whether or not to use the mousewheel to scroll through tabs when mouse is over the navigation pane
        self.defaultImports = []                    #Default import paths for custom formula        
        
        self.initializeSettings()
        
        self.residualsFileList = [] #In case residuals files are being opened
        self.errorsFileList = []    #In case error files are being opened
        self.argFile = ""
        
        self.interpretArgs(args)
        
        master.title("Measurement Model")                       #The program title
        master.geometry('{}x{}'.format(840, 830))               #GUI size
        master.iconbitmap(resource_path('img/elephant3.ico'))   #GUI icon
        
        #---Change the background color depeding on the theme---
        if (self.theme == "dark"):
            master.configure(background='#424242')
        elif (self.theme == "light"):
            master.configure(background='white')
        master.grid_rowconfigure(1, weight=1)       #The second row will expand to fill the space
        master.grid_columnconfigure(2, weight=1)    #The second column will expand to fill the space
        master.bind_all("<F1>", self.helpDown)      #If <F1> is pressed anywhere, it goes to the help tab
        self.currentTab = self.defaultTab           #Keep track of where the program is
        
        #SpringGreen2 - #00ee76
        #SpringGreen3 - #00cd66
        #gray20 - #333333
        
        barColorNoHash = self.backgroundColor.lstrip('#')   #Remove the '#' from the hex value
        barColorRGB = tuple(int(barColorNoHash[i:i+2], 16) for i in (0, 2 ,4))  #Contrast checking formula from https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color
        self.frameLabelColor = "#000000" if (barColorRGB[0]*0.299 + barColorRGB[1]*0.587 + barColorRGB[2]*0.114) > 186 else "#FFFFFF"    #If the background is too light, the title text is black; otherwise it's white
        
        topLabel = "File Tools and Conversion"              #If the default tab is 0; can be changed in the settings tab
        if (self.defaultTab == 1):
            topLabel = "Measurement Model"
        elif (self.defaultTab == 2):
            topLabel = "Error File Preparation"
        elif (self.defaultTab == 3):
            topLabel = "Error Analysis"
        elif (self.defaultTab == 4):
            topLabel = "Custom Formula Fitting"
        elif (self.defaultTab == 5):
            topLabel = "Settings"
        elif (self.defaultTab == 6):
            topLabel = "Help and About"
        self.top_frame = tk.Frame(root, bg=self.backgroundColor, height=50)     #The frame for the tab title
        self.top_frame.grid_columnconfigure(0, weight=1)
        self.frame_label = tk.Label(self.top_frame, text=topLabel, font=("Arial", 15), bg=self.backgroundColor, fg=self.frameLabelColor) #The title that tells the user which tab they are in
        self.frame_label.grid(column=0, row=0, pady=2)
        self.top_frame.grid(row=0, column=0, columnspan=3, sticky="EW")
        
        self.left_frame = tk.Frame(root, bg=self.backgroundColor, width=140, height=190)    #The frame for the navigation icons
        self.left_frame.columnconfigure(1, weight=1)
        self.left_frame.rowconfigure(4, weight=1)        #Makes the help label go to the bottom of the window
        self.left_frame.grid(row=1, column = 1, sticky="nsew")
        
        #---The icons used in the navigation pane---
        self.imgFile = ImageTk.PhotoImage(Image.open(resource_path("img/filedarkgray.png")))
        self.imgFile2 = ImageTk.PhotoImage(Image.open(resource_path("img/filelightgray.png")))
        self.imgFile3 = ImageTk.PhotoImage(Image.open(resource_path("img/file2blue.png")))
        self.imgFile4 = ImageTk.PhotoImage(Image.open(resource_path("img/file1blue.png")))
        
        self.imgModel = ImageTk.PhotoImage(Image.open(resource_path("img/fittingdarkgray.png")))
        self.imgModel2 = ImageTk.PhotoImage(Image.open(resource_path("img/fittinglightgray.png")))
        self.imgModel3 = ImageTk.PhotoImage(Image.open(resource_path("img/fitting2blue.png")))
        self.imgModel4 = ImageTk.PhotoImage(Image.open(resource_path("img/fitting1blue.png")))
        
        self.imgErrorFile = ImageTk.PhotoImage(Image.open(resource_path("img/errorFiledarkgray.png")))
        self.imgErrorFile2 = ImageTk.PhotoImage(Image.open(resource_path("img/errorFilelightgray.png")))
        self.imgErrorFile3 = ImageTk.PhotoImage(Image.open(resource_path("img/errorFile2blue.png")))
        self.imgErrorFile4 = ImageTk.PhotoImage(Image.open(resource_path("img/errorFile1blue.png")))

        self.imgError = ImageTk.PhotoImage(Image.open(resource_path("img/errordarkgray.png")))
        self.imgError2 = ImageTk.PhotoImage(Image.open(resource_path("img/errorlightgray.png")))
        self.imgError3 = ImageTk.PhotoImage(Image.open(resource_path("img/error2blue.png")))
        self.imgError4 = ImageTk.PhotoImage(Image.open(resource_path("img/error1blue.png")))
        
        self.imgSettings = ImageTk.PhotoImage(Image.open(resource_path("img/settingsdarkgray.png")))
        self.imgSettings2 = ImageTk.PhotoImage(Image.open(resource_path("img/settingslightgray.png")))
        self.imgSettings3 = ImageTk.PhotoImage(Image.open(resource_path("img/settings2blue.png")))
        self.imgSettings4 = ImageTk.PhotoImage(Image.open(resource_path("img/settings1blue.png")))

        self.imgHelp = ImageTk.PhotoImage(Image.open(resource_path("img/helpdarkgray.png")))
        self.imgHelp2 = ImageTk.PhotoImage(Image.open(resource_path("img/helplightgray.png")))
        self.imgHelp3 = ImageTk.PhotoImage(Image.open(resource_path("img/help2blue.png")))
        self.imgHelp4 = ImageTk.PhotoImage(Image.open(resource_path("img/help1blue.png")))
        
        self.imgFormula = ImageTk.PhotoImage(Image.open(resource_path("img/customdarkgray.png")))
        self.imgFormula2 = ImageTk.PhotoImage(Image.open(resource_path("img/customlightgray.png")))
        self.imgFormula3 = ImageTk.PhotoImage(Image.open(resource_path("img/custom2blue.png")))
        self.imgFormula4 = ImageTk.PhotoImage(Image.open(resource_path("img/custom1blue.png")))

        #---Depending on the default tab, set the current highlighted icon and setup what should happen on mouseover/click---
        if (self.defaultTab == 0):
            self.fileOpen = tk.Label(self.left_frame, image = self.imgFile3, cursor="hand2", width=100, height=75)
        else:
            self.fileOpen = tk.Label(self.left_frame, image = self.imgFile, cursor="hand2", width=100, height=75)
        self.fileOpen.configure(background=self.backgroundColor)
        self.fileOpen.bind('<Button-1>', self.inputFileDown)
        self.fileOpen.bind('<ButtonRelease-1>', self.inputFileUp)
        self.fileOpen.bind('<Enter>', lambda e: self.fileOpen.configure(image=self.imgFile2 if self.currentTab != 0 else self.imgFile4))
        self.fileOpen.bind('<Leave>', self.leaveInputFile)
        
        if (self.defaultTab == 1):
            self.runModel = tk.Label(self.left_frame, image = self.imgModel3, cursor="hand2", width=100, height=75)
        else:
            self.runModel = tk.Label(self.left_frame, image = self.imgModel, cursor="hand2", width=100, height=75)
        self.runModel.configure(background=self.backgroundColor)
        self.runModel.bind('<Button-1>', self.measureModelDown)
        self.runModel.bind('<ButtonRelease-1>', self.measureModelUp)
        self.runModel.bind('<Enter>', lambda e: self.runModel.configure(image=self.imgModel2 if self.currentTab != 1 else self.imgModel4))
        self.runModel.bind('<Leave>', self.leaveMeasureModel)
        
        if (self.defaultTab == 2):
            self.errorFileLabel = tk.Label(self.left_frame, image=self.imgErrorFile3, cursor="hand2", width=100, height=75)
        else:
            self.errorFileLabel = tk.Label(self.left_frame, image=self.imgErrorFile, cursor="hand2", width=100, height=75)
        self.errorFileLabel.configure(background=self.backgroundColor)
        self.errorFileLabel.bind('<Button-1>', self.errorFileDown)
        self.errorFileLabel.bind('<ButtonRelease-1>', self.errorFileUp)
        self.errorFileLabel.bind('<Enter>', lambda e: self.errorFileLabel.configure(image=self.imgErrorFile2 if self.currentTab != 2 else self.imgErrorFile4))
        self.errorFileLabel.bind('<Leave>', self.leaveErrorFile)
        
        if (self.defaultTab == 3):
            self.errorLabel = tk.Label(self.left_frame, image=self.imgError3, cursor="hand2", width=100, height=75)
        else:
            self.errorLabel = tk.Label(self.left_frame, image=self.imgError, cursor="hand2", width=100, height=75)
        self.errorLabel.configure(background=self.backgroundColor)
        self.errorLabel.bind('<Button-1>', self.errorDown)
        self.errorLabel.bind('<ButtonRelease-1>', self.errorUp)
        self.errorLabel.bind('<Enter>', lambda e: self.errorLabel.configure(image=self.imgError2 if self.currentTab != 3 else self.imgError4))
        self.errorLabel.bind('<Leave>', self.leaveError)
        
        if (self.defaultTab == 4):
            self.formulaLabel = tk.Label(self.left_frame, image = self.imgFormula3, cursor="hand2", width=100, height=75)
        else:
            self.formulaLabel = tk.Label(self.left_frame, image = self.imgFormula, cursor="hand2", width=100, height=75)
        self.formulaLabel.configure(background=self.backgroundColor)
        self.formulaLabel.bind('<Button-1>', self.formulaDown)
        self.formulaLabel.bind('<ButtonRelease-1>', self.formulaUp)
        self.formulaLabel.bind('<Enter>', lambda e: self.formulaLabel.configure(image=self.imgFormula2 if self.currentTab != 6 else self.imgFormula4))
        self.formulaLabel.bind('<Leave>', self.leaveFormula)
        
        if (self.defaultTab == 5):
            self.settingsLabel = tk.Label(self.left_frame, image=self.imgSettings3, cursor="hand2", width=100, height=75)
        else:
            self.settingsLabel = tk.Label(self.left_frame, image=self.imgSettings, cursor="hand2", width=100, height=75)
        self.settingsLabel.configure(back=self.backgroundColor)
        self.settingsLabel.bind('<Button-1>', self.settingsDown)
        self.settingsLabel.bind('<ButtonRelease-1>', self.settingsUp)
        self.settingsLabel.bind('<Enter>', lambda e: self.settingsLabel.configure(image=self.imgSettings2 if self.currentTab != 4 else self.imgSettings4))
        self.settingsLabel.bind('<Leave>', self.leaveSettings)#lambda e: self.settingsLabel.configure(image=self.imgSettingsWhite))
        
        if (self.defaultTab == 6):
            self.helpLabel = tk.Label(self.left_frame, image=self.imgHelp3, cursor="hand2", width=100, height=75)
        else:
            self.helpLabel = tk.Label(self.left_frame, image=self.imgHelp, cursor="hand2", width=100, height=75)
        self.helpLabel.configure(background=self.backgroundColor)
        self.helpLabel.bind('<Button-1>', self.helpDown)
        self.helpLabel.bind('<ButtonRelease-1>', self.helpUp)
        self.helpLabel.bind('<Enter>', lambda e: self.helpLabel.configure(image=self.imgHelp2 if self.currentTab != 5 else self.imgHelp4))
        self.helpLabel.bind('<Leave>', self.leaveHelp)#lambda e: self.helpLabel.configure(image=self.imgHelpWhite))
        
        #---Display the icons---
        self.fileOpen.grid(column=0, row=0)
        self.runModel.grid(column=0, row=1)
        self.errorFileLabel.grid(column=0, row=2)
        self.errorLabel.grid(column=0, row=3)
        self.formulaLabel.grid(column=0, row=4, sticky="N")
        self.settingsLabel.grid(column=0, row=5, sticky="S")
        self.helpLabel.grid(column=0, row=6, sticky="S")
        
        if (self.defaultScroll == 1):
            self.left_frame.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, -1))
            self.fileOpen.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 0))
            self.runModel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 1))
            self.errorFileLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 2))
            self.errorLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 3))
            self.formulaLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 6))
            self.settingsLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 4))
            self.helpLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 5))
        
        self.input_frame = inputFrame.iF(root, self)
        self.help_frame = helpFrame.hF(root, self)
        self.model_frame = modelFrame.mmF(root, self)
        self.error_file_frame = errorFileFrame.eFF(root, self)
        self.error_frame = errorFrame.eF(root, self)
        self.settings_frame = settingsFrame.sF(root, self)
        self.formula_frame = formulaFrame.fF(root, self)
        
        if (self.defaultTab == 0):
            self.input_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
            self.input_frame.bindIt()
        elif (self.defaultTab == 1):
            self.model_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        elif (self.defaultTab == 2):
            self.error_file_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
            self.error_file_frame.bindIt()
        elif (self.defaultTab == 3):
            self.error_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        elif (self.defaultTab == 4):
            self.formula_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        elif (self.defaultTab == 5):
            self.settings_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
            self.settings_frame.bindIt()
        elif (self.defaultTab == 6):
            self.help_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        
        #---Set up the canvases needed to display the small colored bar next to the current tab's icon---
        self.master.update()
        self.master.update_idletasks()
        self.canvasFrame = tk.Frame(root, bg=self.backgroundColor)
        self.canvasFrame.columnconfigure(0, weight=1)
        self.canvasFrame.rowconfigure(4, weight=1)
        self.canvasWidth = 5
        self.inputCanvas = tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.fileOpen.winfo_height())
        self.runCanvas = tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.runModel.winfo_height())
        self.errorFileCanvas = tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.errorFileLabel.winfo_height())
        self.errorCanvas = tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.errorLabel.winfo_height())
        self.formulaCanvas= tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.errorFileLabel.winfo_height())
        self.settingsCanvas = tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.settingsLabel.winfo_height())
        self.helpCanvas = tk.Canvas(self.canvasFrame, background=self.backgroundColor, bd=0, highlightthickness=0, relief='ridge', width=self.canvasWidth, height=self.helpLabel.winfo_height())
        
        if (self.defaultTab == 0):
            self.inputCanvas.configure(background="#0066FF")
        elif (self.defaultTab == 1):
            self.runCanvas.configure(background="#0066FF")
        elif (self.defaultTab == 2):
            self.errorFileCanvas.configure(background="#0066FF")
        elif (self.defaultTab == 3):
            self.errorCanvas.configure(background="#0066FF")
        elif (self.defaultTab == 4):
            self.formulaCanvas.configure(background="#0066FF")
        elif (self.defaultTab == 5):
            self.settingsCanvas.configure(background="#0066FF")
        elif (self.defaultTab == 6):
            self.helpCanvas.configure(background="#0066FF")
        
        self.canvasFrame.grid(column=0, row=1, sticky="NS")
        self.inputCanvas.grid(column=0, row=0)
        self.runCanvas.grid(column=0, row=1)
        self.errorFileCanvas.grid(column=0, row=2)
        self.errorCanvas.grid(column=0, row=3)
        self.formulaCanvas.grid(column=0, row=4, sticky="N")
        self.settingsCanvas.grid(column=0, row=5, sticky="S")
        self.helpCanvas.grid(column=0, row=6, sticky="S")
        
        #---Check the input file arguments and pass it to the appropriate tab---
        if (self.argFile != ""):
            fname, fext = os.path.splitext(self.argFile)
            if (fext == ".mmfile"):
                self.enterMeasureModel(self.argFile)
            elif (fext == ".mmfitting"):
                self.enterMeasureModelFitting(self.argFile)
            elif (fext == ".mmresiduals"):
                self.enterResidual(self.argFile)
            elif (fext == ".mmerrors"):
                self.enterError(self.argFile)
            elif (fext == ".mmcustom"):
                self.enterCustom(self.argFile)
            elif (fext == ".mmformula"):
                self.enterFormula(self.argFile)
            else:
                self.enterInput(self.argFile)
        elif (len(self.residualsFileList) > 0):
            self.enterResiduals(self.residualsFileList)
        elif (len(self.errorsFileList) > 0):
            self.enterErrors(self.errorsFileList)
    
    #---Initialize settings---
    def initializeSettings(self):
        if (os.path.exists(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini")):   #Check if the settings file exists under LOCALAPPDATA
            try:
                config = configparser.ConfigParser()                                        #Begin the config parser to read settings.ini
                config.read(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini")    #Read the file
                
                #---Check the theme---
                if (config['settings']['theme'] == 'light'):
                    self.theme = "light"
                elif (config['settings']['theme'] == 'dark'):
                    self.theme = "dark"
                else:
                    raise Exception     #If not light or dark, that's bad
                
                #---Check the color/appearance---
                if (re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', config['settings']['bar'])):   #Check if the default color is a hex value
                    self.backgroundColor = config['settings']['bar']
                else:
                    raise Exception
                
                if (re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', config['settings']['highlight'])):   #Check if the default color is a hex value
                    self.highlightColor = config['settings']['highlight']
                else:
                    raise Exception
                    
                if (re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', config['settings']['accent'])):   #Check if the default color is a hex value
                    self.activeColor = config['settings']['accent']
                else:
                    raise Exception
                
                #---Default directories---
                self.defaultDirectory = config['settings']['dir']
                self.currentDirectory = self.defaultDirectory
                
                if (config['settings']['tab'] == "Input file"):
                    self.defaultTab = 0
                elif (config['settings']['tab'] == "Measurement model"):
                    self.defaultTab = 1
                elif (config['settings']['tab'] == "Error file preparation"):
                    self.defaultTab = 2
                elif (config['settings']['tab'] == "Error analysis"):
                    self.defaultTab = 3
                elif (config['settings']['tab'] == "Custom fitting"):
                    self.defaultTab = 4
                elif (config['settings']['tab'] == "Settings"):
                    self.defaultTab = 5
                elif (config['settings']['tab'] == "Help and about"):
                    self.defaultTab = 6
                else:
                    self.defaultTab = 0
                    raise Exception
                
                #---Check if changing tab with scrolling is allowed---
                self.defaultScroll = int(config['settings']['scroll'])
                if (self.defaultScroll != 0 and self.defaultScroll != 1):
                    self.defaultScroll = 1
                    raise Exception
                
                self.defaultFormulaDirectory = config['settings']['formulaDir']
                
                #---Check if number of comment lines should be detected---
                self.detectCommentsDefault = int(config['input']['detect'])
                if (self.detectCommentsDefault != 0 and self.detectCommentsDefault != 1):       #Value must be true or false
                    self.detectCommentsDefault = 1
                    raise Exception
                
                #---Default number of comment lines---
                self.commentsDefault = int(config['input']['comments'])
                if (self.commentsDefault < 0):      #If the number of comments is negative, that's bad
                    self.commentsDefault = 1
                    raise Exception
                
                #---Default delimiter---
                if (config['input']['delimiter'] == "Tab" or config['input']['delimiter'] == "Space" or config['input']['delimiter'] == ";" or config['input']['delimiter'] == "," or config['input']['delimiter'] == "|" or config['input']['delimiter'] == ":"):
                    self.delimiterDefault = config['input']['delimiter']
                else:
                    raise Exception
                
                #---Whether to ask on exit from the input tab---
                iEA = int(config['input']['askInputExit'])
                if (iEA != 0 and iEA != 1):     #Value must be true or false
                    self.inputExitAlert = 0
                    raise Exception
                self.inputExitAlert = iEA
                
                #---Number of Monte Carlo simulations---
                self.MCDefault = int(config['model']['mc'])
                if (self.MCDefault < 1):        #If the number of Monte Carlo simulations is negative, that's bad
                    self.MCDefault = 1000
                    raise Exception
                
                #---Default fit type
                if (config['model']['fit'] == "Complex" or config['model']['fit'] == "Real" or config['model']['fit'] == "Imaginary"):
                    self.fitDefault = config['model']['fit']
                else:
                    raise Exception
                
                #---Default weighting---
                if (config['model']['weight'] == "None" or config['model']['weight'] == "Modulus" or config['model']['weight'] == "Proportional" or config['model']['weight'] == "Error model"):
                    self.weightDefault = config['model']['weight']
                else:
                    raise Exception
                
                #---Default ellipse color---
                if (re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', config['model']['ellipse'])):   #Check if the default color is a hex value
                    self.ellipseColor = config['model']['ellipse']
                else:
                    raise Exception
                
                #---Whether frequency range should be kept on loading a new file
                self.freqLoad = int(config['model']['freqLoad'])
                if (self.freqLoad != 0 and self.freqLoad != 1):
                    self.freqLoad = 0
                    raise Exception
                
                #---Whether "undo" should undo a change in frequency range
                self.freqUndo = int(config['model']['freqUndo'])
                if (self.freqUndo != 0 and self.freqUndo != 1):
                    self.freqUndo = 0
                    raise Exception
                
                #---Whether detrending should be used
                self.detrendDefault = config['error']['detrend']
                if (self.detrendDefault != "Off" and self.detrendDefault != "On"):
                    self.detrendDefault = "Off"
                    raise Exception
                
                #---Default error structure options---
                self.errorAlphaDefault = int(config['error']['alphaError'])
                if (self.errorAlphaDefault != 0 and self.errorAlphaDefault != 1):
                    self.errorAlphaDefault = 0
                    raise Exception
                self.errorBetaDefault = int(config['error']['betaError'])
                if (self.errorBetaDefault != 0 and self.errorBetaDefault != 1):
                    self.errorBetaDefault = 0
                    raise Exception
                self.errorReDefault = int(config['error']['reError'])
                if (self.errorReDefault != 0 and self.errorREDefault != 1):
                    self.errorReDefault = 0
                    raise Exception
                self.errorGammaDefault = int(config['error']['gammaError'])
                if (self.errorGammaDefault != 0 and self.errorGammaDefault != 1):
                    self.errorGammaDefault = 1
                    raise Exception
                self.errorDeltaDefault = int(config['error']['deltaError'])
                if (self.errorDeltaDefault != 0 and self.errorDeltaDefault != 1):
                    self.errorDeltaDefault = 1
                    raise Exception
                
                #---The default error weighting---
                if (config['error']['errorWeighting'] == "None" or config['error']['errorWeighting'] == "Variance"):
                    self.errorWeightingDefault = config['error']['errorWeighting']
                else:
                    raise Exception
                
                #---The default moving average for error---
                if (config['error']['errorMA'] == "None" or config['error']['errorMA'] == "3 point" or config['error']['errorMA'] == "5 point"):
                    self.errorMADefault = config['error']['errorMA']
                else:
                    raise Exception
                
                #---Default noise---
                self.alphaDefault = float(config['model']['alpha'])
                if (self.alphaDefault < 0):     #If the default assumed noise is negative, that's bad
                    self.alphaDefault = 1
                    raise Exception
                
                #---Whether to ask on exiting from the custom formula tab if the formula is unsaved---
                aCE = int(config['custom']['askCustomExit'])
                if (aCE != 0 and aCE != 1):
                    raise Exception
                self.customFormulaExitAlert = aCE
                
                self.freqLoadCustom = int(config['custom']['freqLoadCustom'])
                if (self.freqLoadCustom != 0 and self.freqLoadCustom != 1):
                    self.freqLoadCustom = 0
                    raise Exception
                
                #---The default import path list---
                self.defaultImports = config['custom']['imports'].split("*")
                
            except:     #If there's an error loading settings
                pass    #Ignore the error
    
    def interpretArgs(self, args):
        if (len(args) > 1):         #If there's an argument coming in (the first argument is always the file name)
            if (len(args) > 2):     #If there's multiple arguments
                try:
                    whatExt = ""    #For the extension of an incoming file
                    for i in range(1, len(args)):
                        if (not os.path.exists(args[i])):
                            raise Exception
                        fname, fext = os.path.splitext(args[i])                 #Split the argument into its name and extension
                        if (fext != ".mmresiduals" and fext != ".mmerrors"):    #Check if an input file isn't a residual or error file
                            messagebox.showerror("File error", "Error: 2\nTo open multiple files, all of them must be either .mmresiduals or .mmerrors.")
                            self.residualsFileList.clear()
                            self.errorsFileList.clear()
                            break
                        elif (whatExt != "" and whatExt != fext):                  #Check if the extra files don't match the previous files' extensions
                            messagebox.showerror("File error", "Error: 2\nTo open multiple files, all of them must be either .mmresiduals or .mmerrors.")  
                            self.residualsFileList.clear()
                            self.errorsFileList.clear()
                            break
                        elif (fext == ".mmresiduals"):
                            self.residualsFileList.append(args[i])
                            whatExt = fext
                        elif (fext == ".mmerrors"):
                            self.errorsFileList.append(args[i])
                            whatExt = fext
                except:
                    messagebox.showerror("File error", "Error: 3\nError opening files")
            else:       #If there's only one input file
                try:    #Check if the file exists
                    if (not os.path.exists(args[1])):
                        raise Exception
                    self.argFile = args[1]
                except:
                    messagebox.showerror("File error", "Error: 3\nError opening file")
    
    #---Allow the mousewheel to be used to change tabs---
    def on_mouse_wheel(self, event, overTab):
        if (self.currentTab == 0):
            if (event.delta < 0):
                self.measureModelDown(None)
        elif (self.currentTab == 1):
            if (event.delta > 0):
                self.inputFileDown(None)
            else:
                self.errorFileDown(None)
        elif (self.currentTab == 2):
            if (event.delta > 0):
                self.measureModelDown(None)
            else:
                self.errorDown(None)
        elif (self.currentTab == 3):
            if (event.delta > 0):
                self.errorFileDown(None)
            else:
                self.formulaDown(None)
        elif (self.currentTab == 6):
            if (event.delta > 0):
                self.errorDown(None)
            else:
                self.settingsDown(None)
        elif (self.currentTab == 4):
            if (event.delta > 0):
                self.formulaDown(None)
            else:
                self.helpDown(None)
        elif (self.currentTab == 5):
            if (event.delta > 0):
                self.settingsDown(None)
        
        if (overTab == 0):
            self.fileOpen.configure(image=self.imgFile2 if self.currentTab != 0 else self.imgFile4)
        elif (overTab == 1):
            self.runModel.configure(image=self.imgModel2 if self.currentTab != 1 else self.imgModel4)
        elif (overTab == 2):
            self.errorFileLabel.configure(image=self.imgErrorFile2 if self.currentTab != 2 else self.imgErrorFile4)
        elif (overTab == 3):
            self.errorLabel.configure(image=self.imgError2 if self.currentTab != 3 else self.imgError4)
        elif (overTab == 4):
            self.settingsLabel.configure(image=self.imgSettings2 if self.currentTab != 4 else self.imgSettings4)
        elif (overTab == 5):
            self.helpLabel.configure(image=self.imgHelp2 if self.currentTab != 5 else self.imgHelp4)
        elif (overTab == 6):
            self.formulaLabel.configure(image=self.imgFormula2 if self.currentTab != 6 else self.imgFormula4)
    
    #---The functions for what happens when one of the links in the navigation pane is clicked--- 
    def inputFileDown(self, event):
        self.currentTab = 0                                         #Set the current tab to keep track of which one is open
        self.fileOpen.configure(image=self.imgFile3)                #Change the color of the tab being clicked on
        self.inputCanvas.configure(background="#0066FF")
        self.runModel.configure(image=self.imgModel)                #Set all the other tabs to have the standard background
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.help_frame.grid_remove()                                #Remove the other tabs
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="File Tools and Conversion")        #Change the title
        self.input_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)   #Grid the new tab
        self.error_file_frame.unbindIt()
        self.settings_frame.unbindIt()
        self.input_frame.bindIt()
            
    def inputFileUp(self, event):
        self.fileOpen.configure(image=self.imgFile4)             #Change the tab's color so it looks good
    
    #---The other functions follow the same setup as above---
    def measureModelDown(self, event):
        self.currentTab = 1
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel3)
        self.runCanvas.configure(background="#0066FF")
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.runModel.focus_get()
        self.input_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Measurement Model")
        self.model_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_file_frame.unbindIt()
        self.input_frame.unbindIt()
        self.settings_frame.unbindIt()
        
    def measureModelUp(self, event):
        self.runModel.configure(image=self.imgModel4)
    
    def errorFileDown(self, event):
        self.currentTab = 2
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile3)
        self.errorFileCanvas.configure(background="#0066FF")
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.errorFileLabel.focus_get()
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Error File Preparation")
        self.error_file_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.input_frame.unbindIt()
        self.settings_frame.unbindIt()
        self.error_file_frame.bindIt()
    
    def errorFileUp(self, event):
        self.errorFileLabel.configure(image=self.imgErrorFile4)
    
    def errorDown(self, event):
        self.currentTab = 3
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError3)
        self.errorCanvas.configure(background="#0066FF")
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Error Analysis")
        self.error_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_file_frame.unbindIt()
        self.input_frame.unbindIt()
        self.settings_frame.unbindIt()
    
    def errorUp(self, event):
        self.errorLabel.configure(image=self.imgError4)
        
    def formulaDown(self, event):
        self.currentTab = 6
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula3)
        self.formulaCanvas.configure(background="#0066FF")
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_frame.grid_remove()
        self.frame_label.configure(text="Custom Formula Fitting")
        self.formula_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_file_frame.unbindIt()
        self.input_frame.unbindIt()
        self.settings_frame.unbindIt()
    
    def formulaUp(self, event):
        self.formulaLabel.configure(image=self.imgFormula4)
        
    def settingsDown(self, event):
        self.currentTab = 4
        self.fileOpen.configure(background=self.backgroundColor)
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings3)
        self.settingsCanvas.configure(background="#0066FF")
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.help_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.settings_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.frame_label.configure(text="Settings")
        self.error_file_frame.unbindIt()
        self.input_frame.unbindIt()
        self.settings_frame.bindIt()
    
    def settingsUp(self, event):
        self.settingsLabel.configure(image=self.imgSettings4)
    
    def helpDown(self, event):
        self.currentTab = 5
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp3)
        self.helpCanvas.configure(background="#0066FF")
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.helpLabel.focus_get()
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Help and About")
        self.help_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_file_frame.unbindIt()
        self.input_frame.unbindIt()
        self.settings_frame.unbindIt()
        
    def helpUp(self, event):
        self.helpLabel.configure(image=self.imgHelp4)
    
    #---The functions for when a link is no longer hovered over: change the background color back to what it was---
    def leaveInputFile(self, event):
        if (self.currentTab == 0):
            self.fileOpen.configure(image=self.imgFile3)
        else:
            self.fileOpen.configure(image=self.imgFile)
    
    def leaveMeasureModel(self, event):
        if (self.currentTab == 1):
            self.runModel.configure(image=self.imgModel3)
        else:
            self.runModel.configure(image=self.imgModel)
    
    def leaveErrorFile(self, event):
        if (self.currentTab == 2):
            self.errorFileLabel.configure(image=self.imgErrorFile3)
        else:
            self.errorFileLabel.configure(image=self.imgErrorFile)
    
    def leaveError(self, event):
        if (self.currentTab == 3):
            self.errorLabel.configure(image=self.imgError3)
        else:
            self.errorLabel.configure(image=self.imgError)
    
    def leaveFormula(self, event):
        if (self.currentTab == 6):
            self.formulaLabel.configure(image=self.imgFormula3)
        else:
            self.formulaLabel.configure(image=self.imgFormula)
    
    def leaveHelp(self, event):
        if (self.currentTab == 5):
            self.helpLabel.configure(image=self.imgHelp3)
        else:
            self.helpLabel.configure(image=self.imgHelp)
            
    def leaveSettings(self, event):
        if (self.currentTab == 4):
            self.settingsLabel.configure(image=self.imgSettings3)
        else:
            self.settingsLabel.configure(image=self.imgSettings)
    
    #---If a file with an unknown is loaded, send it to the input tab---
    def enterInput(self, fileName):
        self.currentTab = 0
        self.fileOpen.configure(image=self.imgFile3)
        self.inputCanvas.configure(background="#0066FF")
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.help_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="File Tools and Conversion")
        self.input_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.input_frame.inputEnter(fileName)
    
    #---If a .mmfile is opened straight to the measurement model---
    def enterMeasureModel(self, fileName):
        self.currentTab = 1
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel3)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.errorLabel.configure(image=self.imgError)
        self.helpLabel.configure(image=self.imgHelp)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background="#0066FF")
        self.helpCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.input_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Measurment Model")
        self.model_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.model_frame.browseEnter(fileName)
    
    #---If a .mmfitting file is opened straight to the measurement model---
    def enterMeasureModelFitting(self, fileName):
        self.currentTab = 1
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel3)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.errorLabel.configure(image=self.imgError)
        self.helpLabel.configure(image=self.imgHelp)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background="#0066FF")
        self.helpCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.runModel.focus_get()
        self.input_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.error_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Measurment Model")
        self.model_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.model_frame.browseEnterFitting(fileName)
    
    #---If a .mmresiduals file is opened straight to error file preparation---
    def enterResidual(self, fileName):
        self.currentTab = 2
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile3)
        self.errorFileCanvas.configure(background="#0066FF")
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.model_frame.grid_remove()
        self.input_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Error File Preparation")
        self.error_file_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_file_frame.residualEnter(fileName)
    
    #---If multiple .mmresiduals files are opened straight to error file preparation---
    def enterResiduals(self, files):
        self.currentTab = 2
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile3)
        self.errorFileCanvas.configure(background="#0066FF")
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.model_frame.grid_remove()
        self.input_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Error File Preparation")
        self.error_file_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_file_frame.residualsEnter(files)
    
    #---If a .mmerror file is opened straight to error structure fitting---
    def enterError(self, fileName):
        self.currentTab = 3
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError3)
        self.errorCanvas.configure(background="#0066FF")
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.model_frame.grid_remove()
        self.input_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Error Analysis")
        self.error_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_frame.errorEnter(fileName)
    
    #---If multiple .mmerror files are opened straight to error structure fitting---
    def enterErrors(self, files):
        self.currentTab = 3
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError3)
        self.errorCanvas.configure(background="#0066FF")
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula)
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        self.model_frame.grid_remove()
        self.input_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.formula_frame.grid_remove()
        self.frame_label.configure(text="Error Analysis")
        self.error_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.error_frame.errorsEnter(files)
    
    #---If a .mmcustom file is opened straight to custom formula fitting---
    def enterCustom(self, fileName):
        self.currentTab = 6
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula3)
        self.formulaCanvas.configure(background="#0066FF")
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_frame.grid_remove()
        self.frame_label.configure(text="Custom Formula Fitting")
        self.formula_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.formula_frame.formulaEnter(fileName)
    
    def enterFormula(self, fileName):
        self.currentTab = 6
        self.fileOpen.configure(image=self.imgFile)
        self.runModel.configure(image=self.imgModel)
        self.errorFileLabel.configure(image=self.imgErrorFile)
        self.helpLabel.configure(image=self.imgHelp)
        self.errorLabel.configure(image=self.imgError)
        self.settingsLabel.configure(image=self.imgSettings)
        self.formulaLabel.configure(image=self.imgFormula3)
        self.formulaCanvas.configure(background="#0066FF")
        self.runCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.settingsCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.input_frame.grid_remove()
        self.model_frame.grid_remove()
        self.error_file_frame.grid_remove()
        self.help_frame.grid_remove()
        self.settings_frame.grid_remove()
        self.error_frame.grid_remove()
        self.frame_label.configure(text="Custom Formula Fitting")
        self.formula_frame.grid(row=1, column=2, sticky="NW", padx=20, pady=20)
        self.formula_frame.enterFormula(fileName)
    
    #---To check if an "Are you sure you want to close" prompt is needed---
    def canCloseInput(self):
        if (self.currentTab != 0):  #If the current tab isn't the input tab, don't ask
            return True
        else:
            if (self.input_frame.needToSave() and self.inputExitAlert):             #If there is a need to save and alerting is turned on
                return False
            else:
                return True
    
    def canCloseCustom(self):
        if (self.currentTab != 6):  #If the current tab isn't the custom formula tab, don't ask
            return True
        else:
            if (self.formula_frame.needToSave() and self.customFormulaExitAlert):   #If there is a need to save and alerting is turned on
                return False
            else:
                return True
            
    #---Getters and setters for various settings---    
    def getTheme(self):
        return self.theme
    
    def getBarColor(self):
        return self.backgroundColor
    
    def getAccentColor(self):
        return self.activeColor
    
    def getHighlightColor(self):
        return self.highlightColor
    
    def getDetectComments(self):
        return self.detectCommentsDefault
    
    def getComments(self):
        return self.commentsDefault
    
    def getDetectDelimiter(self):
        return self.detectDelimiterDefault
    
    def getDelimiter(self):
        return self.delimiterDefault
    
    def getMC(self):
        return self.MCDefault
    
    def getFit(self):
        return self.fitDefault
    
    def getWeight(self):
        return self.weightDefault
    
    def getAlpha(self):
        return self.alphaDefault
    
    def getErrorAlpha(self):
        return self.errorAlphaDefault
    
    def getErrorBeta(self):
        return self.errorBetaDefault
    
    def getErrorRe(self):
        return self.errorReDefault
    
    def getErrorGamma(self):
        return self.errorGammaDefault
    
    def getErrorDelta(self):
        return self.errorDeltaDefault
    
    def getErrorWeighting(self):
        return self.errorWeightingDefault
    
    def getDetrend(self):
        return self.detrendDefault
    
    def getMovingAverage(self):
        return self.errorMADefault
    
    def getEllipseColor(self):
        return self.ellipseColor
    
    def getCurrentDirectory(self):
        return self.currentDirectory
    
    def setCurrentDirectory(self, directory):
        self.currentDirectory = directory
    
    def getDefaultDirectory(self):
        return self.defaultDirectory
    
    def setDefaultDirectory(self, directory):
        self.defaultDirectory = directory
    
    def getDefaultFormulaDirectory(self):
        return self.defaultFormulaDirectory
    
    def setDefaultFormulaDirectory(self, directory):
        self.defaultFormulaDirectory = directory
    
    def getInputExitAlert(self):
        return self.inputExitAlert
    
    def setInputExitAlert(self, val):
        self.inputExitAlert = val
    
    def getCustomFormulaExitAlert(self):
        return self.customFormulaExitAlert
    
    def setCustomFormulaExitAlert(self, val):
        self.customFormulaExitAlert = val
    
    def getDefaultTab(self):
        return self.defaultTab
    
    def setFreqLoad(self, val):
        self.freqLoad = val
    
    def getFreqLoad(self):
        return self.freqLoad
    
    def setFreqUndo(self, val):
        self.freqUndo = val
        
    def getFreqUndo(self):
        return self.freqUndo
    
    def setFreqLoadCustom(self, val):
        self.freqLoadCustom = val
    
    def getFreqLoadCustom(self):
        return self.freqLoadCustom
    
    def setDefaultImports(self, val):
        self.defaultImports = val
    
    def getDefaultImports(self):
        return self.defaultImports
    
    def setScroll(self, val):
        if (val == 1):
            self.left_frame.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, -1))
            self.fileOpen.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 0))
            self.runModel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 1))
            self.errorFileLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 2))
            self.errorLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 3))
            self.formulaLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 6))
            self.settingsLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 4))
            self.helpLabel.bind("<MouseWheel>", lambda e: self.on_mouse_wheel(e, 5))
        else:
            self.left_frame.unbind("<MouseWheel>")
            self.fileOpen.unbind("<MouseWheel>")
            self.runModel.unbind("<MouseWheel>")
            self.errorFileLabel.unbind("<MouseWheel>")
            self.errorLabel.unbind("<MouseWheel>")
            self.formulaLabel.unbind("<MouseWheel>")
            self.settingsLabel.unbind("<MouseWheel>")
            self.helpLabel.unbind("<MouseWheel>")
        self.defaultScroll = val
    
    def getScroll(self):
        return self.defaultScroll
    
    #---Change the colors/appearance---
    def setBarColor(self, color):
        self.backgroundColor = color
        self.left_frame.configure(bg=color)
        self.top_frame.configure(bg=color)
        self.frame_label.configure(bg=color)
        self.fileOpen.configure(background=self.backgroundColor)
        self.runModel.configure(background=self.backgroundColor)
        self.helpLabel.configure(background=self.backgroundColor)
        self.errorLabel.configure(background=self.backgroundColor)
        self.errorFileLabel.configure(background=self.backgroundColor)
        self.formulaLabel.configure(background=self.backgroundColor)
        self.settingsLabel.configure(background=self.backgroundColor)
        self.canvasFrame.configure(background=self.backgroundColor)
        self.inputCanvas.configure(background=self.backgroundColor)
        self.runCanvas.configure(background=self.backgroundColor)
        self.helpCanvas.configure(background=self.backgroundColor)
        self.errorCanvas.configure(background=self.backgroundColor)
        self.errorFileCanvas.configure(background=self.backgroundColor)
        self.formulaCanvas.configure(background=self.backgroundColor)
        barColorNoHash = self.backgroundColor.lstrip('#')
        barColorRGB = tuple(int(barColorNoHash[i:i+2], 16) for i in (0, 2 ,4))      #See earlier comment for citation/explanation
        self.frameLabelColor = "#000000" if (barColorRGB[0]*0.299 + barColorRGB[1]*0.587 + barColorRGB[2]*0.114) > 186 else "#FFFFFF"   #See earlier comment for citation
        self.frame_label.configure(fg=self.frameLabelColor)
    
    def setHighlightColor(self, color):
        self.highlightColor = color
    
    def setActiveColor(self, color):
        self.activeColor = color
        
    def setEllipseColor(self, color):
        self.ellipseColor = color
        self.model_frame.setEllipseColor(color)
        self.formula_frame.setEllipseColor(color)
    
    def setThemeLight(self):
        self.theme = "light"
        self.master.configure(background="#FFFFFF")
        self.input_frame.setThemeLight()
        self.help_frame.setThemeLight()
        self.error_frame.setThemeLight()
        self.model_frame.setThemeLight()
        self.error_file_frame.setThemeLight()
        self.formula_frame.setThemeLight()
    
    def setThemeDark(self):
        self.theme = "dark"
        self.master.configure(background="#424242")
        self.input_frame.setThemeDark()
        self.help_frame.setThemeDark()
        self.error_frame.setThemeDark()
        self.model_frame.setThemeDark()
        self.error_file_frame.setThemeDark()
        self.formula_frame.setThemeDark()

    def closeAllPopups(self, e=None):
        self.input_frame.closeWindows()
        self.error_file_frame.closeWindows()
        self.model_frame.closeWindows()
        self.error_frame.closeWindows()
        self.formula_frame.closeWindows()

if (__name__ == "__main__"):            #If the code is being run as the main module
    multiprocessing.freeze_support()    #Allows multiprocessing to work on Windows
    try:
        if 'win' in sys.platform:       #Makes the program sharper on Windows
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    
    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
    
        return os.path.join(base_path, relative_path)
    
    root = tk.Tk()                  #The root window
    root.withdraw()                 #Hide the root window for the copyright notice to apppear
    gui = myGUI(root, sys.argv)     #The GUI object
    
    #---Set up the copyright splash screen---
    splashScreen = tk.Toplevel(bg="white")
    splashScreen.after(1, lambda: splashScreen.focus_force())
    splashScreen.title("Measurement Model")
    splashScreen.iconbitmap(resource_path('img/elephant3.ico'))
    copyrightLabel = tk.Label(splashScreen, text="©Copyright 2020 University of Florida Research Foundation, Inc. All Rights Reserved.\n\nThis program is free software: you can redistribute it and/or modify \
it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.\n\n\
\
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the \
GNU General Public License for more details.\n\n\
You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.", justify=tk.LEFT, wraplength=600, bg="white", fg="black")
        
    gnuLicenseLink = tk.Label(splashScreen, text="GNU General Public License", cursor="hand2", bg="white", fg="blue")
    gnuLicenseLink.bind('<Button-1>', lambda e: webbrowser.open_new("gpl-3.0-standalone.html"))
    buttonFrame = tk.Frame(splashScreen, bg="white")
    
    #---If Start program is clicked in the splash screen---
    def on_start_splash():
        splashScreen.destroy()
        root.deiconify()
        root.lift()
    okButton = ttk.Button(buttonFrame, text="Start program", command=on_start_splash)
    okButton.bind("<Return>", lambda e: on_start_splash())
    
    #---If Cancel is clicked in the splash screen---
    def on_cancel_splash():
        splashScreen.destroy()
        root.destroy()
    
    #---Prepare the splash screen---
    cancelButton = ttk.Button(buttonFrame, text="Cancel", command=on_cancel_splash)
    cancelButton.bind("<Return>", lambda e: on_cancel_splash())
    copyrightLabel.pack(side=tk.TOP, fill=tk.X, expand=True)
    gnuLicenseLink.pack(side=tk.TOP, fill=tk.X, expand=True)
    okButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
    cancelButton.pack(side=tk.RIGHT, fill=tk.X, expand=True)
    buttonFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=True)
    splashScreen.protocol("WM_DELETE_WINDOW", on_cancel_splash)
    okButton.focus_set()
    splashScreen.resizable(False, False)
    
    def on_closing():                               #Check what's happening on closing
        if (not gui.canCloseInput() or not gui.canCloseCustom()):   #If there is a need to save and alerting is turned on, ask before closing
            if (messagebox.askokcancel("Exit?", "You have not saved. Do you still wish to exit?")):
                gui.model_frame.cancelThreads()     #Close running threads before exit
                gui.formula_frame.cancelThreads()
                root.destroy()
        else:
            gui.model_frame.cancelThreads()         #Close running threads before exit
            gui.formula_frame.cancelThreads()
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)   #Catch the closing command
    root.mainloop()