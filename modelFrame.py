# -*- coding: utf-8 -*-
"""
Created on Mon Jul  9 16:30:37 2018

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

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolledtext
from tkinter.filedialog import askopenfilename, asksaveasfile, askdirectory
from tkinter import messagebox
import mmFittingA
import mmFittingCap
import autoFitter
import numpy as np
import threading
import ctypes
import comtypes.client as cc
import queue
import pyperclip
#--------------------------------pyperclip-----------------------------------
#     Source: https://pypi.org/project/pyperclip/
#     Email: al@inventwithpython.com
#     By: Al Sweigart
#     License: BSD
#----------------------------------------------------------------------------
import os.path
import sys
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.patches as patches
import matplotlib.pyplot as plt
#from matplotlib.transforms import Bbox
from rangeSlider import RangeSlider
#------------------------------Range Slider----------------------------------
#  Source: https://github.com/halsafar/rangeslider
#  "THE BEER-WARE LICENSE" (Revision 42):
#  <shinhalsafar@gmail.com> wrote this file. As long as you retain this notice you
#  can do whatever you want with this stuff. If we meet some day, and you think
#  this stuff is worth it, you can buy me a beer in return.
#
#	Stephen Damm (shinhalsafar@gmail.com)
#----------------------------------------------------------------------------
from functools import partial

class ThreadedTaskAuto(threading.Thread):
    def __init__(self, queue, max_nve, fRe, fReV, nMC, w, r, j, per, c, t, errA, errB, errBRe, errG, errD):
        threading.Thread.__init__(self)
        self.queue = queue
        self.nve_max = max_nve
        self.reFix = fReV
        self.reFixVal = fReV
        self.nMC = nMC
        self.wdat = w
        self.rdat = r
        self.jdat = j
        self.listPer = per
        self.choice = c
        self.autoType = t
        self.errAlpha = errA
        self.errBeta = errB
        self.errBetaRe = errBRe
        self.errGamma = errG
        self.errDelta = errD
        self.fittingObject = autoFitter.autoFitter()
    def run(self):
        try:
            r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic, ft, fw = self.fittingObject.autoFit(self.nve_max, self.reFix, self.reFixVal, self.nMC, self.wdat, self.rdat, self.jdat, self.listPer, self.choice, self.autoType, self.errAlpha, self.errBeta, self.errBetaRe, self.errGamma, self.errDelta)  
            self.queue.put((r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic, ft, fw))
        except:
            self.queue.put(("@", "@", "@", "@", "@", "@", "@", "@", "@", "@", "@", "@", "@"))
    def get_id(self):
        # From https://www.geeksforgeeks.org/python-different-ways-to-kill-a-thread/
        # returns id of the respective thread 
        if hasattr(self, '_thread_id'): 
            return self._thread_id 
        for id, thread in threading._active.items(): 
            if thread is self: 
                return id  
    def terminate(self):
        self.fittingObject.terminateProcesses()
        # From https://www.geeksforgeeks.org/python-different-ways-to-kill-a-thread/
        thread_id = self.get_id() 
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, ctypes.py_object(SystemExit))
        if res > 1: 
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0) 
            #print('Exception raise failure')
        
#---Calls the fitting module; threaded to prevent the GUI from freezing---
class ThreadedTask(threading.Thread):
    def __init__(self, queue, w, r, j, nVE, nMC, c, an, rm, fT, bL, g, cnst, bU, per, errA, errB, errBRe, errG, errD):
        threading.Thread.__init__(self)
        self.queue = queue
        self.wdat = w
        self.rdat = r
        self.jdat = j
        self.nVE = nVE
        self.nMC = nMC
        self.choice = c
        self.assumeNoise = an
        self.rM= rm
        self.fitType = fT
        self.boundLower = bL
        self.boundUpper = bU
        self.guesses = g
        self.const = cnst
        self.listPer = per
        self.errAlpha = errA
        self.errBeta = errB
        self.errBetaRe = errBRe
        self.errGamma = errG
        self.errDelta = errD
        self.fittingObject = mmFittingA.fitter()
    def run(self):
        try:
            r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic = self.fittingObject.findFit(self.wdat, self.rdat, self.jdat, self.nVE, self.nMC, self.choice, self.assumeNoise, self.rM, self.fitType, self.boundLower, self.guesses, self.const, self.boundUpper, self.listPer, self.errAlpha, self.errBeta, self.errBetaRe, self.errGamma, self.errDelta)
            self.queue.put((r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic))
        except:
            self.queue.put(("@", "@", "@", "@", "@", "@", "@", "@", "@", "@", "@"))
    def get_id(self):
        # From https://www.geeksforgeeks.org/python-different-ways-to-kill-a-thread/
        # returns id of the respective thread 
        if hasattr(self, '_thread_id'): 
            return self._thread_id 
        for id, thread in threading._active.items(): 
            if thread is self: 
                return id
    
    def terminate(self):
        self.fittingObject.terminateProcesses()
        # From https://www.geeksforgeeks.org/python-different-ways-to-kill-a-thread/
        thread_id = self.get_id() 
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, ctypes.py_object(SystemExit))
        if res > 1: 
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0) 
            #print('Exception raise failure')

class ThreadedTaskCap(threading.Thread):
    def __init__(self, queue, w, r, j, nVE, nMC, c, an, rm, fT, bL, g, cnst, bU, per, errA, errB, errBRe, errG, errD, cG, bLC, bUC, fC):
        threading.Thread.__init__(self)
        self.queue = queue
        self.wdat = w
        self.rdat = r
        self.jdat = j
        self.nVE = nVE
        self.nMC = nMC
        self.choice = c
        self.assumeNoise = an
        self.rM= rm
        self.fitType = fT
        self.boundLower = bL
        self.boundUpper = bU
        self.guesses = g
        self.const = cnst
        self.listPer = per
        self.errAlpha = errA
        self.errBeta = errB
        self.errBetaRe = errBRe
        self.errGamma = errG
        self.errDelta = errD
        self.capG = cG
        self.bLCap = bLC
        self.bUCap = bUC
        self.fCap = fC
        self.fittingObjectCap = mmFittingCap.fitter()
    def run(self):
        try:
            r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic, capR, capS, corCap = self.fittingObjectCap.findFit(self.wdat, self.rdat, self.jdat, self.nVE, self.nMC, self.choice, self.assumeNoise, self.rM, self.fitType, self.boundLower, self.guesses, self.const, self.boundUpper, self.listPer, self.capG, self.bLCap, self.bUCap, self.fCap, self.errAlpha, self.errBeta, self.errBetaRe, self.errGamma, self.errDelta)
            self.queue.put((r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic, capR, capS, corCap))
        except:
            self.queue.put(("@", "@", "@", "@", "@", "@", "@", "@", "@", "@", "@"))
    
    def get_id(self): 
        # From https://www.geeksforgeeks.org/python-different-ways-to-kill-a-thread/
        # returns id of the respective thread 
        if hasattr(self, '_thread_id'): 
            return self._thread_id 
        for id, thread in threading._active.items(): 
            if thread is self: 
                return id
    
    def terminate(self):
        self.fittingObjectCap.terminateProcesses()
        # From https://www.geeksforgeeks.org/python-different-ways-to-kill-a-thread/
        thread_id = self.get_id() 
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, ctypes.py_object(SystemExit))
        if res > 1: 
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0) 
            #print('Exception raise failure')

class CreateToolTip(object):
    """
    create a tooltip for a given widget
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

class mmF(tk.Frame):
    def __init__(self, parent, topOne):
        
        def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
        
            return os.path.join(base_path, relative_path)
        
        self.parent = parent
        self.topGUI = topOne
        if (self.topGUI.getTheme() == "dark"):
            self.backgroundColor = "#424242"
            self.foregroundColor = "white"
        elif (self.topGUI.getTheme() == "light"):
            self.backgroundColor = "white"
            self.foregroundColor = "black"
        tk.Frame.__init__(self, parent, background=self.backgroundColor)
        
        s = ttk.Style()
        if (self.topGUI.getTheme() == "light"):
            s.configure('TCheckbutton', background='white')
            s.configure('TCombobox', background='white')
            s.configure('TEntry', background='white')
            s.configure('TRadiobutton', background='white', foreground='black')
        elif (self.topGUI.getTheme() == "dark"):
            s.configure('TCheckbutton', background='#424242')
            s.configure('TCombobox', background='#424242')
            s.configure('TEntry', background='#424242')
            s.configure('TLabelFrame', background='#424242', foreground="#FFFFFF")
            s.configure('TRadiobutton', background='#424242', foreground='#FFFFFF')
        self.ellipseColor = self.topGUI.getEllipseColor()
                        
        
        self.omega = np.zeros(1)
        self.nVoigt = 1
        self.resultsWindows = []
        self.resultPlotBigs = []
        self.resultPlotBigFigs = []
        self.updateWindows = []
        self.updateFigs = []
        self.simplePopups = []
        self.pltFigs = []
        self.resultPlots = []
        self.saveNyCanvasButtons = []
        self.saveNyCanvasButton_ttps = []
        self.paramLabels = []
        self.paramEntries = []
        self.paramEntryVariables = []
        self.paramEntryVariables.append(tk.StringVar(self, ""))
        self.paramComboboxes = []
        self.paramComboboxVariables = []
        self.paramComboboxVariables.append(tk.StringVar(self, "+"))
        self.tauComboboxes = []
        self.tauComboboxVariables = []
        #self.paramRemoveLabels = []
        #self.tauComboboxVariables.append(tk.StringVar(self, "+"))
        self.paramFrames = []
        self.paramDeleteButtons = []
        self.capacitanceCheckboxVariable = tk.IntVar(self, 0)
        self.capacitanceComboboxVariable = tk.StringVar(self, "+ or -")
        self.capacitanceEntryVariable = tk.StringVar(self, "0")
        self.paramPopup = tk.Toplevel(bg=self.backgroundColor)
        self.paramPopup.withdraw()
        self.pcanvas = tk.Canvas(self.paramPopup, borderwidth=0, highlightthickness=0, background=self.backgroundColor)
        self.pframe = tk.Frame(self.pcanvas, background=self.backgroundColor)
        self.vsb = tk.Scrollbar(self.paramPopup, orient="vertical", command=self.pcanvas.yview)
        self.howMany = tk.Label(self.paramPopup, text="Number of elements: ", bg=self.backgroundColor)
        self.advancedOptions = ttk.Button(self.paramPopup, text="Advanced options")
        self.advancedOptionsPopup = tk.Toplevel(bg=self.backgroundColor)
        self.advancedOptionsPopup.withdraw()
        self.advancedOptionsPopup.grid_rowconfigure(1, weight=1)
        self.paramListbox = tk.Listbox(self.advancedOptionsPopup, selectmode=tk.BROWSE, activestyle='none', background=self.backgroundColor, foreground=self.foregroundColor, exportselection=False)
        self.paramListboxScrollbar = ttk.Scrollbar(self.advancedOptionsPopup, orient=tk.VERTICAL)
        self.paramListboxScrollbar['command'] = self.paramListbox.yview
        self.paramListbox.configure(yscrollcommand=self.paramListboxScrollbar.set)
        self.advancedChoiceLabel = tk.Label(self.advancedOptionsPopup, text="Re", bg=self.backgroundColor, fg=self.foregroundColor, font=('Helvetica', 10, 'bold'))
        self.advancedParamFrame = tk.Frame(self.advancedOptionsPopup, background=self.backgroundColor)
        self.multistartEnabled = []
        self.multistartSelection = []
        self.multistartLower = []
        self.multistartUpper = []
        self.multistartNumber = []
        self.multistartCustom = []
        self.tauDefault = 1
        self.rDefault = 1
        self.lengthOfData = 0
        self.aR = ""
        self.wdata = []
        self.rdata = []
        self.jdata = []
        self.wdataRaw = []
        self.rdataRaw = []
        self.jdataRaw = []
        self.fits = []
        self.sigmas = []
        self.sdrReal = []
        self.sdrImag = []
        self.fitWeightR = []
        self.fitWeightJ = []
        self.alreadyPlotted = False
        self.currentNVEPlotted = 0
        self.currentNVEPlotNeeded = 0
        self.undoStack = []
        self.undoSigmaStack = []
        self.undoCorStack = []
        self.undoAICStack = []
        self.undoChiStack = []
        self.undoCapStack = []
        self.undoCapSigmaStack = []
        self.undoCapNeededStack = []
        self.undoCorCapStack = []
        self.undoUpDeleteStack = []
        self.undoLowDeleteStack = []
        self.undoSdrRealStack = []
        self.undoSdrImagStack = []
        self.undoAlphaStack = []
        self.pastAlphaVariable = 1
        self.undoFitTypeStack = []
        self.pastFitType = "Complex"
        self.undoWeightingStack = []
        self.pastWeighting = "Modulus"
        self.undoErrorModelChoicesStack = []
        self.pastErrorModelChoices = [False, False, True, True]
        self.undoErrorModelValuesStack = []
        self.pastErrorValues = [0.1, 0.1, 0.1, 0.1]
        self.undoNumSimulationsStack = []
        self.pastNumSimulations = 1000
        self.capCor = 0
        self.needUndo = False
        self.whatFit = "C"
        self.whatFitStack = []
        self.alreadyChanged = False
        self.updateReVariable = tk.StringVar(self, "0")
        self.autoFitWindow = tk.Toplevel(background=self.backgroundColor)
        self.autoFitWindow.withdraw()
        self.autoFitWindow.title("Automatic Fitting")
        self.autoFitWindow.iconbitmap(resource_path('img/elephant3.ico'))
        self.autoFitWindow.resizable(False, False)
        self.autoMaxNVEVariable = tk.StringVar(self, "15")
        self.autoNMCVariable = tk.StringVar(self, str(self.topGUI.getMC()))
        self.autoReFixVariable = tk.IntVar(self, 0)
        self.autoReFixEntryVariable = tk.StringVar(self, "0")
        self.autoWeightingVariable = tk.StringVar(self, "Modulus")
        self.errorAlphaCheckboxVariableAuto = tk.IntVar(self, 0)
        self.errorAlphaVariableAuto = tk.StringVar(self, "0.1")
        self.errorBetaCheckboxVariableAuto = tk.IntVar(self, 0)
        self.errorBetaVariableAuto = tk.StringVar(self, "0.1")
        self.errorBetaReCheckboxVariableAuto = tk.IntVar(self, 0)
        self.errorBetaReVariableAuto = tk.StringVar(self, "1")
        self.errorGammaCheckboxVariableAuto = tk.IntVar(self, 1)
        self.errorGammaVariableAuto = tk.StringVar(self, "0.1")
        self.errorDeltaCheckboxVariableAuto = tk.IntVar(self, 1)
        self.errorDeltaVariableAuto = tk.StringVar(self, "0.1")
        self.radioVal = tk.IntVar(self, 2)
        self.freqWindow = tk.Toplevel(background=self.backgroundColor)
        self.freqWindow.withdraw()
        self.freqWindow.title("Change Frequency Range")
        self.freqWindow.iconbitmap(resource_path('img/elephant3.ico'))
        #self.rs = RangeSlider(self.freqWindow, lowerBound=-4, upperBound=7, background=self.backgroundColor, tickColor=self.foregroundColor)
        #self.rs.grid(column=1, row=1, rowspan=5)
        self.lowestUndeleted = tk.Label(self.freqWindow, text="Lowest Remaining Frequency: 0", background=self.backgroundColor, foreground=self.foregroundColor)
        self.highestUndeleted = tk.Label(self.freqWindow, text="Highest Remaining Frequency: 0", background=self.backgroundColor, foreground=self.foregroundColor)
        self.lowerSpinboxVariable = tk.StringVar(self, "0")
        self.upperSpinboxVariable = tk.StringVar(self, "0")
        self.upDelete = 0
        self.lowDelete = 0
        self.pastUpDelete = 0
        self.pastLowDelete = 0
        self.capUsed = False
        self.resultCap = 0
        self.sigmaCap = 0
        self.listPercent = []
        self.autoPercent = []
        self.ellipsisPercent = 0
        self.currentThreads = []
        self.magicPlot = tk.Toplevel()
        self.magicPlot.withdraw()
        self.magicPlot.title("Magic Finger Nyquist Plot")     #Plot window title
        self.magicPlot.iconbitmap(resource_path('img/elephant3.ico'))
        self.magicElementFrame = tk.Frame(self.magicPlot, background=self.backgroundColor)
        self.magicElementRVariable = tk.StringVar(self, '0')
        self.magicElementTVariable = tk.StringVar(self, '0')
        self.magicElementREntry = ttk.Entry(self.magicElementFrame, textvariable=self.magicElementRVariable, width=10)
        self.magicElementRLabel = tk.Label(self.magicElementFrame, background=self.backgroundColor, foreground=self.foregroundColor, text="R: ")
        self.magicElementTEntry = ttk.Entry(self.magicElementFrame, textvariable=self.magicElementTVariable, width=10)
        self.magicElementTLabel = tk.Label(self.magicElementFrame, background=self.backgroundColor, foreground=self.foregroundColor, text="T: ")
        self.magicElementEnterButton = ttk.Button(self.magicElementFrame, text="Add")
        self.magicElementRLabel.grid(row=0, column=2, padx=(5, 0))
        self.magicElementREntry.grid(row=0, column=3)
        self.magicElementTLabel.grid(row=0, column=4)
        self.magicElementTEntry.grid(row=0, column=5, padx=(0, 5))
        self.magicElementEnterButton.grid(row=0, column=6, padx=5)
        self.magicElementFrame.pack(side=tk.BOTTOM, fill=tk.X, pady=10, expand=True)
        self.magicInput = Figure(figsize=(5,4), dpi=100)
        toolbarFrame = tk.Frame(master=self.magicPlot)
        toolbarFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.magicCanvasInput = FigureCanvasTkAgg(self.magicInput, self.magicPlot)
        self.magicCanvasInput.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        
        toolbar = NavigationToolbar2Tk(self.magicCanvasInput, toolbarFrame)    #Enables the zoom and move toolbar for the plot
        toolbar.update()
        self.magicCanvasInput._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        def OpenFile():
            name = askopenfilename(initialdir=self.topGUI.getCurrentDirectory, filetypes = [("Measurement model files", "*.mmfile *.mmfitting"), ("Measurement model data (*.mmfile)", "*.mmfile"), ("Measurement model fit (*.mmfitting)", "*.mmfitting")], title = "Choose a file")
            try:
                with open(name,'r') as UseFile:
                    return name
            except:
                return "+"
            
        def browse():
#            if (self.browseEntry.get() == ""):
            n = OpenFile()
            if (n != '+'):
                fname, fext = os.path.splitext(n)
                directory = os.path.dirname(str(n))
                self.topGUI.setCurrentDirectory(directory)
                if (fext == ".mmfile"):
                    try:
                        data = np.loadtxt(n)
                        w_in = data[:,0]
                        r_in = data[:,1]
                        j_in = data[:,2]
#                        self.wdata = w_in
#                        self.rdata = r_in
#                        self.jdata = j_in
                        self.wdataRaw = w_in
                        self.rdataRaw = r_in
                        self.jdataRaw = j_in
                        doneSorting = False
                        while (not doneSorting):
                            doneSorting = True
                            for i in range(len(self.wdataRaw)-1):
                                if (self.wdataRaw[i] > self.wdataRaw[i+1]):
                                    temp = self.wdataRaw[i]
                                    self.wdataRaw[i] = self.wdataRaw[i+1]
                                    self.wdataRaw[i+1] = temp
                                    tempR = self.rdataRaw[i]
                                    self.rdataRaw[i] = self.rdataRaw[i+1]
                                    self.rdataRaw[i+1] = tempR
                                    tempJ = self.jdataRaw[i]
                                    self.jdataRaw[i] = self.jdataRaw[i+1]
                                    self.jdataRaw[i+1] = tempJ
                                    doneSorting = False
                        if (self.topGUI.getFreqLoad() == 1):
                            try:
                                if (self.upDelete == 0):
                                    self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                                    self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                                    self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                                else:
                                    self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                    self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                    self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.lowerSpinboxVariable.set(str(self.lowDelete))
                                self.upperSpinboxVariable.set(str(self.upDelete))
                            except:
                                messagebox.showwarning("Frequency error", "There are more frequencies set to be deleted than data points. The number of frequencies to delete has been reset to 0.")
                                self.upDelete = 0
                                self.lowDelete = 0
                                self.lowerSpinboxVariable.set(str(self.lowDelete))
                                self.upperSpinboxVariable.set(str(self.upDelete))
                                self.wdata = self.wdataRaw.copy()
                                self.rdata = self.rdataRaw.copy()
                                self.jdata = self.jdataRaw.copy()
                            #self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                            #self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                            #self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                            #self.rs.setNumberOfMajorTicks(10)
                            #self.rs.showMinorTicks(False)
                            #self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                            #self.rs.setLower(np.log10(min(self.wdata)))
                            #self.rs.setUpper(np.log10(max(self.wdata)))
                            self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata))) #%f" % round_to_n(min(self.wdata), 6)).strip("0"))
                            self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata))) #%f" % round_to_n(max(self.wdata), 6)).strip("0"))
                            #self.wdata = self.wdataRaw.copy()
                            #self.rdata = self.rdataRaw.copy()
                            #self.jdata = self.jdataRaw.copy()
                        else:
                            self.upDelete = 0
                            self.lowDelete = 0
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                            self.wdata = self.wdataRaw.copy()
                            self.rdata = self.rdataRaw.copy()
                            self.jdata = self.jdataRaw.copy()
                            #self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                            #self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                            #self.rs.setNumberOfMajorTicks(10)
                            #self.rs.showMinorTicks(False)
                            #self.rs.setLower(np.log10(min(self.wdata)))
                            #self.rs.setUpper(np.log10(max(self.wdata)))
                            self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata))) #%f" % round_to_n(min(self.wdata), 6)).strip("0"))
                            self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata))) #%f" % round_to_n(max(self.wdata), 6)).strip("0"))
                        try:
                            self.figFreq.clear()
                            dataColor = "tab:blue"
                            deletedColor = "#A9CCE3"
                            if (self.topGUI.getTheme() == "dark"):
                                dataColor = "cyan"
                            else:
                                dataColor = "tab:blue"
                            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                                self.figFreq = Figure(figsize=(5, 5), dpi=100)
                                self.canvasFreq = FigureCanvasTkAgg(self.figFreq, master=self.freqWindow)
                                self.canvasFreq.get_tk_widget().grid(row=1,column=1, rowspan=5)
                                self.realFreq = self.figFreq.add_subplot(211)
                                self.realFreq.set_facecolor(self.backgroundColor)
                                self.realFreqDeletedLow, = self.realFreq.plot(self.wdataRaw[:self.lowDelete], self.rdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqDeletedHigh, = self.realFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqPlot, = self.realFreq.plot(self.wdata, self.rdata, "o", color=dataColor)
                                self.realFreq.set_xscale("log")
                                self.realFreq.get_xaxis().set_visible(False)
                                self.realFreq.set_title("Real Impedance", color=self.foregroundColor)
                                self.imagFreq = self.figFreq.add_subplot(212)
                                self.imagFreq.set_facecolor(self.backgroundColor)
                                self.imagFreqDeletedLow, = self.imagFreq.plot(self.wdataRaw[:self.lowDelete], -1*self.jdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqDeletedHigh, = self.imagFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], -1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqPlot, = self.imagFreq.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                                self.imagFreq.set_xscale("log")
                                self.imagFreq.set_title("Imaginary Impedance", color=self.foregroundColor)
                                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                                self.canvasFreq.draw()
                        except:
                            pass
                        self.lengthOfData = len(self.wdata)
                        self.tauDefault = 1/((max(w_in)-min(w_in))/(np.log10(abs(max(w_in)/min(w_in)))))
                        self.rDefault = (max(r_in)+min(r_in))/2
                        self.cDefault = 1/(self.jdata[0] * -2 * np.pi * min(w_in))
                        self.resultCap = self.cDefault
                        self.parametersButton.configure(state="normal")
                        self.freqRangeButton.configure(state="normal")
                        self.parametersLoadButton.configure(state="normal")
                        self.measureModelButton.configure(state="normal")
                        self.autoButton.configure(state="normal")
                        self.magicButton.configure(state="normal")
                        #self.saveCurrent.configure(state="normal")
                        self.numVoigtSpinbox.configure(state="readonly")
                        self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
                        self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtMinus.bind("<ButtonRelease-1>", minusNVE)
                        self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
                        self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
                        self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtPlus.bind("<ButtonRelease-1>", plusNVE)
                        self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
                        if (self.browseEntry.get() == ""):
                            self.paramComboboxVariables.append(tk.StringVar(self, "+ or -"))
                            self.tauComboboxVariables.append(tk.StringVar(self, "+"))
                            self.paramEntryVariables[0] = tk.StringVar(self, str(round(self.rDefault, -int(np.floor(np.log10(abs(self.rDefault)))) + (4 - 1))))
                            self.paramEntryVariables.append(tk.StringVar(self, str(round(self.rDefault, -int(np.floor(np.log10(abs(self.rDefault)))) + (4 - 1)))))
                            self.paramEntryVariables.append(tk.StringVar(self, str(round(self.tauDefault, -int(np.floor(np.log10(abs(self.tauDefault)))) + (4 - 1)))))
                            self.capacitanceEntryVariable.set(str(round(self.cDefault, -int(np.floor(np.log10(abs(self.cDefault)))) + (4 - 1))))
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, n)
                        self.browseEntry.configure(state="readonly")
                    except:
                        messagebox.showerror("File error", "Error 9:\nThere was an error loading or reading the file")
                elif (fext == ".mmfitting"):
                    try:
                        toLoad = open(n)
                        fileToLoad = toLoad.readline().split("filename: ")[1][:-1]
                        numberOfVoigt = int(toLoad.readline().split("number_voigt: ")[1][:-1])
                        if (numberOfVoigt < 1):
                            raise ValueError
                        numberOfSimulations = int(toLoad.readline().split("number_simulations: ")[1][:-1])
                        if (numberOfSimulations < 1):
                            raise ValueError
                        fitType = toLoad.readline().split("fit_type: ")[1][:-1]
                        if (fitType != "Real" and fitType != "Complex" and fitType != "Imaginary"):
                            raise ValueError
                        uD = int(toLoad.readline().split("upDelete: ")[1][:-1])
                        if (uD < 0):
                            raise ValueError
                        lD = int(toLoad.readline().split("lowDelete: ")[1][:-1])
                        if (lD < 0):
                            raise ValueError
                        weightType = toLoad.readline().split("weight_type: ")[1][:-1]
                        if (weightType != "Modulus" and weightType != "None" and weightType != "Proportional" and weightType != "Error model"):
                            raise ValueError
                        errorAlpha = toLoad.readline().split("error_alpha: ")[1][:-1]
                        errorBeta = toLoad.readline().split("error_beta: ")[1][:-1]
                        errorBetaRe = toLoad.readline().split("error_betaRe: ")[1][:-1]
                        errorGamma = toLoad.readline().split("error_gamma: ")[1][:-1]
                        errorDelta = toLoad.readline().split("error_delta: ")[1][:-1]
                        useAlpha = False
                        useBeta = False
                        useBetaRe = False
                        useGamma = False
                        useDelta = False
                        if (errorAlpha[0] != "n"):
                            errorAlpha = float(errorAlpha)
                            useAlpha = True
                        else:
                            errorAlpha = float(errorAlpha[1:])
                        if (errorBeta[0] != "n"):
                            errorBeta = float(errorBeta)
                            useBeta = True
                        else:
                            errorBeta = float(errorBeta[1:])
                        if (errorBetaRe[0] != "n" and useBeta):
                            errorBetaRe = float(errorBetaRe)
                            useBetaRe = True
                        else:
                            errorBetaRe = float(errorBetaRe[1:])
                        if (errorGamma[0] != "n"):
                            errorGamma = float(errorGamma)
                            useGamma = True
                        else:
                            errorGamma = float(errorGamma[1:])
                        if (errorDelta[0] != "n"):
                            errorDelta = float(errorDelta)
                            useDelta = True
                        else:
                            errorDelta = float(errorDelta[1:])
                        alphaValue = float(toLoad.readline().split("alpha: ")[1][:-1])
                        if (alphaValue <= 0):
                            raise ValueError
                        capValue = toLoad.readline().split("cap: ")[1][:-1]
                        if (capValue == "N"):
                            pass
                        else:
                            capValue = float(capValue)
                        capCombo = toLoad.readline().split("cap_combo: ")[1][:-1]
                        if (capCombo != "+" and capCombo != "+ or -" and capCombo != "-" and capCombo != "fixed"):
                            raise ValueError
                        toLoad.readline()
                        paramValues = []
                        while (True):
                            p = toLoad.readline()[:-1]
                            if ("paramComboboxes" in p):
                                break
                            else:
                                paramValues.append(float(p))
                        pComboboxes = []
                        while (True):
                            p = toLoad.readline()[:-1]
                            if ("tauComboboxes" in p):
                                break
                            else:
                                if (p != "+" and p != "+ or -" and p != "-" and p != "fixed"):
                                    raise ValueError
                                pComboboxes.append(p)
                        tComboboxes = []
                        while (True):
                            p = toLoad.readline()
                            if not p:
                                break
                            else:
                                p = p[:-1]
                                if (p != "+" and p != "fixed"):
                                    raise ValueError
                                tComboboxes.append(p)
                        toLoad.close()
                        data = np.loadtxt(fileToLoad)
                        w_in = data[:,0]
                        r_in = data[:,1]
                        j_in = data[:,2]
    #                        self.wdata = w_in
    #                        self.rdata = r_in
    #                        self.jdata = j_in
                        self.wdataRaw = w_in
                        self.rdataRaw = r_in
                        self.jdataRaw = j_in
                        #---Bubble sort the data---
                        doneSorting = False
                        while (not doneSorting):
                            doneSorting = True
                            for i in range(len(self.wdataRaw)-1):
                                if (self.wdataRaw[i] > self.wdataRaw[i+1]):
                                    temp = self.wdataRaw[i]
                                    self.wdataRaw[i] = self.wdataRaw[i+1]
                                    self.wdataRaw[i+1] = temp
                                    tempR = self.rdataRaw[i]
                                    self.rdataRaw[i] = self.rdataRaw[i+1]
                                    self.rdataRaw[i+1] = tempR
                                    tempJ = self.jdataRaw[i]
                                    self.jdataRaw[i] = self.jdataRaw[i+1]
                                    self.jdataRaw[i+1] = tempJ
                                    doneSorting = False
                        self.upDelete = uD
                        self.lowDelete = lD
                        self.lowerSpinboxVariable.set(str(self.lowDelete))
                        self.upperSpinboxVariable.set(str(self.upDelete))
                        if (self.upDelete == 0):
                            self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                            self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                            self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                        else:
                            self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                            self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                            self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                        #self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                        #self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                        #self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                        #self.rs.setNumberOfMajorTicks(10)
                        #self.rs.showMinorTicks(False)
                        #self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                        #self.rs.setLower(np.log10(min(self.wdata)))
                        #self.rs.setUpper(np.log10(max(self.wdata)))
                        try:
                            self.figFreq.clear()
                            dataColor = "tab:blue"
                            deletedColor = "#A9CCE3"
                            if (self.topGUI.getTheme() == "dark"):
                                dataColor = "cyan"
                            else:
                                dataColor = "tab:blue"
                            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                                self.figFreq = Figure(figsize=(5, 5), dpi=100)
                                self.canvasFreq = FigureCanvasTkAgg(self.figFreq, master=self.freqWindow)
                                self.canvasFreq.get_tk_widget().grid(row=1,column=1, rowspan=5)
                                self.realFreq = self.figFreq.add_subplot(211)
                                self.realFreq.set_facecolor(self.backgroundColor)
                                self.realFreqDeletedLow, = self.realFreq.plot(self.wdataRaw[:self.lowDelete], self.rdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqDeletedHigh, = self.realFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqPlot, = self.realFreq.plot(self.wdata, self.rdata, "o", color=dataColor)
                                self.realFreq.set_xscale("log")
                                self.realFreq.get_xaxis().set_visible(False)
                                self.realFreq.set_title("Real Impedance", color=self.foregroundColor)
                                self.imagFreq = self.figFreq.add_subplot(212)
                                self.imagFreq.set_facecolor(self.backgroundColor)
                                self.imagFreqDeletedLow, = self.imagFreq.plot(self.wdataRaw[:self.lowDelete], -1*self.jdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqDeletedHigh, = self.imagFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], -1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqPlot, = self.imagFreq.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                                self.imagFreq.set_xscale("log")
                                self.imagFreq.set_title("Imaginary Impedance", color=self.foregroundColor)
                                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                                self.canvasFreq.draw()
                        except:
                            pass
                        self.lengthOfData = len(self.wdata)
                        self.tauDefault = 1/((max(w_in)-min(w_in))/(np.log10(abs(max(w_in)/min(w_in)))))
                        self.rDefault = (max(r_in)+min(r_in))/2
                        self.cDefault = 1/(self.jdata[0] * -2 * np.pi * min(w_in))
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, fileToLoad)
                        self.browseEntry.configure(state="readonly")
                        self.parametersButton.configure(state="normal")
                        self.parametersLoadButton.configure(state="normal")
                        self.measureModelButton.configure(state="normal")
                        self.freqRangeButton.configure(state="normal")
                        self.autoButton.configure(state="normal")
                        self.magicButton.configure(state="normal")
                        #self.saveCurrent.configure(state="normal")
                        self.numMonteVariable.set(str(numberOfSimulations))
                        self.alphaVariable.set(str(alphaValue))
                        self.fitTypeVariable.set(fitType)
                        self.weightingVariable.set(weightType)
                        if (weightType == "Error model"):
                            if (not useAlpha):
                                self.errorAlphaCheckboxVariable.set(0)
                                self.errorAlphaEntry.configure(state="disabled")
                                self.errorAlphaVariable.set(errorAlpha)
                            else:
                                self.errorAlphaCheckboxVariable.set(1)
                                self.errorAlphaEntry.configure(state="normal")
                                self.errorAlphaVariable.set(errorAlpha)
                            if (not useBeta):
                                self.errorBetaCheckboxVariable.set(0)
                                self.errorBetaEntry.configure(state="disabled")
                                self.errorBetaVariable.set(errorBeta)
                            else:
                                self.errorBetaCheckboxVariable.set(1)
                                self.errorBetaEntry.configure(state="normal")
                                self.errorBetaVariable.set(errorBeta)
                            if (not useBetaRe):
                                self.errorBetaReCheckboxVariable.set(0)
                                self.errorBetaReEntry.configure(state="disabled")
                                self.errorBetaReVariable.set(errorBetaRe)
                            else:
                                self.errorBetaReCheckboxVariable.set(1)
                                self.errorBetaReCheckbox.configure(state="normal")
                                self.errorBetaReEntry.configure(state="normal")
                                self.errorBetaReVariable.set(errorBetaRe)
                            if (not useGamma):
                                self.errorGammaCheckboxVariable.set(0)
                                self.errorGammaEntry.configure(state="disabled")
                                self.errorGammaVariable.set(errorGamma)
                            else:
                                self.errorGammaCheckboxVariable.set(1)
                                self.errorGammaEntry.configure(state="normal")
                                self.errorGammaVariable.set(errorGamma)
                            if (not useDelta):
                                self.errorDeltaCheckboxVariable.set(0)
                                self.errorDeltaEntry.configure(state="disabled")
                                self.errorDeltaVariable.set(errorDelta)
                            else:
                                self.errorDeltaCheckboxVariable.set(1)
                                self.errorDeltaEntry.configure(state="normal")
                                self.errorDeltaVariable.set(errorDelta)
                        if (self.weightingVariable.get() == "None"):
                            self.alphaLabel.grid_remove()
                            self.alphaEntry.grid_remove()
                            self.errorStructureFrame.grid_remove()
                        elif (self.weightingVariable.get() == "Error model"):
                            self.alphaLabel.grid_remove()
                            self.alphaEntry.grid_remove()
                            self.errorStructureFrame.grid(column=0, row=1, pady=5, sticky="W", columnspan=5)
                            checkErrorStructure()
                        else:
                            self.alphaLabel.grid(column=4, row=0)
                            self.alphaEntry.grid(column=5, row=0)
                            self.errorStructureFrame.grid_remove()
                        self.numVoigtVariable.set(numberOfVoigt)
                        self.numVoigtSpinbox.configure(state="readonly")
                        self.numVoigtTextVariable.set(self.numVoigtVariable.get())
                        self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
                        self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtMinus.bind("<ButtonRelease-1>", minusNVE)
                        self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
                        self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
                        self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                        self.numVoigtPlus.bind("<ButtonRelease-1>", plusNVE)
                        self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
                        self.nVoigt = numberOfVoigt
                        self.paramComboboxVariables.clear()
                        self.tauComboboxVariables.clear()
                        self.paramEntryVariables.clear()
                        for pvar in pComboboxes:
                            self.paramComboboxVariables.append(tk.StringVar(self, pvar))
                        for tvar in tComboboxes:
                            self.tauComboboxVariables.append(tk.StringVar(self, tvar))
                        self.paramEntryVariables.append(tk.StringVar(self, str(paramValues[0])))
                        for pval in paramValues[1:]:
                            self.paramEntryVariables.append(tk.StringVar(self, str(pval)))
                        if (capValue == "N"):
                            self.capacitanceCheckboxVariable.set(0)
                            self.capacitanceEntryVariable.set(str(round(self.cDefault, -int(np.floor(np.log10(abs(self.cDefault)))) + (4 - 1))))
                            self.capacitanceComboboxVariable.set("+ or -")
                        else:
                            self.capacitanceCheckboxVariable.set(1)
                            self.capacitanceEntryVariable.set(str(capValue))
                            self.capacitanceComboboxVariable.set(capCombo)
                    except:
                        messagebox.showerror("File error", "Error 9:\nThere was an error loading or reading the file")
                else:
                    messagebox.showerror("File error", "Error 10:\nThe file has an unknown extension")

        def validateMC(P):
            if (P == ''):
                return False
            if ' ' in P:
                return False
            if '\t' in P:
                return False
            if '\n' in P:
                return False
            try:
                int(P)
                if (int(P) < 1):
                    return False
                else:
                    return True
            except:
                return False
            
        def validateFreqLow(P):
            if (P == ''):
                return True
            if ' ' in P:
                return False
            if '\t' in P:
                return False
            if '\n' in P:
                return False
            try:
                int(P)
                if (int(P) < 0):
                    return False
                elif (self.upperSpinboxVariable.get() == ''):
                    if (int(P) > len(self.wdataRaw)):
                        return False
                    else:
                        return True
                elif (int(P) + int(self.upperSpinboxVariable.get()) > len(self.wdataRaw)-1):
                    return False
                else:
                    return True
            except:
                return False
            
        def validateFreqHigh(P):
            if (P == ''):
                return True
            if ' ' in P:
                return False
            if '\t' in P:
                return False
            if '\n' in P:
                return False
            try:
                int(P)
                if (int(P) < 0):
                    return False
                elif (self.lowerSpinboxVariable.get() == ''):
                    if (int(P) > len(self.wdataRaw)):
                        return False
                    else:
                        return True
                elif (int(P) + int(self.lowerSpinboxVariable.get()) > len(self.wdataRaw)-1):
                    return False
                else:
                    return True
            except:
                return False
            
        def validateEither(P):
            if (P == ''):
                return True
            if (P == '.'):
                return True
            if (P == '-'):
                return True
            if ' ' in P:
                return False
            if '\t' in P:
                return False
            if '\n' in P:
                return False
            if ((P.count("E") == 1 and P.count("e") == 0) or (P.count("e") == 1 and P.count("E") == 0)):
                if (P[0] != "E" and P[0] != "e"):
                    if (P.count("-") == 0 and P.count("+") == 0):
                        return True
                    else:
                        if ((P.count("-") == 1 and P.count("+") == 0) or (P.count("+") == 1 and P.count("-") == 0)):
                            try:
                                if (P.count("-") == 1):
                                    if (P[P.find("-")-1] == "E" or P[P.find("-")-1] == "e"):
                                        return True
                                elif (P.count("+") == 1):
                                    if (P[P.find("+")-1] == "E" or P[P.find("+")-1] == "e"):
                                        return True
                            except:
                                return False                   
            try:
                float(P)
                return True
            except:
                return False
        
        valcom = (self.register(validateEither), '%P')
        valfreqlow = (self.register(validateFreqLow), '%P')
        valfreqhigh = (self.register(validateFreqHigh), '%P')
        
        def validateFreqHigh2(P):
            if (P == ""):
                return True
            try:
                val = int(P)
            except:
                return False
            if (val < 0):
                return False
            elif (val + self.lowDelete > len(self.wdataRaw)):
                return False
            return True
        valfreqHigh2 = (self.register(validateFreqHigh2), '%P')
        
        def validateFreqLow2(P):
            if (P == ""):
                return True
            try:
                val = int(P)
            except:
                return False
            if (val < 0):
                return False
            elif (val + self.upDelete > len(self.wdataRaw)):
                return False
            return True
        valfreqLow2 = (self.register(validateFreqLow2), '%P')
        
        def removeElement(element_number):
            #---Copy each variable from "above" down by 1 (2 paramEntryVariables, 1 paramComboboxVariables, 1 tauComboboxVariables)---
            for i in range(element_number, int(self.numVoigtVariable.get())):
                self.paramComboboxVariables[i].set(self.paramComboboxVariables[i+1].get())
                self.tauComboboxVariables[i-1].set(self.tauComboboxVariables[i].get())
                self.paramEntryVariables[i*2 - 1].set(self.paramEntryVariables[i*2 - 1 + 2].get())
                self.paramEntryVariables[i*2].set(self.paramEntryVariables[i*2 + 2].get())
            for i in range(2*element_number-1, len(self.multistartEnabled)-2, 2):
                self.multistartEnabled[i].set(self.multistartEnabled[i+2].get())
                self.multistartEnabled[i+1].set(self.multistartEnabled[i+3].get())
                self.multistartLower[i].set(self.multistartLower[i+2].get())
                self.multistartLower[i+1].set(self.multistartLower[i+3].get())
                self.multistartUpper[i].set(self.multistartUpper[i+2].get())
                self.multistartUpper[i+1].set(self.multistartUpper[i+3].get())
                self.multistartNumber[i].set(self.multistartNumber[i+2].get())
                self.multistartNumber[i+1].set(self.multistartNumber[i+3].get())
                self.multistartSelection[i].set(self.multistartSelection[i+2].get())
                self.multistartSelection[i+1].set(self.multistartSelection[i+3].get())
                self.multistartCustom[i].set(self.multistartCustom[i+2].get())
                self.multistartCustom[i+1].set(self.multistartCustom[i+3].get())
            self.numVoigtSpinbox.invoke("buttondown")

        def changeNVE():
            num_VE = int(self.numVoigtVariable.get())
            self.numVoigtTextVariable.set(self.numVoigtVariable.get())
            defaultRGuess = 0
            defaultTauGuess = 0
            if (len(self.fits) <= 3):
                defaultRGuess = self.rDefault
                defaultTauGuess = self.tauDefault
            else:
                for i in range(1, len(self.fits), 2):
                    defaultRGuess += self.fits[i]
                    defaultTauGuess += np.log10(self.fits[i+1])
                defaultRGuess /= (len(self.fits)-1)/2
                defaultTauGuess /= (len(self.fits)-1)/2
                defaultTauGuess = 10**defaultTauGuess
            if (self.paramPopup.state() != "withdrawn"):
                self.howMany.configure(text="Number of elements: " + self.numVoigtVariable.get())
                if (self.nVoigt > num_VE):
                    self.paramComboboxes.pop().grid_remove()
                    self.tauComboboxes.pop().grid_remove()
                    self.paramLabels.pop().grid_remove()
                    self.paramLabels.pop().grid_remove()
                    self.paramEntries.pop().grid_remove()
                    self.paramEntries.pop().grid_remove()
                    self.paramFrames.pop().grid_remove()
                    self.paramDeleteButtons.pop().grid_remove()
                    self.paramComboboxVariables.pop()
                    self.tauComboboxVariables.pop()
                    self.paramEntryVariables.pop()
                    self.paramEntryVariables.pop()
                    if (num_VE == 1 and len(self.paramDeleteButtons) == 1):
                        self.paramDeleteButtons.pop().grid_remove()
                elif (self.nVoigt < num_VE):
                    self.paramEntryVariables.append(tk.StringVar(self, str(round(defaultRGuess, -int(np.floor(np.log10(defaultRGuess))) + (4 - 1)))))
                    self.paramEntryVariables.append(tk.StringVar(self, str(round(defaultTauGuess, -int(np.floor(np.log10(defaultTauGuess))) + (4 - 1)))))
                    #self.paramFrames.append(tk.LabelFrame(self.pframe, text="Element " + str(num_VE), bg="white", padx=4, pady=4))
                    self.paramFrames.append(tk.Label(self.pframe, text="Element " + str(num_VE) + ": ", bg=self.backgroundColor, fg=self.foregroundColor))
                    self.paramLabels.append(tk.Label(self.pframe, text="R ", bg=self.backgroundColor, fg=self.foregroundColor))
                    self.paramLabels.append(tk.Label(self.pframe, text="  Tau ", bg=self.backgroundColor, fg=self.foregroundColor))
                    self.paramEntries.append(ttk.Entry(self.pframe, textvariable=self.paramEntryVariables[len(self.paramEntryVariables)-2], width=10))
                    self.paramEntries.append(ttk.Entry(self.pframe, textvariable=self.paramEntryVariables[len(self.paramEntryVariables)-1], width=10))
                    self.paramComboboxVariables.append(tk.StringVar(self, "+ or -"))
                    self.tauComboboxVariables.append(tk.StringVar(self, "+"))
                    self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxVariables[len(self.paramComboboxVariables)-1], value=("+", "+ or -", "-", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                    self.tauComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.tauComboboxVariables[len(self.tauComboboxVariables)-1], value=("+", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                    self.paramFrames[len(self.paramFrames)-1].grid(column=0, row=num_VE, sticky="W", pady=(0,2))
                    self.paramComboboxes[len(self.paramComboboxes)-1].grid(column=2, row=num_VE, sticky="W", padx=(0,2), pady=(0,2))
                    self.paramLabels[len(self.paramLabels)-2].grid(column=1, row=num_VE, pady=(0,2))
                    self.paramEntries[len(self.paramEntries)-2].grid(column=3, row=num_VE, sticky="W", pady=(0,2))
                    self.tauComboboxes[len(self.tauComboboxes)-1].grid(column=5, row=num_VE, padx=(0,2), pady=(0,2))
                    self.paramLabels[len(self.paramLabels)-1].grid(column=4, row=num_VE, pady=(0,2))
                    self.paramEntries[len(self.paramEntries)-1].grid(column=6, row=num_VE, pady=(0,2))
                    self.paramLabels[len(self.paramLabels)-2].bind("<MouseWheel>", _on_mousewheel)
                    self.paramEntries[len(self.paramEntries)-1].bind("<MouseWheel>", _on_mousewheel)
                    self.paramLabels[len(self.paramLabels)-1].bind("<MouseWheel>", _on_mousewheel)
                    self.paramEntries[len(self.paramEntries)-2].bind("<MouseWheel>", _on_mousewheel)
                    self.paramFrames[len(self.paramFrames)-1].bind("<MouseWheel>", _on_mousewheel)
                    if (self.nVoigt == 1):
                        self.paramDeleteButtons.append(ttk.Button(self.pframe, text="Remove", command=partial(removeElement, 1)))
                        self.paramDeleteButtons[len(self.paramDeleteButtons)-1].grid(column=7, row=1, sticky="E", pady=(0,2), padx=(2, 0))
                        self.paramDeleteButtons[len(self.paramDeleteButtons)-1].bind("<MouseWheel>", _on_mousewheel)
                    self.paramDeleteButtons.append(ttk.Button(self.pframe, text="Remove", command=partial(removeElement, num_VE)))
                    self.paramDeleteButtons[len(self.paramDeleteButtons)-1].grid(column=7, row=num_VE, sticky="E", pady=(0,2), padx=(2, 0))
                    self.paramDeleteButtons[len(self.paramDeleteButtons)-1].bind("<MouseWheel>", _on_mousewheel)
                    
            else:
                if (self.nVoigt > num_VE):
                    self.paramComboboxVariables.pop()
                    self.tauComboboxVariables.pop()
                    self.paramEntryVariables.pop()
                    self.paramEntryVariables.pop()
                elif (self.nVoigt < num_VE):
                    self.paramComboboxVariables.append(tk.StringVar(self, "+ or -"))
                    self.tauComboboxVariables.append(tk.StringVar(self, "+"))
                    self.paramEntryVariables.append(tk.StringVar(self, str(round(defaultRGuess, -int(np.floor(np.log10(abs(defaultRGuess)))) + (4 - 1)))))
                    self.paramEntryVariables.append(tk.StringVar(self, str(round(defaultTauGuess, -int(np.floor(np.log10(defaultTauGuess))) + (4 - 1)))))
            if (self.advancedOptionsPopup.state() != "withdrawn"):
                if (self.nVoigt < num_VE):
                    for i in range(2):  #Once for R, once for T
                        self.multistartEnabled.append(tk.IntVar(self, 0))
                        self.multistartSelection.append(tk.StringVar(self, "Logarithmic"))
                        self.multistartLower.append(tk.StringVar(self, "1E-5"))
                        self.multistartUpper.append(tk.StringVar(self, "1E5"))
                        self.multistartNumber.append(tk.StringVar(self, "10"))
                        self.multistartCustom.append(tk.StringVar(self, "0,1,10,100,1000"))
                    self.paramListbox.insert(tk.END, "R" + str(num_VE))
                    self.paramListbox.insert(tk.END, "T" + str(num_VE))
                elif (self.nVoigt > num_VE):
                    for i in range(2):  #Once for R, once for T
                        self.multistartEnabled.pop()
                        self.multistartSelection.pop()
                        self.multistartLower.pop()
                        self.multistartUpper.pop()
                        self.multistartNumber.pop()
                        self.multistartCustom.pop()
                    self.paramListbox.delete(tk.END)
                    self.paramListbox.delete(tk.END)
                    self.paramListbox.selection_clear(0, tk.END)
                    self.paramListbox.select_set(0)
                    self.advancedChoiceLabel.configure(text="Re")
                    self.paramListbox.event_generate("<<ListboxSelect>>")
                    self.advCheckbox.configure(variable=self.multistartEnabled[0])
                    self.advSelection.configure(textvariable=self.multistartSelection[0])
                    self.advUpperEntry.configure(textvariable=self.multistartUpper[0])
                    self.advLowerEntry.configure(textvariable=self.multistartLower[0])
                    self.advNumberEntry.configure(textvariable=self.multistartNumber[0])
                    self.advCustomEntry.configure(textvariable=self.multistartCustom[0])
                    if(self.multistartEnabled[0].get()):
                        self.advSelection.configure(state="readonly")
                        self.advUpperEntry.configure(state="normal")
                        self.advLowerEntry.configure(state="normal")
                        self.advNumberEntry.configure(state="normal")
                        self.advCustomEntry.configure(state="normal")
                    else:
                        self.advSelection.configure(state="disabled")
                        self.advUpperEntry.configure(state="disabled")
                        self.advLowerEntry.configure(state="disabled")
                        self.advNumberEntry.configure(state="disabled")
                        self.advCustomEntry.configure(state="disabled")
                    if (self.multistartSelection[0].get() == "Custom"):
                        self.advLowerLabel.grid_remove()
                        self.advLowerEntry.grid_remove()
                        self.advUpperLabel.grid_remove()
                        self.advUpperEntry.grid_remove()
                        self.advNumberLabel.grid_remove()
                        self.advNumberEntry.grid_remove()
                        self.advCustomLabel.grid(column=0, row=2, pady=5, sticky="W")
                        self.advCustomEntry.grid(column=0, row=3, pady=5, columnspan=2, sticky="W")
                    else:
                        self.advCustomLabel.grid_remove()
                        self.advCustomEntry.grid_remove()
                        self.advLowerLabel.grid(column=0, row=2, pady=5, sticky="W")
                        self.advLowerEntry.grid(column=1, row=2, pady=5, sticky="W")
                        self.advUpperLabel.grid(column=0, row=3, sticky="W")
                        self.advUpperEntry.grid(column=1, row=3, sticky="W")
                        self.advNumberLabel.grid(column=0, row=4, pady=5, sticky="W")
                        self.advNumberEntry.grid(column=1, row=4, pady=5, sticky="W")
            self.nVoigt = num_VE
        
        def capCommand():
            if (self.capacitanceCheckboxVariable.get() == 0):
                self.capacitanceCombobox.configure(state="disabled")
                self.capacitanceEntry.configure(state="disabled")
            else:
                self.capacitanceCombobox.configure(state="readonly")
                self.capacitanceEntry.configure(state="normal")
            
        def _on_mousewheel(event):
            xpos, ypos = self.vsb.get()
            xpos_round = round(xpos, 2)     #Avoid floating point errors
            ypos_round = round(ypos, 2)
            if (xpos_round != 0.00 or ypos_round != 1.00):
                self.pcanvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def paramsPopup():
            
            if (self.paramPopup.state() == "withdrawn"):
                num_VE = int(self.numVoigtVariable.get())
                self.pcanvas = tk.Canvas(self.paramPopup, borderwidth=0, highlightthickness=0, background=self.backgroundColor)
                self.pframe = tk.Frame(self.pcanvas, background=self.backgroundColor)
                self.vsb = tk.Scrollbar(self.paramPopup, orient="vertical", command=self.pcanvas.yview)
                def populate(frame):
                    self.reFrame = tk.Frame(frame, bg=self.backgroundColor)
                    self.reLabel = tk.Label(self.reFrame, text="Re (Rsol) = ", bg=self.backgroundColor, fg=self.foregroundColor)
                    reCombobox = ttk.Combobox(self.reFrame, value=("+", "+ or -", "fixed"),  textvariable=self.paramComboboxVariables[0], justify=tk.CENTER, state="readonly", width=6)
                    reEntry = ttk.Entry(self.reFrame, width=10, textvariable=self.paramEntryVariables[0], validate="all", validatecommand=valcom)
                    self.capacitanceCheckbox = ttk.Checkbutton(self.reFrame, text="Capacitance", variable=self.capacitanceCheckboxVariable, command=capCommand)
                    self.capacitanceCombobox = ttk.Combobox(self.reFrame, value=("+", "+ or -", "-", "fixed"), textvariable=self.capacitanceComboboxVariable, justify=tk.CENTER, state="disabled", exportselection=0, width=6)
                    self.capacitanceEntry = ttk.Entry(self.reFrame, width=10, textvariable=self.capacitanceEntryVariable, state="disabled", validate="all", validatecommand=valcom)
                    if (self.capacitanceCheckboxVariable.get() == 1):
                        self.capacitanceCombobox.configure(state="readonly")
                        self.capacitanceEntry.configure(state="normal")
                    reCombobox.grid(column=1, row=0, padx=5)
                    self.reLabel.grid(column=0, row=0)
                    reEntry.grid(column=2, row=0)
                    self.capacitanceCheckbox.grid(column=3, row=0, padx=(15, 5))
                    self.capacitanceCombobox.grid(column=4, row=0)
                    self.capacitanceEntry.grid(column=5, row=0, padx=5)
                    self.reFrame.grid(column=0, row=0, pady=(1, 8), columnspan=10)
                    for i in range(num_VE-(len(self.tauComboboxVariables))):
                        self.tauComboboxVariables.append(tk.StringVar(self, "+"))
                    for i in range(num_VE-(len(self.paramComboboxVariables)-1)):
                        self.paramComboboxVariables.append(tk.StringVar(self, "+ or -"))
                    for i in range(0, num_VE*2-(len(self.paramEntryVariables)-1), 2):
                        self.paramEntryVariables.append(tk.StringVar(self, str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1)))))
                        self.paramEntryVariables.append(tk.StringVar(self, str(round(self.tauDefault, -int(np.floor(np.log10(self.tauDefault))) + (4 - 1)))))
                    for i in range(num_VE):
                        #self.paramFrames.append(tk.LabelFrame(frame, bg="white", padx=4, pady=4, text="Element " + str(i+1)))
                        self.paramFrames.append(tk.Label(frame, text="Element " + str(i+1) + ":  ", bg=self.backgroundColor, fg=self.foregroundColor))
                        self.paramLabels.append(tk.Label(frame, text="R ", bg=self.backgroundColor, fg=self.foregroundColor))
                        self.paramLabels.append(tk.Label(frame, text="  Tau ", bg=self.backgroundColor, fg=self.foregroundColor))
                        self.paramEntries.append(ttk.Entry(frame, textvariable=self.paramEntryVariables[i*2+1], width=10))
                        self.paramEntries.append(ttk.Entry(frame, textvariable=self.paramEntryVariables[i*2+2], width=10))
                        self.paramComboboxes.append(ttk.Combobox(frame, textvariable=self.paramComboboxVariables[i+1], value=("+", "+ or -", "-", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                        self.tauComboboxes.append(ttk.Combobox(frame, textvariable=self.tauComboboxVariables[i], value=("+", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                        self.paramFrames[i].grid(column=0, row=i+1, sticky="W", pady=(0,2))
                        self.paramComboboxes[len(self.paramComboboxes)-1].grid(column=2, row=i+1, sticky="W", padx=(0, 2), pady=(0,2))
                        self.paramLabels[len(self.paramLabels)-2].grid(column=1, row=i+1, pady=(0,2))
                        self.paramEntries[len(self.paramEntries)-2].grid(column=3, row=i+1, sticky="W", pady=(0,2))
                        self.tauComboboxes[len(self.tauComboboxes)-1].grid(column=5, row=i+1, padx=(0,2), pady=(0,2))
                        self.paramLabels[len(self.paramLabels)-1].grid(column=4, row=i+1, pady=(0,2))
                        self.paramEntries[len(self.paramEntries)-1].grid(column=6, row=i+1, pady=(0,2))
                        if (i != 0 or (i == 0 and int(self.numVoigtVariable.get()) > 1)):
                            self.paramDeleteButtons.append(ttk.Button(frame, text="Remove", command=partial(removeElement, i+1)))
                            self.paramDeleteButtons[len(self.paramDeleteButtons)-1].grid(column=7, row=i+1, sticky="E", pady=(0,2), padx=(2, 0))
                            self.paramDeleteButtons[len(self.paramDeleteButtons)-1].bind("<MouseWheel>", _on_mousewheel)
                        self.paramLabels[len(self.paramLabels)-2].bind("<MouseWheel>", _on_mousewheel)
                        self.paramEntries[len(self.paramEntries)-1].bind("<MouseWheel>", _on_mousewheel)
                        self.paramLabels[len(self.paramLabels)-1].bind("<MouseWheel>", _on_mousewheel)
                        self.paramEntries[len(self.paramEntries)-2].bind("<MouseWheel>", _on_mousewheel)
                        self.paramFrames[len(self.paramFrames)-1].bind("<MouseWheel>", _on_mousewheel)
            
                def onFrameConfigure(canvas):
                    #Reset the scroll region to encompass the inner frame
                    canvas.configure(scrollregion=canvas.bbox("all"))
                
                def onClose():
                    self.paramPopup.withdraw()
                    clearAll()
    
                self.paramPopup.title("Model parameters")
                self.paramPopup.iconbitmap(resource_path("img/elephant3.ico"))
                self.paramPopup.geometry("700x400")
                self.paramPopup.deiconify()
                self.paramPopup.protocol("WM_DELETE_WINDOW", onClose)
                #canvas = tk.Canvas(self.paramPopup, borderwidth=0, highlightthickness=0, background=self.backgroundColor)
                #frame = tk.Frame(canvas, background="white")
                #vsb = tk.Scrollbar(self.paramPopup, orient="vertical", command=self.pcanvas.yview)
                self.pcanvas.configure(yscrollcommand=self.vsb.set)
                self.pcanvas.bind("<MouseWheel>", _on_mousewheel)
                self.pframe.bind("<MouseWheel>", _on_mousewheel)
    
                self.buttonFrame = tk.Frame(self.paramPopup, bg=self.backgroundColor)
                self.addButton = ttk.Button(self.buttonFrame, text="Add Element", command=lambda: self.numVoigtSpinbox.invoke("buttonup"))
                self.removeButton = ttk.Button(self.buttonFrame, text="Remove Last Element", command= lambda: self.numVoigtSpinbox.invoke("buttondown"))
                addButton_ttp = CreateToolTip(self.addButton, "Add a Voigt element")
                removeButton_ttp = CreateToolTip(self.removeButton, "Remove a Voigt element")
                self.howMany = tk.Label(self.paramPopup, text="Number of elements: " + str(num_VE), bg=self.backgroundColor, fg=self.foregroundColor)
                self.advancedOptions = ttk.Button(self.paramPopup, text="Advanced options", command=advancedParamsPopup)
                advancedOptions_ttp = CreateToolTip(self.advancedOptions, "Advanced parameter options, including multistart")
                self.addButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
                self.removeButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
                self.advancedOptions.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
                self.howMany.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
                self.buttonFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=False, pady=1)
                
                self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
                self.pcanvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                self.pcanvas.create_window((4,4), window=self.pframe, anchor="nw")
                
                self.pframe.bind("<Configure>", lambda event, canvas=self.pcanvas: onFrameConfigure(canvas))
                
                populate(self.pframe)
                
                def clearAll():
                    self.pframe.destroy()
                    self.pcanvas.destroy()
                    self.vsb.destroy()
                    #vsb.destroy()
                    self.addButton.destroy()
                    self.removeButton.destroy()
                    self.buttonFrame.destroy()
                    self.howMany.destroy()
                    self.advancedOptions.destroy()
                    self.paramComboboxes.clear()
                    self.tauComboboxes.clear()
                    self.paramLabels.clear()
                    self.paramEntries.clear()
                    self.paramFrames.clear()
                    self.paramDeleteButtons.clear()
                    if (self.paramEntryVariables[0].get() == ''):
                        self.paramEntryVariables[0].set(str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1))))
                    for i in range(1, len(self.paramEntryVariables), 2):
                        if (self.paramEntryVariables[i].get() == ''):
                            self.paramEntryVariables[i].set(str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1))))
                        if (self.paramEntryVariables[i+1].get() == ''):
                            self.paramEntryVariables[i+1].set(str(round(self.tauDefault, -int(np.floor(np.log10(self.tauDefault))) + (4 - 1))))
            else:
                self.paramPopup.deiconify()
                self.paramPopup.lift()  
        
        def advancedParamsPopup():
            def onCloseAdvanced():
                self.advancedOptionsPopup.withdraw()
                self.paramListbox.delete(0, tk.END)
                self.advCustomLabel.destroy()
                self.advCustomEntry.destroy()
                self.advCheckbox.destroy()
                self.advSelectionLabel.destroy()
                self.advSelection.destroy()
                self.advLowerLabel.destroy()
                self.advLowerEntry.destroy()
                self.advUpperLabel.destroy()
                self.advUpperEntry.destroy()
                self.advNumberLabel.destroy()
                self.advNumberEntry.destroy()
                
            def onSelect(e):
                num_selected = self.paramListbox.curselection()[0]
                self.advCheckbox.configure(variable=self.multistartEnabled[num_selected])
                self.advSelection.configure(textvariable=self.multistartSelection[num_selected])
                self.advUpperEntry.configure(textvariable=self.multistartUpper[num_selected])
                self.advLowerEntry.configure(textvariable=self.multistartLower[num_selected])
                self.advNumberEntry.configure(textvariable=self.multistartNumber[num_selected])
                self.advCustomEntry.configure(textvariable=self.multistartCustom[num_selected])
                self.advancedChoiceLabel.configure(text=self.paramListbox.get(0, num_selected)[-1])
                if (self.multistartEnabled[num_selected].get()):
                    self.advSelection.configure(state="readonly")
                    self.advUpperEntry.configure(state="normal")
                    self.advLowerEntry.configure(state="normal")
                    self.advNumberEntry.configure(state="normal")
                    self.advCustomEntry.configure(state="normal")
                else:
                    self.advSelection.configure(state="disabled")
                    self.advUpperEntry.configure(state="disabled")
                    self.advLowerEntry.configure(state="disabled")
                    self.advNumberEntry.configure(state="disabled")
                    self.advCustomEntry.configure(state="disabled")
                if (self.multistartSelection[num_selected].get() == "Custom"):
                    self.advLowerLabel.grid_remove()
                    self.advLowerEntry.grid_remove()
                    self.advUpperLabel.grid_remove()
                    self.advUpperEntry.grid_remove()
                    self.advNumberLabel.grid_remove()
                    self.advNumberEntry.grid_remove()
                    self.advCustomLabel.grid(column=0, row=2, pady=5, sticky="W")
                    self.advCustomEntry.grid(column=0, row=3, pady=5, columnspan=2, sticky="W")
                else:
                    self.advCustomLabel.grid_remove()
                    self.advCustomEntry.grid_remove()
                    self.advLowerLabel.grid(column=0, row=2, pady=5, sticky="W")
                    self.advLowerEntry.grid(column=1, row=2, pady=5, sticky="W")
                    self.advUpperLabel.grid(column=0, row=3, sticky="W")
                    self.advUpperEntry.grid(column=1, row=3, sticky="W")
                    self.advNumberLabel.grid(column=0, row=4, pady=5, sticky="W")
                    self.advNumberEntry.grid(column=1, row=4, pady=5, sticky="W")
            
            def advCheck():
                num_selected = self.paramListbox.curselection()[0]
                if (self.multistartEnabled[num_selected].get()):
                    self.advSelection.configure(state="readonly")
                    self.advUpperEntry.configure(state="normal")
                    self.advLowerEntry.configure(state="normal")
                    self.advNumberEntry.configure(state="normal")
                    self.advCustomEntry.configure(state="normal")
                else:
                    self.advSelection.configure(state="disabled")
                    self.advUpperEntry.configure(state="disabled")
                    self.advLowerEntry.configure(state="disabled")
                    self.advNumberEntry.configure(state="disabled")
                    self.advCustomEntry.configure(state="disabled")
            
            def advSelectionChange(e):
                num_selected = self.paramListbox.curselection()[0]
                if (self.multistartSelection[num_selected].get() == "Custom"):
                    self.advLowerLabel.grid_remove()
                    self.advLowerEntry.grid_remove()
                    self.advUpperLabel.grid_remove()
                    self.advUpperEntry.grid_remove()
                    self.advNumberLabel.grid_remove()
                    self.advNumberEntry.grid_remove()
                    self.advCustomLabel.grid(column=0, row=2, pady=5, sticky="W")
                    self.advCustomEntry.grid(column=0, row=3, pady=5, columnspan=2, sticky="W")
                else:
                    self.advCustomLabel.grid_remove()
                    self.advCustomEntry.grid_remove()
                    self.advLowerLabel.grid(column=0, row=2, pady=5, sticky="W")
                    self.advLowerEntry.grid(column=1, row=2, pady=5, sticky="W")
                    self.advUpperLabel.grid(column=0, row=3, sticky="W")
                    self.advUpperEntry.grid(column=1, row=3, sticky="W")
                    self.advNumberLabel.grid(column=0, row=4, pady=5, sticky="W")
                    self.advNumberEntry.grid(column=1, row=4, pady=5, sticky="W")

            if (self.advancedOptionsPopup.state() == "withdrawn"):
                num_VE = int(self.numVoigtVariable.get())
                self.advancedOptionsPopup.title("Advanced parameter options")
                self.advancedOptionsPopup.iconbitmap(resource_path("img/elephant3.ico"))
                self.advancedOptionsPopup.geometry("700x400")
                self.advancedOptionsPopup.protocol("WM_DELETE_WINDOW", onCloseAdvanced)
                #self.advancedOptionsPopup.grid_rowconfigure(0, weight=1)
                self.paramListbox.grid(column=0, row=0, rowspan=2, sticky="NS")
                self.paramListboxScrollbar.grid(column=1, row=0, rowspan=2, sticky="NS")
                self.paramListbox.insert(tk.END, "Re")
                for i in range(1, num_VE+1):
                    self.paramListbox.insert(tk.END, "R" + str(i))
                    self.paramListbox.insert(tk.END, "T" + str(i))

                for i in range(2*num_VE+1-len(self.multistartEnabled)):
                    self.multistartEnabled.append(tk.IntVar(self, 0))
                    self.multistartSelection.append(tk.StringVar(self, "Logarithmic"))
                    self.multistartLower.append(tk.StringVar(self, "1E-5"))
                    self.multistartUpper.append(tk.StringVar(self, "1E5"))
                    self.multistartNumber.append(tk.StringVar(self, "10"))
                    self.multistartCustom.append(tk.StringVar(self, "0,1,10,100,1000"))
                self.paramListbox.select_set(0)
                self.paramListbox.bind("<<ListboxSelect>>", onSelect)
                self.advCheckbox = ttk.Checkbutton(self.advancedParamFrame, variable=self.multistartEnabled[0], text="Enable multistart", command=advCheck)
                self.advSelectionLabel = tk.Label(self.advancedParamFrame, text="Point spacing: ", background=self.backgroundColor, foreground=self.foregroundColor)
                self.advSelection = ttk.Combobox(self.advancedParamFrame, textvariable=self.multistartSelection[0], value=("Logarithmic", "Linear", "Random", "Custom"), justify=tk.CENTER, state="disabled", exportselection=0, width=12)
                self.advSelection.bind("<<ComboboxSelected>>", advSelectionChange)
                self.advLowerLabel = tk.Label(self.advancedParamFrame, text="Lower bound: ", background=self.backgroundColor, foreground=self.foregroundColor)
                self.advLowerEntry = ttk.Entry(self.advancedParamFrame, textvariable=self.multistartLower[0], state="disabled", width=15)
                self.advUpperLabel = tk.Label(self.advancedParamFrame, text="Upper bound: ", background=self.backgroundColor, foreground=self.foregroundColor)
                self.advUpperEntry = ttk.Entry(self.advancedParamFrame, textvariable=self.multistartUpper[0], state="disabled", width=15)
                self.advNumberLabel = tk.Label(self.advancedParamFrame, text="Number of points: ", background=self.backgroundColor, foreground=self.foregroundColor)
                self.advNumberEntry = ttk.Entry(self.advancedParamFrame, textvariable=self.multistartNumber[0], state="disabled", width=15)
                self.advCustomLabel = tk.Label(self.advancedParamFrame, text="Enter values separated by commas:", background=self.backgroundColor, foreground=self.foregroundColor)
                self.advCustomEntry = ttk.Entry(self.advancedParamFrame, textvariable=self.multistartCustom[0], state="disabled", width=45)
                self.advCheckbox.grid(column=0, row=0, sticky="W")
                self.advSelectionLabel.grid(column=0, row=1, sticky="W")
                self.advSelection.grid(column=1, row=1, sticky="W")
                if (self.multistartSelection[0].get() == "Custom"):
                    self.advCustomLabel.grid(column=0, row=2, pady=5, sticky="W")
                    self.advCustomEntry.grid(column=0, row=3, pady=5, columnspan=2, sticky="W")
                else:
                    self.advLowerLabel.grid(column=0, row=2, pady=5, sticky="W")
                    self.advLowerEntry.grid(column=1, row=2, pady=5, sticky="W")
                    self.advUpperLabel.grid(column=0, row=3, sticky="W")
                    self.advUpperEntry.grid(column=1, row=3, sticky="W")
                    self.advNumberLabel.grid(column=0, row=4, pady=5, sticky="W")
                    self.advNumberEntry.grid(column=1, row=4, pady=5, sticky="W")
                if (self.multistartEnabled[0].get()):
                    self.advSelection.configure(state="readonly")
                    self.advUpperEntry.configure(state="normal")
                    self.advLowerEntry.configure(state="normal")
                    self.advNumberEntry.configure(state="normal")
                    self.advCustomEntry.configure(state="normal")
                self.advancedChoiceLabel.grid(column=2, row=0, sticky="EW")
                self.advancedParamFrame.grid(column=2, row=1, padx=10, sticky="NSEW")
                self.advancedOptionsPopup.deiconify()
            else:
                self.advancedOptionsPopup.deiconify()
                self.advancedOptionsPopup.lift()
    
        def loadParams():
            self.paramPopup.withdraw()
            self.advancedOptionsPopup.withdraw()
            try:
                self.pframe.destroy()
                self.pcanvas.destroy()
                self.vsb.destroy()
                self.howMany.destroy()
                self.advancedOptions.destroy()
                self.addButton.destroy()
                self.removeButton.destroy()
                self.buttonFrame.destroy()
                self.paramComboboxes.clear()
                self.tauComboboxes.clear()
                self.paramLabels.clear()
                self.paramEntries.clear()
                self.paramFrames.clear()
                self.paramDeleteButtons.clear()
                if (self.paramEntryVariables[0].get() == ''):
                    self.paramEntryVariables[0].set(str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1))))
                for i in range(1, len(self.paramEntryVariables), 2):
                    if (self.paramEntryVariables[i].get() == ''):
                        self.paramEntryVariables[i].set(str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1))))
                    if (self.paramEntryVariables[i+1].get() == ''):
                        self.paramEntryVariables[i+1].set(str(round(self.tauDefault, -int(np.floor(np.log10(self.tauDefault))) + (4 - 1))))
            except:
                pass
            name = askopenfilename(initialdir=self.topGUI.getCurrentDirectory, filetypes = [("Measurement model fitting", "*.mmfitting")], title = "Choose a file")
            if (len(name) == 0):
                return
            try:
                directory = os.path.dirname(str(name))
                self.topGUI.setCurrentDirectory(directory)
                with open(name,'r') as toLoad:
                    for i in range(12):
                        toLoad.readline()
    #                    fileToLoad = toLoad.readline().split("filename: ")[1][:-1]
    #                    numberOfVoigt = int(toLoad.readline().split("number_voigt: ")[1][:-1])
    #                    if (numberOfVoigt < 1):
    #                        raise ValueError
    #                    numberOfSimulations = int(toLoad.readline().split("number_simulations: ")[1][:-1])
    #                    if (numberOfSimulations < 1):
    #                        raise ValueError
    #                    fitType = toLoad.readline().split("fit_type: ")[1][:-1]
    #                    if (fitType != "Real" and fitType != "Complex" and fitType != "Imaginary"):
    #                        raise ValueError
    #                    weightType = toLoad.readline().split("weight_type: ")[1][:-1]
    #                    if (weightType != "Modulus" and weightType != "None" and weightType != "Proportional" and weightType != "Error Model"):
    #                        raise ValueError
                    alphaValue = float(toLoad.readline().split("alpha: ")[1][:-1])
                    if (alphaValue < 0):
                        raise ValueError
                    capValue = toLoad.readline().split("cap: ")[1][:-1]
                    if (capValue == "N"):
                        pass
                    else:
                        capValue = float(capValue)
                    capCombo = toLoad.readline().split("cap_combo: ")[1][:-1]
                    if (capCombo != "+" and capCombo != "+ or -" and capCombo != "-" and capCombo != "fixed"):
                        raise ValueError
                    toLoad.readline()
                    paramValues = []
                    while (True):
                        p = toLoad.readline()[:-1]
                        if ("paramComboboxes" in p):
                            break
                        else:
                            paramValues.append(float(p))
                    pComboboxes = []
                    while (True):
                        p = toLoad.readline()[:-1]
                        if ("tauComboboxes" in p):
                            break
                        else:
                            if (p != "+" and p != "+ or -" and p != "fixed"):
                                raise ValueError
                            pComboboxes.append(p)
                    tComboboxes = []
                    while (True):
                        p = toLoad.readline()
                        if not p:
                            break
                        else:
                            p = p[:-1]
                            if (p != "+" and p != "fixed"):
                                raise ValueError
                            tComboboxes.append(p)
                if (len(pComboboxes)-1 != len(tComboboxes)):
                    raise ValueError
                elif (len(pComboboxes) + len(tComboboxes) != len(paramValues)):
                    raise ValueError
                self.paramEntries.clear()
                self.paramEntryVariables.clear()
                self.paramComboboxes.clear()
                self.tauComboboxes.clear()
                self.paramComboboxes.clear()
                self.paramComboboxVariables.clear()
                self.tauComboboxVariables.clear()
                for pvar in pComboboxes:
                    self.paramComboboxVariables.append(tk.StringVar(self, pvar))
                for tvar in tComboboxes:
                    self.tauComboboxVariables.append(tk.StringVar(self, tvar))
                self.paramEntryVariables.append(tk.StringVar(self, str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1)))))
                self.paramEntryVariables[0] = tk.StringVar(self, str(paramValues[0]))
                for pval in paramValues[1:]:
                    self.paramEntryVariables.append(tk.StringVar(self, str(pval)))
                if (capValue == "N"):
                    self.capacitanceCheckboxVariable.set(0)
                    self.capacitanceEntryVariable.set("0")
                    self.capacitanceComboboxVariable.set("+ or -")
                else:
                    self.capacitanceCheckboxVariable.set(1)
                    self.capacitanceEntryVariable.set(str(capValue))
                    self.capacitanceComboboxVariable.set(capCombo)
                self.multistartCustom.clear()
                self.multistartSelection.clear()
                self.multistartEnabled.clear()
                self.multistartLower.clear()
                self.multistartUpper.clear()
                self.multistartNumber.clear()
                self.paramListbox.delete(0, tk.END)
                """
                for i in range(3):
                    self.multistartEnabled.append(tk.IntVar(self, 0))
                    self.multistartSelection.append(tk.StringVar(self, "Logarithmic"))
                    self.multistartLower.append(tk.StringVar(self, "1"))
                    self.multistartUpper.append(tk.StringVar(self, "1E5"))
                    self.multistartNumber.append(tk.StringVar(self, "10"))
                    self.multistartCustom.append(tk.StringVar(self, "0,1,10,100,1000"))
                """
                self.numVoigtTextVariable.set(str(len(tComboboxes)))
                self.numVoigtVariable.set(str(len(tComboboxes)))
                self.nVoigt = int(self.numVoigtVariable.get())
            except:
                messagebox.showerror("File error", "Error 9: \nThere was an error loading or reading the file")

        def checkWeight(event):
            if (self.weightingVariable.get() == "None"):
                self.alphaLabel.grid_remove()
                self.alphaEntry.grid_remove()
                self.errorStructureFrame.grid_remove()
            elif (self.weightingVariable.get() == "Error model"):
                self.alphaLabel.grid_remove()
                self.alphaEntry.grid_remove()
                self.errorStructureFrame.grid(column=0, row=1, pady=5, sticky="W", columnspan=5)
            else:
                self.alphaLabel.grid(column=4, row=0)
                self.alphaEntry.grid(column=5, row=0)
                self.errorStructureFrame.grid_remove()
        
        def process_queue_auto():
            try:
                r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic, ft, fw = self.queueAuto.get(0)
                self.measureModelButton.configure(state="normal")
                self.autoButton.configure(state="normal")
                self.magicButton.configure(state="normal")
                self.parametersButton.configure(state="normal")
                self.freqRangeButton.configure(state="normal")
                self.parametersLoadButton.configure(state="normal")
                self.numVoigtSpinbox.configure(state="readonly")
                self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
                self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtMinus.bind("<ButtonRelease-1>", minusNVE)
                self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
                self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
                self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtPlus.bind("<ButtonRelease-1>", plusNVE)
                self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
                self.browseButton.configure(state="normal")
                self.numMonteEntry.configure(state="normal")
                self.fitTypeCombobox.configure(state="readonly")
                self.weightingCombobox.configure(state="readonly")
                self.autoNMC.configure(state="normal")
                self.autoMaxNVE.configure(state="normal")
                self.autoWeighting.configure(state="readonly")
                self.autoRunButton.configure(state="normal")
                self.autoCancelButton.configure(state="disabled")
                self.autoReFixEntry.configure(state="normal")
                self.autoReFix.configure(state="normal")
                self.autoPercent = []
                self.autoStatusLabel.configure(text="")
                self.ellipsisPercent = 0
                #self.topGUI.master.focus_force()
                try:
                    self.taskbar.SetProgressState(self.masterWindowId, 0x0)
                except:
                    pass
                try:
                    ctypes.windll.user32.FlashWindow(self.masterWindowId, True)
                except:
                    pass
                #self.numVoigtVariable.set((len(r)-1)//2)
                #changeNVE()
                if (len(r) == 1 and len(s) == 1 and len(sdr) == 1 and len(sdi)==1 and len(zz)==1 and len(zzs)==1 and len(chi)==1):
                    if (r=="^" and s=="^" and sdr=="^" and sdi=="^" and zz=="^" and zzs=="^" and zp=="^" and zps=="^" and chi=="^"):
                        messagebox.showerror("Fitting error", "Error 11:\nThe fitting failed")
                    elif (r=="@" and s=="@" and sdr=="@" and sdi=="@" and zz=="@" and zzs=="@" and zp=="@" and zps=="@" and chi=="@"):
                        pass    #Fitting was cancelled
                else:
                    nveNeeded = (len(r)-1)//2
                    if (nveNeeded < self.nVoigt):
                        while (nveNeeded-self.nVoigt != 0):
                            self.numVoigtSpinbox.invoke("buttondown")
                    elif (nveNeeded > self.nVoigt):
                        while(nveNeeded-self.nVoigt != 0):
                            self.numVoigtSpinbox.invoke("buttonup")
                    self.fitWeightR = np.zeros(len(self.wdata))
                    self.fitWeightJ = np.zeros(len(self.wdata))
                    if (self.autoReFixVariable.get()):
                        self.paramComboboxVariables[0].set("fixed")
                        
                    if (fw == 2):
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = float(self.alphaVariable.get()) * self.rdata[i]
                        for i in range(len(self.wdata)):
                            self.fitWeightJ[i] = float(self.alphaVariable.get()) * self.jdata[i]
                    elif (fw == 1):
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = float(self.alphaVariable.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                            self.fitWeightJ[i] = float(self.alphaVariable.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                    elif (fw == 3):
                        if (self.errorAlphaCheckboxVariableAuto.get() == 0):
                            errA = 0
                        else:
                            errA = float(self.errorAlphaEntryAuto.get())
                        if (self.errorBetaCheckboxVariableAuto.get() == 0):
                            errB = 0
                        else:
                            errB = float(self.errorBetaEntryAuto.get())
                        if (self.errorBetaReCheckboxVariableAuto.get() == 0):
                            errBRe = 0
                        else:
                            errBRe = float(self.errorBetaReEntryAuto.get())
                        if (self.errorGammaCheckboxVariableAuto.get() == 0):
                            errG = 0
                        else:
                            errG = float(self.errorGammaEntryAuto.get())
                        if (self.errorDeltaCheckboxVariableAuto.get() == 0):
                            errD = 0
                        else:                        
                            errD = float(self.errorDeltaEntryAuto.get())
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = errA*self.jdata[i] + errB*(self.rdata[i] - errBRe) + errG*np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)**2 + errD
                            self.fitWeightJ[i] = errA*self.jdata[i] + errB*(self.rdata[i] - errBRe) + errG*np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)**2 + errD
                    self.currentNVEPlotNeeded = (len(self.fits)-1)/2
                    if (len(self.fits) > 0):
                        self.undoStack.append(self.fits.copy())
                        self.undoSigmaStack.append(self.sigmas.copy())
                        self.undoAICStack.append(self.aicResult)
                        self.undoChiStack.append(self.chiSquared)
                        self.undoCorStack.append(self.corResult.copy())
                        self.undoCapNeededStack.append(self.capUsed)
                        self.undoCapStack.append(self.resultCap)
                        self.undoCapSigmaStack.append(self.sigmaCap)
                        self.undoCorCapStack.append(self.capCor)
                        self.undoUpDeleteStack.append(self.pastUpDelete)
                        self.undoLowDeleteStack.append(self.pastLowDelete)
                        self.undoSdrRealStack.append(self.sdrReal)
                        self.undoSdrImagStack.append(self.sdrImag)
                        self.pastUpDelete = self.upDelete
                        self.pastLowDelete = self.lowDelete
                        self.undoAlphaStack.append(self.pastAlphaVariable)
                        self.pastAlphaVariable = self.alphaVariable.get()
                        self.undoFitTypeStack.append(self.pastFitType)
                        self.pastFitType = self.fitTypeVariable.get()
                        self.undoNumSimulationsStack.append(self.pastNumSimulations)
                        self.pastNumSimulations = self.numMonteVariable.get()
                        self.undoWeightingStack.append(self.pastWeighting)
                        self.pastWeighting = self.weightingVariable.get()
                        self.undoErrorModelChoicesStack.append(self.pastErrorModelChoices)
                        self.pastErrorModelChoices = [self.errorAlphaCheckboxVariable.get(), self.errorBetaCheckboxVariable.get(), self.errorBetaReCheckboxVariable.get(), self.errorGammaCheckboxVariable.get(), self.errorDeltaCheckboxVariable.get()]
                        self.undoErrorModelValuesStack.append(self.pastErrorValues)
                        self.pastErrorValues = [self.errorAlphaVariable.get(), self.errorBetaVariable.get(), self.errorBetaReVariable.get(), self.errorGammaVariable.get(), self.errorDeltaVariable.get()]
                        self.undoButton.configure(state="normal")
#                        self.redoStack = []
#                        self.redoButton.configure(state="disabled")
                    else:
                        self.pastUpDelete = self.upDelete
                        self.pastLowDelete = self.lowDelete
                        self.pastAlphaVariable = self.alphaVariable.get()
                        self.pastFitType = self.fitTypeVariable.get()
                        self.pastNumSimulations = self.numMonteVariable.get()
                        self.pastWeighting = self.weightingVariable.get()
                        self.pastErrorModelChoices = [self.errorAlphaCheckboxVariable.get(), self.errorBetaCheckboxVariable.get(), self.errorBetaReCheckboxVariable.get(), self.errorGammaCheckboxVariable.get(), self.errorDeltaCheckboxVariable.get()]
                        self.pastErrorValues = [self.errorAlphaVariable.get(), self.errorBetaVariable.get(), self.errorBetaReVariable.get(), self.errorGammaVariable.get(), self.errorDeltaVariable.get()]
                    self.capUsed = False
                    self.alreadyChanged = False
                    self.fits = r
                    self.sigmas = s
                    self.sdrReal = sdr
                    self.sdrImag = sdi
                    self.chiSquared = chi
                    self.aicResult = aic
                    self.corResult = cor
#                    self.resultData = r
#                    self.sigmaData = s
#                    self.sdrData = sdr
#                    self.sdiData = sdi
                    self.aR = ""
                    self.resultsView.delete(*self.resultsView.get_children())
                    for i in range(len(self.paramEntryVariables)):
                        self.paramEntryVariables[i].set("%.5g"%r[i])
                    self.resultsFrame.grid(column=0, row=6, sticky="W", pady=5)
                    self.graphFrame.grid(column=0, row=7, sticky="W", pady=5)
                    self.resultRe.configure(text="Ohmic Resistance = %.4g"%r[0] + " ± %.2g"%s[0])
                    self.resultRp.configure(text="Polarization Impedance = %.4g"%zp + " ± %.2g"%zps)
                    capacitances = np.zeros(int((len(r)-1)/2))
                    for i in range(1, len(r), 2):
                        if (r[1] == 0):
                            capacitances[int(i/2)] = 0
                        else:
                            capacitances[int(i/2)] = r[i+1]/r[i]
                    ceff = 1/np.sum(1/capacitances)
                    sigmaCapacitances = np.zeros(len(capacitances))
                    for i in range(1, len(r), 2):
                        if (r[i] == 0 or r[i+1] == 0):
                            sigmaCapacitances[int(i/2)] = 0
                        else:
                            sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((s[i+1]/r[i+1])**2 + (s[i]/r[i])**2)
                    partCap = 0
                    for i in range(len(capacitances)):
                        partCap += sigmaCapacitances[i]**2/capacitances[i]**4
                    sigmaCeff = ceff**2 * np.sqrt(partCap)
                    self.resultC.configure(text="Capacitance = %.4g"%ceff + " ± %.2g"%sigmaCeff)
                    self.aR += "File name: " + self.browseEntry.get() + "\n"
                    self.aR += "Number of data: " + str(self.lengthOfData) + "\n"
                    self.aR += "Number of parameters: " + str(len(r)) + "\n"
                    self.aR += "\nRe (Rsol) = %.8g"%r[0] + " ± %.4g"%s[0] + "\n"
                    for i in range(1, len(r), 2):
                        self.aR += "R" + str(int(i/2 + 1)) + " = %.8g"%r[i] + " ± %.4g"%s[i] + "\n"
                        self.aR += "Tau" + str(int(i/2 + 1)) + " = %.8g"%r[i+1] + " ± %.4g"%s[i+1] + "\n"
                        self.aR += "C" + str(int(i/2 + 1)) + " = %.8g"%capacitances[int(i/2)] + " ± %.4g"%sigmaCapacitances[int(i/2)] + "\n"
                    self.aR += "\nZero frequency impedance = %.8g"%zz + " ± %.4g"%zzs + "\n"
                    self.aR += "Polarization Impedance = %.8g"%zp + " ± %.4g"%zps + "\n"
                    self.aR += "Capacitance = %.8g"%ceff + " ± %.4g"%sigmaCeff + "\n"
                    self.aR += "\nCorrelation matrix:\n"
                    self.aR += "         Re    "
                    for i in range(1, len(r), 2):
                        if (int(i/2 +1) < 10):
                            self.aR += "  R" + str(int(i/2 + 1)) + "    " + " Tau" + str(int(i/2 + 1)) + "   "
                        else:
                            self.aR += "  R" + str(int(i/2 + 1)) + "   " + " Tau" + str(int(i/2 + 1)) + "  "
                    self.aR += "\n      \u250C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                    #for i in range(1, len(r), 2):
                    #    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                    self.aR += "\n"
                    self.aR += "Re    "
                    for i in range(len(r)):
                        if (i != 0):
                            if (i%2 != 0):
                                if (i/2 + 1 < 10):
                                    self.aR += "R" + str(int(i/2+1)) + "    "
                                else:
                                    self.aR += "R" + str(int(i/2+1)) + "   "
                            else:
                                if (i/2 < 10):
                                    self.aR += "Tau" + str(int(i/2)) + "  "
                                else:
                                    self.aR += "Tau" + str(int(i/2)) + " "
                        self.aR += "\u2502"
                        for j in range(len(r)):
                            if (j <= i):
                                if ("%05.2f"%cor[i][j] == "01.00"):
                                    self.aR += "   1   \u2502"
                                elif (np.isnan(cor[i][j])):
                                    self.aR += "   0   \u2502"
                                else:
                                    self.aR += " %05.2f"%cor[i][j] + " \u2502"
                        if (i != len(r)-1):
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                            self.aR += "\u253C"
                            if (i == 0):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        else:
                            self.aR += "\n      \u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u2518"
                                else:
                                    self.aR += "\u2534"
                        self.aR += "\n"
                        
                    self.aR += "\nChi-squared statistic = %.4g"%chi + "\n"
                    self.aR += "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-len(r))) + "\n"
                    try:
                        self.aR += "Akaike Information Criterion = %.4g"%aic#(np.log(chi*((1+2*len(r))/self.lengthOfData))) + "\n"
                    except:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    #self.aR += "Akaike Performance Index = %.4g"%(chi*((1+(len(r)/self.lengthOfData))/(1-(len(r)/self.lengthOfData))))
                    """
                    self.results.configure(state="normal")
                    self.results.insert(tk.INSERT, "Re (Rsol) = %.4g"%r[0] + " ± %.1g"%s[0] +"\n")
                    for i in range(1, len(r), 2):
                        self.results.insert(tk.INSERT, "Rt" + str(int(i/2 + 1)) + " = %.4g"%r[i] + " ± %.1g"%s[i] + "\n")
                        self.results.insert(tk.INSERT, "Tau" + str(int(i/2+1)) + " = %.4g"%r[i+1] + " ± %.1g"%s[i+1] + "\n")
                    self.results.insert(tk.INSERT, "\nZero frequency impedance = %.4g"%zz + " ± %.1g"%zzs + "\n")
                    self.results.insert(tk.INSERT, "Polarization impedance = %.4g"%zp + " ± %.1g"%zps + "\n")
                    self.results.insert(tk.INSERT, "Chi-squared statistic = %.4g"%chi + "\n")
                    self.results.insert(tk.INSERT, "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-(int(self.numVoigtVariable.get())*2+1))))
                    self.results.configure(state="disabled")
                    """
                    self.resultAlert.grid_remove()
                    if (len(s) == 1):
                        self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "nan", "nan"))
                        for i in range(1, len(r), 2):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "nan", "nan"))
                            else:
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "nan", "nan"))
                            if (r[i+1] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "nan", "nan"))
                            else:
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "nan", "nan"))
                    else:
                        if (abs(s[0]*2*100/r[0]) > 100 or np.isnan(s[0])):
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "%.3g"%(s[0]*2*100/r[0])+"%"), tags=("bad",))
                            self.resultAlert.grid(column=0, row=2, sticky="E")
                        else:
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "%.3g"%(s[0]*2*100/r[0])+"%"))
                        for i in range(1, len(r), 2):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "0%"))
                            else:
                                if (abs(s[i]*2*100/r[i]) > 100 or np.isnan(s[i])):
                                    self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"), tags=("bad",))
                                    self.resultAlert.grid(column=0, row=2, sticky="E")
                                else:
                                    self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"))
                            if (r[i+1] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "0%"))
                            else:
                                if (abs(s[i+1]*2*100/r[i+1]) > 100 or np.isnan(s[i])):
                                    self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "%.3g"%(s[i+1]*2*100/r[i+1])+"%"), tags=("bad",))
                                    self.resultAlert.grid(column=0, row=2, sticky="E")
                                else:
                                    self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "%.3g"%(s[i+1]*2*100/r[i+1])+"%"))
                    self.resultsView.tag_configure("bad", background="yellow")
                    self.whatFitStack.append(self.whatFit)
                    if (ft == 0):
                        self.whatFit = "C"
                        self.fitTypeVariable.set("Complex")
                    elif (ft == 1):
                        self.whatFit = "J"
                        self.fitTypeVariable.set("Imaginary")
                    elif (ft == 3):
                        self.whatFit = "R"
                        self.fitTypeVariable.set("Real")
                    if (self.magicPlot.state() != "withdrawn"):
                        self.magicInput.clf()
                        Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                        if (not self.capUsed):
                            for i in range(len(self.wdata)):
                                Zfit[i] = self.fits[0]
                                for k in range(1, len(self.fits), 2):
                                    Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        else:
                            for i in range(len(self.wdata)):
                                Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                                for k in range(1, len(self.fits), 2):
                                    Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        
                        x = np.array(self.rdata)    #Doesn't plot without this
                        y = np.array(self.jdata)
                        dataColor = "tab:blue"
                        fitColor = "orange"
                        if (self.topGUI.getTheme() == "dark"):
                            dataColor = "cyan"
                            fitColor = "gold"
                        else:
                            dataColor = "tab:blue"
                            fitColor = "orange"
                        with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                            self.magicSubplot = self.magicInput.add_subplot(111)
                            self.magicSubplot.set_facecolor(self.backgroundColor)
                            self.magicSubplot.yaxis.set_ticks_position("both")
                            self.magicSubplot.yaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                            self.magicSubplot.xaxis.set_ticks_position("both")
                            self.magicSubplot.xaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                            pointsPlot, = self.magicSubplot.plot(x, -1*y, "o", color=dataColor)
                            topPoint = max(-1*y)
                            rightPoint = max(x)
                            if (len(self.fits) > 0):
                                self.magicSubplot.plot(np.array(Zfit.real), np.array(-1*Zfit.imag), color=fitColor)
                            self.magicSubplot.axis("equal")
                            self.magicSubplot.set_title("Magic Finger Nyquist Plot", color=self.foregroundColor)
                            self.magicSubplot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                            self.magicSubplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                            self.magicInput.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                        self.magicCanvasInput.draw()
                        annot = self.magicSubplot.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
                        annot.set_visible(False)
                        def update_annot(ind):
                            x,y = pointsPlot.get_data()
                            xval = x[ind["ind"][0]]
                            yval = y[ind["ind"][0]]
                            annot.xy = (xval, yval)
                            text = "Zr=%.3g"%xval + "\nZj=-%.3g"%yval + "\nf=%.5g"%self.wdata[np.where(self.rdata == xval)][0]
                            annot.set_text(text)
                            #---Check if we're within 5% of the right or top edges, and adjust label positions accordingly
                            if (rightPoint != 0):
                                if (abs(xval - rightPoint)/rightPoint <= 0.2):
                                    annot.set_position((-90, -10))
                            if (topPoint != 0):
                                if (abs(yval - topPoint)/topPoint <= 0.05):
                                    annot.set_position((10, -20))
                            else:
                                annot.set_position((10, -20))
                        def hover(event):
                            vis = annot.get_visible()
                            if event.inaxes == self.magicSubplot:
                                cont, ind = pointsPlot.contains(event)
                                if cont:
                                    update_annot(ind)
                                    annot.set_visible(True)
                                    self.magicInput.canvas.draw_idle()
                                else:
                                    if vis:
                                        annot.set_position((10,10))
                                        annot.set_visible(False)
                                        self.magicInput.canvas.draw_idle()
                        self.magicInput.canvas.mpl_connect("motion_notify_event", hover)
                    self.autoFitWindow.withdraw()
                    runMeasurementModel()
            except queue.Empty:
                if (len(self.autoPercent) > 0):
                    currentVal = self.autoPercent[len(self.autoPercent) - 1]
                    textToDisplay = ""
                    if (currentVal[0] == "a"):
                        textToDisplay = "Trying " + currentVal[1:] + " Voigt elements with imaginary fit"
                    elif (currentVal[0] == "b"):
                        textToDisplay = "Trying " + currentVal[1:] + " Voigt elements with multistart fit"
                    elif (currentVal[0] == "c"):
                        textToDisplay = "Calculating confidence intervals"
                    elif (currentVal[0] == "d"):
                        textToDisplay = "Trying " + currentVal[1:] + " Voigt elements with proportional weighting"
                    else:
                        textToDisplay = "Trying " + currentVal + " Voigt elements"
                    ellipsisList = ["   ", "   ", "   ", ".  ", ".  ", ".  ", ".. ", ".. ", ".. ", "...", "...", "..."]
                    textToDisplay += ellipsisList[self.ellipsisPercent%12]
                    self.autoStatusLabel.configure(text=textToDisplay)
                    self.ellipsisPercent += 1
                self.after(100, process_queue_auto)                
        
        def process_queue():
            try:
                r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic = self.queue.get(0)
                self.measureModelButton.configure(state="normal")
                self.autoButton.configure(state="normal")
                self.magicButton.configure(state="normal")
                self.parametersButton.configure(state="normal")
                self.freqRangeButton.configure(state="normal")
                self.parametersLoadButton.configure(state="normal")
                self.numVoigtSpinbox.configure(state="readonly")
                self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
                self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtMinus.bind("<ButtonRelease-1>", minusNVE)
                self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
                self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
                self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtPlus.bind("<ButtonRelease-1>", plusNVE)
                self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
                self.browseButton.configure(state="normal")
                self.numMonteEntry.configure(state="normal")
                self.fitTypeCombobox.configure(state="readonly")
                self.weightingCombobox.configure(state="readonly")
                self.listPercent = []
                try:
                    self.progStatus.configure(text="Initializing...")
                    self.progStatus.grid_remove()
                except:
                    pass
                for delButton in self.paramDeleteButtons:
                    delButton.configure(state="normal")
                try:
                    self.addButton.configure(state="normal")
                    self.removeButton.configure(state="normal")
                except:
                    pass
                try:
                    self.autoRunButton.configure(state="normal")
                except:
                    pass
                self.prog_bar.stop()
                self.prog_bar.destroy()
                try:
                    self.taskbar.SetProgressState(self.masterWindowId, 0x0)
                except:
                    pass
                self.cancelButton.grid_remove()
                if (len(r) == 1 and len(s) == 1 and len(sdr) == 1 and len(sdi)==1 and len(zz)==1 and len(zzs)==1 and len(chi)==1):
                    if (r=="^" and s=="^" and sdr=="^" and sdi=="^" and zz=="^" and zzs=="^" and zp=="^" and zps=="^" and chi=="^"):
                        messagebox.showerror("Fitting error", "Error 11:\nThe fitting failed")
                    elif (r=="@" and s=="@" and sdr=="@" and sdi=="@" and zz=="@" and zzs=="@" and zp=="@" and zps=="@" and chi=="@"):
                        pass    #Fitting was cancelled
                else:
                    if (len(r) != 1 and len(s) == 1 and len(sdr) == 1 and len(sdi)==1 and len(zzs)==1):
                        if (s=="-" and sdr=="-" and sdi=="-" and zzs=="-" and zps=="-"):
                            s = np.zeros(len(r))
                            sdr = np.zeros(len(self.wdata))
                            sdi = np.zeros(len(self.wdata))
                            zzs = 0
                            zps = 0
                            messagebox.showwarning("Variance error", "Error 12:\nA minimization was found, but the standard deviations and confidence intervals could not be estimated. Try fitting again with these values or try fewer line shapes")
                    self.fitWeightR = np.zeros(len(self.wdata))
                    self.fitWeightJ = np.zeros(len(self.wdata))
                    if (self.weightingVariable.get() == "None"):
                        self.fitWeightR = np.ones(len(self.wdata))
                        self.fitWeightJ = np.ones(len(self.wdata))
                    elif (self.weightingVariable.get() == "Proportional"):
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = float(self.alphaVariable.get()) * self.rdata[i]
                        for i in range(len(self.wdata)):
                            self.fitWeightJ[i] = float(self.alphaVariable.get()) * self.jdata[i]
                    elif (self.weightingVariable.get() == "Modulus"):
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = float(self.alphaVariable.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                            self.fitWeightJ[i] = float(self.alphaVariable.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                    elif (self.weightingVariable.get() == "Error model"):
                        for i in range(len(self.wdata)):
                            if (self.errorAlphaCheckboxVariable.get() == 1):
                                self.fitWeightR[i] += float(self.errorAlphaVariable.get())*abs(self.jdata[i])
                            if (self.errorBetaCheckboxVariable.get() == 1):
                                if (self.errorBetaReCheckboxVariable.get() == 1):
                                    self.fitWeightR[i] += (float(self.errorBetaVariable.get())*abs(self.rdata[i]) - float(self.errorBetaReVariable.get()))
                                else:
                                    self.fitWeightR[i] += float(self.errorBetaVariable.get())*abs(self.rdata[i])
                            if (self.errorGammaCheckboxVariable.get() == 1):
                                self.fitWeightR[i] += float(self.errorGammaVariable.get())*np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)**2
                            if (self.errorDeltaCheckboxVariable.get() == 1):
                                self.fitWeightR[i] += float(self.errorDeltaVariable.get())
                            self.fitWeightJ[i] = self.fitWeightR[i]
                    self.currentNVEPlotNeeded = (len(self.fits)-1)/2
                    if (len(self.fits) > 0):
                        self.undoStack.append(self.fits.copy())
                        self.undoSigmaStack.append(self.sigmas.copy())
                        self.undoAICStack.append(self.aicResult)
                        self.undoChiStack.append(self.chiSquared)
                        self.undoCorStack.append(self.corResult.copy())
                        self.undoCapNeededStack.append(self.capUsed)
                        self.undoCapStack.append(self.resultCap)
                        self.undoCapSigmaStack.append(self.sigmaCap)
                        self.undoCorCapStack.append(self.capCor)
                        self.undoUpDeleteStack.append(self.pastUpDelete)
                        self.undoLowDeleteStack.append(self.pastLowDelete)
                        self.undoSdrRealStack.append(self.sdrReal)
                        self.undoSdrImagStack.append(self.sdrImag)
                        self.pastUpDelete = self.upDelete
                        self.pastLowDelete = self.lowDelete
                        self.undoAlphaStack.append(self.pastAlphaVariable)
                        self.pastAlphaVariable = self.alphaVariable.get()
                        self.undoFitTypeStack.append(self.pastFitType)
                        self.pastFitType = self.fitTypeVariable.get()
                        self.undoNumSimulationsStack.append(self.pastNumSimulations)
                        self.pastNumSimulations = self.numMonteVariable.get()
                        self.undoWeightingStack.append(self.pastWeighting)
                        self.pastWeighting = self.weightingVariable.get()
                        self.undoErrorModelChoicesStack.append(self.pastErrorModelChoices)
                        self.pastErrorModelChoices = [self.errorAlphaCheckboxVariable.get(), self.errorBetaCheckboxVariable.get(), self.errorBetaReCheckboxVariable.get(), self.errorGammaCheckboxVariable.get(), self.errorDeltaCheckboxVariable.get()]
                        self.undoErrorModelValuesStack.append(self.pastErrorValues)
                        self.pastErrorValues = [self.errorAlphaVariable.get(), self.errorBetaVariable.get(), self.errorBetaReVariable.get(), self.errorGammaVariable.get(), self.errorDeltaVariable.get()]
                        self.undoButton.configure(state="normal")
#                        self.redoStack = []
#                        self.redoButton.configure(state="disabled")
                    else:
                        self.pastUpDelete = self.upDelete
                        self.pastLowDelete = self.lowDelete
                        self.pastAlphaVariable = self.alphaVariable.get()
                        self.pastFitType = self.fitTypeVariable.get()
                        self.pastNumSimulations = self.numMonteVariable.get()
                        self.pastWeighting = self.weightingVariable.get()
                        self.pastErrorModelChoices = [self.errorAlphaCheckboxVariable.get(), self.errorBetaCheckboxVariable.get(), self.errorBetaReCheckboxVariable.get(), self.errorGammaCheckboxVariable.get(), self.errorDeltaCheckboxVariable.get()]
                        self.pastErrorValues = [self.errorAlphaVariable.get(), self.errorBetaVariable.get(), self.errorBetaReVariable.get(), self.errorGammaVariable.get(), self.errorDeltaVariable.get()]
                    self.capUsed = False
                    self.alreadyChanged = False
                    self.fits = r
                    self.sigmas = s
                    self.sdrReal = sdr
                    self.sdrImag = sdi
                    self.chiSquared = chi
                    self.aicResult = aic
                    self.corResult = cor
#                    self.resultData = r
#                    self.sigmaData = s
#                    self.sdrData = sdr
#                    self.sdiData = sdi
                    self.aR = ""
                    self.resultsView.delete(*self.resultsView.get_children())
                    for i in range(len(self.paramEntryVariables)):
                        self.paramEntryVariables[i].set("%.5g"%r[i])
                    self.resultsFrame.grid(column=0, row=6, sticky="W", pady=5)
                    self.graphFrame.grid(column=0, row=7, sticky="W", pady=5)
                    self.resultRe.configure(text="Ohmic Resistance = %.4g"%r[0] + " ± %.2g"%s[0])
                    self.resultRp.configure(text="Polarization Impedance = %.4g"%zp + " ± %.2g"%zps)
                    capacitances = np.zeros(int((len(r)-1)/2))
                    for i in range(1, len(r), 2):
                        if (r[1] == 0):
                            capacitances[int(i/2)] = 0
                        else:
                            capacitances[int(i/2)] = r[i+1]/r[i]
                    ceff = 1/np.sum(1/capacitances)
                    sigmaCapacitances = np.zeros(len(capacitances))
                    for i in range(1, len(r), 2):
                        if (r[i] == 0 or r[i+1] == 0):
                            sigmaCapacitances[int(i/2)] = 0
                        else:
                            sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((s[i+1]/r[i+1])**2 + (s[i]/r[i])**2)
                    partCap = 0
                    for i in range(len(capacitances)):
                        partCap += sigmaCapacitances[i]**2/capacitances[i]**4
                    sigmaCeff = ceff**2 * np.sqrt(partCap)
                    self.resultC.configure(text="Capacitance = %.4g"%ceff + " ± %.2g"%sigmaCeff)
                    self.aR += "File name: " + self.browseEntry.get() + "\n"
                    self.aR += "Number of data: " + str(self.lengthOfData) + "\n"
                    self.aR += "Number of parameters: " + str(len(r)) + "\n"
                    self.aR += "\nRe (Rsol) = %.8g"%r[0] + " ± %.4g"%s[0] + "\n"
                    for i in range(1, len(r), 2):
                        self.aR += "R" + str(int(i/2 + 1)) + " = %.8g"%r[i] + " ± %.4g"%s[i] + "\n"
                        self.aR += "Tau" + str(int(i/2 + 1)) + " = %.8g"%r[i+1] + " ± %.4g"%s[i+1] + "\n"
                        self.aR += "C" + str(int(i/2 + 1)) + " = %.8g"%capacitances[int(i/2)] + " ± %.4g"%sigmaCapacitances[int(i/2)] + "\n"
                    self.aR += "\nZero frequency impedance = %.8g"%zz + " ± %.4g"%zzs + "\n"
                    self.aR += "Polarization Impedance = %.8g"%zp + " ± %.4g"%zps + "\n"
                    self.aR += "Capacitance = %.8g"%ceff + " ± %.4g"%sigmaCeff + "\n"
                    self.aR += "\nCorrelation matrix:\n"
                    self.aR += "         Re    "
                    for i in range(1, len(r), 2):
                        if (int(i/2 +1) < 10):
                            self.aR += "  R" + str(int(i/2 + 1)) + "    " + " Tau" + str(int(i/2 + 1)) + "   "
                        else:
                            self.aR += "  R" + str(int(i/2 + 1)) + "   " + " Tau" + str(int(i/2 + 1)) + "  "
                    self.aR += "\n      \u250C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                    #for i in range(1, len(r), 2):
                    #    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                    self.aR += "\n"
                    self.aR += "Re    "
                    for i in range(len(r)):
                        if (i != 0):
                            if (i%2 != 0):
                                if (int(i/2 + 1) < 10):
                                    self.aR += "R" + str(int(i/2+1)) + "    "
                                else:
                                    self.aR += "R" + str(int(i/2+1)) + "   "
                            else:
                                if (i/2 < 10):
                                    self.aR += "Tau" + str(int(i/2)) + "  "
                                else:
                                    self.aR += "Tau" + str(int(i/2)) + " "
                        self.aR += "\u2502"
                        for j in range(len(r)):
                            if (j <= i):
                                if ("%05.2f"%cor[i][j] == "01.00"):
                                    self.aR += "   1   \u2502"
                                elif (np.isnan(cor[i][j])):
                                    self.aR += "   0   \u2502"
                                else:
                                    self.aR += " %05.2f"%cor[i][j] + " \u2502"
                        if (i != len(r)-1):
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                            self.aR += "\u253C"
                            if (i == 0):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        else:
                            self.aR += "\n      \u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u2518"
                                else:
                                    self.aR += "\u2534"
                        self.aR += "\n"
                        
                    self.aR += "\nChi-squared statistic = %.4g"%chi + "\n"
                    self.aR += "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-len(r))) + "\n"
                    try:
                        self.aR += "Akaike Information Criterion = %.4g"%aic#(np.log(chi*((1+2*len(r))/self.lengthOfData))) + "\n"
                    except:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    #self.aR += "Akaike Performance Index = %.4g"%(chi*((1+(len(r)/self.lengthOfData))/(1-(len(r)/self.lengthOfData))))
                    """
                    self.results.configure(state="normal")
                    self.results.insert(tk.INSERT, "Re (Rsol) = %.4g"%r[0] + " ± %.1g"%s[0] +"\n")
                    for i in range(1, len(r), 2):
                        self.results.insert(tk.INSERT, "Rt" + str(int(i/2 + 1)) + " = %.4g"%r[i] + " ± %.1g"%s[i] + "\n")
                        self.results.insert(tk.INSERT, "Tau" + str(int(i/2+1)) + " = %.4g"%r[i+1] + " ± %.1g"%s[i+1] + "\n")
                    self.results.insert(tk.INSERT, "\nZero frequency impedance = %.4g"%zz + " ± %.1g"%zzs + "\n")
                    self.results.insert(tk.INSERT, "Polarization impedance = %.4g"%zp + " ± %.1g"%zps + "\n")
                    self.results.insert(tk.INSERT, "Chi-squared statistic = %.4g"%chi + "\n")
                    self.results.insert(tk.INSERT, "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-(int(self.numVoigtVariable.get())*2+1))))
                    self.results.configure(state="disabled")
                    """
                    self.resultAlert.grid_remove()
                    
                    if (len(s) == 1):
                        self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "nan", "nan"))
                        for i in range(1, len(r), 2):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "nan", "nan"))
                            else:
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "nan", "nan"))
                            if (r[i+1] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "nan", "nan"))
                            else:
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "nan", "nan"))
                    else:
                        if (r[0] == 0):
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "0%"))
                        elif (abs(s[0]*2*100/r[0]) > 100 or np.isnan(s[0])):
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "%.3g"%(s[0]*2*100/r[0])+"%"), tags=("bad",))
                            self.resultAlert.grid(column=0, row=2, sticky="E")
                        else:
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "%.3g"%(s[0]*2*100/r[0])+"%"))
                        for i in range(1, len(r), 2):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "0%"))
                            else:
                                if (abs(s[i]*2*100/r[i]) > 100 or np.isnan(s[i])):
                                    self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"), tags=("bad",))
                                    self.resultAlert.grid(column=0, row=2, sticky="E")
                                else:
                                    self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"))
                            if (r[i+1] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "0%"))
                            else:
                                if (abs(s[i+1]*2*100/r[i+1]) > 100 or np.isnan(s[i])):
                                    self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "%.3g"%(s[i+1]*2*100/r[i+1])+"%"), tags=("bad",))
                                    self.resultAlert.grid(column=0, row=2, sticky="E")
                                else:
                                    self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "%.3g"%(s[i+1]*2*100/r[i+1])+"%"))
                    self.resultsView.tag_configure("bad", background="yellow")
                    self.whatFitStack.append(self.whatFit)
                    if (self.fitTypeVariable.get() == "Complex"):
                        self.whatFit = "C"
                    elif (self.fitTypeVariable.get() == "Imaginary"):
                        self.whatFit = "J"
                    elif (self.fitTypeVariable.get() == "Real"):
                        self.whatFit = "R"
                    if (self.magicPlot.state() != "withdrawn"):
                        self.magicInput.clf()
                        Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                        if (not self.capUsed):
                            for i in range(len(self.wdata)):
                                Zfit[i] = self.fits[0]
                                for k in range(1, len(self.fits), 2):
                                    Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        else:
                            for i in range(len(self.wdata)):
                                Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                                for k in range(1, len(self.fits), 2):
                                    Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        
                        x = np.array(self.rdata)
                        y = np.array(self.jdata)
                        dataColor = "tab:blue"
                        fitColor = "orange"
                        if (self.topGUI.getTheme() == "dark"):
                            dataColor = "cyan"
                            fitColor = "gold"
                        else:
                            dataColor = "tab:blue"
                            fitColor = "orange"
                        with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                            self.magicSubplot = self.magicInput.add_subplot(111)
                            self.magicSubplot.set_facecolor(self.backgroundColor)
                            self.magicSubplot.yaxis.set_ticks_position("both")
                            self.magicSubplot.yaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                            self.magicSubplot.xaxis.set_ticks_position("both")
                            self.magicSubplot.xaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                            pointsPlot, = self.magicSubplot.plot(x, -1*y, "o", color=dataColor)
                            rightPoint = max(x)
                            topPoint = max(-1*y)
                            if (len(self.fits) > 0):
                                self.magicSubplot.plot(np.array(Zfit.real), np.array(-1*Zfit.imag), color=fitColor)
                            self.magicSubplot.axis("equal")
                            self.magicSubplot.set_title("Magic Finger Nyquist Plot", color=self.foregroundColor)
                            self.magicSubplot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                            self.magicSubplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                            self.magicInput.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                        self.magicCanvasInput.draw()
                        annot = self.magicSubplot.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
                        annot.set_visible(False)
                        def update_annot(ind):
                            x,y = pointsPlot.get_data()
                            xval = x[ind["ind"][0]]
                            yval = y[ind["ind"][0]]
                            annot.xy = (xval, yval)
                            text = "Zr=%.3g"%xval + "\nZj=-%.3g"%yval + "\nf=%.5g"%self.wdata[np.where(self.rdata == xval)][0]
                            annot.set_text(text)
                            #---Check if we're within 5% of the right or top edges, and adjust label positions accordingly
                            if (rightPoint != 0):
                                if (abs(xval - rightPoint)/rightPoint <= 0.2):
                                    annot.set_position((-90, -10))
                            if (topPoint != 0):
                                if (abs(yval - topPoint)/topPoint <= 0.05):
                                    annot.set_position((10, -20))
                            else:
                                annot.set_position((10, -20))
                        def hover(event):
                            vis = annot.get_visible()
                            if event.inaxes == self.magicSubplot:
                                cont, ind = pointsPlot.contains(event)
                                if cont:
                                    update_annot(ind)
                                    annot.set_visible(True)
                                    self.magicInput.canvas.draw_idle()
                                else:
                                    if vis:
                                        annot.set_position((10,10))
                                        annot.set_visible(False)
                                        self.magicInput.canvas.draw_idle()
                        self.magicInput.canvas.mpl_connect("motion_notify_event", hover)
#                    self.editReVariable.set("%.4g"%r[0])
#                    self.needUndo = True
#                    if (self.justUndid):
#                        afterUndoRun()
            except queue.Empty:
                if (self.listPercent[0] != 1):
                    percent_to_step = (len(self.listPercent) - 1)/self.listPercent[0]
                    self.prog_bar.config(value=percent_to_step)
                    try:
                        self.taskbar.SetProgressValue(self.masterWindowId, (len(self.listPercent) - 1), self.listPercent[0])
                        self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                    except:
                        pass
                    if (self.listPercent[len(self.listPercent)-1] == 2):
                        self.progStatus.configure(text="Starting processes...")
                    elif (percent_to_step >= 1):
                        self.progStatus.configure(text="Calculating CIs...")
                    elif (percent_to_step > 0.001):
                        self.progStatus.configure(text="{:.1%}".format(percent_to_step))
                self.after(100, process_queue)
        
        def process_queue_cap():
            try:
                r, s, sdr, sdi, zz, zzs, zp, zps, chi, cor, aic, rc, sc, cor_cap = self.queue.get(0)
                self.measureModelButton.configure(state="normal")
                self.autoButton.configure(state="normal")
                self.magicButton.configure(state="normal")
                self.parametersButton.configure(state="normal")
                self.freqRangeButton.configure(state="normal")
                self.parametersLoadButton.configure(state="normal")
                self.numVoigtSpinbox.configure(state="readonly")
                self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
                self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtMinus.bind("<ButtonRelease-1>", minusNVE)
                self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
                self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
                self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
                self.numVoigtPlus.bind("<ButtonRelease-1>", plusNVE)
                self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
                self.browseButton.configure(state="normal")
                self.numMonteEntry.configure(state="normal")
                self.fitTypeCombobox.configure(state="readonly")
                self.weightingCombobox.configure(state="readonly")
                self.listPercent = []
                try:
                    self.progStatus.configure(text="Initializing...")
                    self.progStatus.grid_remove()
                except:
                    pass
                try:
                    self.taskbar.SetProgressState(self.masterWindowId, 0x0)
                except:
                    pass
                for delButton in self.paramDeleteButtons:
                    delButton.configure(state="normal")
                try:
                    self.addButton.configure(state="normal")
                    self.removeButton.configure(state="normal")
                except:
                    pass
                try:
                    self.autoRunButton.configure(state="normal")
                except:
                    pass
                self.prog_bar.stop()
                self.prog_bar.destroy()
                self.cancelButton.grid_remove()
                if (len(r) == 1 and len(s) == 1 and len(sdr) == 1 and len(sdi)==1 and len(zz)==1 and len(zzs)==1 and len(chi)==1):
                    if (r=="^" and s=="^" and sdr=="^" and sdi=="^" and zz=="^" and zzs=="^" and zp=="^" and zps=="^" and chi=="^"):
                        messagebox.showerror("Fitting error", "Error 11:\nThe fitting failed")
                    elif (r=="@" and s=="@" and sdr=="@" and sdi=="@" and zz=="@" and zzs=="@" and zp=="@" and zps=="@" and chi=="@"):
                        pass    #Fitting was cancelled
                else:
                    if (len(r) != 1 and len(s) == 1 and len(sdr) == 1 and len(sdi)==1 and len(zzs)==1):
                        if (s=="-" and sdr=="-" and sdi=="-" and zzs=="-" and zps=="-"):
                            s = np.zeros(len(r))
                            sdr = np.zeros(len(self.wdata))
                            sdi = np.zeros(len(self.wdata))
                            zzs = 0
                            zps = 0
                            messagebox.showwarning("Variance error", "Error 12:\nA minimization was found, but the standard deviations and confidence intervals could not be estimated. Try fitting again with these values or try fewer line shapes")
                    self.fitWeightR = np.zeros(len(self.wdata))
                    self.fitWeightJ = np.zeros(len(self.wdata))
                    if (self.weightingVariable.get() == "None"):
                        self.fitWeightR = np.ones(len(self.wdata))
                        self.fitWeightJ = np.ones(len(self.wdata))
                    elif (self.weightingVariable.get() == "Proportional"):
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = float(self.alphaVariable.get()) * self.rdata[i]
                        for i in range(len(self.wdata)):
                            self.fitWeightJ[i] = float(self.alphaVariable.get()) * self.jdata[i]
                    elif (self.weightingVariable.get() == "Modulus"):
                        for i in range(len(self.wdata)):
                            self.fitWeightR[i] = float(self.alphaVariable.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                            self.fitWeightJ[i] = float(self.alphaVariable.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                    elif (self.weightingVariable.get() == "Error model"):
                        for i in range(len(self.wdata)):
                            if (self.errorAlphaCheckboxVariable.get() == 1):
                                self.fitWeightR[i] += float(self.errorAlphaVariable.get())*abs(self.jdata[i])
                            if (self.errorBetaCheckboxVariable.get() == 1):
                                if (self.errorBetaReCheckboxVariable.get() == 1):
                                    self.fitWeightR[i] += (float(self.errorBetaVariable.get())*abs(self.rdata[i]) - float(self.errorBetaReVariable.get()))
                                else:
                                    self.fitWeightR[i] += float(self.errorBetaVariable.get())*abs(self.rdata[i])
                            if (self.errorGammaCheckboxVariable.get() == 1):
                                self.fitWeightR[i] += float(self.errorGammaVariable.get())*np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)**2
                            if (self.errorDeltaCheckboxVariable.get() == 1):
                                self.fitWeightR[i] += float(self.errorDeltaVariable.get())
                            self.fitWeightJ[i] = self.fitWeightR[i]
                    self.currentNVEPlotNeeded = (len(self.fits)-1)/2
                    if (len(self.fits) > 0):
                        self.undoStack.append(self.fits.copy())
                        self.undoSigmaStack.append(self.sigmas.copy())
                        self.undoAICStack.append(self.aicResult)
                        self.undoChiStack.append(self.chiSquared)
                        self.undoCorStack.append(self.corResult.copy())
                        self.undoCapNeededStack.append(self.capUsed)
                        self.undoCapStack.append(self.resultCap)
                        self.undoCapSigmaStack.append(self.sigmaCap)
                        self.undoCorCapStack.append(self.capCor)
                        self.undoUpDeleteStack.append(self.pastUpDelete)
                        self.undoLowDeleteStack.append(self.pastLowDelete)
                        self.undoSdrRealStack.append(self.sdrReal)
                        self.undoSdrImagStack.append(self.sdrImag)
                        self.pastUpDelete = self.upDelete
                        self.pastLowDelete = self.lowDelete
                        self.undoAlphaStack.append(self.pastAlphaVariable)
                        self.pastAlphaVariable = self.alphaVariable.get()
                        self.undoFitTypeStack.append(self.pastFitType)
                        self.pastFitType = self.fitTypeVariable.get()
                        self.undoNumSimulationsStack.append(self.pastNumSimulations)
                        self.pastNumSimulations = self.numMonteVariable.get()
                        self.undoWeightingStack.append(self.pastWeighting)
                        self.pastWeighting = self.weightingVariable.get()
                        self.undoErrorModelChoicesStack.append(self.pastErrorModelChoices)
                        self.pastErrorModelChoices = [self.errorAlphaCheckboxVariable.get(), self.errorBetaCheckboxVariable.get(), self.errorBetaReCheckboxVariable.get(), self.errorGammaCheckboxVariable.get(), self.errorDeltaCheckboxVariable.get()]
                        self.undoErrorModelValuesStack.append(self.pastErrorValues)
                        self.pastErrorValues = [self.errorAlphaVariable.get(), self.errorBetaVariable.get(), self.errorBetaReVariable.get(), self.errorGammaVariable.get(), self.errorDeltaVariable.get()]
                        self.undoButton.configure(state="normal")
#                        self.redoStack = []
#                        self.redoButton.configure(state="disabled")
                    else:
                        self.pastUpDelete = self.upDelete
                        self.pastLowDelete = self.lowDelete
                        self.pastAlphaVariable = self.alphaVariable.get()
                        self.pastFitType = self.fitTypeVariable.get()
                        self.pastNumSimulations = self.numMonteVariable.get()
                        self.pastWeighting = self.weightingVariable.get()
                        self.pastErrorModelChoices = [self.errorAlphaCheckboxVariable.get(), self.errorBetaCheckboxVariable.get(), self.errorGammaCheckboxVariable.get(), self.errorDeltaCheckboxVariable.get()]
                        self.pastErrorValues = [self.errorAlphaVariable.get(), self.errorBetaVariable.get(), self.errorGammaVariable.get(), self.errorDeltaVariable.get()]
                    self.capUsed = True
                    self.alreadyChanged = False
                    self.fits = r
                    self.sigmas = s
                    self.sdrReal = sdr
                    self.sdrImag = sdi
                    self.chiSquared = chi
                    self.aicResult = aic
                    self.corResult = cor
                    self.resultCap = rc
                    self.sigmaCap = sc
                    self.capCor = cor_cap.copy()
#                    self.resultData = r
#                    self.sigmaData = s
#                    self.sdrData = sdr
#                    self.sdiData = sdi
                    self.aR = ""
                    self.resultsView.delete(*self.resultsView.get_children())
                    for i in range(len(self.paramEntryVariables)):
                        self.paramEntryVariables[i].set("%.5g"%r[i])
                    self.capacitanceEntryVariable.set("%.5g"%rc)
                    self.resultsFrame.grid(column=0, row=6, sticky="W", pady=5)
                    self.graphFrame.grid(column=0, row=7, sticky="W", pady=5)
                    self.resultRe.configure(text="Ohmic Resistance = %.4g"%r[0] + " ± %.2g"%s[0])
                    self.resultRp.configure(text="Polarization Impedance = %.4g"%zp + " ± %.2g"%zps)
                    capacitances = np.zeros(int((len(r)-1)/2))
                    for i in range(1, len(r), 2):
                        if (r[1] == 0):
                            capacitances[int(i/2)] = 0
                        else:
                            capacitances[int(i/2)] = r[i+1]/r[i]
                    ceff = 1/(np.sum(1/capacitances) + 1/rc)
                    sigmaCapacitances = np.zeros(len(capacitances))
                    for i in range(1, len(r), 2):
                        if (r[i] == 0 or r[i+1] == 0):
                            sigmaCapacitances[int(i/2)] = 0
                        else:
                            sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((s[i+1]/r[i+1])**2 + (s[i]/r[i])**2)
                    #sigmaOtherC = (1/rc)*(sc/rc)
                    partCap = 0
                    for i in range(len(capacitances)):
                        partCap += sigmaCapacitances[i]**2/capacitances[i]**4
                    try:
                        partCap += sc**2/rc**4
                    except:
                        pass
                    sigmaCeff = ceff**2 * np.sqrt(partCap)
                    self.resultC.configure(text="Overall Capacitance = %.4g"%ceff + " ± %.2g"%sigmaCeff)
                    self.aR += "File name: " + self.browseEntry.get() + "\n"
                    self.aR += "Number of data: " + str(self.lengthOfData) + "\n"
                    self.aR += "Number of parameters: " + str(len(r)) + "\n"
                    self.aR += "\nRe (Rsol) = %.8g"%r[0] + " ± %.4g"%s[0] + "\n"
                    try:
                        self.aR += "Capacitance = %.8g"%rc + " ± %.4g"%sc + "\n"
                    except:
                        self.aR += "Capacitance = %.8g"%rc + " ± %.4g"%0 + "\n"
                    for i in range(1, len(r), 2):
                        self.aR += "R" + str(int(i/2 + 1)) + " = %.8g"%r[i] + " ± %.4g"%s[i] + "\n"
                        self.aR += "Tau" + str(int(i/2 + 1)) + " = %.8g"%r[i+1] + " ± %.4g"%s[i+1] + "\n"
                        self.aR += "C" + str(int(i/2 + 1)) + " = %.8g"%capacitances[int(i/2)] + " ± %.4g"%sigmaCapacitances[int(i/2)] + "\n"
                    self.aR += "\nZero frequency impedance = %.8g"%zz + " ± %.4g"%zzs + "\n"
                    self.aR += "Polarization Impedance = %.8g"%zp + " ± %.4g"%zps + "\n"
                    self.aR += "Overall Capacitance = %.8g"%ceff + " ± %.4g"%sigmaCeff + "\n"
                    self.aR += "\nCorrelation matrix:\n"
                    self.aR += "         Re    "
                    for i in range(1, len(r), 2):
                        if (int(i/2 +1) < 10):
                            self.aR += "  R" + str(int(i/2 + 1)) + "    " + " Tau" + str(int(i/2 + 1)) + "   "
                        else:
                            self.aR += "  R" + str(int(i/2 + 1)) + "   " + " Tau" + str(int(i/2 + 1)) + "  "
                    self.aR += "  Cap "
                    self.aR += "\n      \u250C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                    #for i in range(1, len(r), 2):
                    #    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                    self.aR += "\n"
                    self.aR += "Re    "
                    for i in range(len(r)):
                        if (i != 0):
                            if (i%2 != 0):
                                if (i/2 + 1 < 10):
                                    self.aR += "R" + str(int(i/2+1)) + "    "
                                else:
                                    self.aR += "R" + str(int(i/2+1)) + "   "
                            else:
                                if (i/2 < 10):
                                    self.aR += "Tau" + str(int(i/2)) + "  "
                                else:
                                    self.aR += "Tau" + str(int(i/2)) + " "
                        self.aR += "\u2502"
                        for j in range(len(r)):
                            if (j <= i):
                                if ("%05.2f"%cor[i][j] == "01.00"):
                                    self.aR += "   1   \u2502"
                                elif (np.isnan(cor[i][j])):
                                    self.aR += "   0   \u2502"
                                else:
                                    self.aR += " %05.2f"%cor[i][j] + " \u2502"
                        if (i != len(r)-1):
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                            self.aR += "\u253C"
                            if (i == 0):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        else:
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253C"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        self.aR += "\n"
                    self.aR += "Cap   \u2502"
                    for i in range(len(cor_cap)):
                        if ("%05.2f"%cor_cap[i] == "01.00"):
                            self.aR += "   1   \u2502"
                        elif (np.isnan(cor_cap[i])):
                            self.aR += "   0   \u2502"
                        else:
                            self.aR += " %05.2f"%cor_cap[i] + " \u2502"
                    self.aR += "   1   \u2502"
                    self.aR += "\n      \u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                    for i in range(len(cor_cap)-1):
                        self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2518\n"
                    
                    self.aR += "\nChi-squared statistic = %.4g"%chi + "\n"
                    self.aR += "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-len(r))) + "\n"
                    try:
                        self.aR += "Akaike Information Criterion = %.4g"%aic#(np.log(chi*((1+2*len(r))/self.lengthOfData))) + "\n"
                    except:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    #self.aR += "Akaike Performance Index = %.4g"%(chi*((1+(len(r)/self.lengthOfData))/(1-(len(r)/self.lengthOfData))))
                    """
                    self.results.configure(state="normal")
                    self.results.insert(tk.INSERT, "Re (Rsol) = %.4g"%r[0] + " ± %.1g"%s[0] +"\n")
                    for i in range(1, len(r), 2):
                        self.results.insert(tk.INSERT, "Rt" + str(int(i/2 + 1)) + " = %.4g"%r[i] + " ± %.1g"%s[i] + "\n")
                        self.results.insert(tk.INSERT, "Tau" + str(int(i/2+1)) + " = %.4g"%r[i+1] + " ± %.1g"%s[i+1] + "\n")
                    self.results.insert(tk.INSERT, "\nZero frequency impedance = %.4g"%zz + " ± %.1g"%zzs + "\n")
                    self.results.insert(tk.INSERT, "Polarization impedance = %.4g"%zp + " ± %.1g"%zps + "\n")
                    self.results.insert(tk.INSERT, "Chi-squared statistic = %.4g"%chi + "\n")
                    self.results.insert(tk.INSERT, "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-(int(self.numVoigtVariable.get())*2+1))))
                    self.results.configure(state="disabled")
                    """
                    self.resultAlert.grid_remove()
                    if (len(s) == 1):
                        self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "nan", "nan"))
                        self.resultsView.insert("", tk.END, text="", values=("Capacitance", "%.5g"%rc, "nan", "nan"))
                        for i in range(1, len(r), 2):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "nan", "nan"))
                            else:
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "nan", "nan"))
                            if (r[i+1] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "nan", "nan"))
                            else:
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "nan", "nan"))
                    else:
                        if (r[0] == 0):
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "0%"))
                        elif (abs(s[0]*2*100/r[0]) > 100 or np.isnan(s[0])):
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "%.3g"%(s[0]*2*100/r[0])+"%"), tags=("bad",))
                            self.resultAlert.grid(column=0, row=2, sticky="E")
                        else:
                            self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%r[0], "%.3g"%s[0], "%.3g"%(s[0]*2*100/r[0])+"%"))
                        try:
                            if (abs(sc*2*100/rc) > 100 or np.isnan(sc)):
                                self.resultsView.insert("", tk.END, text="", values=("Capacitance", "%.5g"%rc, "%.3g"%sc, "%.3g"%(sc*2*100/rc)+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=("Capacitance", "%.5g"%rc, "%.3g"%sc, "%.3g"%(sc*2*100/rc)+"%"))
                        except:
                            self.resultsView.insert("", tk.END, text="", values=("Capacitance", "%.5g"%rc, "%.3g"%0, "%.3g"%(0*2*100/rc)+"%"))
                        for i in range(1, len(r), 2):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "0%"))
                            else:
                                if (abs(s[i]*2*100/r[i]) > 100 or np.isnan(s[i])):
                                    self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"), tags=("bad",))
                                    self.resultAlert.grid(column=0, row=2, sticky="E")
                                else:
                                    self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"))
                            if (r[i+1] == 0):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "0%"))
                            else:
                                if (abs(s[i+1]*2*100/r[i+1]) > 100 or np.isnan(s[i])):
                                    self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "%.3g"%(s[i+1]*2*100/r[i+1])+"%"), tags=("bad",))
                                    self.resultAlert.grid(column=0, row=2, sticky="E")
                                else:
                                    self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%r[i+1], "%.3g"%s[i+1], "%.3g"%(s[i+1]*2*100/r[i+1])+"%"))
                    self.resultsView.tag_configure("bad", background="yellow")
                    self.whatFitStack.append(self.whatFit)
                    if (self.fitTypeVariable.get() == "Complex"):
                        self.whatFit = "C"
                    elif (self.fitTypeVariable.get() == "Imaginary"):
                        self.whatFit = "J"
                    elif (self.fitTypeVariable.get() == "Real"):
                        self.whatFit = "R"
                    if (self.magicPlot.state() != "withdrawn"):
                        self.magicInput.clf()
                        Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                        if (not self.capUsed):
                            for i in range(len(self.wdata)):
                                Zfit[i] = self.fits[0]
                                for k in range(1, len(self.fits), 2):
                                    Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        else:
                            for i in range(len(self.wdata)):
                                Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                                for k in range(1, len(self.fits), 2):
                                    Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        
                        x = np.array(self.rdata)    #Doesn't plot without this
                        y = np.array(self.jdata)
                        dataColor = "tab:blue"
                        fitColor = "orange"
                        if (self.topGUI.getTheme() == "dark"):
                            dataColor = "cyan"
                            fitColor = "gold"
                        else:
                            dataColor = "tab:blue"
                            fitColor = "orange"
                        with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                            self.magicSubplot = self.magicInput.add_subplot(111)
                            self.magicSubplot.set_facecolor(self.backgroundColor)
                            self.magicSubplot.yaxis.set_ticks_position("both")
                            self.magicSubplot.yaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                            self.magicSubplot.xaxis.set_ticks_position("both")
                            self.magicSubplot.xaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                            pointsPlot, = self.magicSubplot.plot(x, -1*y, "o", color=dataColor)
                            rightPoint = max(x)
                            topPoint = max(-1*y)
                            if (len(self.fits) > 0):
                                self.magicSubplot.plot(np.array(Zfit.real), np.array(-1*Zfit.imag), color=fitColor)
                            self.magicSubplot.axis("equal")
                            self.magicSubplot.set_title("Magic Finger Nyquist Plot", color=self.foregroundColor)
                            self.magicSubplot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                            self.magicSubplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                            self.magicInput.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                        self.magicCanvasInput.draw()
                        annot = self.magicSubplot.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
                        annot.set_visible(False)
                        def update_annot(ind):
                            x,y = pointsPlot.get_data()
                            xval = x[ind["ind"][0]]
                            yval = y[ind["ind"][0]]
                            annot.xy = (xval, yval)
                            text = "Zr=%.3g"%xval + "\nZj=-%.3g"%yval + "\nf=%.5g"%self.wdata[np.where(self.rdata == xval)][0]
                            annot.set_text(text)
                            #---Check if we're within 5% of the right or top edges, and adjust label positions accordingly
                            if (rightPoint != 0):
                                if (abs(xval - rightPoint)/rightPoint <= 0.2):
                                    annot.set_position((-90, -10))
                            if (topPoint != 0):
                                if (abs(yval - topPoint)/topPoint <= 0.05):
                                    annot.set_position((10, -20))
                            else:
                                annot.set_position((10, -20))
                        def hover(event):
                            vis = annot.get_visible()
                            if event.inaxes == self.magicSubplot:
                                cont, ind = pointsPlot.contains(event)
                                if cont:
                                    update_annot(ind)
                                    annot.set_visible(True)
                                    self.magicInput.canvas.draw_idle()
                                else:
                                    if vis:
                                        annot.set_position((10,10))
                                        annot.set_visible(False)
                                        self.magicInput.canvas.draw_idle()
                        self.magicInput.canvas.mpl_connect("motion_notify_event", hover)
#                    self.editReVariable.set("%.4g"%r[0])
#                    self.needUndo = True
#                    if (self.justUndid):
#                        afterUndoRun()
            except queue.Empty:
                if (self.listPercent[0] != 1):
                    percent_to_step = (len(self.listPercent) - 1)/self.listPercent[0]
                    self.prog_bar.config(value=percent_to_step)
                    try:
                        self.taskbar.SetProgressValue(self.masterWindowId, (len(self.listPercent) - 1), self.listPercent[0])
                        self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                    except:
                        pass
                    if (self.listPercent[len(self.listPercent)-1] == 2):
                        self.progStatus.configure(text="Starting processes...")
                    elif (percent_to_step >= 1):
                        self.progStatus.configure(text="Calculating CIs...")
                    elif (percent_to_step > 0.001):
                        self.progStatus.configure(text="{:.1%}".format(percent_to_step))
                self.after(100, process_queue_cap)
    
        def terminateRun(running_thread):
            #self.keepRunning = False
            running_thread.terminate()
        
        def runMeasurementModel():
            try:
                self.masterWindowId = ctypes.windll.user32.GetForegroundWindow()
            except:
                pass
            if (self.browseEntry.get() != ""):
                if (self.capacitanceCheckboxVariable.get() == 0):
                    #self.results.configure(state="normal")
                    #self.results.delete("1.0", tk.END)
                    #self.results.configure(state="disabled")
                    nVE = int(self.numVoigtVariable.get())
                    nMC = int(self.numMonteEntry.get())
                    c = 0
                    fT = 0
                    bL = np.zeros(int(self.numVoigtVariable.get())*2 + 1)
                    bU = np.zeros(int(self.numVoigtVariable.get())*2 + 1)
                    try:
                        if (self.paramComboboxVariables[0].get() == "+" and float(self.paramEntryVariables[0].get()) < 0):
                            messagebox.showerror("Value error", "Error 18:\nRe has a negative value but is constrained to be positive")
                            return
                    except:
                        pass
                    bL[0] = np.NINF if (self.paramComboboxVariables[0].get() == "+ or -" or self.paramComboboxVariables[0].get() == "fixed") else 0
                    bU[0] = np.inf if (self.paramComboboxVariables[0].get() == "+ or -" or self.paramComboboxVariables[0].get() == "fixed" or self.paramComboboxVariables[0].get() == "+") else 0
                    const = [-1]
                    numCombo = 1
                    const[0] = 0 if (self.paramComboboxVariables[0].get() == "fixed") else -1
                    #print(self.paramComboboxVariables)
                    for comboboxVariable in self.paramComboboxVariables[1:]:
                        bU[numCombo+1] = np.inf
                        if (comboboxVariable.get() == "+ or -" or comboboxVariable.get() == "+"):
                            bU[numCombo] = np.inf
                        if (comboboxVariable.get() == "+ or -" or comboboxVariable.get() == "-"):
                            bL[numCombo] = np.NINF
                        elif (comboboxVariable.get() == "fixed"):
                            bL[numCombo] = np.NINF
                            bU[numCombo] = np.inf
                            if (const[0] == -1):
                                const[0] = numCombo
                            else:
                                const.append(numCombo)
                        numCombo += 2
                    numCombo2 = 2
                    for comboboxVariable in self.tauComboboxVariables:
                        if (comboboxVariable.get() == "fixed"):
                            if (const[0] == -1):
                                const[0] = numCombo2
                            else:
                                const.append(numCombo2)
                        numCombo2 += 2
                    ig = []
                    for i in range(int(self.numVoigtVariable.get())*2+1):
                        ig.append([])
                    for i in range(len(self.paramEntryVariables)):
                        possibleVal = self.paramEntryVariables[i].get()
                        if (possibleVal == '' or possibleVal == '.' or possibleVal == 'E' or possibleVal == 'e'):
                            messagebox.showerror("Value error", "Error 19:\nOne of the parameters is missing a value")
                            return
                        try:
                            ig[i].append(float(possibleVal))
                            if (ig[i][0] < 0 and bL[i] == 0):
                                messagebox.showerror("Value error", "Error 20:\nOne of the parameters has a negative value but is constrained to be positive")
                                return
                            elif (ig[i][0] > 0 and bU[i] == 0):
                                messagebox.showerror("Value error", "Error 20: \nOne of the parameters has a positive value but is constrained to be negative")
                                return
                        except:
                            messagebox.showerror("Value error", "Error 21:\nOne of the parameters has an invalid value: " + str(possibleVal))
                            return
                    try:
                        for i in range(len(self.multistartEnabled)):
                            if (self.multistartEnabled[i].get()):
                                if (self.multistartSelection[i].get() == "Logarithmic"):
                                    if (np.sign(float(self.multistartLower[i].get())) != np.sign(float(self.multistartUpper[i].get()))):
                                        messagebox.showerror("Value error", "Error 58: \nFor logarithmic spacing, the lower and upper bound must have the same sign and both be nonzero")
                                        return
                                    if (float(self.multistartLower[i].get()) == 0 or float(self.multistartUpper[i].get()) == 0):
                                        messagebox.showerror("Value error", "Error 59: \nFor logarithmic spacing, 0 cannot be the upper or lower bound")
                                        return
                                    ig[i].extend(np.geomspace(float(self.multistartLower[i].get()), float(self.multistartUpper[i].get()), int(self.multistartNumber[i].get())))
                                if (self.multistartSelection[i].get() == "Linear"):
                                    ig[i].extend(np.linspace(float(self.multistartLower[i].get()), float(self.multistartUpper[i].get()), int(self.multistartNumber[i].get())))
                                if (self.multistartSelection[i].get() == "Random"):
                                    ig[i].extend((float(self.multistartUpper[i].get())-float(self.multistartLower[i].get()))*np.random.rand(int(self.multistartNumber[i].get()))+float(self.multistartLower[i].get()))
                                if (self.multistartSelection[i].get() == "Custom"):
                                    ig[i].extend(float(val) for val in [x.strip() for x in self.multistartCustom[i].get().split(',')])
                    except:
                        messagebox.showerror("Value error", "There was an invalid value in the multistart options")
                        return
                    for i in range(len(ig)):
                        if (len(ig[i]) != 1 and i in const):
                            messagebox.showerror("Value error", "Multistart is enabled for a fixed parameter")
                            return
                            for j in range(len(ig[i])):
                                if (ig[i][j] > 0 and bL[i] == 0):
                                    messagebox.showerror("Value error", "One of the multistart parameters has a negative value, but the parameter is constrained to be positive")
                                    return
                                elif (ig[i][j] < 0 and bU[i] == 0):
                                    messagebox.showerror("Value error", "One of the multistart parameters has a positive value, but the parameter is constrained to be negative")
                                    return
                    error_alpha = 0
                    error_beta = 0
                    error_betaRe = 0
                    error_gamma = 0
                    error_delta = 0
                    if (self.weightingVariable.get() == "Modulus"):
                        c = 1
                    elif (self.weightingVariable.get() == "Error model"):
                        if (self.errorAlphaCheckboxVariable.get() != 1 and self.errorBetaCheckboxVariable.get() != 1 and self.errorGammaCheckboxVariable.get() != 1 and self.errorDeltaCheckboxVariable.get() != 1):
                            messagebox.showerror("Error structure error", "Error 22:\nAt least one parameter must be chosen to use error structure weighting")
                            return
                        try:
                            if (self.errorAlphaCheckboxVariable.get() == 1):
                                if (self.errorAlphaVariable.get() == "" or self.errorAlphaVariable.get() == " " or self.errorAlphaVariable.get() == "."):
                                    error_alpha = 0
                                else:
                                    error_alpha = float(self.errorAlphaVariable.get())
                            if (self.errorBetaCheckboxVariable.get() == 1):
                                if (self.errorBetaVariable.get() == "" or self.errorBetaVariable.get() == " " or self.errorBetaVariable.get() == "."):
                                    error_beta = 0
                                else:
                                    error_beta = float(self.errorBetaVariable.get())
                            if (self.errorBetaReCheckboxVariable.get() == 1):
                                if (self.errorBetaReVariable.get() == "" or self.errorBetaReVariable.get() == " " or self.errorBetaReVariable.get() == "."):
                                    error_betaRe = 0
                                else:
                                    error_betaRe = float(self.errorBetaReVariable.get())
                            if (self.errorGammaCheckboxVariable.get() == 1):
                                if (self.errorGammaVariable.get() == "" or self.errorGammaVariable.get() == " " or self.errorGammaVariable.get() == "."):
                                    error_gamma = 0
                                else:
                                    error_gamma = float(self.errorGammaVariable.get())
                            if (self.errorDeltaCheckboxVariable.get() == 1):
                                if (self.errorDeltaVariable.get() == "" or self.errorDeltaVariable.get() == " " or self.errorDeltaVariable.get() == "."):
                                    error_delta = 0
                                else:
                                    error_delta = float(self.errorDeltaVariable.get())
                        except:
                            messagebox.showerror("Value error", "Error 23:\nOne of the error structure parameters has an invalid value")
                            return
                        if (error_alpha == 0 and error_beta == 0 and error_gamma == 0 and error_delta == 0):
                            messagebox.showerror("Value error", "Error 24:\nAt least one of the error structure parameters must be nonzero")
                            return
                        c = 3
                    elif (self.weightingVariable.get() == "Proportional"):
                        c = 2
                    if (self.fitTypeVariable.get() == "Real"):
                        fT = 2
                    elif (self.fitTypeVariable.get() == "Imaginary"):
                        fT = 1
                    try:
                        aN = float(self.alphaVariable.get()) if (self.alphaVariable.get() != "") else 1
                    except:
                        messagebox.showerror("Value error", "Error 25:\nThe assumed noise (alpha) has an invalid value")
                        return
                    if (aN <= 0):
                        messagebox.showerror("Value error", "The assumed noise (alpha) must be greater than 0")
                        return
                    self.measureModelButton.configure(state="disabled")
                    self.autoButton.configure(state="disabled")
                    self.magicButton.configure(state="disabled")
                    self.parametersButton.configure(state="disabled")
                    self.freqRangeButton.configure(state="disabled")
                    self.parametersLoadButton.configure(state="disabled")
                    self.numVoigtSpinbox.configure(state="disabled")
                    self.browseButton.configure(state="disabled")
                    self.numMonteEntry.configure(state="disabled")
                    self.fitTypeCombobox.configure(state="disabled")
                    self.weightingCombobox.configure(state="disabled")
                    self.undoButton.configure(state="disabled")
                    self.numVoigtMinus.configure(cursor="arrow", bg="gray60")
                    self.numVoigtMinus.unbind("<Enter>")
                    self.numVoigtMinus.unbind("<Leave>")
                    self.numVoigtMinus.unbind("<Button-1>")
                    self.numVoigtMinus.unbind("<ButtonRelease-1>")
                    self.numVoigtPlus.configure(cursor="arrow", bg="gray60")
                    self.numVoigtPlus.unbind("<Enter>")
                    self.numVoigtPlus.unbind("<Leave>")
                    self.numVoigtPlus.unbind("<Button-1>")
                    self.numVoigtPlus.unbind("<ButtonRelease-1>")
                    for delButton in self.paramDeleteButtons:
                        delButton.configure(state="disabled")
                    try:
                        self.addButton.configure(state="disabled")
                        self.removeButton.configure(state="disabled")
                    except:
                        pass
                    try:
                        self.autoRunButton.configure(state="disabled")
                    except:
                        pass
                    if (len(max(ig, key=len)) > 1):
                        self.prog_bar = ttk.Progressbar(self.measureModelFrame, orient="horizontal", length=130, maximum=1, mode="determinate")
                        self.progStatus = tk.Label(self.measureModelFrame, text="Initializing...", bg=self.backgroundColor, fg=self.foregroundColor)
                        self.progStatus.grid(column=5, row=0, padx=5)
                        try:
                            cc.GetModule('tl.tlb')
                            import comtypes.gen.TaskbarLib as tbl
                            self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                            self.taskbar.HrInit()
                            self.taskbar.ActivateTab(self.masterWindowId)
                            self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                        except:
                            pass
                    else:
                        self.prog_bar = ttk.Progressbar(self.measureModelFrame, orient="horizontal", length=130, mode="indeterminate")
                        try:
                            cc.GetModule('tl.tlb')
                            import comtypes.gen.TaskbarLib as tbl
                            self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                            self.taskbar.HrInit()
                            self.taskbar.ActivateTab(self.masterWindowId)
                            self.taskbar.SetProgressState(self.masterWindowId, 0x1)
                        except:
                            pass
                    self.cancelButton = tk.Label(self.measureModelFrame, text="×", font=('Helvetica', 12, 'bold'), fg=self.foregroundColor, bg=self.backgroundColor, cursor="hand2")
                    self.cancelButton.grid(column=4, row=0)
                    self.prog_bar.grid(column=3, row=0, padx=5)
                    if (len(max(ig, key=len)) > 1):
                        pass
                    else:
                        self.prog_bar.start(40)
                    
                    cancel_ttp = CreateToolTip(self.cancelButton, text="Cancel current fitting", color="red", resetColor=self.foregroundColor)
    
                    self.queue = queue.Queue()
                    numCombos = 1
                    for g in ig:
                        numCombos *= len(g)
                    self.listPercent.append(numCombos)
                    self.currentThreads.append(ThreadedTask(self.queue, self.wdata, self.rdata, self.jdata, nVE, nMC, c, aN, 1, fT, bL, ig, const, bU, self.listPercent, error_alpha, error_beta, error_betaRe, error_gamma, error_delta))
                    self.currentThreads[len(self.currentThreads)-1].start()
                    self.cancelButton.bind("<Button-1>", lambda e: terminateRun(self.currentThreads[len(self.currentThreads)-1]))
                    self.after(100, process_queue)
                else:   #Capacitance is chosen
                    nVE = int(self.numVoigtVariable.get())
                    nMC = int(self.numMonteEntry.get())
                    c = 0
                    fT = 0
                    bL = np.zeros(int(self.numVoigtVariable.get())*2 + 1)
                    bU = np.zeros(int(self.numVoigtVariable.get())*2 + 1)
                    try:
                        if (self.paramComboboxVariables[0].get() == "+" and float(self.paramEntryVariables[0].get()) < 0):
                            messagebox.showerror("Value error", "Error 18:\nRe has a negative value but is constrained to be positive")
                            return
                    except:
                        pass
                    bL[0] = np.NINF if (self.paramComboboxVariables[0].get() == "+ or -" or self.paramComboboxVariables[0].get() == "fixed") else 0
                    bU[0] = np.inf if (self.paramComboboxVariables[0].get() == "+ or -" or self.paramComboboxVariables[0].get() == "fixed" or self.paramComboboxVariables[0].get() == "+") else 0
                    const = [-1]
                    numCombo = 1
                    const[0] = 0 if (self.paramComboboxVariables[0].get() == "fixed") else -1
                    #print(self.paramComboboxVariables)
                    for comboboxVariable in self.paramComboboxVariables[1:]:
                        bU[numCombo+1] = np.inf
                        if (comboboxVariable.get() == "+ or -" or comboboxVariable.get() == "+"):
                            bU[numCombo] = np.inf
                        if (comboboxVariable.get() == "+ or -" or comboboxVariable.get() == "-"):
                            bL[numCombo] = np.NINF
                        elif (comboboxVariable.get() == "fixed"):
                            bL[numCombo] = np.NINF
                            bU[numCombo] = np.inf
                            if (const[0] == -1):
                                const[0] = numCombo
                            else:
                                const.append(numCombo)
                        numCombo += 2
                    numCombo2 = 2
                    for comboboxVariable in self.tauComboboxVariables:
                        if (comboboxVariable.get() == "fixed"):
                            if (const[0] == -1):
                                const[0] = numCombo2
                            else:
                                const.append(numCombo2)
                        numCombo2 += 2
                    
                    ig = []
                    for i in range(int(self.numVoigtVariable.get())*2+1):
                        ig.append([])
                    for i in range(len(self.paramEntryVariables)):
                        possibleVal = self.paramEntryVariables[i].get()
                        if (possibleVal == '' or possibleVal == '.' or possibleVal == 'E' or possibleVal == 'e'):
                            messagebox.showerror("Value error", "Error 19:\nOne of the parameters is missing a value")
                            return
                        try:
                            ig[i].append(float(possibleVal))
                            if (ig[i][0] < 0 and bL[i] == 0):
                                messagebox.showerror("Value error", "Error 20:\nOne of the parameters has a negative value but is constrained to be positive")
                                return
                            elif (ig[i][0] > 0 and bU[i] == 0):
                                messagebox.showerror("Value error", "One of the parameters has a positive value but is constrained to be negative")
                                return
                        except:
                            messagebox.showerror("Value error", "Error 21:\nOne of the parameters has an invalid value: " + str(possibleVal))
                            return
                    try:
                        for i in range(len(self.multistartEnabled)):
                            if (self.multistartEnabled[i].get()):
                                if (self.multistartSelection[i].get() == "Logarithmic"):
                                    if (np.sign(float(self.multistartLower[i].get())) != np.sign(float(self.multistartUpper[i].get()))):
                                        messagebox.showerror("Value error", "For logarithmic spacing, the lower and upper bound must have the same sign and both be nonzero")
                                        return
                                    if (float(self.multistartLower[i].get()) == 0 or float(self.multistartUpper[i].get()) == 0):
                                        messagebox.showerror("Value error", "For logarithmic spacing, 0 cannot be the upper or lower bound")
                                        return
                                    ig[i].extend(np.geomspace(float(self.multistartLower[i].get()), float(self.multistartUpper[i].get()), int(self.multistartNumber[i].get())))
                                if (self.multistartSelection[i].get() == "Linear"):
                                    ig[i].extend(np.linspace(float(self.multistartLower[i].get()), float(self.multistartUpper[i].get()), int(self.multistartNumber[i].get())))
                                if (self.multistartSelection[i].get() == "Random"):
                                    ig[i].extend((float(self.multistartUpper[i].get())-float(self.multistartLower[i].get()))*np.random.rand(int(self.multistartNumber[i].get()))+float(self.multistartLower[i].get()))
                                if (self.multistartSelection[i].get() == "Custom"):
                                    ig[i].extend(int(val) for val in [x.strip() for x in self.multistartCustom[i].get().split(',')])
                    except:
                        messagebox.showerror("Value error", "There was an invalid value in the multistart options")
                        return
                    for i in range(len(ig)):
                        if (len(ig[i]) != 1 and i in const):
                            messagebox.showerror("Value error", "Multistart is enabled for a fixed parameter")
                            return
                            for j in range(len(ig[i])):
                                if (ig[i][j] > 0 and bL[i] == 0):
                                    messagebox.showerror("Value error", "One of the multistart parameters has a negative value, but the parameter is constrained to be positive")
                                    return
                                elif (ig[i][j] < 0 and bU[i] == 0):
                                    messagebox.showerror("Value error", "One of the multistart parameters has a positive value, but the parameter is constrained to be negative")
                                    return
                    
                    g = np.zeros(int(self.numVoigtVariable.get())*2+1)
                    g[0] = self.rDefault
                    for i in range(1, len(g), 2):
                        g[i] = self.rDefault
                        g[i+1] = self.tauDefault
                    for i in range(len(self.paramEntryVariables)):
                        possibleVal = self.paramEntryVariables[i].get()
                        if (possibleVal == '' or possibleVal == '.' or possibleVal == 'E' or possibleVal == 'e'):
                            messagebox.showerror("Value error", "Error 19:\nOne of the parameters is missing a value")
                            return
                        try:
                            g[i] = float(possibleVal)
                            if (g[i] < 0 and bL[i] == 0):
                                messagebox.showerror("Value error", "Error 20:\nOne of the parameters has a negative value but is constrained to be positive")
                                return
                            elif (g[i] > 0 and bU[i] == 0):
                                messagebox.showerror("Value error", "One of the parameters has a positive value but is constrained to be negative")
                                return
                        except:
                            messagebox.showerror("Value error", "Error 21:\nOne of the parameters has an invalid value: " + str(possibleVal))
                            return
                    error_alpha = 0
                    error_beta = 0
                    error_betaRe = 0
                    error_gamma = 0
                    error_delta = 0
                    if (self.weightingVariable.get() == "Modulus"):
                        c = 1
                    elif (self.weightingVariable.get() == "Error model"):
                        if (self.errorAlphaCheckboxVariable.get() != 1 and self.errorBetaCheckboxVariable.get() != 1 and self.errorGammaCheckboxVariable.get() != 1 and self.errorDeltaCheckboxVariable.get() != 1):
                            messagebox.showerror("Error structure error", "Error 22:\nAt least one parameter must be chosen to use error structure weighting")
                            return
                        try:
                            if (self.errorAlphaCheckboxVariable.get() == 1):
                                if (self.errorAlphaVariable.get() == "" or self.errorAlphaVariable.get() == " " or self.errorAlphaVariable.get() == "."):
                                    error_alpha = 0
                                else:
                                    error_alpha = float(self.errorAlphaVariable.get())
                            if (self.errorBetaCheckboxVariable.get() == 1):
                                if (self.errorBetaVariable.get() == "" or self.errorBetaVariable.get() == " " or self.errorBetaVariable.get() == "."):
                                    error_beta = 0
                                else:
                                    error_beta = float(self.errorBetaVariable.get())
                            if (self.errorBetaReCheckboxVariable.get() == 1):
                                if (self.errorBetaReVariable.get() == "" or self.errorBetaReVariable.get() == " " or self.errorBetaReVariable.get() == "."):
                                    error_betaRe = 0
                                else:
                                    error_betaRe = float(self.errorBetaReVariable.get())
                            if (self.errorGammaCheckboxVariable.get() == 1):
                                if (self.errorGammaVariable.get() == "" or self.errorGammaVariable.get() == " " or self.errorGammaVariable.get() == "."):
                                    error_gamma = 0
                                else:
                                    error_gamma = float(self.errorGammaVariable.get())
                            if (self.errorDeltaCheckboxVariable.get() == 1):
                                if (self.errorDeltaVariable.get() == "" or self.errorDeltaVariable.get() == " " or self.errorDeltaVariable.get() == "."):
                                    error_delta = 0
                                else:
                                    error_delta = float(self.errorDeltaVariable.get())
                        except:
                            messagebox.showerror("Value error", "Error 23:\nOne of the error structure parameters has an invalid value")
                            return
                        if (error_alpha == 0 and error_beta == 0 and error_gamma == 0 and error_delta == 0):
                            messagebox.showerror("Value error", "Error 24:\nAt least one of the error structure parameters must be nonzero")
                            return
                        c = 3
                    elif (self.weightingVariable.get() == "Proportional"):
                        c = 2
                    if (self.fitTypeVariable.get() == "Real"):
                        fT = 2
                    elif (self.fitTypeVariable.get() == "Imaginary"):
                        fT = 1
                    try:
                        aN = float(self.alphaVariable.get()) if (self.alphaVariable.get() != "") else 1
                    except:
                        messagebox.showerror("Value error", "Error 25:\nThe assumed noise (alpha) has an invalid value")
                        return
                    try:
                        if (float(self.capacitanceEntryVariable.get()) < 0 and self.capacitanceComboboxVariable.get() == "+"):
                            messagebox.showerror("Value error", "The capacitance has a negative value but is constrained to be positive")
                            return
                        elif (float(self.capacitanceEntryVariable.get()) > 0 and self.capacitanceComboboxVariable.get() == "-"):
                            messagebox.showerror("Value error", "The capacitance has a positive value but is constrained to be negative")
                            return
                        cG = float(self.capacitanceEntryVariable.get())
                        bLC = 0
                        bUC = 0
                        if (self.capacitanceComboboxVariable.get() == "+"):
                            bUC = np.inf
                        elif (self.capacitanceComboboxVariable.get() == "-"):
                            bLC = np.NINF
                        elif (self.capacitanceComboboxVariable.get() == "+ or -" or self.capacitanceComboboxVariable.get() == "fixed"):
                            bLC = np.NINF
                            bUC = np.inf                       
                        fC = True if self.capacitanceComboboxVariable.get() == "fixed" else False
                    except:
                        messagebox.showerror("Value error", "The capacitance has an invalid value")
                        return
                    self.measureModelButton.configure(state="disabled")
                    self.autoButton.configure(state="disabled")
                    self.magicButton.configure(state="disabled")
                    self.parametersButton.configure(state="disabled")
                    self.freqRangeButton.configure(state="disabled")
                    self.parametersLoadButton.configure(state="disabled")
                    self.numVoigtSpinbox.configure(state="disabled")
                    self.browseButton.configure(state="disabled")
                    self.numMonteEntry.configure(state="disabled")
                    self.fitTypeCombobox.configure(state="disabled")
                    self.weightingCombobox.configure(state="disabled")
                    self.undoButton.configure(state="disabled")
                    self.numVoigtMinus.configure(cursor="arrow", bg="gray60")
                    self.numVoigtMinus.unbind("<Enter>")
                    self.numVoigtMinus.unbind("<Leave>")
                    self.numVoigtMinus.unbind("<Button-1>")
                    self.numVoigtMinus.unbind("<ButtonRelease-1>")
                    self.numVoigtPlus.configure(cursor="arrow", bg="gray60")
                    self.numVoigtPlus.unbind("<Enter>")
                    self.numVoigtPlus.unbind("<Leave>")
                    self.numVoigtPlus.unbind("<Button-1>")
                    self.numVoigtPlus.unbind("<ButtonRelease-1>")
                    for delButton in self.paramDeleteButtons:
                        delButton.configure(state="disabled")
                    try:
                        self.addButton.configure(state="disabled")
                        self.removeButton.configure(state="disabled")
                    except:
                        pass
                    try:
                        self.autoRunButton.configure(state="disabled")
                    except:
                        pass
                    if (len(max(ig, key=len)) > 1):
                        self.prog_bar = ttk.Progressbar(self.measureModelFrame, orient="horizontal", length=130, maximum=1, mode="determinate")
                        self.progStatus = tk.Label(self.measureModelFrame, text="Initializing...", bg=self.backgroundColor, fg=self.foregroundColor)
                        self.progStatus.grid(column=5, row=0, padx=5)
                        try:
                            cc.GetModule('tl.tlb')
                            import comtypes.gen.TaskbarLib as tbl
                            self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                            self.taskbar.HrInit()
                            self.taskbar.ActivateTab(self.masterWindowId)
                            self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                        except:
                            pass
                    else:
                        self.prog_bar = ttk.Progressbar(self.measureModelFrame, orient="horizontal", length=130, mode="indeterminate")
                        try:
                            cc.GetModule('tl.tlb')
                            import comtypes.gen.TaskbarLib as tbl
                            self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                            self.taskbar.HrInit()
                            self.taskbar.ActivateTab(self.masterWindowId)
                            self.taskbar.SetProgressState(self.masterWindowId, 0x1)
                        except:
                            pass
                    self.cancelButton = tk.Label(self.measureModelFrame, text="×", font=('Helvetica', 12, 'bold'), fg=self.foregroundColor, bg=self.backgroundColor, cursor="hand2")
                    self.cancelButton.grid(column=4, row=0)
                    self.prog_bar.grid(column=3, row=0, padx=5)
                    if (len(max(ig, key=len)) > 1):
                        pass
                    else:
                        self.prog_bar.start(40)
                    
                    cancel_ttp = CreateToolTip(self.cancelButton, text="Cancel current fitting", color="red", resetColor=self.foregroundColor)
    
                    self.queue = queue.Queue()
                    numCombos = 1
                    for g in ig:
                        numCombos *= len(g)
                    self.listPercent.append(numCombos)
                    self.currentThreads.append(ThreadedTaskCap(self.queue, self.wdata, self.rdata, self.jdata, nVE, nMC, c, aN, 1, fT, bL, ig, const, bU, self.listPercent, error_alpha, error_beta, error_betaRe, error_gamma, error_delta, cG, bLC, bUC, fC))
                    self.currentThreads[len(self.currentThreads)-1].start()
                    self.cancelButton.bind("<Button-1>", lambda e: terminateRun(self.currentThreads[len(self.currentThreads)-1]))
                    self.after(100, process_queue_cap)
        
        def checkErrorStructureAuto():
            if (self.errorAlphaCheckboxVariableAuto.get() == 1):
                self.errorAlphaEntryAuto.configure(state="normal")
            else:
                self.errorAlphaEntryAuto.configure(state="disabled")
            if (self.errorBetaCheckboxVariableAuto.get() == 1):
                self.errorBetaEntryAuto.configure(state="normal")
                self.errorBetaReCheckboxAuto.configure(state="normal")
            else:
                self.errorBetaEntryAuto.configure(state="disabled")
                self.errorBetaReCheckboxVariableAuto.set(0)
                self.errorBetaReEntryAuto.configure(state="disabled")
                self.errorBetaReCheckboxAuto.configure(state="disabled")
            if (self.errorBetaReCheckboxVariableAuto.get() == 1):
                self.errorBetaReEntryAuto.configure(state="normal")
            else:
                self.errorBetaReEntryAuto.configure(state="disabled")
            if (self.errorGammaCheckboxVariableAuto.get() == 1):
                self.errorGammaEntryAuto.configure(state="normal")
            else:
                self.errorGammaEntryAuto.configure(state="disabled")
            if (self.errorDeltaCheckboxVariableAuto.get() == 1):
                self.errorDeltaEntryAuto.configure(state="normal")
            else:
                self.errorDeltaEntryAuto.configure(state="disabled")
        self.autoErrorFrame = tk.Frame(self.autoFitWindow, bg=self.backgroundColor)
        self.errorAlphaCheckboxAuto = ttk.Checkbutton(self.autoErrorFrame, variable=self.errorAlphaCheckboxVariableAuto, text="\u03B1 = ", command=checkErrorStructureAuto)
        self.errorAlphaEntryAuto = ttk.Entry(self.autoErrorFrame, textvariable=self.errorAlphaVariableAuto, state="disabled", width=6)
        self.errorBetaCheckboxAuto = ttk.Checkbutton(self.autoErrorFrame, variable=self.errorBetaCheckboxVariableAuto, text="\u03B2 = ", command=checkErrorStructureAuto)
        self.errorBetaEntryAuto = ttk.Entry(self.autoErrorFrame, textvariable=self.errorBetaVariableAuto, state="disabled", width=6)
        self.errorBetaReCheckboxAuto = ttk.Checkbutton(self.autoErrorFrame, variable=self.errorBetaReCheckboxVariableAuto, text="Re = ", state="disabled", command=checkErrorStructureAuto)
        self.errorBetaReEntryAuto = ttk.Entry(self.autoErrorFrame, textvariable=self.errorBetaReVariableAuto, state="disabled", width=6)
        self.errorGammaCheckboxAuto = ttk.Checkbutton(self.autoErrorFrame, variable=self.errorGammaCheckboxVariableAuto, text="\u03B3 = ", command=checkErrorStructureAuto)
        self.errorGammaEntryAuto = ttk.Entry(self.autoErrorFrame, textvariable=self.errorGammaVariableAuto, width=6)
        self.errorDeltaCheckboxAuto = ttk.Checkbutton(self.autoErrorFrame, variable=self.errorDeltaCheckboxVariableAuto, text="\u03B4 = ", command=checkErrorStructureAuto)
        self.errorDeltaEntryAuto = ttk.Entry(self.autoErrorFrame, textvariable=self.errorDeltaVariableAuto, width=6)
        self.autoRunButton = ttk.Button(self.autoFitWindow, text="Run", width=10)
        self.autoCancelButton = ttk.Button(self.autoFitWindow, text="Cancel", width=10, state="disabled")
        
        def runAutoFitter():
            try:
                self.masterWindowId = ctypes.windll.user32.GetForegroundWindow()
            except:
                pass
            def onClose():
                self.autoFitWindow.withdraw()
            
            def runAuto(e=None):
                a = True
                if (len(self.fits) > 0):
                    a = messagebox.askokcancel("Continue?", "Auto fitting will remove currently fitted parameters. Continue?", parent=self.autoFitWindow)
                if (not a):
                    return
                if (str(self.measureModelButton["state"]) == "disabled"):
                    messagebox.showwarning("Currently running", "Another fitting is currently running. Please wait until it finishes and try again.", parent=self.autoFitWindow)
                    return
                self.queueAuto = queue.Queue()
                try:
                    nve = int(self.autoMaxNVEVariable.get())
                    if (nve <= 0):
                        raise Exception
                except:
                    messagebox.showerror("Bad value", "The maximum number of Voigt elements must be a positive integer.")
                    return
                try:
                    nmc = int(self.autoNMCVariable.get())
                    if (nmc <= 0):
                        raise Exception
                except:
                    messagebox.showerror("Bad value", "The number of Monte Carlo simulations must be a positive integer.")
                    return
                try:
                    if (self.errorAlphaCheckboxVariableAuto.get() == 0):
                        errA = 0
                    else:
                        errA = float(self.errorAlphaEntryAuto.get())
                    if (self.errorBetaCheckboxVariableAuto.get() == 0):
                        errB = 0
                    else:
                        errB = float(self.errorBetaEntryAuto.get())
                    if (self.errorBetaReCheckboxVariableAuto.get() == 0):
                        errBRe = 0
                    else:
                        errBRe = float(self.errorBetaReEntryAuto.get())
                    if (self.errorGammaCheckboxVariableAuto.get() == 0):
                        errG = 0
                    else:
                        errG = float(self.errorGammaEntryAuto.get())
                    if (self.errorDeltaCheckboxVariableAuto.get() == 0):
                        errD = 0
                    else:                        
                        errD = float(self.errorDeltaEntryAuto.get())
                except:
                    messagebox.showerror("Bad error value", "The values for the error structure must be real numbers.", parent=self.autoFitWindow)
                    return
                if (self.autoWeightingVariable.get() == "Error structure" and errA == 0 and errB == 0 and errG == 0 and errD == 0):
                    messagebox.showerror("Need error variable", "At least one error structure variable must be chosen", parent=self.autoFitWindow)
                    return
                if (self.autoWeightingVariable.get() == "Modulus"):
                    choice = 1
                elif (self.autoWeightingVariable.get() == "Proportional"):
                    choice = 2
                else:
                    choice = 3
                try:
                    fixReValue = float(self.autoReFixEntryVariable.get())
                except:
                    messagebox.showerror("Bad R\u2091 value", "The value for R\u2091 must be a real number.", parent=self.autoFitWindow)
                    return
                self.autoStatusLabel.configure(text="Initializing...")
                self.autoPercent = []
                self.measureModelButton.configure(state="disabled")
                self.autoButton.configure(state="normal")
                self.magicButton.configure(state="disabled")
                self.parametersButton.configure(state="disabled")
                self.freqRangeButton.configure(state="disabled")
                self.parametersLoadButton.configure(state="disabled")
                self.numVoigtSpinbox.configure(state="disabled")
                self.browseButton.configure(state="disabled")
                self.numMonteEntry.configure(state="disabled")
                self.fitTypeCombobox.configure(state="disabled")
                self.weightingCombobox.configure(state="disabled")
                self.undoButton.configure(state="disabled")
                self.numVoigtMinus.configure(cursor="arrow", bg="gray60")
                self.numVoigtMinus.unbind("<Enter>")
                self.numVoigtMinus.unbind("<Leave>")
                self.numVoigtMinus.unbind("<Button-1>")
                self.numVoigtMinus.unbind("<ButtonRelease-1>")
                self.numVoigtPlus.configure(cursor="arrow", bg="gray60")
                self.numVoigtPlus.unbind("<Enter>")
                self.numVoigtPlus.unbind("<Leave>")
                self.numVoigtPlus.unbind("<Button-1>")
                self.numVoigtPlus.unbind("<ButtonRelease-1>")
                self.autoNMC.configure(state="disabled")
                self.autoMaxNVE.configure(state="disabled")
                self.autoWeighting.configure(state="disabled")
                self.autoRunButton.configure(state="disabled")
                self.autoReFixEntry.configure(state="disabled")
                self.autoReFix.configure(state="disabled")
                self.autoCancelButton.configure(state="normal")
                self.freqWindow.withdraw()
                self.paramPopup.withdraw()
                self.advancedOptionsPopup.withdraw()
                try:
                    cc.GetModule('tl.tlb')
                    import comtypes.gen.TaskbarLib as tbl
                    self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                    self.taskbar.HrInit()
                    self.taskbar.ActivateTab(self.masterWindowId)
                    self.taskbar.SetProgressState(self.masterWindowId, 0x1)
                except:
                    pass
                
                self.currentThreads.append(ThreadedTaskAuto(self.queueAuto, nve, self.autoReFixVariable.get(), fixReValue, nmc, self.wdata, self.rdata, self.jdata, self.autoPercent, choice, self.radioVal.get(), errA, errB, errBRe, errG, errD))
                self.currentThreads[len(self.currentThreads)-1].start()
                self.autoCancelButton.bind("<Button-1>", lambda e: terminateRun(self.currentThreads[len(self.currentThreads)-1]))
                self.after(100, process_queue_auto)
            
            def checkAutoFix(e=None):
                if (self.autoReFixVariable.get()):
                    self.autoReFixEntry.configure(state="normal")
                else:
                    self.autoReFixEntry.configure(state="disabled")
            
            def checkWeightAuto(e):
                if (self.autoWeightingVariable.get() == "Modulus" or self.autoWeightingVariable.get() == "Proportional"):
                    self.autoErrorFrame.grid_remove()
                else:
                    self.autoErrorFrame.grid(column=0, row=2, columnspan=4, sticky="W")
            
            if (self.autoFitWindow.state() != "withdrawn"):
                self.autoFitWindow.deiconify()
                self.autoFitWindow.lift()
            
            self.autoMaxLabel = tk.Label(self.autoFitWindow, text="Max Voigt Elements: ", bg=self.backgroundColor, fg=self.foregroundColor)
            self.autoMaxNVE = ttk.Entry(self.autoFitWindow, textvariable=self.autoMaxNVEVariable, width=7)
            self.autoNMCLabel = tk.Label(self.autoFitWindow, text="Number of Simulations: ", bg=self.backgroundColor, fg=self.foregroundColor)
            self.autoNMC = ttk.Entry(self.autoFitWindow, textvariable=self.autoNMCVariable, width=7)
            self.autoReFrame = tk.Frame(self.autoFitWindow, bg=self.backgroundColor)
            self.autoReFix = ttk.Checkbutton(self.autoReFrame, text="Fix R\u2091 = ", variable=self.autoReFixVariable, command=checkAutoFix)
            self.autoReFixEntry = ttk.Entry(self.autoReFrame, width=7, textvariable=self.autoReFixEntryVariable, state="disabled")
            if (self.autoReFixVariable.get()):
                self.autoReFixEntry.configure(state="normal")
            self.autoWeightingLabel = tk.Label(self.autoFitWindow, text="Weighting: ", bg=self.backgroundColor, fg=self.foregroundColor)
            self.autoWeighting = ttk.Combobox(self.autoFitWindow, value=("Modulus", "Proportional", "Error structure"), textvariable=self.autoWeightingVariable, width=15, state="readonly")
            self.autoStatusLabel = tk.Label(self.autoFitWindow, text="", bg=self.backgroundColor, fg=self.foregroundColor)
            self.errorAlphaCheckboxAuto.grid(column=0, row=0)
            self.errorAlphaEntryAuto.grid(column=1, row=0, padx=(2, 15))
            self.errorBetaCheckboxAuto.grid(column=2, row=0)
            self.errorBetaEntryAuto.grid(column=3, row=0, padx=(2, 15))
            self.errorBetaReCheckboxAuto.grid(column=4, row=0)
            self.errorBetaReEntryAuto.grid(column=5, row=0, padx=(2, 15))
            self.errorGammaCheckboxAuto.grid(column=6, row=0)
            self.errorGammaEntryAuto.grid(column=7, row=0, padx=(2, 15))
            self.errorDeltaCheckboxAuto.grid(column=8, row=0)
            self.errorDeltaEntryAuto.grid(column=9, row=0, padx=(2, 15))
            self.autoRunButton.configure(command=runAuto)
            self.autoWeighting.bind("<<ComboboxSelected>>", checkWeightAuto)
            self.autoSliderFrame = tk.Frame(self.autoFitWindow, background=self.backgroundColor)
            self.autoRadioLabel = tk.Label(self.autoSliderFrame, text="Fit type: ", bg=self.backgroundColor, fg=self.foregroundColor)
            #self.radio1 = ttk.Radiobutton(self.autoSliderFrame, text="Fastest", variable=self.radioVal, value=1)
            self.radio2 = ttk.Radiobutton(self.autoSliderFrame, text="Complex", variable=self.radioVal, value=2)
            self.radio3 = ttk.Radiobutton(self.autoSliderFrame, text="Real", variable=self.radioVal, value=3)
            self.radio4 = ttk.Radiobutton(self.autoSliderFrame, text="Imaginary", variable=self.radioVal, value=4)
            autoMaxNVE_ttp = CreateToolTip(self.autoMaxNVE, "Number of Voigt elements to stop at")
            autoNMCLabel_ttp = CreateToolTip(self.autoNMC, "Number of Monte Carlo simulations for confidence intervals")
            autoWeighting_ttp = CreateToolTip(self.autoWeighting, "Weighting to be used")
            autoRunButton_ttp = CreateToolTip(self.autoRunButton, "Run automatic fitting")
            autoCancelButton_ttp = CreateToolTip(self.autoCancelButton, "Cancel automatic fitting")
            #radio1_ttp = CreateToolTip(self.radio1, "Use complex fit with modulus weighting only")
            radio2_ttp = CreateToolTip(self.radio2, "Fit both real and imaginary parts")
            radio3_ttp = CreateToolTip(self.radio3, "Fit real part only")
            radio4_ttp = CreateToolTip(self.radio4, "Fit imaginary part only")
            autoReFix_ttp = CreateToolTip(self.autoReFix, "Fix (i.e. don't fit) R\u2091")
            autoReFixEntry_ttp = CreateToolTip(self.autoReFixEntry, "The value of R\u2091 if it is fixed")
            
            self.autoMaxLabel.grid(column=0, row=0, sticky="W", padx=(3, 0))
            self.autoMaxNVE.grid(column=1, row=0, sticky="W", padx=3)
            self.autoNMCLabel.grid(column=2, row=0, padx=(5, 0))
            self.autoNMC.grid(column=3, row=0, padx=3)
            self.autoWeightingLabel.grid(column=0, row=1, sticky="W", pady=5, padx=(3, 0))
            self.autoWeighting.grid(column=1, row=1, sticky="W", columnspan=3, padx=3, pady=5)
            self.autoRadioLabel.grid(column=0, row=0, sticky="W", padx=3)
            #self.radio1.grid(column=0, row=0, sticky="W", padx=3)
            self.radio2.grid(column=1, row=0, sticky="W", padx=3)
            self.radio3.grid(column=2, row=0, sticky="W", padx=3)
            self.radio4.grid(column=3, row=0, sticky="W", padx=3)
            self.autoSliderFrame.grid(column=0, row=4, sticky="W", columnspan=5, padx=3, pady=5)
            self.autoReFrame.grid(column=0, row=3, sticky="W", columnspan=2, pady=5)
            self.autoReFix.grid(column=0, row=3, sticky="W", padx=3)
            self.autoReFixEntry.grid(column=1, row=3, sticky="W")
            self.autoStatusLabel.grid(column=0, row=5, sticky="EW", columnspan=5, pady=5)
            self.autoRunButton.grid(column=0, row=6, sticky="EW", columnspan=2, pady=10, padx=(3, 0))
            self.autoCancelButton.grid(column=2, row=6, sticky="EW", columnspan=2, pady=10, padx=3)
            
            #self.autoFitWindow.geometry("500x200")
            self.autoFitWindow.deiconify()
            self.autoFitWindow.bind("<Return>", runAuto)
            self.autoFitWindow.protocol("WM_DELETE_WINDOW", onClose)

        def scrollCanvas(event):
            self.parametersCanvas.yview_scroll(int(-1*(event.delta/120)), "units")
       
        def copyVals():
            self.copyButton.configure(text="Copied")
            self.clipboard_clear()
            stringToCopy = str(self.browseEntry.get()) + "\nT\tT Std. Dev.\tR\tR Std. Dev.\tC\tC Std. Dev.\n"
            resultValues = []
            resultStdDevs = []
            cap = 0
            capS = 0
            for child in self.resultsView.get_children():
                resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                resultValues.append(resultValue)
                resultStdDevs.append(resultStdDev)
                #Discard the name and 95% CI
            stringToCopy += "\t\t" + str(resultValues[0]) + "\t" + str(resultStdDevs[0]) + "\n"
            if (self.capUsed):
                cap = str(resultValues[1])
                capS = str(resultStdDevs[1])
                resultValues = resultValues[2:]
                resultStdDevs = resultStdDevs[2:]
            else:
                resultValues = resultValues[1:]
                resultStdDevs = resultStdDevs[1:]
            #---Bubble sort the results by increasing time constant---
            doneSwapping = False
            while (not doneSwapping):
                doneSwapping = True
                for i in range(1, len(resultValues)-1, 2):
                    if (float(resultValues[i]) > float(resultValues[i+2])):
                        tempT = resultValues[i]
                        tempR = resultValues[i-1]
                        tempTStd = resultStdDevs[i]
                        tempRStd = resultStdDevs[i-1]
                        resultValues[i] = resultValues[i+2]
                        resultValues[i-1] = resultValues[i+1]
                        resultStdDevs[i] = resultStdDevs[i+2]
                        resultStdDevs[i-1] = resultStdDevs[i+1]
                        resultValues[i+2] = tempT
                        resultValues[i+1] = tempR
                        resultStdDevs[i+2] = tempTStd
                        resultStdDevs[i+1] = tempRStd
                        doneSwapping = False
            for i in range(0, len(resultValues)-1, 2):
                c = float(resultValues[i+1])/float(resultValues[i])
                cs = c*np.sqrt((float(resultStdDevs[i+1])/float(resultValues[i+1]))**2 + (float(resultStdDevs[i])/float(resultValues[i]))**2)
                stringToCopy += str(resultValues[i+1]) + "\t" + str(resultStdDevs[i+1]) + "\t" + str(resultValues[i]) + "\t" + str(resultStdDevs[i]) + "\t" + str(c) + "\t" + str(cs) + "\n"
            stringToCopy += "\n"
            if (self.capUsed):
                stringToCopy += "C\t" + cap + "\tC Std. Dev.\t" + capS + "\n"
            capacitances = np.zeros(int((len(self.fits)-1)/2))
            for i in range(1, len(self.fits), 2):
                if (self.fits[1] == 0):
                    capacitances[int(i/2)] = 0
                else:
                    capacitances[int(i/2)] = self.fits[i+1]/self.fits[i]
            ceff = 0
            if (self.capUsed):
                ceff = 1/(np.sum(1/capacitances) + 1/self.resultCap)
            else:
                ceff = 1/(np.sum(1/capacitances))
            sigmaCapacitances = np.zeros(len(capacitances))
            for i in range(1, len(self.fits), 2):
                if (self.fits[i] == 0 or self.fits[i+1] == 0):
                    sigmaCapacitances[int(i/2)] = 0
                else:
                    sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((self.sigmas[i+1]/self.fits[i+1])**2 + (self.sigmas[i]/self.fits[i])**2)
            #sigmaOtherC = (1/rc)*(sc/rc)
            partCap = 0
            for i in range(len(capacitances)):
                partCap += sigmaCapacitances[i]**2/capacitances[i]**4
            try:
                partCap += self.sigmaCap**2/self.resultCap**4
            except:
                pass
            sigmaCeff = ceff**2 * np.sqrt(partCap)
            Zzero = self.fits[0] + self.fits[1]
            ZzeroSigma = self.sigmas[0]**2 + self.sigmas[1]**2
            for i in range(3, len(self.fits), 2):
                Zzero += self.fits[i]
                ZzeroSigma += self.sigmas[i]**2
            ZzeroSigma = np.sqrt(ZzeroSigma)
            Zpolar = 0
            ZpolarSigma = 0
            for i in range(1, len(self.fits), 2):
                Zpolar += self.fits[i]
                ZpolarSigma += self.sigmas[i]**2
            ZpolarSigma = np.sqrt(ZpolarSigma)
            stringToCopy += "Z zero freq.\t" + str(Zzero) + "\tZ zero freq. std. dev.\t" + str(ZzeroSigma) + "\nZ polarization\t" + str(Zpolar) + "\tZ polarization std. dev.\t" + str(ZpolarSigma) + "\nOverall C\t" + str(ceff) + "\tOverall C std. dev.\t" + str(sigmaCeff) + "\n"
            pyperclip.copy(stringToCopy)
            self.after(500, lambda : self.copyButton.configure(text="Copy values and std. devs. as spreadsheet"))
            
        def advancedResults():
            self.resultsWindow = tk.Toplevel()
            self.resultsWindows.append(self.resultsWindow)
            self.resultsWindow.title("Advanced results")
            self.resultsWindow.iconbitmap(resource_path("img/elephant3.ico"))
            self.resultsWindow.geometry("1000x550")
            self.advResults = scrolledtext.ScrolledText(self.resultsWindow, width=50, height=10, bg=self.backgroundColor, fg=self.foregroundColor, state="normal")
            self.advResults.insert(tk.INSERT, self.aR)
            self.advResults.configure(state="disabled")
            self.advResults.pack(fill=tk.BOTH, expand=True)
            
        def saveCurrent():
            if (not self.capUsed):
                valuesMatcher = []
                a = True
                for child in self.resultsView.get_children():
                    resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                    valuesMatcher.append(resultValue)
                if (len(valuesMatcher) != 0):
                    if (len(valuesMatcher) != len(self.paramEntryVariables)):
                        a = messagebox.askokcancel("Values do not match", "The values under \"Edit Parameters\" do not match the current fitted values. Only values under \"Edit Parameters\" will be saved. Continue?")
                    elif (self.capacitanceCheckboxVariable.get() == 1):
                        a = messagebox.askokcancel("Values do not match", "The capacitance checkbox is checked, but the capacitance was not fit. The current value under \"Edit Parameters\" will be saved. Continue?")
                    else:
                        dontMatch = False
                        for i in range(len(valuesMatcher)):
                            if (valuesMatcher[i] != self.paramEntryVariables[i].get()):
                                dontMatch = True
                                break
                        if (dontMatch):
                            a = messagebox.askokcancel("Values do not match", "The values under \"Edit Parameters\" do not match the current fitted values. Only values under \"Edit Parameters\" will be saved. Continue?")
                if (a):
                    stringToSave = "filename: " + self.browseEntry.get() + "\n"
                    stringToSave += "number_voigt: " + self.numVoigtVariable.get() + "\n"
                    stringToSave += "number_simulations: " + self.numMonteVariable.get() + "\n"
                    stringToSave += "fit_type: " + self.fitTypeVariable.get() + "\n"
                    stringToSave += "upDelete: " + str(self.upDelete) + "\n"
                    stringToSave += "lowDelete: " + str(self.lowDelete) + "\n"
                    stringToSave += "weight_type: " + self.weightingVariable.get() + "\n"
                    if (self.errorAlphaCheckboxVariable.get() == 0):
                        stringToSave += "error_alpha: n" + self.errorAlphaVariable.get() + "\n"
                    else:
                        stringToSave += "error_alpha: " + self.errorAlphaVariable.get() + "\n"
                    if (self.errorBetaCheckboxVariable.get() == 0):
                        stringToSave += "error_beta: n" + self.errorBetaVariable.get() + "\n"
                    else:
                        stringToSave += "error_beta: " + self.errorBetaVariable.get() + "\n"
                    if (self.errorBetaReCheckboxVariable.get() == 0):
                        stringToSave += "error_betaRe: n" + self.errorBetaReVariable.get() + "\n"
                    else:
                        stringToSave += "error_betaRe: " + self.errorBetaReVariable.get() + "\n"
                    if (self.errorGammaCheckboxVariable.get() == 0):
                        stringToSave += "error_gamma: n" + self.errorGammaVariable.get() + "\n"
                    else:
                        stringToSave += "error_gamma: " + self.errorGammaVariable.get() + "\n"
                    if (self.errorDeltaCheckboxVariable.get() == 0):
                        stringToSave += "error_delta: n" + self.errorDeltaVariable.get() + "\n"
                    else:
                        stringToSave += "error_delta: " + self.errorDeltaVariable.get() + "\n"
                    stringToSave += "alpha: " + self.alphaVariable.get() + "\n"
                    if (self.capacitanceCheckboxVariable.get() == 1):
                        stringToSave += "cap: " + self.capacitanceEntryVariable.get() + "\n"
                    else:
                        stringToSave += "cap: N\n"
                    stringToSave += "cap_combo: " + self.capacitanceComboboxVariable.get() + "\n"
                    stringToSave += "params: \n"
                    for p in self.paramEntryVariables:
                        stringToSave += p.get() + "\n"
                    stringToSave += "paramComboboxes: \n"
                    for p in self.paramComboboxVariables:
                        stringToSave += p.get() + "\n"
                    stringToSave += "tauComboboxes: \n"
                    for p in self.tauComboboxVariables:
                        stringToSave += p.get() + "\n"
                    defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
                    defaultSaveName += "-" + self.whatFit + str((len(self.fits)-1)//2)
                    saveName = asksaveasfile(mode='w', defaultextension=".mmfitting", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Measurement model fitting", ".mmfitting")])
                    directory = os.path.dirname(str(saveName))
                    self.topGUI.setCurrentDirectory(directory)
                    if saveName is None:     #If save is cancelled
                        return
                    saveName.write(stringToSave)
                    saveName.close()
                    self.saveCurrent.configure(text="Saved")
                    self.after(1000, lambda : self.saveCurrent.configure(text="Save Current Options and Parameters"))
            else:                  
                valuesMatcher = []
                a = True
                for child in self.resultsView.get_children():
                    resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                    valuesMatcher.append(resultValue)
                valuesMatcherCap = valuesMatcher[1]
                del valuesMatcher[1]
                if (len(valuesMatcher) != 0):
                    if (len(valuesMatcher) != len(self.paramEntryVariables)):
                        a = messagebox.askokcancel("Values do not match", "The values under \"Edit Parameters\" do not match the current fitted values. Only values under \"Edit Parameters\" will be saved. Continue?")
                    elif (self.capacitanceCheckboxVariable.get() != 1):
                        a = messagebox.askokcancel("Values do not match", "The capacitance was fit, but the capacitance checkbox is not checked. The capacitance will not be saved. Continue?")
                    else:
                        dontMatch = False
                        for i in range(len(valuesMatcher)):
                            if (valuesMatcher[i] != self.paramEntryVariables[i].get()):
                                dontMatch = True
                                break
                        if (valuesMatcherCap != self.capacitanceEntryVariable.get()):
                            dontMatch = True
                        if (dontMatch):
                            a = messagebox.askokcancel("Values do not match", "The values under \"Edit Parameters\" do not match the current fitted values. Only values under \"Edit Parameters\" will be saved. Continue?")
                if (a):
                    stringToSave = "filename: " + self.browseEntry.get() + "\n"
                    stringToSave += "number_voigt: " + self.numVoigtVariable.get() + "\n"
                    stringToSave += "number_simulations: " + self.numMonteVariable.get() + "\n"
                    stringToSave += "fit_type: " + self.fitTypeVariable.get() + "\n"
                    stringToSave += "upDelete: " + str(self.upDelete) + "\n"
                    stringToSave += "lowDelete: " + str(self.lowDelete) + "\n"
                    stringToSave += "weight_type: " + self.weightingVariable.get() + "\n"
                    if (self.errorAlphaCheckboxVariable.get() == 0):
                        stringToSave += "error_alpha: n" + self.errorAlphaVariable.get() + "\n"
                    else:
                        stringToSave += "error_alpha: " + self.errorAlphaVariable.get() + "\n"
                    if (self.errorBetaCheckboxVariable.get() == 0):
                        stringToSave += "error_beta: n" + self.errorBetaVariable.get() + "\n"
                    else:
                        stringToSave += "error_beta: " + self.errorBetaVariable.get() + "\n"
                    if (self.errorBetaReCheckboxVariable.get() == 0):
                        stringToSave += "error_betaRe: n" + self.errorBetaReVariable.get() + "\n"
                    else:
                        stringToSave += "error_betaRe: " + self.errorBetaReVariable.get() + "\n"
                    if (self.errorGammaCheckboxVariable.get() == 0):
                        stringToSave += "error_gamma: n" + self.errorGammaVariable.get() + "\n"
                    else:
                        stringToSave += "error_gamma: " + self.errorGammaVariable.get() + "\n"
                    if (self.errorDeltaCheckboxVariable.get() == 0):
                        stringToSave += "error_delta: n" + self.errorDeltaVariable.get() + "\n"
                    else:
                        stringToSave += "error_delta: " + self.errorDeltaVariable.get() + "\n"
                    stringToSave += "alpha: " + self.alphaVariable.get() + "\n"
                    if (self.capacitanceCheckboxVariable.get() != 1):
                        stringToSave += "cap: N\n"
                    else:
                        stringToSave += "cap: " + self.capacitanceEntryVariable.get() + "\n"
                    stringToSave += "cap_combo: " + self.capacitanceComboboxVariable.get() + "\n"
                    stringToSave += "params: \n"
                    for p in self.paramEntryVariables:
                        stringToSave += p.get() + "\n"
                    stringToSave += "paramComboboxes: \n"
                    for p in self.paramComboboxVariables:
                        stringToSave += p.get() + "\n"
                    stringToSave += "tauComboboxes: \n"
                    for p in self.tauComboboxVariables:
                        stringToSave += p.get() + "\n"
                    defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
                    defaultSaveName += "-" + self.whatFit + str((len(self.fits)-1)//2)
                    saveName = asksaveasfile(mode='w', defaultextension=".mmfitting", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Measurement model fitting", ".mmfitting")])
                    directory = os.path.dirname(str(saveName))
                    self.topGUI.setCurrentDirectory(directory)
                    if saveName is None:     #If save is cancelled
                        return
                    saveName.write(stringToSave)
                    saveName.close()
                    self.saveCurrent.configure(text="Saved")
                    self.after(1000, lambda : self.saveCurrent.configure(text="Save Current Options and Parameters"))
        
        def onclick(event):
            try:
                if (not event.inaxes):      #If a subplot isn't clicked
                    return
                resultPlotBig = tk.Toplevel()
                self.resultPlotBigs.append(resultPlotBig)
                resultPlotBig.iconbitmap(resource_path('img/elephant3.ico'))
                w, h = self.winfo_screenwidth(), self.winfo_screenheight()
                resultPlotBig.geometry("%dx%d+0+0" % (w/2, h/2))
                with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                    fig = Figure()
                    self.resultPlotBigFigs.append(fig)
                    fig.set_facecolor(self.backgroundColor)
                    Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                    if (not self.capUsed):
                        for i in range(len(self.wdata)):
                            Zfit[i] = self.fits[0]
                            for k in range(1, len(self.fits), 2):
                                Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)
                    else:
                        for i in range(len(self.wdata)):
                            Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                            for k in range(1, len(self.fits), 2):
                                Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                        phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)                
                    
                    larger = fig.add_subplot(111)
                    larger.set_facecolor(self.backgroundColor)
                    larger.yaxis.set_ticks_position("both")
                    larger.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                    larger.xaxis.set_ticks_position("both")
                    larger.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                    whichPlot = "a"
                    dataColor = "tab:blue"
                    fitColor = "orange"
                    if (self.topGUI.getTheme() == "dark"):
                        dataColor = "cyan"
                        fitColor = "gold"
                    else:
                        dataColor = "tab:blue"
                        fitColor = "orange"
                    
                    if (event.inaxes == self.aplot):
                        resultPlotBig.title("Nyquist Plot")
                        pointsPlot, = larger.plot(self.rdata, -1*self.jdata, "o", color=dataColor)
                        larger.plot(Zfit.real, -1*Zfit.imag, color=fitColor, marker="x", markeredgecolor="black")
                        if (self.confInt):
                            for i in range(len(self.wdata)):
                                ellipse = patches.Ellipse(xy=(Zfit[i].real, -1*Zfit[i].imag), width=4*self.sdrReal[i], height=4*self.sdrImag[i])
                                larger.add_artist(ellipse)
                                ellipse.set_alpha(0.1)
                                ellipse.set_facecolor(self.ellipseColor)
                        rightPoint = max(self.rdata)
                        topPoint = max(-1*self.jdata)
                        larger.axis("equal")
                        larger.set_title("Nyquist Plot", color=self.foregroundColor)
                        larger.set_xlabel("Zr / Ω", color=self.foregroundColor)
                        larger.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                        whichPlot = "a"
                    elif (event.inaxes == self.bplot):
                        resultPlotBig.title("Real Impedance Plot")
                        pointsPlot, = larger.plot(self.wdata, self.rdata, "o", color=dataColor, zorder=2)
                        larger.plot(self.wdata, Zfit.real, color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(Zfit.real[i] + 2*self.sdrReal[i], Zfit.real[i] - 2*self.sdrReal[i])
                                error_below[i] = min(Zfit.real[i] + 2*self.sdrReal[i], Zfit.real[i] - 2*self.sdrReal[i])
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        rightPoint = max(self.wdata)
                        topPoint = max(self.rdata)
                        larger.set_xscale("log")
                        larger.set_title("Real Impedance Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Zr / Ω", color=self.foregroundColor)
                        whichPlot = "b"
                    elif (event.inaxes == self.cplot):
                        resultPlotBig.title("Imaginary Impedance Plot")
                        pointsPlot, = larger.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                        larger.plot(self.wdata, -1*Zfit.imag, color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(Zfit.imag[i] + 2*self.sdrImag[i], Zfit.imag[i] - 2*self.sdrImag[i])
                                error_below[i] = min(Zfit.imag[i] + 2*self.sdrImag[i], Zfit.imag[i] - 2*self.sdrImag[i])
                            larger.plot(self.wdata, -1*error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, -1*error_below, "--", color=self.ellipseColor)
                        rightPoint = max(self.wdata)
                        topPoint = max(-1*self.jdata)
                        larger.set_xscale("log")
                        larger.set_title("Imaginary Impedance Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                        whichPlot = "c"
                    elif (event.inaxes == self.dplot):
                        resultPlotBig.title("|Z| Bode Plot")
                        pointsPlot, = larger.plot(self.wdata, np.sqrt(self.jdata**2 + self.rdata**2), "o", color=dataColor)
                        larger.plot(self.wdata, np.sqrt(Zfit.imag**2 + Zfit.real**2), color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = np.sqrt(max((Zfit.real[i]+2*self.sdrReal[i])**2, (Zfit.real[i]-2*self.sdrReal[i])**2) + max((Zfit.imag[i]+2*self.sdrImag[i])**2, (Zfit.imag[i]-2*self.sdrImag[i])**2))
                                error_below[i] = np.sqrt(min((Zfit.real[i]+2*self.sdrReal[i])**2, (Zfit.real[i]-2*self.sdrReal[i])**2) + min((Zfit.imag[i]+2*self.sdrImag[i])**2, (Zfit.imag[i]-2*self.sdrImag[i])**2))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        rightPoint = max(self.wdata)
                        topPoint = max(np.sqrt(self.jdata**2 + self.rdata**2))
                        larger.set_xscale("log")
                        larger.set_yscale("log")
                        larger.set_title("|Z| Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                        whichPlot = "d"
                    elif (event.inaxes == self.eplot):
                        phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)
                        actual_phase = np.arctan2(self.jdata, self.rdata) * (180/np.pi)
                        resultPlotBig.title("Phase Angle Bode Plot")
                        pointsPlot, = larger.plot(self.wdata, actual_phase, "o", color=dataColor)
                        larger.plot(self.wdata, phase_fit, color=fitColor)
                        if (self.confInt):
                            error_above = np.arctan2((Zfit.imag+2*self.sdrImag), (Zfit.real+2*self.sdrReal)) * (180/np.pi)
                            error_below = np.arctan2((Zfit.imag-2*self.sdrImag), (Zfit.real-2*self.sdrReal)) * (180/np.pi)
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        larger.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                        larger.set_ylim(bottom=0, top=-90)
                        rightPoint = max(self.wdata)
                        topPoint = -90
                        larger.set_xscale("log")
                        larger.set_title("Phase Angle Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Phase Angle / Degrees", color=self.foregroundColor)
                        whichPlot = "e"
                    elif (event.inaxes == self.fplot):
                        resultPlotBig.title("Re-adjusted |Z| Bode Plot")
                        pointsPlot, = larger.plot(self.wdata, np.sqrt(self.jdata**2 + (self.rdata-self.fits[0])**2), "o", color=dataColor)
                        larger.plot(self.wdata, np.sqrt(Zfit.imag**2 + (Zfit.real-self.fits[0])**2), color=fitColor)
                        rightPoint = max(self.wdata)
                        topPoint = max(np.sqrt(self.jdata**2 + (self.rdata-self.fits[0])**2))
                        larger.set_xscale("log")
                        larger.set_yscale("log")
                        larger.set_title("Re-adjusted |Z| Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel(r'$\sqrt{(Zr-Re)^2 + Zj^2}$ / Ω', color=self.foregroundColor)
                        whichPlot = "f"
                    elif (event.inaxes == self.gplot):
                        phase_fit = np.arctan2(Zfit.imag, (Zfit.real-self.fits[0])) * (180/np.pi)
                        actual_phase = np.arctan2(self.jdata, (self.rdata-self.fits[0])) * (180/np.pi)
                        resultPlotBig.title("Re-adjusted Phase Angle Bode Plot")
                        pointsPlot, = larger.plot(self.wdata, actual_phase, "o", color=dataColor)
                        larger.plot(self.wdata, phase_fit, color=fitColor)
                        larger.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                        larger.set_ylim(bottom=0, top=-90)
                        rightPoint = max(self.wdata)
                        topPoint = -90
                        larger.set_xscale("log")
                        larger.set_title("Re-adjusted Phase Angle Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel(r'$tan^{-1}({\frac{Zj}{(Zr-Re)}})$ / Degrees', color=self.foregroundColor)
                        whichPlot = "g"
                    elif (event.inaxes == self.hplot):
                        resultPlotBig.title("Log(Zr) vs f")
                        pointsPlot, = larger.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                        larger.plot(self.wdata, np.log10(abs(Zfit.real)), color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(np.log10(abs(Zfit.real[i]+2*self.sdrReal[i])), np.log10(abs(Zfit.real[i]-2*self.sdrReal[i]))) #np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                                error_below[i] = min(np.log10(abs(Zfit.real[i]+2*self.sdrReal[i])), np.log10(abs(Zfit.real[i]-2*self.sdrReal[i])))#np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
    #                            error_above[i] = np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
    #                            error_below[i] = np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        rightPoint = max(self.wdata)
                        topPoint = max(np.log10(abs(self.rdata)))
                        larger.set_xscale("log")
                        larger.set_title("Log|Zr|", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Log|Zr|", color=self.foregroundColor)
                        whichPlot = "h"
                    elif (event.inaxes == self.iplot):
                         resultPlotBig.title("Log|Zj| vs f")
                         pointsPlot, = larger.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                         larger.plot(self.wdata, np.log10(abs(Zfit.imag)), color=fitColor)
                         if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(np.log10(abs(Zfit.imag[i]+2*self.sdrImag[i])), np.log10(abs(Zfit.imag[i]-2*self.sdrImag[i])))#np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                                error_below[i] = min(np.log10(abs(Zfit.imag[i]+2*self.sdrImag[i])), np.log10(abs(Zfit.imag[i]-2*self.sdrImag[i])))#np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
    #                            error_above[i] = np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
    #                            error_below[i] = np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                         rightPoint = max(self.wdata)
                         topPoint = max(np.log10(abs(self.jdata)))
                         larger.set_xscale("log")
                         larger.set_title("Log|Zj|", color=self.foregroundColor)
                         larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                         larger.set_ylabel("Log|Zj|", color=self.foregroundColor)
                         whichPlot = "i"
                    elif (event.inaxes == self.jplot):
                        def logModel(freq):         #Returns the log of the imaginary part of the fitted model at a given frequency
                            if (not self.capUsed):
                                calculated = self.fits[0]
                            else:
                                calculated = self.fits[0] + 1/(1j*2*np.pi*freq*self.resultCap)
                            for k in range(1, len(self.fits), 2):
                                calculated += (self.fits[k]/(1+(1j*2*np.pi*freq*self.fits[k+1])))
                            return np.log(abs(calculated.imag))
                        
                        deriv = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            Dw = self.wdata[i]*1E-6
                            deriv[i] = (logModel(self.wdata[i]+Dw)-logModel(self.wdata[i]))/Dw    #Numerically calculate the derivative
                            deriv[i] *= self.wdata[i]       #Multiply by the frequency as per chain rule of d/dlog(omega)
                        resultPlotBig.title("dlog|Zj|/dlog(f)")#ω)")
                        larger.plot(self.wdata, deriv, color=fitColor)
                        larger.set_xscale("log")
                        larger.set_title("dlog|Zj|/dlog(f)", color=self.foregroundColor)#ω)", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel(r'$\frac{dlog|Zj|}{dlog(f)}$', color=self.foregroundColor)#ω)}$', color=self.foregroundColor)
                    elif (event.inaxes == self.kplot):
                        normalized_residuals_real = np.zeros(len(self.wdata))
                        normalized_error_real_below = np.zeros(len(self.wdata))
                        normalized_error_real_above = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            normalized_residuals_real[i] = (self.rdata[i] - Zfit[i].real)/Zfit[i].real
                            normalized_error_real_below[i] = 2*self.sdrReal[i]/Zfit[i].real
                            normalized_error_real_above[i] = -2*self.sdrReal[i]/Zfit[i].real
                        resultPlotBig.title("Real Normalized Residuals")
                        pointsPlot, = larger.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                        if (self.confInt):
                            larger.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                        larger.axhline(0, color="black", linewidth=1.0)
                        rightPoint = max(self.wdata)
                        topPoint = max(normalized_residuals_real)
                        larger.set_xscale("log")
                        larger.set_title("Real Normalized Residuals", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("(Zr-Z\u0302rmodel)/Zr", color=self.foregroundColor)
                        whichPlot = "k"
                    elif (event.inaxes == self.lplot):
                        normalized_residuals_imag = np.zeros(len(self.wdata))
                        normalized_error_imag_below = np.zeros(len(self.wdata))
                        normalized_error_imag_above = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            normalized_residuals_imag[i] = (self.jdata[i] - Zfit[i].imag)/Zfit[i].imag
                            normalized_error_imag_below[i] = 2*self.sdrImag[i]/Zfit[i].imag
                            normalized_error_imag_above[i] = -2*self.sdrImag[i]/Zfit[i].imag
                        resultPlotBig.title("Imaginary Normalized Residuals")
                        pointsPlot, = larger.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                        if (self.confInt):
                            larger.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                        larger.axhline(0, color="black", linewidth=1.0)
                        rightPoint = max(self.wdata)
                        topPoint = max(normalized_residuals_imag)
                        larger.set_xscale("log")
                        larger.set_title("Imaginary Normalized Residuals", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("(Zj-Z\u0302jmodel)/Zj", color=self.foregroundColor)
                        whichPlot = "l"
                    larger.xaxis.label.set_fontsize(20)
                    larger.yaxis.label.set_fontsize(20)
                    larger.title.set_fontsize(30)    
                largerCanvas = FigureCanvasTkAgg(fig, resultPlotBig)
                largerCanvas.draw()
                largerCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                toolbar = NavigationToolbar2Tk(largerCanvas, resultPlotBig)
                toolbar.update()
                largerCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                if (self.mouseOverCheckboxVariable.get() == 1):
                    annot = larger.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
                    annot.set_visible(False)
                    def update_annot(ind):
                        x,y = pointsPlot.get_data()
                        xval = x[ind["ind"][0]]
                        yval = y[ind["ind"][0]]
                        annot.xy = (xval, yval)
                        if (whichPlot == "a"):
                            text = "Zr=%.5g"%xval + "\nZj=-%.5g"%yval + "\nf=%.5g"%self.wdata[np.where(self.rdata == xval)][0]
                        elif (whichPlot == "b"):
                            text = "Zr=%.5g"%yval + "\nZj=-%.5g"%self.jdata[np.where(self.rdata == yval)][0] + "\nf=%.5g"%xval
                        elif (whichPlot == "c"):
                            text = "Zr=%.5g"%self.rdata[np.where(-1*self.jdata == yval)][0] + "\nZj=-%.5g"%yval + "\nf=%.5g"%xval
                        elif (whichPlot == "d"):
                            text = "|Z|=%.5g"%yval + "\nf=%.5g"%xval
                        elif (whichPlot == "e"):
                            text = "\u03D5=%.5g\u00B0"%yval + "\nf=%.5g"%xval
                        elif (whichPlot == "f"):
                            text = "R\u2091-adj. |Z|=%.5g"%yval + "\nf=%.5g"%xval
                        elif (whichPlot == "g"):
                            text = "R\u2091-adj. \u03D5=%.5g\u00B0"%yval + "\nf=%.5g"%xval
                        elif (whichPlot == "h"):
                            text = "Zr=%.5g"%10**yval + "\nf=%.5g"%xval
                        elif (whichPlot == "i"):
                            text = "Zj=-%.5g"%10**yval + "\nf=%.5g"%xval
                        elif (whichPlot == "k"):
                            text = "Real res.=%.5g"%yval + "\nf=%.5g"%xval
                        elif (whichPlot == "l"):
                            text = "Imag. res.=%.5g"%yval + "\nf=%.5g"%xval  
                        annot.set_text(text)
                        #---Check if we're within 5% of the right or top edges, and adjust label positions accordingly
                        if (rightPoint != 0):
                            if (abs(xval - rightPoint)/rightPoint <= 0.05):
                                annot.set_position((-10, -20))
                        if (topPoint != 0):
                            if (whichPlot != "e" and whichPlot != "g"):
                                if (abs(yval - topPoint)/topPoint <= 0.05):
                                    annot.set_position((10, -20))
                            else:
                                if (yval <= -85.5):
                                    annot.set_position((10, -20))
                        else:
                            annot.set_position((10, -20))
                    def hover(event):
                        vis = annot.get_visible()
                        if event.inaxes == larger:
                            nonlocal pointsPlot
                            cont, ind = pointsPlot.contains(event)
                            if cont:
                                update_annot(ind)
                                annot.set_visible(True)
                                fig.canvas.draw_idle()
                            else:
                                if vis:
                                    annot.set_position((10,10))
                                    annot.set_visible(False)
                                    fig.canvas.draw_idle()
                    fig.canvas.mpl_connect("motion_notify_event", hover)
                def on_closing():   #Clear the figure before closing the popup
                    fig.clear()
                    resultPlotBig.destroy()
                resultPlotBig.protocol("WM_DELETE_WINDOW", on_closing)
            except:     #A subplot wasn't clicked
                pass
        
        def graphOut(event):
            self.nyCanvas._tkcanvas.config(cursor="arrow")
#            self.nyCanvas._tkcanvas.delete(self.rec)
            #event.inaxes.set_facecolor("white")
            #self.nyCanvas.draw_idle()
        
        def graphOver(event):
            #axes = event.inaxes
            #autoAxis = event.inaxes
            whichCan = self.nyCanvas._tkcanvas
            he = int(whichCan.winfo_height())
            wi = int(whichCan.winfo_width())
            whichCan.config(cursor="hand2")
        
        def plotResults():
            if (self.alreadyPlotted and self.currentNVEPlotted == self.currentNVEPlotNeeded):
                self.resultPlot.deiconify()
                self.resultPlot.lift()
                return
            elif (self.alreadyPlotted and self.currentNVEPlotted != self.currentNVEPlotNeeded):
                self.resultPlot.destroy()
            self.alreadyPlotted = True
            self.currentNVEPlotted = self.currentNVEPlotNeeded
            self.resultPlot = tk.Toplevel(background=self.backgroundColor)
            self.resultPlots.append(self.resultPlot)
            self.resultPlot.title("Measurement Model Plots")
            self.resultPlot.iconbitmap(resource_path('img/elephant3.ico'))
            #w, h = self.winfo_screenwidth(), self.winfo_screenheight()
            #resultPlot.geometry("%dx%d+0+0" % (w, h))
            self.resultPlot.state("zoomed")
            self.confInt = False if self.confidenceIntervalCheckboxVariable.get() == 0 else True

            Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
            if (not self.capUsed):
                for i in range(len(self.wdata)):
                    Zfit[i] = self.fits[0]
                    for k in range(1, len(self.fits), 2):
                        Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)
            else:
                for i in range(len(self.wdata)):
                    Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                    for k in range(1, len(self.fits), 2):
                        Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)  
            
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                pltFig = Figure()   #figsize=((5,4), dpi=100)
                self.pltFigs.append(pltFig)
                pltFig.set_facecolor(self.backgroundColor)
                dataColor = "tab:blue"
                fitColor = "orange"
                if (self.topGUI.getTheme() == "dark"):
                    dataColor = "cyan"
                    fitColor = "gold"
                else:
                    dataColor = "tab:blue"
                    fitColor = "orange"
                
                self.aplot = pltFig.add_subplot(341)
                self.aplot.set_facecolor(self.backgroundColor)
                self.aplot.yaxis.set_ticks_position("both")
                self.aplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.aplot.xaxis.set_ticks_position("both")
                self.aplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.aplot.plot(self.rdata, -1*self.jdata, "o",  color=dataColor)
                self.aplot.plot(Zfit.real, -1*Zfit.imag, color=fitColor, marker="x", markeredgecolor="black")
                if (self.confInt):
                    for i in range(len(self.wdata)):
                        ellipse = patches.Ellipse(xy=(Zfit[i].real, -1*Zfit[i].imag), width=4*self.sdrReal[i], height=4*self.sdrImag[i])
                        self.aplot.add_artist(ellipse)
                        ellipse.set_alpha(0.1)
                        ellipse.set_facecolor(self.ellipseColor)
                self.aplot.axis("equal")
                self.aplot.set_title("Nyquist", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #extra = patches.Rectangle((0, 0), 1, 1, fc="w", fill=False, edgecolor='none', linewidth=0)
                #legA = self.aplot.legend([extra], ["Nyquist"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legA.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.aplot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                self.aplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                
                self.bplot = pltFig.add_subplot(342)
                self.bplot.set_facecolor(self.backgroundColor)
                self.bplot.yaxis.set_ticks_position("both")
                self.bplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.bplot.xaxis.set_ticks_position("both")
                self.bplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.bplot.plot(self.wdata, self.rdata, "o", color=dataColor)
                self.bplot.plot(self.wdata, Zfit.real, color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(Zfit.real[i] + 2*self.sdrReal[i], Zfit.real[i] - 2*self.sdrReal[i])
                        error_below[i] = min(Zfit.real[i] + 2*self.sdrReal[i], Zfit.real[i] - 2*self.sdrReal[i])
                    self.bplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.bplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.bplot.set_xscale('log')
                self.bplot.set_title("Real Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legB = self.bplot.legend([extra], ["Real Impedance"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legB.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.bplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.bplot.set_ylabel("Zr / Ω", color=self.foregroundColor)
                
                self.cplot = pltFig.add_subplot(343)
                self.cplot.set_facecolor(self.backgroundColor)
                self.cplot.yaxis.set_ticks_position("both")
                self.cplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.cplot.xaxis.set_ticks_position("both")
                self.cplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.cplot.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                self.cplot.plot(self.wdata, -1*Zfit.imag, color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(Zfit.imag[i] + 2*self.sdrImag[i], Zfit.imag[i] - 2*self.sdrImag[i])
                        error_below[i] = min(Zfit.imag[i] + 2*self.sdrImag[i], Zfit.imag[i] - 2*self.sdrImag[i])
                    self.cplot.plot(self.wdata, -1*error_above, "--", color=self.ellipseColor)
                    self.cplot.plot(self.wdata, -1*error_below, "--", color=self.ellipseColor)
                self.cplot.set_xscale('log')
                self.cplot.set_title("Imaginary Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legC = self.cplot.legend([extra], ["Imaginary Impedance"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legC.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.cplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.cplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                
                self.dplot = pltFig.add_subplot(344)
                self.dplot.set_facecolor(self.backgroundColor)
                self.dplot.yaxis.set_ticks_position("both")
                self.dplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.dplot.xaxis.set_ticks_position("both")
                self.dplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.dplot.plot(self.wdata, np.sqrt(self.rdata**2 + self.jdata**2), "o", color=dataColor)
                self.dplot.plot(self.wdata, np.sqrt(Zfit.real**2 + Zfit.imag**2), color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = np.sqrt(max((Zfit.real[i]+2*self.sdrReal[i])**2, (Zfit.real[i]-2*self.sdrReal[i])**2) + max((Zfit.imag[i]+2*self.sdrImag[i])**2, (Zfit.imag[i]-2*self.sdrImag[i])**2))
                        error_below[i] = np.sqrt(min((Zfit.real[i]+2*self.sdrReal[i])**2, (Zfit.real[i]-2*self.sdrReal[i])**2) + min((Zfit.imag[i]+2*self.sdrImag[i])**2, (Zfit.imag[i]-2*self.sdrImag[i])**2))
                    self.dplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.dplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.dplot.set_xscale('log')
                self.dplot.set_yscale('log')
                self.dplot.set_title("|Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legD = self.dplot.legend([extra], ["|Z| Bode"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legD.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.dplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.dplot.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                
                self.eplot = pltFig.add_subplot(345)
                self.eplot.set_facecolor(self.backgroundColor)
                self.eplot.yaxis.set_ticks_position("both")
                self.eplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.eplot.xaxis.set_ticks_position("both")
                self.eplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.eplot.plot(self.wdata, np.arctan2(self.jdata, self.rdata)*(180/np.pi), "o", color=dataColor)
                self.eplot.plot(self.wdata, phase_fit, color=fitColor)
                if (self.confInt):
    #                error_above = np.zeros(len(self.wdata))
    #                error_below = np.zeros(len(self.wdata))
    #                for i in range(len(self.wdata)):
    #                    error_above[i] = np.arctan2(max((Zfit.imag[i]+2*self.sdrImag[i]), (Zfit.imag[i]-2*self.sdrImag[i])), min((Zfit.real[i]+2*self.sdrReal[i]), (Zfit.real[i]-2*self.sdrReal[i])))
    #                    error_below[i] = np.arctan2(min((Zfit.imag[i]+2*self.sdrImag[i]), (Zfit.imag[i]-2*self.sdrImag[i])), max((Zfit.real[i]+2*self.sdrReal[i]), (Zfit.real[i]-2*self.sdrReal[i])))
                    error_above = np.arctan2((Zfit.imag+2*self.sdrImag) , (Zfit.real+2*self.sdrReal)) * (180/np.pi)
                    error_below = np.arctan2((Zfit.imag-2*self.sdrImag) , (Zfit.real-2*self.sdrReal)) * (180/np.pi)
                    self.eplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.eplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.eplot.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                self.eplot.set_ylim(bottom=0, top=-90)
                self.eplot.set_xscale('log')
                self.eplot.set_title("Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legE = self.eplot.legend([extra], ["Phase Angle Bode"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legE.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.eplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.eplot.set_ylabel("Phase angle / Deg.", color=self.foregroundColor)
                
                self.fplot = pltFig.add_subplot(346)
                self.fplot.set_facecolor(self.backgroundColor)
                self.fplot.yaxis.set_ticks_position("both")
                self.fplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.fplot.xaxis.set_ticks_position("both")
                self.fplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.fplot.plot(self.wdata, np.sqrt((self.rdata-self.fits[0])**2 + self.jdata**2), "o", color=dataColor)
                self.fplot.plot(self.wdata, np.sqrt((Zfit.real-self.fits[0])**2 + Zfit.imag**2), color=fitColor)
                self.fplot.set_xscale('log')
                self.fplot.set_yscale('log')
                self.fplot.set_title("Re-adj. |Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legF = self.fplot.legend([extra], ["Re-adj. |Z| Bode"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legF.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.fplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.fplot.set_ylabel("Re-adj. |Z| / Ω", color=self.foregroundColor)
                
                self.gplot = pltFig.add_subplot(347)
                self.gplot.set_facecolor(self.backgroundColor)
                self.gplot.yaxis.set_ticks_position("both")
                self.gplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.gplot.xaxis.set_ticks_position("both")
                self.gplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.gplot.plot(self.wdata, np.arctan2(self.jdata, (self.rdata-self.fits[0]))*(180/np.pi), "o", color=dataColor)
                self.gplot.plot(self.wdata, np.arctan2(Zfit.imag, (Zfit.real-self.fits[0]))*(180/np.pi), color=fitColor)
                self.gplot.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                self.gplot.set_ylim(bottom=0, top=-90)
                self.gplot.set_xscale('log')
                self.gplot.set_title("Re-adj. Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legG = self.gplot.legend([extra], ["Re-adj. Phase Angle Bode"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legG.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.gplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.gplot.set_ylabel("Re-adj. Phase angle / Deg.", color=self.foregroundColor)
                
                self.hplot = pltFig.add_subplot(348)
                self.hplot.set_facecolor(self.backgroundColor)
                self.hplot.yaxis.set_ticks_position("both")
                self.hplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.hplot.xaxis.set_ticks_position("both")
                self.hplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.hplot.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                self.hplot.plot(self.wdata, np.log10(abs(Zfit.real)), color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(np.log10(abs(Zfit.real[i]+2*self.sdrReal[i])), np.log10(abs(Zfit.real[i]-2*self.sdrReal[i]))) #np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                        error_below[i] = min(np.log10(abs(Zfit.real[i]+2*self.sdrReal[i])), np.log10(abs(Zfit.real[i]-2*self.sdrReal[i])))#np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                    self.hplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.hplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.hplot.set_xscale('log')
                self.hplot.set_title("Log|Zr|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legH = self.hplot.legend([extra], ["Log|Zr|"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legH.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.hplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.hplot.set_ylabel("Log|Zr|", color=self.foregroundColor)
                
                self.iplot = pltFig.add_subplot(349)
                self.iplot.set_facecolor(self.backgroundColor)
                self.iplot.yaxis.set_ticks_position("both")
                self.iplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.iplot.xaxis.set_ticks_position("both")
                self.iplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.iplot.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                self.iplot.plot(self.wdata, np.log10(abs(Zfit.imag)), color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(np.log10(abs(Zfit.imag[i]+2*self.sdrImag[i])), np.log10(abs(Zfit.imag[i]-2*self.sdrImag[i])))#np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                        error_below[i] = min(np.log10(abs(Zfit.imag[i]+2*self.sdrImag[i])), np.log10(abs(Zfit.imag[i]-2*self.sdrImag[i])))#np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                    self.iplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.iplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.iplot.set_xscale('log')
                self.iplot.set_title("Log|Zj|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legI = self.iplot.legend([extra], ["Log|Zj|"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legI.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.iplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.iplot.set_ylabel("Log|Zj|", color=self.foregroundColor)
                
                def logModel(freq):
                    if (not self.capUsed):
                        calculated = self.fits[0]
                    else:
                        calculated = self.fits[0] + 1/(1j*2*np.pi*freq*self.resultCap)
                    for k in range(1, len(self.fits), 2):
                        calculated += (self.fits[k]/(1+(1j*2*np.pi*freq*self.fits[k+1])))
                    return np.log(abs(calculated.imag))
                
                deriv = np.zeros(len(self.wdata))
                for i in range(len(self.wdata)):
                    Dw = self.wdata[i]*1E-6
                    deriv[i] = (logModel(self.wdata[i]+Dw)-logModel(self.wdata[i]))/Dw    #Numerically calculate the derivative
                    deriv[i] *= self.wdata[i]       #Multiply by the frequency as per chain rule of d/dlog(omega)
                
                self.jplot = pltFig.add_subplot(3, 4, 10)
                self.jplot.set_facecolor(self.backgroundColor)
                self.jplot.yaxis.set_ticks_position("both")
                self.jplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.jplot.xaxis.set_ticks_position("both")
                self.jplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                #jplot.plot(self.wdata, np.log10(-1*self.jdata), "o")
                self.jplot.plot(self.wdata, deriv, color=fitColor)
                self.jplot.set_xscale('log')
                self.jplot.set_title("dlog|Zj|/dlog(f)", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legJ = self.jplot.legend([extra], ["dlog|Zj|/dlog(ω)"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legJ.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.jplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.jplot.set_ylabel(r'$\frac{dlog|Zj|}{dlog(f)}$', color=self.foregroundColor)
                
                normalized_residuals_real = np.zeros(len(self.wdata))
                normalized_error_real_above = np.zeros(len(self.wdata))
                normalized_error_real_below = np.zeros(len(self.wdata))
                normalized_residuals_imag = np.zeros(len(self.wdata))
                normalized_error_imag_above = np.zeros(len(self.wdata))
                normalized_error_imag_below = np.zeros(len(self.wdata))
                for i in range(len(self.wdata)):
                    normalized_residuals_real[i] = (self.rdata[i] - Zfit[i].real)/Zfit[i].real
                    normalized_error_real_above[i] = 2*self.sdrReal[i]/Zfit[i].real
                    normalized_error_real_below[i] = -2*self.sdrReal[i]/Zfit[i].real
                    normalized_residuals_imag[i] = (self.jdata[i] - Zfit[i].imag)/Zfit[i].imag
                    normalized_error_imag_below[i] = 2*self.sdrImag[i]/Zfit[i].imag
                    normalized_error_imag_above[i] = -2*self.sdrImag[i]/Zfit[i].imag
                
                self.kplot = pltFig.add_subplot(3, 4, 11)
                self.kplot.set_facecolor(self.backgroundColor)
                self.kplot.yaxis.set_ticks_position("both")
                self.kplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.kplot.xaxis.set_ticks_position("both")
                self.kplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.kplot.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                if (self.confInt):
                    self.kplot.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                    self.kplot.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                #fplot.plot(self.wdata, phase_fit, color="orange")
                self.kplot.axhline(0, color="black", linewidth=1.0)
                self.kplot.set_xscale('log')
                self.kplot.set_title("Real Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legK = self.kplot.legend([extra], ["Real Residuals"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legK.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.kplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.kplot.set_ylabel("(Zr-Zrmodel)/Zr", color=self.foregroundColor)
                
                self.lplot = pltFig.add_subplot(3, 4, 12)
                self.lplot.set_facecolor(self.backgroundColor)
                self.lplot.yaxis.set_ticks_position("both")
                self.lplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                self.lplot.xaxis.set_ticks_position("both")
                self.lplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                self.lplot.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                if (self.confInt):
                    self.lplot.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                    self.lplot.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                #fplot.plot(self.wdata, phase_fit, color="orange")
                self.lplot.axhline(0, color="black", linewidth=1.0)
                self.lplot.set_xscale('log')
                self.lplot.set_title("Imaginary Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                #legL = self.lplot.legend([extra], ["Imaginary Residuals"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                #for text in legL.get_texts():
                #    text.set_color(self.foregroundColor)
                #self.lplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.lplot.set_ylabel("(Zj-Zjmodel)/Zj", color=self.foregroundColor)
            
            self.nyCanvas = FigureCanvasTkAgg(pltFig, self.resultPlot)
            pltFig.subplots_adjust(top=0.95, bottom=0.1, right=0.95, left=0.05)
            self.nyCanvas.draw()
            self.nyCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
            self.nyCanvas.mpl_connect('button_press_event', onclick)
            #toolbar = NavigationToolbar2Tk(nyCanvas, resultPlot)
            #toolbar.update()
            self.nyCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            def saveAllPlots():
                folder = askdirectory(parent=self.resultPlot, initialdir=self.topGUI.getCurrentDirectory())
                folder_str = str(folder)
                if (len(folder_str) == 0):
                    pass
                else:
                    self.saveNyCanvasButton.configure(text="Saving...")
                    defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
                    defaultSaveName += "-" + self.whatFit + str((len(self.fits)-1)//2)
                    pltSaveFig = plt.Figure()
                    pltSaveFig.set_facecolor(self.backgroundColor)
                    dataColor = "tab:blue"
                    fitColor = "orange"
                    if (self.topGUI.getTheme() == "dark"):
                        dataColor = "cyan"
                        fitColor = "gold"
                    else:
                        dataColor = "tab:blue"
                        fitColor = "orange"

                    save_canvas = FigureCanvas(pltSaveFig)
                    ax_save = pltSaveFig.add_subplot(111)
                    ax_save.set_facecolor(self.backgroundColor)
                    #---aplot---
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    ax_save.plot(self.rdata, -1*self.jdata, "o",  color=dataColor)
                    ax_save.plot(Zfit.real, -1*Zfit.imag, color=fitColor, marker="x", markeredgecolor="black")
                    if (self.confInt):
                        for i in range(len(self.wdata)):
                            ellipse = patches.Ellipse(xy=(Zfit[i].real, -1*Zfit[i].imag), width=4*self.sdrReal[i], height=4*self.sdrImag[i])
                            ax_save.add_artist(ellipse)
                            ellipse.set_alpha(0.1)
                            ellipse.set_facecolor(self.ellipseColor)
                    ax_save.axis("equal")
                    ax_save.set_title("Nyquist", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Zr / Ω", color=self.foregroundColor)
                    ax_save.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_nyquist.png", dpi=300)
                    ax_save.clear()
                    ax_save.axis("auto")
                    #---bplot---
                    ax_save.set_facecolor(self.backgroundColor)
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.plot(self.wdata, self.rdata, "o", color=dataColor)
                    ax_save.plot(self.wdata, Zfit.real, color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(Zfit.real[i] + 2*self.sdrReal[i], Zfit.real[i] - 2*self.sdrReal[i])
                            error_below[i] = min(Zfit.real[i] + 2*self.sdrReal[i], Zfit.real[i] - 2*self.sdrReal[i])
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Real Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_ylabel("Zr / Ω", color=self.foregroundColor)
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_real_impedance.png", dpi=300)
                    ax_save.clear()
                    #---cplot---
                    ax_save.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                    ax_save.plot(self.wdata, -1*Zfit.imag, color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(Zfit.imag[i] + 2*self.sdrImag[i], Zfit.imag[i] - 2*self.sdrImag[i])
                            error_below[i] = min(Zfit.imag[i] + 2*self.sdrImag[i], Zfit.imag[i] - 2*self.sdrImag[i])
                        ax_save.plot(self.wdata, -1*error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, -1*error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Imaginary Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_imag_impedance.png", dpi=300)
                    ax_save.clear()
                    #---dplot---
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.plot(self.wdata, np.sqrt(self.rdata**2 + self.jdata**2), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.sqrt(Zfit.real**2 + Zfit.imag**2), color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = np.sqrt(max((Zfit.real[i]+2*self.sdrReal[i])**2, (Zfit.real[i]-2*self.sdrReal[i])**2) + max((Zfit.imag[i]+2*self.sdrImag[i])**2, (Zfit.imag[i]-2*self.sdrImag[i])**2))
                            error_below[i] = np.sqrt(min((Zfit.real[i]+2*self.sdrReal[i])**2, (Zfit.real[i]-2*self.sdrReal[i])**2) + min((Zfit.imag[i]+2*self.sdrImag[i])**2, (Zfit.imag[i]-2*self.sdrImag[i])**2))
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_yscale('log')
                    ax_save.set_title("|Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_Z_bode.png", dpi=300)
                    ax_save.clear()
                    #---eplot---
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.plot(self.wdata, np.arctan2(self.jdata, self.rdata)*(180/np.pi), "o", color=dataColor)
                    ax_save.plot(self.wdata, phase_fit, color=fitColor)
                    if (self.confInt):
        #                error_above = np.zeros(len(self.wdata))
        #                error_below = np.zeros(len(self.wdata))
        #                for i in range(len(self.wdata)):
        #                    error_above[i] = np.arctan2(max((Zfit.imag[i]+2*self.sdrImag[i]), (Zfit.imag[i]-2*self.sdrImag[i])), min((Zfit.real[i]+2*self.sdrReal[i]), (Zfit.real[i]-2*self.sdrReal[i])))
        #                    error_below[i] = np.arctan2(min((Zfit.imag[i]+2*self.sdrImag[i]), (Zfit.imag[i]-2*self.sdrImag[i])), max((Zfit.real[i]+2*self.sdrReal[i]), (Zfit.real[i]-2*self.sdrReal[i])))
                        error_above = np.arctan2((Zfit.imag+2*self.sdrImag) , (Zfit.real+2*self.sdrReal)) * (180/np.pi)
                        error_below = np.arctan2((Zfit.imag-2*self.sdrImag) , (Zfit.real-2*self.sdrReal)) * (180/np.pi)
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                    ax_save.set_ylim(bottom=0, top=-90)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Phase angle / Deg.", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_phase_angle_bode.png", dpi=300)
                    ax_save.clear()
                    #---fplot---
                    ax_save.plot(self.wdata, np.sqrt((self.rdata-self.fits[0])**2 + self.jdata**2), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.sqrt((Zfit.real-self.fits[0])**2 + Zfit.imag**2), color=fitColor)
                    ax_save.set_xscale('log')
                    ax_save.set_yscale('log')
                    ax_save.set_title("Re-adj. |Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Re-adj. |Z| / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_re-adj_Z_bode.png", dpi=300)
                    ax_save.clear()
                    #---gplot---
                    ax_save.plot(self.wdata, np.arctan2(self.jdata, (self.rdata-self.fits[0]))*(180/np.pi), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.arctan2(Zfit.imag, (Zfit.real-self.fits[0]))*(180/np.pi), color=fitColor)
                    ax_save.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                    ax_save.set_ylim(bottom=0, top=-90)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Re-adj. Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Re-adj. Phase angle / Deg.", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_re-adj_phase_angle_bode.png", dpi=300)
                    ax_save.clear()
                    #---hplot---
                    ax_save.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.log10(abs(Zfit.real)), color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(np.log10(abs(Zfit.real[i]+2*self.sdrReal[i])), np.log10(abs(Zfit.real[i]-2*self.sdrReal[i]))) #np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                            error_below[i] = min(np.log10(abs(Zfit.real[i]+2*self.sdrReal[i])), np.log10(abs(Zfit.real[i]-2*self.sdrReal[i])))#np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Log|Zr|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Log|Zr|", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_log_zr.png", dpi=300)
                    ax_save.clear()
                    #---iplot---
                    ax_save.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.log10(abs(Zfit.imag)), color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(np.log10(abs(Zfit.imag[i]+2*self.sdrImag[i])), np.log10(abs(Zfit.imag[i]-2*self.sdrImag[i])))#np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                            error_below[i] = min(np.log10(abs(Zfit.imag[i]+2*self.sdrImag[i])), np.log10(abs(Zfit.imag[i]-2*self.sdrImag[i])))#np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Log|Zj|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Log|Zj|", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_log_zj.png", dpi=300)
                    ax_save.clear()
                    #---jplot---
                    def logModel(freq):
                        if (not self.capUsed):
                            calculated = self.fits[0]
                        else:
                            calculated = self.fits[0] + 1/(1j*2*np.pi*freq*self.resultCap)
                        for k in range(1, len(self.fits), 2):
                            calculated += (self.fits[k]/(1+(1j*2*np.pi*freq*self.fits[k+1])))
                        return np.log(abs(calculated.imag))
                    
                    deriv = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        Dw = self.wdata[i]*1E-6
                        deriv[i] = (logModel(self.wdata[i]+Dw)-logModel(self.wdata[i]))/Dw    #Numerically calculate the derivative
                        deriv[i] *= self.wdata[i]       #Multiply by the frequency as per chain rule of d/dlog(omega)
                    
                    ax_save.plot(self.wdata, deriv, color=fitColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("dlog|Zj|/dlog(f)", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("dlog|Zj|/dlog(f)", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_dlogzj_dlogf.png", dpi=300)
                    ax_save.clear()
                    #---kplot---
                    ax_save.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                    if (self.confInt):
                        ax_save.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                    #fplot.plot(self.wdata, phase_fit, color="orange")
                    ax_save.axhline(0, color="black", linewidth=1.0)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Real Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("(Zr-Zrmodel)/Zr", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_real_residuals.png", dpi=300)
                    ax_save.clear()
                    #---lplot---
                    ax_save.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                    if (self.confInt):
                        ax_save.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                    ax_save.axhline(0, color="black", linewidth=1.0)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Imaginary Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("(Zj-Zjmodel)/Zj", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_imag_residuals.png", dpi=300)
                    ax_save.clear()
                    self.saveNyCanvasButton.configure(text="Saved")
                    self.after(500, lambda: self.saveNyCanvasButton.configure(text="Save All"))
            self.saveNyCanvasButton = ttk.Button(self.resultPlot, text="Save All", command=saveAllPlots)
            self.saveNyCanvasButtons.append(self.saveNyCanvasButton)
            saveNyCanvasButton_ttp = CreateToolTip(self.saveNyCanvasButton, "Save all plots")
            self.saveNyCanvasButton_ttps.append(saveNyCanvasButton_ttp)
            self.saveNyCanvasButton.pack(side=tk.BOTTOM, pady=5, expand=False)
            #self.nyCanvas._tkcanvas.config(cursor="hand2")
            enterAxes = pltFig.canvas.mpl_connect('axes_enter_event', graphOver)
            leaveAxes = pltFig.canvas.mpl_connect('axes_leave_event', graphOut)
            def on_closing():   #Clear the figure before closing the popup
                pltFig.clear()
                self.resultPlot.destroy()
                self.saveNyCanvasButton.destroy()
                nonlocal saveNyCanvasButton_ttp
                del saveNyCanvasButton_ttp
                self.alreadyPlotted = False
            
            self.resultPlot.protocol("WM_DELETE_WINDOW", on_closing)
            
        def plusNVE(event):
            self.numVoigtSpinbox.invoke("buttonup")
            self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white")
        def minusNVE(event):
            self.numVoigtSpinbox.invoke("buttondown")
            self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white")
        
        def saveResiduals():
            def model(freq):
                if (self.capUsed):
                    valueToReturn = self.fits[0] + 1/(1j*2*np.pi*freq*self.resultCap)
                else:
                    valueToReturn = self.fits[0]
                for k in range(1, len(self.fits), 2):
                    valueToReturn += (self.fits[k]/(1+(1j*2*np.pi*freq*self.fits[k+1])))
                return valueToReturn
            
            stringToSave = str(self.fits[0]) + "\t" + str(self.resultCap) + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\n"
            for i in range(len(self.wdata)):
                stringToSave += str(self.wdata[i]) + "\t" + str(self.rdata[i] - model(self.wdata[i]).real) + "\t" + str(self.jdata[i] - model(self.wdata[i]).imag) + "\t" + str(self.rdata[i]) + "\t" + str(self.jdata[i]) + "\n"
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
            defaultSaveName += "-" + self.whatFit + str((len(self.fits)-1)//2)
            saveName = asksaveasfile(mode='w', defaultextension=".mmresiduals", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Measurement model residuals", ".mmresiduals")])
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if saveName is None:     #If save is cancelled
                return
            saveName.write(stringToSave)
            saveName.close()
            self.saveResiduals.configure(text="Saved")
            self.after(1000, lambda : self.saveResiduals.configure(text="Save Residuals for Error Analysis"))
        
        def saveAll():
            def model(freq):
                valueToReturn = self.fits[0]
                if (self.capUsed):
                    valueToReturn += 1/(1j*2*np.pi*freq*self.resultCap)
                for k in range(1, len(self.fits), 2):
                    valueToReturn += (self.fits[k]/(1+(1j*2*np.pi*freq*self.fits[k+1])))
                return valueToReturn
            
            stringToSave = "freq\tZr data\tZj data\tZr model\tZj model\tReal Weight\tImag. Weight\tReal conf. interv.\tImag. conf. interv.\n"
            for i in range(len(self.wdata)):
                stringToSave += str(self.wdata[i]) + "\t" + str(self.rdata[i]) + "\t" + str(self.jdata[i]) + "\t" + str(model(self.wdata[i]).real) + "\t" + str(model(self.wdata[i]).imag) + "\t" + str(self.fitWeightR[i]) + "\t" + str(self.fitWeightJ[i]) + "\t" + str(self.sdrReal[i]) + "\t" + str(self.sdrImag[i]) + "\n"
            stringToSave += "----------------------------------------------------------------------------------\n"
            stringToSave += "Re = " + str(self.fits[0]) + "\tStd. Dev. = " + str(self.sigmas[0]) + "\n"
            if (self.capUsed):
                stringToSave += "Capacitance = " + str(self.resultCap) + "\tStd. Dev. = " + str(self.sigmaCap) + "\n"
            elementNumberCorrector = 0
            for i in range(1, len(self.fits), 2):
                stringToSave += "R" + str(i-elementNumberCorrector) + " = " + str(self.fits[i]) + "\tStd. Dev. = " + str(self.sigmas[i]) + "\n"
                stringToSave += "Tau" + str(i-elementNumberCorrector) + " = " + str(self.fits[i+1]) + "\tStd. Dev. = " + str(self.sigmas[i+1]) + "\n"
                stringToSave += "C" + str(i-elementNumberCorrector) + " = " + str(self.fits[i+1]/self.fits[i]) + "\tStd. Dev. = " +  str((self.fits[i+1]/self.fits[i])*np.sqrt((self.sigmas[i+1]/self.fits[i+1])**2 + (self.sigmas[i]/self.fits[i])**2)) + "\n"
                elementNumberCorrector += 1
            stringToSave += "\nOhmic Resistance = " + str(self.fits[0]) + "\tStd. Dev. = " + str(self.sigmas[0]) + "\n"
            Zzero = self.fits[0] + self.fits[1]
            ZzeroSigma = self.sigmas[0]**2 + self.sigmas[1]**2
            for i in range(3, len(self.fits), 2):
                Zzero += self.fits[i]
                ZzeroSigma += self.sigmas[i]**2
            ZzeroSigma = np.sqrt(ZzeroSigma)
            stringToSave += "Zero frequency Impedance = " + str(Zzero) + "\tStd. Dev. = " + str(ZzeroSigma) + "\n"
            Zpolar = 0
            ZpolarSigma = 0
            for i in range(1, len(self.fits), 2):
                Zpolar += self.fits[i]
                ZpolarSigma += self.sigmas[i]**2
            ZpolarSigma = np.sqrt(ZpolarSigma)
            stringToSave += "Polarization Impedance = " + str(Zpolar) + "\tStd. Dev. = " + str(ZpolarSigma) + "\n"
            capacitances = np.zeros(int((len(self.fits)-1)/2))
            for i in range(1, len(self.fits), 2):
                capacitances[int(i/2)] = self.fits[i+1]/self.fits[i]
            ceff = 1/np.sum(1/capacitances)
            sigmaCapacitances = np.zeros(len(capacitances))
            for i in range(1, len(self.fits), 2):
                sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((self.sigmas[i+1]/self.fits[i+1])**2 + (self.sigmas[i]/self.fits[i])**2)
            partCap = 0
            for i in range(len(capacitances)):
                partCap += sigmaCapacitances[i]**2/capacitances[i]**4
            sigmaCeff = ceff**2 * np.sqrt(partCap)
            stringToSave += "Capacitance = " + str(ceff) + "\tStd. Dev. = " + str(sigmaCeff) + "\n"
            stringToSave += "Chi-squared = " + str(self.chiSquared)
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
            defaultSaveName += "-" + self.whatFit + str((len(self.fits)-1)//2)
            saveName = asksaveasfile(title="Save All Results", mode='w', defaultextension=".txt", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Text file (*.txt)", ".txt")])
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if saveName is None:
                return
            saveName.write(stringToSave)
            saveName.close()
            self.saveAll.configure(text="Saved")
            self.after(1000, lambda : self.saveAll.configure(text="Export All Results"))
        
        def handle_click(event):
            if self.resultsView.identify_region(event.x, event.y) == "separator":
                if (self.resultsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def handle_motion(event):
            if self.resultsView.identify_region(event.x, event.y) == "separator":
                if (self.resultsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def checkErrorStructure():
            if (self.errorAlphaCheckboxVariable.get() == 1):
                self.errorAlphaEntry.configure(state="normal")
            else:
                self.errorAlphaEntry.configure(state="disabled")
            if (self.errorBetaCheckboxVariable.get() == 1):
                self.errorBetaEntry.configure(state="normal")
                self.errorBetaReCheckbox.configure(state="normal")
            else:
                self.errorBetaEntry.configure(state="disabled")
                self.errorBetaReCheckboxVariable.set(0)
                self.errorBetaReEntry.configure(state="disabled")
                self.errorBetaReCheckbox.configure(state="disabled")
            if (self.errorBetaReCheckboxVariable.get() == 1):
                self.errorBetaReEntry.configure(state="normal")
            else:
                self.errorBetaReEntry.configure(state="disabled")
            if (self.errorGammaCheckboxVariable.get() == 1):
                self.errorGammaEntry.configure(state="normal")
            else:
                self.errorGammaEntry.configure(state="disabled")
            if (self.errorDeltaCheckboxVariable.get() == 1):
                self.errorDeltaEntry.configure(state="normal")
            else:
                self.errorDeltaEntry.configure(state="disabled")
        """
        def simplexFit():
            choice = 0
            fT = 0
            error_alpha = 0
            error_beta = 0
            error_gamma = 0
            error_delta = 0
            try:
                aN = float(self.alphaVariable.get()) if (self.alphaVariable.get() != "") else 1
            except:
                messagebox.showerror("Value error", "Error 25:\nThe assumed noise (alpha) has an invalid value")
                return
            if (self.weightingVariable.get() == "Modulus"):
                choice = 1
            elif (self.weightingVariable.get() == "Proportional"):
                choice = 2
            elif (self.weightingVariable.get() == "Error model"):
                if (self.errorAlphaCheckboxVariable.get() != 1 and self.errorBetaCheckboxVariable.get() != 1 and self.errorGammaCheckboxVariable.get() != 1 and self.errorDeltaCheckboxVariable.get() != 1):
                    messagebox.showerror("Error structure error", "Error 22:\nAt least one parameter must be chosen to use error structure weighting")
                    return
                try:
                    if (self.errorAlphaCheckboxVariable.get() == 1):
                        if (self.errorAlphaVariable.get() == "" or self.errorAlphaVariable.get() == " " or self.errorAlphaVariable.get() == "."):
                            error_alpha = 0
                        else:
                            error_alpha = float(self.errorAlphaVariable.get())
                    if (self.errorBetaCheckboxVariable.get() == 1):
                        if (self.errorBetaVariable.get() == "" or self.errorBetaVariable.get() == " " or self.errorBetaVariable.get() == "."):
                            error_beta = 0
                        else:
                            error_beta = float(self.errorBetaVariable.get())
                    if (self.errorGammaCheckboxVariable.get() == 1):
                        if (self.errorGammaVariable.get() == "" or self.errorGammaVariable.get() == " " or self.errorGammaVariable.get() == "."):
                            error_gamma = 0
                        else:
                            error_gamma = float(self.errorGammaVariable.get())
                    if (self.errorDeltaCheckboxVariable.get() == 1):
                        if (self.errorDeltaVariable.get() == "" or self.errorDeltaVariable.get() == " " or self.errorDeltaVariable.get() == "."):
                            error_delta = 0
                        else:
                            error_delta = float(self.errorDeltaVariable.get())
                except:
                    messagebox.showerror("Value error", "Error 23:\nOne of the error structure parameters has an invalid value")
                    return
                if (error_alpha == 0 and error_beta == 0 and error_gamma == 0 and error_delta == 0):
                    messagebox.showerror("Valid error", "Error 24:\nAt least one of the error structure parameters must be nonzero")
                    return
                choice = 3
            if (self.fitTypeVariable.get() == "Real"):
                fT = 1
            elif (self.fitTypeVariable.get() == "Imaginary"):
                fT = 2
            
            g = np.zeros(int(self.numVoigtVariable.get())*2+1)
            g[0] = self.rDefault
            for i in range(1, len(g), 2):
                g[i] = self.rDefault
                g[i+1] = self.tauDefault
            for i in range(len(self.paramEntryVariables)):
                possibleVal = self.paramEntryVariables[i].get()
                if (possibleVal == '' or possibleVal == '.' or possibleVal == 'E' or possibleVal == 'e'):
                    messagebox.showerror("Value error", "Error 19:\nOne of the parameters is missing a value")
                    return
                try:
                    g[i] = float(possibleVal)
                except:
                    messagebox.showerror("Value error", "Error 21:\nOne of the parameters has an invalid value: " + str(possibleVal))
                    return
            
            bounds = []
#            bounds.append((-1E9 if self.paramComboboxVariables[0].get() == "+ or -" else 0, 1E9))
#            for comboboxVariable in self.paramComboboxVariables[1:]:
#                if (comboboxVariable.get() == "+ or -"):
#                    bounds.append((-1E9, 1E9))
#                    bounds.append((0, 1E9))
#                elif (comboboxVariable.get() == "fixed"):
#                    messagebox.showerror("Fitting error", "One of the variables is set to \"fixed\". Evolutionary fitting cannot be used with fixed variables.")
#                else:
#                    bounds.append((0, 1E9))
#                    bounds.append((0, 1E9))
            self.queue_simplex = queue.Queue()
            
            self.measureModelButton.configure(state="disabled")
            self.freqRangeButton.configure(state="disabled")
            self.magicButton.configure(state="disabled")
            self.parametersButton.configure(state="disabled")
            self.parametersLoadButton.configure(state="disabled")
            self.numVoigtSpinbox.configure(state="disabled")
            self.browseButton.configure(state="disabled")
            self.numMonteEntry.configure(state="disabled")
            self.fitTypeCombobox.configure(state="disabled")
            self.weightingCombobox.configure(state="disabled")
            if (str(self.undoButton['state']) == "disabled"):
                self.undoDisabled = True
            else:
                self.undoDisabled = False
            self.undoButton.configure(state="disabled")
            self.numVoigtMinus.configure(cursor="arrow", bg="gray60")
            self.numVoigtMinus.unbind("<Enter>")
            self.numVoigtMinus.unbind("<Leave>")
            self.numVoigtMinus.unbind("<Button-1>")
            self.numVoigtMinus.unbind("<ButtonRelease-1>")
            self.numVoigtPlus.configure(cursor="arrow", bg="gray60")
            self.numVoigtPlus.unbind("<Enter>")
            self.numVoigtPlus.unbind("<Leave>")
            self.numVoigtPlus.unbind("<Button-1>")
            self.numVoigtPlus.unbind("<ButtonRelease-1>")
            self.prog_bar_simplex = ttk.Progressbar(self.measureModelFrame, orient="horizontal", length=150, mode="indeterminate")
            self.prog_bar_simplex.grid(column=2, row=0, padx=5)
            self.prog_bar_simplex.start(40)
            
            ThreadedTaskSimplex(self.queue_simplex, self.wdata, self.rdata, self.jdata, int(self.numVoigtVariable.get()), choice, aN, fT, g, bounds, error_alpha, error_beta, error_gamma, error_delta).start()
            self.after(100, process_queue_simplex)
        """
        def undoFit():
            if (len(self.undoStack) > 0):
                capNeeded = self.undoCapNeededStack.pop()
                self.capUsed = capNeeded    #Needed to prevent breaking if capacitance isn't fit first, then is fit and undone, then not fit again
                if (not capNeeded):         #If capacitance is not needed (i.e. last run didn't include capacitance)
                    NVE_undo = int((len(self.undoStack[len(self.undoStack)-1])-1)/2)
                    if (NVE_undo < self.nVoigt):
                        while (NVE_undo-self.nVoigt != 0):
                            self.numVoigtSpinbox.invoke("buttondown")
                    elif (NVE_undo > self.nVoigt):
                        while(NVE_undo-self.nVoigt != 0):
                            self.numVoigtSpinbox.invoke("buttonup")
                    self.fits = []
                    self.sigmas = []
                    self.sdrReal = self.undoSdrRealStack.pop()
                    self.sdrImag = self.undoSdrImagStack.pop()
                    self.resultCap = self.undoCapStack.pop()
                    self.sigmaCap = self.undoCapSigmaStack.pop()
                    self.alphaVariable.set(str(self.undoAlphaStack.pop()))
                    self.fitTypeVariable.set(str(self.undoFitTypeStack.pop()))
                    self.numMonteVariable.set(str(self.undoNumSimulationsStack.pop()))
                    self.weightingVariable.set(str(self.undoWeightingStack.pop()))
                    self.whatFit = self.whatFitStack.pop()
                    errorValuesToUse = self.undoErrorModelValuesStack.pop()
                    errorChoicesToUse = self.undoErrorModelChoicesStack.pop()
                    self.errorAlphaCheckboxVariable.set(errorChoicesToUse[0])
                    self.errorBetaCheckboxVariable.set(errorChoicesToUse[1])
                    self.errorBetaReCheckboxVariable.set(errorChoicesToUse[2])
                    self.errorGammaCheckboxVariable.set(errorChoicesToUse[3])
                    self.errorDeltaCheckboxVariable.set(errorChoicesToUse[4])
                    self.errorAlphaVariable.set(str(errorValuesToUse[0]))
                    self.errorBetaVariable.set(str(errorValuesToUse[1]))
                    self.errorBetaReVariable.set(str(errorValuesToUse[2]))
                    self.errorGammaVariable.set(str(errorValuesToUse[3]))
                    self.errorDeltaVariable.set(str(errorValuesToUse[4]))
                    if (self.weightingVariable.get() == "None"):
                        self.alphaLabel.grid_remove()
                        self.alphaEntry.grid_remove()
                        self.errorStructureFrame.grid_remove()
                    elif (self.weightingVariable.get() == "Error model"):
                        self.alphaLabel.grid_remove()
                        self.alphaEntry.grid_remove()
                        self.errorStructureFrame.grid(column=0, row=1, pady=5, sticky="W", columnspan=5)
                        self.checkErrorStructure()
                    else:
                        self.alphaLabel.grid(column=4, row=0)
                        self.alphaEntry.grid(column=5, row=0)
                        self.errorStructureFrame.grid_remove()
                    if (self.topGUI.getFreqUndo() == 1):
                        self.upDelete = self.undoUpDeleteStack.pop()
                        self.lowDelete = self.undoLowDeleteStack.pop()
                        try:
                            if (self.upDelete == 0):
                                self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                                self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                                self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                            else:
                                self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                        except:
                            messagebox.showwarning("Frequency error", "There are more frequencies set to be deleted than data points. The number of frequencies to delete has been reset to 0.")
                            self.upDelete = 0
                            self.lowDelete = 0
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                            self.wdata = self.wdataRaw.copy()
                            self.rdata = self.rdataRaw.copy()
                            self.jdata = self.jdataRaw.copy()
                        #self.rs.setLower(np.log10(min(self.wdata)))
                        #self.rs.setUpper(np.log10(max(self.wdata)))
                        try:
                            self.figFreq.clear()
                            dataColor = "tab:blue"
                            deletedColor = "#A9CCE3"
                            if (self.topGUI.getTheme() == "dark"):
                                dataColor = "cyan"
                            else:
                                dataColor = "tab:blue"
                            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                                self.figFreq = Figure(figsize=(5, 5), dpi=100)
                                self.canvasFreq = FigureCanvasTkAgg(self.figFreq, master=self.freqWindow)
                                self.canvasFreq.get_tk_widget().grid(row=1,column=1, rowspan=5)
                                self.realFreq = self.figFreq.add_subplot(211)
                                self.realFreq.set_facecolor(self.backgroundColor)
                                self.realFreqDeletedLow, = self.realFreq.plot(self.wdataRaw[:self.lowDelete], self.rdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqDeletedHigh, = self.realFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqPlot, = self.realFreq.plot(self.wdata, self.rdata, "o", color=dataColor)
                                self.realFreq.set_xscale("log")
                                self.realFreq.get_xaxis().set_visible(False)
                                self.realFreq.set_title("Real Impedance", color=self.foregroundColor)
                                self.imagFreq = self.figFreq.add_subplot(212)
                                self.imagFreq.set_facecolor(self.backgroundColor)
                                self.imagFreqDeletedLow, = self.imagFreq.plot(self.wdataRaw[:self.lowDelete], -1*self.jdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqDeletedHigh, = self.imagFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], -1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqPlot, = self.imagFreq.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                                self.imagFreq.set_xscale("log")
                                self.imagFreq.set_title("Imaginary Impedance", color=self.foregroundColor)
                                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                                self.canvasFreq.draw()
                        except:
                            pass
                    for i in range(NVE_undo*2 + 1):
                        self.paramEntryVariables[i].set(str(self.undoStack[len(self.undoStack)-1][i]))
                        self.fits.append(self.undoStack[len(self.undoStack)-1][i])
                        self.sigmas.append(self.undoSigmaStack[len(self.undoSigmaStack)-1][i])
                    self.capacitanceCheckboxVariable.set(0)
                    self.capacitanceEntryVariable.set(self.resultCap)
                    try:    #Ignore these lines if the parameter popup has never been used
                        self.capacitanceEntry.configure(state="disabled")
                        self.capacitanceCombobox.configure(state="disabled")
                    except:
                        pass
                    self.resultAlert.grid_remove()
                    self.resultsView.delete(*self.resultsView.get_children())
                    self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%self.fits[0], "%.3g"%self.sigmas[0], "%.3g"%(self.sigmas[0]*2*100/self.fits[0])+"%"))
                    for i in range(1, len(self.fits), 2):
                        if (self.fits[i] == 0):
                            self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%self.fits[i], "%.3g"%self.sigmas[i], "0%"))
                        else:
                            if (abs(self.sigmas[i]*2*100/self.fits[i]) > 100 or np.isnan(self.sigmas[i])):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%self.fits[i], "%.3g"%self.sigmas[i], "%.3g"%(self.sigmas[i]*2*100/self.fits[i])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%self.fits[i], "%.3g"%self.sigmas[i], "%.3g"%(self.sigmas[i]*2*100/self.fits[i])+"%"))
                        if (self.fits[i+1] == 0):
                            self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%self.fits[i+1], "%.3g"%self.sigmas[i+1], "0%"))
                        else:
                            if (abs(self.sigmas[i+1]*2*100/self.fits[i+1]) > 100 or np.isnan(self.sigmas[i+1])):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%self.fits[i+1], "%.3g"%self.sigmas[i+1], "%.3g"%(self.sigmas[i+1]*2*100/self.fits[i+1])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%self.fits[i+1], "%.3g"%self.sigmas[i+1], "%.3g"%(self.sigmas[i+1]*2*100/self.fits[i+1])+"%"))
                    self.resultsView.tag_configure("bad", background="yellow")
                    self.resultRe.configure(text="Ohmic Resistance = %.4g"%self.fits[0] + " ± %.2g"%self.sigmas[0])
                    self.updateReVariable.set("%.4g"%self.fits[0])
                    aic = self.undoAICStack.pop()
                    chi = self.undoChiStack.pop()
                    cor = self.undoCorStack.pop()
                    Zpolar = 0
                    ZpolarSigma = 0
                    for i in range(1, len(self.fits), 2):
                        Zpolar += self.fits[i]
                        ZpolarSigma += self.sigmas[i]**2
                    ZpolarSigma = np.sqrt(ZpolarSigma)
                    self.resultRp.configure(text="Polarization Impedance = %.4g"%Zpolar + " ± %.2g"%ZpolarSigma)
                    capacitances = np.zeros(int((len(self.fits)-1)/2))
                    for i in range(1, len(self.fits), 2):
                        if (self.fits[1] == 0):
                            capacitances[int(i/2)] = 0
                        else:
                            capacitances[int(i/2)] = self.fits[i+1]/self.fits[i]
                    ceff = 1/np.sum(1/capacitances)
                    sigmaCapacitances = np.zeros(len(capacitances))
                    for i in range(1, len(self.fits), 2):
                        if (self.fits[i] == 0 or self.fits[i+1] == 0):
                            sigmaCapacitances[int(i/2)] = 0
                        else:
                            sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((self.sigmas[i+1]/self.fits[i+1])**2 + (self.sigmas[i]/self.fits[i])**2)
                    partCap = 0
                    for i in range(len(capacitances)):
                        partCap += sigmaCapacitances[i]**2/capacitances[i]**4
                    sigmaCeff = ceff**2 * np.sqrt(partCap)
                    self.resultC.configure(text="Capacitance = %.4g"%ceff + " ± %.2g"%sigmaCeff)
                    Zzero = self.fits[0] + self.fits[1]
                    ZzeroSigma = self.sigmas[0]**2 + self.sigmas[1]**2
                    for i in range(3, len(self.fits), 2):
                        Zzero += self.fits[i]
                        ZzeroSigma += self.sigmas[i]**2
                    ZzeroSigma = np.sqrt(ZzeroSigma)
                    self.aR = ""
                    self.aR += "File name: " + self.browseEntry.get() + "\n"
                    self.aR += "Number of data: " + str(self.lengthOfData) + "\n"
                    self.aR += "Number of parameters: " + str(len(self.fits)) + "\n"
                    self.aR += "\nRe (Rsol) = %.8g"%self.fits[0] + " ± %.4g"%self.sigmas[0] + "\n"
                    for i in range(1, len(self.fits), 2):
                        self.aR += "R" + str(int(i/2 + 1)) + " = %.8g"%self.fits[i] + " ± %.4g"%self.sigmas[i] + "\n"
                        self.aR += "Tau" + str(int(i/2 + 1)) + " = %.8g"%self.fits[i+1] + " ± %.4g"%self.sigmas[i+1] + "\n"
                        self.aR += "C" + str(int(i/2 + 1)) + " = %.8g"%capacitances[int(i/2)] + " ± %.4g"%sigmaCapacitances[int(i/2)] + "\n"
                    self.aR += "\nZero frequency impedance = %.8g"%Zzero + " ± %.4g"%ZzeroSigma + "\n"
                    self.aR += "Polarization Impedance = %.8g"%Zpolar + " ± %.4g"%ZpolarSigma + "\n"
                    self.aR += "Capacitance = %.8g"%ceff + " ± %.4g"%sigmaCeff + "\n"
                    self.aR += "\nCorrelation matrix:\n"
                    self.aR += "         Re    "
                    for i in range(1, len(self.fits), 2):
                        if (int(i/2 +1) < 10):
                            self.aR += "  R" + str(int(i/2 + 1)) + "    " + " Tau" + str(int(i/2 + 1)) + "   "
                        else:
                            self.aR += "  R" + str(int(i/2 + 1)) + "   " + " Tau" + str(int(i/2 + 1)) + "  "
                    self.aR += "\n      \u250C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                    #for i in range(1, len(r), 2):
                    #    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                    self.aR += "\n"
                    self.aR += "Re    "
                    for i in range(len(self.fits)):
                        if (i != 0):
                            if (i%2 != 0):
                                if (i/2 + 1 < 10):
                                    self.aR += "R" + str(int(i/2+1)) + "    "
                                else:
                                    self.aR += "R" + str(int(i/2+1)) + "   "
                            else:
                                if (i/2 < 10):
                                    self.aR += "Tau" + str(int(i/2)) + "  "
                                else:
                                    self.aR += "Tau" + str(int(i/2)) + " "
                        self.aR += "\u2502"
                        for j in range(len(self.fits)):
                            if (j <= i):
                                if ("%05.2f"%cor[i][j] == "01.00"):
                                    self.aR += "   1   \u2502"
                                else:
                                    self.aR += " %05.2f"%cor[i][j] + " \u2502"
                        if (i != len(self.fits)-1):
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                            self.aR += "\u253C"
                            if (i == 0):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        else:
                            self.aR += "\n      \u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u2518"
                                else:
                                    self.aR += "\u2534"
                        self.aR += "\n"
                        
                    self.aR += "\nChi-squared statistic = %.4g"%chi + "\n"
                    self.aR += "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-len(self.fits))) + "\n"
                    try:
                        self.aR += "Akaike Information Criterion = %.4g"%aic#(np.log(chi*((1+2*len(r))/self.lengthOfData))) + "\n"
                    except:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    
                    self.undoStack.pop()
                    self.undoSigmaStack.pop()
                    self.needUndo = False
                    self.alreadyChanged = False
                    if (len(self.undoStack) == 0):
                        self.undoButton.configure(state="disabled")
                    #runMeasurementModel()
                else:       #If capacitance is needed
                    NVE_undo = int((len(self.undoStack[len(self.undoStack)-1])-1)/2)
                    if (NVE_undo < self.nVoigt):
                        while (NVE_undo-self.nVoigt != 0):
                            self.numVoigtSpinbox.invoke("buttondown")
                    elif (NVE_undo > self.nVoigt):
                        while(NVE_undo-self.nVoigt != 0):
                            self.numVoigtSpinbox.invoke("buttonup")
                    self.fits = []
                    self.sigmas = []
                    self.sdrReal = self.undoSdrRealStack.pop()
                    self.sdrImag = self.undoSdrImagStack.pop()
                    self.resultCap = self.undoCapStack.pop()
                    self.sigmaCap = self.undoCapSigmaStack.pop()
                    self.alphaVariable.set(str(self.undoAlphaStack.pop()))
                    self.fitTypeVariable.set(str(self.undoFitTypeStack.pop()))
                    self.numMonteVariable.set(str(self.undoNumSimulationsStack.pop()))
                    self.weightingVariable.set(str(self.undoWeightingStack.pop()))
                    errorValuesToUse = self.undoErrorModelValuesStack.pop()
                    errorChoicesToUse = self.undoErrorModelChoicesStack.pop()
                    self.errorAlphaCheckboxVariable.set(errorChoicesToUse[0])
                    self.errorBetaCheckboxVariable.set(errorChoicesToUse[1])
                    self.errorBetaReCheckboxVariable.set(errorChoicesToUse[2])
                    self.errorGammaCheckboxVariable.set(errorChoicesToUse[3])
                    self.errorDeltaCheckboxVariable.set(errorChoicesToUse[4])
                    self.errorAlphaVariable.set(str(errorValuesToUse[0]))
                    self.errorBetaVariable.set(str(errorValuesToUse[1]))
                    self.errorBetaReVariable.set(str(errorValuesToUse[2]))
                    self.errorGammaVariable.set(str(errorValuesToUse[3]))
                    self.errorDeltaVariable.set(str(errorValuesToUse[4]))
                    if (self.weightingVariable.get() == "None"):
                        self.alphaLabel.grid_remove()
                        self.alphaEntry.grid_remove()
                        self.errorStructureFrame.grid_remove()
                    elif (self.weightingVariable.get() == "Error model"):
                        self.alphaLabel.grid_remove()
                        self.alphaEntry.grid_remove()
                        self.errorStructureFrame.grid(column=0, row=1, pady=5, sticky="W", columnspan=5)
                        self.checkErrorStructure()
                    else:
                        self.alphaLabel.grid(column=4, row=0)
                        self.alphaEntry.grid(column=5, row=0)
                        self.errorStructureFrame.grid_remove()
                    if (self.topGUI.getFreqUndo() == 1):
                        self.upDelete = self.undoUpDeleteStack.pop()
                        self.lowDelete = self.undoLowDeleteStack.pop()
                        try:
                            if (self.upDelete == 0):
                                self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                                self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                                self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                            else:
                                self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                        except:
                            messagebox.showwarning("Frequency error", "There are more frequencies set to be deleted than data points. The number of frequencies to delete has been reset to 0.")
                            self.upDelete = 0
                            self.lowDelete = 0
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                            self.wdata = self.wdataRaw.copy()
                            self.rdata = self.rdataRaw.copy()
                            self.jdata = self.jdataRaw.copy()
                        #self.rs.setLower(np.log10(min(self.wdata)))
                        #self.rs.setUpper(np.log10(max(self.wdata)))
                        try:
                            self.figFreq.clear()
                            dataColor = "tab:blue"
                            deletedColor = "#A9CCE3"
                            if (self.topGUI.getTheme() == "dark"):
                                dataColor = "cyan"
                            else:
                                dataColor = "tab:blue"
                            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                                self.figFreq = Figure(figsize=(5, 5), dpi=100)
                                self.canvasFreq = FigureCanvasTkAgg(self.figFreq, master=self.freqWindow)
                                self.canvasFreq.get_tk_widget().grid(row=1,column=1, rowspan=5)
                                self.realFreq = self.figFreq.add_subplot(211)
                                self.realFreq.set_facecolor(self.backgroundColor)
                                self.realFreqDeletedLow, = self.realFreq.plot(self.wdataRaw[:self.lowDelete], self.rdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqDeletedHigh, = self.realFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.realFreqPlot, = self.realFreq.plot(self.wdata, self.rdata, "o", color=dataColor)
                                self.realFreq.set_xscale("log")
                                self.realFreq.get_xaxis().set_visible(False)
                                self.realFreq.set_title("Real Impedance", color=self.foregroundColor)
                                self.imagFreq = self.figFreq.add_subplot(212)
                                self.imagFreq.set_facecolor(self.backgroundColor)
                                self.imagFreqDeletedLow, = self.imagFreq.plot(self.wdataRaw[:self.lowDelete], -1*self.jdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqDeletedHigh, = self.imagFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], -1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                                self.imagFreqPlot, = self.imagFreq.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                                self.imagFreq.set_xscale("log")
                                self.imagFreq.set_title("Imaginary Impedance", color=self.foregroundColor)
                                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                                self.canvasFreq.draw()
                        except:
                            pass
                    for i in range(NVE_undo*2 + 1):
                        self.paramEntryVariables[i].set(str(self.undoStack[len(self.undoStack)-1][i]))
                        self.fits.append(self.undoStack[len(self.undoStack)-1][i])
                        self.sigmas.append(self.undoSigmaStack[len(self.undoSigmaStack)-1][i])
                    self.capacitanceCheckboxVariable.set(1)
                    self.capacitanceEntryVariable.set(self.resultCap)
                    self.capacitanceEntry.configure(state="normal")
                    self.capacitanceCombobox.configure(state="readonly")
                    self.resultAlert.grid_remove()
                    self.resultsView.delete(*self.resultsView.get_children())
                    self.resultsView.insert("", tk.END, text="", values=("Re (Rsol)", "%.5g"%self.fits[0], "%.3g"%self.sigmas[0], "%.3g"%(self.sigmas[0]*2*100/self.fits[0])+"%"))
                    for i in range(1, len(self.fits), 2):
                        if (self.fits[i] == 0):
                            self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%self.fits[i], "%.3g"%self.sigmas[i], "0%"))
                        else:
                            if (abs(self.sigmas[i]*2*100/self.fits[i]) > 100 or np.isnan(self.sigmas[i])):
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%self.fits[i], "%.3g"%self.sigmas[i], "%.3g"%(self.sigmas[i]*2*100/self.fits[i])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=("R"+str(int(i/2)+1), "%.5g"%self.fits[i], "%.3g"%self.sigmas[i], "%.3g"%(self.sigmas[i]*2*100/self.fits[i])+"%"))
                        if (self.fits[i+1] == 0):
                            self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%self.fits[i+1], "%.3g"%self.sigmas[i+1], "0%"))
                        else:
                            if (abs(self.sigmas[i+1]*2*100/self.fits[i+1]) > 100 or np.isnan(self.sigmas[i+1])):
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%self.fits[i+1], "%.3g"%self.sigmas[i+1], "%.3g"%(self.sigmas[i+1]*2*100/self.fits[i+1])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=("Tau"+str(int(i/2)+1), "%.5g"%self.fits[i+1], "%.3g"%self.sigmas[i+1], "%.3g"%(self.sigmas[i+1]*2*100/self.fits[i+1])+"%"))
                    self.resultsView.tag_configure("bad", background="yellow")
                    self.resultRe.configure(text="Ohmic Resistance = %.4g"%self.fits[0] + " ± %.2g"%self.sigmas[0])
                    self.updateReVariable.set("%.4g"%self.fits[0])
                    aic = self.undoAICStack.pop()
                    chi = self.undoChiStack.pop()
                    cor = self.undoCorStack.pop()
                    Zpolar = 0
                    ZpolarSigma = 0
                    for i in range(1, len(self.fits), 2):
                        Zpolar += self.fits[i]
                        ZpolarSigma += self.sigmas[i]**2
                    ZpolarSigma = np.sqrt(ZpolarSigma)
                    self.resultRp.configure(text="Polarization Impedance = %.4g"%Zpolar + " ± %.2g"%ZpolarSigma)
                    Zzero = self.fits[0] + self.fits[1]
                    ZzeroSigma = self.sigmas[0]**2 + self.sigmas[1]**2
                    for i in range(3, len(self.fits), 2):
                        Zzero += self.fits[i]
                        ZzeroSigma += self.sigmas[i]**2
                    ZzeroSigma = np.sqrt(ZzeroSigma)
                    capacitances = np.zeros(int((len(self.fits)-1)/2))
                    for i in range(1, len(self.fits), 2):
                        if (self.fits[1] == 0):
                            capacitances[int(i/2)] = 0
                        else:
                            capacitances[int(i/2)] = self.fits[i+1]/self.fits[i]
                    ceff = 1/(np.sum(1/capacitances) + 1/self.resultCap)
                    sigmaCapacitances = np.zeros(len(capacitances))
                    for i in range(1, len(self.fits), 2):
                        if (self.fits[i] == 0 or self.fits[i+1] == 0):
                            sigmaCapacitances[int(i/2)] = 0
                        else:
                            sigmaCapacitances[int(i/2)] = capacitances[int(i/2)]*np.sqrt((self.sigmas[i+1]/self.fits[i+1])**2 + (self.sigmas[i]/self.fits[i])**2)
                    #sigmaOtherC = (1/rc)*(sc/rc)
                    partCap = 0
                    for i in range(len(capacitances)):
                        partCap += sigmaCapacitances[i]**2/capacitances[i]**4
                    try:
                        partCap += self.sigmaCap**2/self.resultCap**4
                    except:
                        pass
                    sigmaCeff = ceff**2 * np.sqrt(partCap)
                    self.resultC.configure(text="Overall Capacitance = %.4g"%ceff + " ± %.2g"%sigmaCeff)
                    self.aR += "File name: " + self.browseEntry.get() + "\n"
                    self.aR += "Number of data: " + str(self.lengthOfData) + "\n"
                    self.aR += "Number of parameters: " + str(len(self.fits)) + "\n"
                    self.aR += "\nRe (Rsol) = %.8g"%self.fits[0] + " ± %.4g"%self.sigmas[0] + "\n"
                    try:
                        self.aR += "Capacitance = %.8g"%self.resultCap + " ± %.4g"%self.sigmaCap + "\n"
                    except:
                        self.aR += "Capacitance = %.8g"%self.resultCap + " ± NA" + "\n"
                    for i in range(1, len(self.fits), 2):
                        self.aR += "R" + str(int(i/2 + 1)) + " = %.8g"%self.fits[i] + " ± %.4g"%self.sigmas[i] + "\n"
                        self.aR += "Tau" + str(int(i/2 + 1)) + " = %.8g"%self.fits[i+1] + " ± %.4g"%self.sigmas[i+1] + "\n"
                        self.aR += "C" + str(int(i/2 + 1)) + " = %.8g"%capacitances[int(i/2)] + " ± %.4g"%sigmaCapacitances[int(i/2)] + "\n"
                    self.aR += "\nZero frequency impedance = %.8g"%Zzero + " ± %.4g"%ZzeroSigma + "\n"
                    self.aR += "Polarization Impedance = %.8g"%Zpolar + " ± %.4g"%ZpolarSigma + "\n"
                    self.aR += "Overall Capacitance = %.8g"%ceff + " ± %.4g"%sigmaCeff + "\n"
                    self.aR += "\nCorrelation matrix:\n"
                    self.aR += "         Re    "
                    for i in range(1, len(self.fits), 2):
                        if (int(i/2 +1) < 10):
                            self.aR += "  R" + str(int(i/2 + 1)) + "    " + " Tau" + str(int(i/2 + 1)) + "   "
                        else:
                            self.aR += "  R" + str(int(i/2 + 1)) + "   " + " Tau" + str(int(i/2 + 1)) + "  "
                    self.aR += "  Cap "
                    self.aR += "\n      \u250C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                    #for i in range(1, len(r), 2):
                    #    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                    self.aR += "\n"
                    self.aR += "Re    "
                    for i in range(len(self.fits)):
                        if (i != 0):
                            if (i%2 != 0):
                                if (i/2 + 1 < 10):
                                    self.aR += "R" + str(int(i/2+1)) + "    "
                                else:
                                    self.aR += "R" + str(int(i/2+1)) + "   "
                            else:
                                if (i/2 < 10):
                                    self.aR += "Tau" + str(int(i/2)) + "  "
                                else:
                                    self.aR += "Tau" + str(int(i/2)) + " "
                        self.aR += "\u2502"
                        for j in range(len(self.fits)):
                            if (j <= i):
                                if ("%05.2f"%cor[i][j] == "01.00"):
                                    self.aR += "   1   \u2502"
                                elif (np.isnan(cor[i][j])):
                                    self.aR += "   0   \u2502"
                                else:
                                    self.aR += " %05.2f"%cor[i][j] + " \u2502"
                        if (i != len(self.fits)-1):
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                            self.aR += "\u253C"
                            if (i == 0):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        else:
                            self.aR += "\n      \u251C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253C"
                            for k in range(i):
                                self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500"
                                if (k == i-1):
                                    self.aR += "\u253C\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510"
                                else:
                                    self.aR += "\u253C"
                        self.aR += "\n"
                    self.aR += "Cap   \u2502"
                    cor_cap = self.undoCorCapStack.pop()
                    for i in range(len(cor_cap)):
                        if ("%05.2f"%cor_cap[i] == "01.00"):
                            self.aR += "   1   \u2502"
                        elif (np.isnan(cor_cap[i])):
                            self.aR += "   0   \u2502"
                        else:
                            self.aR += " %05.2f"%cor_cap[i] + " \u2502"
                    self.aR += "   1   \u2502"
                    self.aR += "\n      \u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                    for i in range(len(cor_cap)-1):
                        self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2534"
                    self.aR += "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2518\n"
                    
                    self.aR += "\nChi-squared statistic = %.4g"%chi + "\n"
                    self.aR += "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-len(self.fits))) + "\n"
                    try:
                        self.aR += "Akaike Information Criterion = %.4g"%aic#(np.log(chi*((1+2*len(r))/self.lengthOfData))) + "\n"
                    except:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    self.undoStack.pop()
                    self.undoSigmaStack.pop()
                    self.needUndo = False
                    self.alreadyChanged = False
                    if (len(self.undoStack) == 0):
                        self.undoButton.configure(state="disabled")
        
#        def afterUndoRun():
#            self.justUndid = False
#            if (len(self.undoStack) == 0):
#                self.needUndo = False
#                self.undoButton.configure(state="disabled")
#            else:
#                self.undoButton.configure(state="normal")
#                self.needUndo = True
        
#        def redoFit():
#            if (len(self.redoStack) > 0):
#                NVE_redo = int((len(self.redoStack[len(self.redoStack)-1])-1)/2)
#                if (NVE_redo < self.nVoigt):
#                    while (NVE_redo-self.nVoigt != 0):
#                        self.numVoigtSpinbox.invoke("buttondown")
#                elif (NVE_redo > self.nVoigt):
#                    while(NVE_redo-self.nVoigt != 0):
#                        self.numVoigtSpinbox.invoke("buttonup")
#                self.undoStack.append(self.redoStack.pop())
#                self.undoButton.configure(state="normal")
#                if (len(self.redoStack) == 0):
#                    self.redoButton.configure(state="disabled")
        
        def updateRe():
            updateWindow = tk.Toplevel(background=self.backgroundColor)
            self.updateWindows.append(updateWindow)
            updateWindow.title("Update R\u2091")
            updateWindow.iconbitmap(resource_path('img/elephant3.ico'))
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                updateFig = Figure()
                self.updateFigs.append(updateFig)
                updatePlot = updateFig.add_subplot(111)
                updateFig.set_facecolor(self.backgroundColor)
                updatePlot.set_facecolor(self.backgroundColor)
                updatePlot.yaxis.set_ticks_position("both")
                updatePlot.yaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                updatePlot.xaxis.set_ticks_position("both")
                updatePlot.xaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                #updatePlot.spines['bottom'].set_color(self.foregroundColor)
                #updatePlot.spines['top'].set_color(self.foregroundColor)
                #updatePlot.spines['left'].set_color(self.foregroundColor)
                #updatePlot.spines['right'].set_color(self.foregroundColor)
                self.updateReVariable.set(self.paramEntryVariables[0].get())
                Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                for i in range(len(self.wdata)):
                    if (not self.capUsed):
                        Zfit[i] = self.fits[0]
                    else:
                        Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                    for k in range(1, len(self.fits), 2):
                        Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                normalized_residuals_real = np.zeros(len(self.wdata))
                normalized_error_real_below = np.zeros(len(self.wdata))
                normalized_error_real_above = np.zeros(len(self.wdata))
                for i in range(len(self.wdata)):
                    normalized_residuals_real[i] = (self.rdata[i] - Zfit[i].real)/Zfit[i].real
                    normalized_error_real_below[i] = 2*self.sdrReal[i]/Zfit[i].real
                    normalized_error_real_above[i] = -2*self.sdrReal[i]/Zfit[i].real
                residualColor = "tab:blue"
                if (self.topGUI.getTheme() == "dark"):
                    residualColor = "cyan"
                else:
                    residualColor = "tab:blue"
                (fitLine,) = updatePlot.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=residualColor)
                (aboveLine,) = updatePlot.plot(self.wdata, normalized_error_real_above, "--", color="red")
                (belowLine,) = updatePlot.plot(self.wdata, normalized_error_real_below, "--", color="red")
                updatePlot.axhline(0, color="black", linewidth=1.0)
                updatePlot.set_xscale("log")
                updatePlot.set_title("Real Normalized Residuals", color=self.foregroundColor)
                updatePlot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                updatePlot.set_ylabel("(Zr-Z\u0302rmodel)/Zr", color=self.foregroundColor)
            def updatePlotNow(e):
                updateReVar = self.updateReVariable.get()
                try:
                    entryGood = False
                    if (len(updateReVar) >= 2):
                        if (updateReVar[-2] == "E" or updateReVar[-2] == "e"):
                            if (updateReVar[-1] == "-" or updateReVar[-1] == "+"):
                                float(updateReVar[:-2])
                                entryGood = True
                    if (len(updateReVar) >= 1 and not entryGood):
                        if (updateReVar[-1] == "E" or updateReVar[-1] == "e"):
                            float(updateReVar[:-1])
                            entryGood = True
                    if (updateReVar == "" or updateReVar == " " or updateReVar == "-"):
                        entryGood = True
                    elif (not entryGood):
                        float(updateReVar)
                except:
                    messagebox.showerror("Value error", "The value for ohmic resistance must be a real number", parent=updateWindow)
                    return
                cont = False
                if (not self.alreadyChanged):
                    cont = messagebox.askyesno("Update R\u2091?", "This will update the value of the ohmic resistance and remove the confidence interval. Do you wish to continue?", parent=updateWindow)
                if (self.alreadyChanged or cont):
                    self.alreadyChanged = True
                    entryChanged = False
                    if (len(updateReVar) >= 2):
                        if (updateReVar[-2] == "E" or updateReVar[-2] == "e"):
                            if (updateReVar[-1] == "-" or updateReVar[-1] == "+"):
                                updateReVar = updateReVar[:-2]
                                entryChanged = True
                    if (len(updateReVar) >= 1):
                        if (updateReVar[-1] == "E" or updateReVar[-1] == "e"):
                            updateReVar = updateReVar[:-1]
                            entryChanged = True
                    if (updateReVar == "" or updateReVar == " " or updateReVar == "-"):
                        updateReVar = 0
                        entryChanged = True
                    elif (not entryChanged):
                        pass
                    Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                    for i in range(len(self.wdata)):
                        if (not self.capUsed):
                            Zfit[i] = float(updateReVar)
                        else:
                            Zfit[i] = float(updateReVar) + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                        for k in range(1, len(self.fits), 2):
                            Zfit[i] += (self.fits[k]/(1+(1j*2*np.pi*self.wdata[i]*self.fits[k+1])))
                    normalized_residuals_real = np.zeros(len(self.wdata))
                    normalized_error_real_below = np.zeros(len(self.wdata))
                    normalized_error_real_above = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        normalized_residuals_real[i] = (self.rdata[i] - Zfit[i].real)/Zfit[i].real
                        normalized_error_real_below[i] = 2*self.sdrReal[i]/Zfit[i].real
                        normalized_error_real_above[i] = -2*self.sdrReal[i]/Zfit[i].real
                    fitLine.set_ydata(normalized_residuals_real)
                    aboveLine.set_ydata(normalized_error_real_above)
                    belowLine.set_ydata(normalized_error_real_below)   
                    updateCanvas.draw()
                    self.paramEntryVariables[0].set(updateReVar)
                    self.resultRe.configure(text="Ohmic Resistance = %.4g"%float(updateReVar) + " ± 0")
                    self.fits[0] = float(updateReVar)
                    self.sigmas[0] = 0
                    self.resultsView.delete(self.resultsView.get_children()[0])
                    self.resultsView.insert("", 0, text="", values=("Re (Rsol)", "%.5g"%float(updateReVar), "0", "0%"))
                    keepResultAlert = False
                    for i in range(len(self.fits)):
                        if (self.fits[i] != 0):
                            if (abs(self.sigmas[i]*2*100/self.fits[i]) > 100 or np.isnan(self.sigmas[i])):
                                keepResultAlert = True
                                break
                    if (self.capUsed):
                        if (abs(self.sigmaCap*2*100/self.resultCap) > 100 or np.isnan(self.sigmaCap)):
                            keepResultAlert = True
                    if (not keepResultAlert):
                        self.resultAlert.grid_remove()
            
            def updatePlotUp5(percent):
                try:
                    if (not self.alreadyChanged):
                        cont = messagebox.askyesno("Update R\u2091?", "This will update the value of the ohmic resistance and remove the confidence interval. Do you wish to continue?", parent=updateWindow)
                    if (self.alreadyChanged or cont):
                        self.alreadyChanged = True
                        self.updateReVariable.set('{:.5f}'.format(float(self.updateReVariable.get()) + float(self.updateReVariable.get())*0.05))
                        updatePlotNow(None)
                except:
                    messagebox.showwarning("Re update failed", "There was an issue updating Re")
            
            def updatePlotDown5(percent):
                try:
                    if (not self.alreadyChanged):
                        cont = messagebox.askyesno("Update R\u2091?", "This will update the value of the ohmic resistance and remove the confidence interval. Do you wish to continue?", parent=updateWindow)
                    if (self.alreadyChanged or cont):
                        self.alreadyChanged = True
                        self.updateReVariable.set('{:.5f}'.format(float(self.updateReVariable.get()) - float(self.updateReVariable.get())*0.05))
                        updatePlotNow(None)
                except:
                    messagebox.showwarning("Re update failed", "There was an issue updating Re")
                
            def updatePlotUp1(percent):
                try:
                    if (not self.alreadyChanged):
                        cont = messagebox.askyesno("Update R\u2091?", "This will update the value of the ohmic resistance and remove the confidence interval. Do you wish to continue?", parent=updateWindow)
                    if (self.alreadyChanged or cont):
                        self.alreadyChanged = True
                        self.updateReVariable.set('{:.5f}'.format(float(self.updateReVariable.get()) + float(self.updateReVariable.get())*0.01))
                        updatePlotNow(None)
                except:
                    messagebox.showwarning("Re update failed", "There was an issue updating Re")
            
            def updatePlotDown1(e):
                try:
                    if (not self.alreadyChanged):
                        cont = messagebox.askyesno("Update R\u2091?", "This will update the value of the ohmic resistance and remove the confidence interval. Do you wish to continue?", parent=updateWindow)
                    if (self.alreadyChanged or cont):
                        self.alreadyChanged = True
                        self.updateReVariable.set('{:.5f}'.format(float(self.updateReVariable.get()) - float(self.updateReVariable.get())*0.01))
                        updatePlotNow(None)
                except:
                    messagebox.showwarning("Re update failed", "There was an issue updating Re")
            
            updateFrame = tk.Frame(updateWindow, background=self.backgroundColor)
            updateReFrame = tk.Frame(updateFrame, background=self.backgroundColor)
            update5Arrows = tk.Frame(updateFrame, background=self.backgroundColor)
            update1Arrows = tk.Frame(updateFrame, background=self.backgroundColor)
            #updatePlotButton = ttk.Button(updateFrame, text="Update", command=updatePlotNow)
            #updatePlotButton.pack(side=tk.BOTTOM, pady=3)
            #updatePlotButton_ttp = CreateToolTip(updatePlotButton, "Update the value for the ohmic resistance")
            updateLabel = tk.Label(updateReFrame, text="R\u2091 = ", background=self.backgroundColor, foreground=self.foregroundColor)
            updateEntry = ttk.Entry(updateReFrame, textvariable=self.updateReVariable, width=10, validate="all", validatecommand=valcom)
            update5Label = tk.Label(updateFrame, text="5%: ", background=self.backgroundColor, foreground=self.foregroundColor)
            update5Up = tk.Label(update5Arrows, text="\u25B2", background=self.backgroundColor, foreground=self.foregroundColor, cursor="hand2")
            update5Down = tk.Label(update5Arrows, text="\u25BC", background=self.backgroundColor, foreground=self.foregroundColor, cursor="hand2")
            update1Label = tk.Label(updateFrame, text="1%: ", background=self.backgroundColor, foreground=self.foregroundColor)
            update1Up = tk.Label(update1Arrows, text="\u25B2", background=self.backgroundColor, foreground=self.foregroundColor, cursor="hand2")
            update1Down = tk.Label(update1Arrows, text="\u25BC", background=self.backgroundColor, foreground=self.foregroundColor, cursor="hand2")
            updateEntry.bind("<KeyRelease>", updatePlotNow)
            update5Up.bind("<Enter>", lambda e: update5Up.configure(foreground="gray50"))
            update5Up.bind("<Leave>", lambda e: update5Up.configure(foreground=update5Label["foreground"]))
            update5Up.bind("<Button-1>", updatePlotUp5)
            update5Down.bind("<Enter>", lambda e: update5Down.configure(foreground="gray50"))
            update5Down.bind("<Leave>", lambda e: update5Down.configure(foreground=update5Label["foreground"]))
            update5Down.bind("<Button-1>", updatePlotDown5)
            update1Up.bind("<Enter>", lambda e: update1Up.configure(foreground="gray50"))
            update1Up.bind("<Leave>", lambda e: update1Up.configure(foreground=update5Label["foreground"]))
            update1Up.bind("<Button-1>", updatePlotUp1)
            update1Down.bind("<Enter>", lambda e: update1Down.configure(foreground="gray50"))
            update1Down.bind("<Leave>", lambda e: update1Down.configure(foreground=update5Label["foreground"]))
            update1Down.bind("<Button-1>", updatePlotDown1)
            #updateTicker = tk.Spinbox(updateFrame, format="%.2f", textvariable=self.updateReVariable, exportselection=0, from_=0.0, to_=np.inf, width=10)
            #updateTicker.pack(side=tk.RIGHT)
            updateLabel.grid(column=0, row=0)
            updateEntry.grid(column=1, row=0)
            updateReFrame.grid(column=0, row=0, columnspan=4)
            update5Label.grid(column=0, row=1, sticky="W")
            update5Up.grid(column=0, row=0, sticky="S")
            update5Down.grid(column=0, row=1, sticky="N")
            update5Arrows.grid(column=1, row=1, sticky="W")
            update1Up.grid(column=0, row=0, sticky="S")
            update1Down.grid(column=0, row=1, sticky="N")
            update1Label.grid(column=2, row=1, sticky="E")
            update1Arrows.grid(column=3, row=1, sticky="E")
            updateFrame.pack(side=tk.RIGHT, padx=(2, 5))
            updatePlot.xaxis.label.set_fontsize(20)
            updatePlot.yaxis.label.set_fontsize(20)
            updatePlot.title.set_fontsize(30)    
            updateCanvas = FigureCanvasTkAgg(updateFig, updateWindow)
            updateCanvas.draw()
            updateCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, pady=5, padx=5, expand=True)
#            toolbar = NavigationToolbar2Tk(updateCanvas, updateWindow)
#            toolbar.configure(background="white")
#            toolbar.update()
            updateCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            def on_closing():   #Clear the figure before closing the popup
                updateFig.clear()
                updateWindow.destroy()
            updateWindow.protocol("WM_DELETE_WINDOW", on_closing)
        
#        def validateFreqsLow(P):
#            print(P)
#            if P == '':
#                return True
#            if " " in P:
#                return False
#            if "\t" in P:
#                return False
#            if "\n" in P:
#                return False 
#            try:
#                int(P)
#                if (int(P) < 0):     #No negative numbers of frequencies deleted
#                    return False
#                if (self.upperSpinboxVariable.get() != ""):
#                    if (int(self.upperSpinboxVariable.get()) + int(P) >= len(self.wdataRaw)):
##                        messagebox.showwarning("Value error", "The total number of frequencies deleted cannot exceed the total number of frequencies", parent=self.freqWindow)
#                        return False
#                else:
#                    if (int(P) >= len(self.wdataRaw)):
##                        messagebox.showwarning("Value error", "The total number of frequencies deleted cannot exceed the total number of frequencies", parent=self.freqWindow)
#                        return False
#                return True
#            except:
#                return False
#        vcmdFreqsLow = (self.register(validateFreqsLow), '%P')
#        
#        def validateFreqsUp(P):
#            if P == '':
#                return True
#            if " " in P:
#                return False
#            if "\t" in P:
#                return False
#            if "\n" in P:
#                return False 
#            try:
#                int(P)
#                if (int(P) < 0):     #No negative numbers of frequencies deleted
#                    return False
#                else:
#                    return True
#            except:
#                return False
#        vcmdFreqsUp = (self.register(validateFreqsUp), '%P')
        
        def graphOutMagic(event):
            self.magicInput.canvas._tkcanvas.config(cursor="arrow")
#            self.nyCanvas._tkcanvas.delete(self.rec)
            #event.inaxes.set_facecolor("white")
            #self.nyCanvas.draw_idle()
        
        def graphOverMagic(event):
            #axes = event.inaxes
            #autoAxis = event.inaxes
            whichCan = self.magicInput.canvas._tkcanvas
            whichCan.config(cursor="hand2")
        
        def magic_click(event):
            if (event.inaxes is not None):
                min_dist = np.inf
                min_index = 0
                for i in range(len(self.rdata)):
                    if np.sqrt((self.rdata[i] - event.xdata)**2 + (self.jdata[i] + event.ydata)**2) < min_dist:
                        min_dist = np.sqrt((self.rdata[i] - event.xdata)**2 + (self.jdata[i] + event.ydata)**2)
                        min_index = i
                #print("freq:" + str(1/(2*np.pi*self.wdata[min_index])))
                #print("R:" + str(-1*self.jdata[min_index]))          
                #self.numVoigtSpinbox.invoke("buttonup")
                #round(1/(2*np.pi*self.wdata[min_index]), -int(np.floor(np.log10(abs(1/(2*np.pi*self.wdata[min_index]))))) + (5 - 1))
                self.magicElementRVariable.set(round(-1*self.jdata[min_index], -int(np.floor(np.log10(abs(-1*self.jdata[min_index])))) + (5 - 1)))
                self.magicElementTVariable.set(round(1/(2*np.pi*self.wdata[min_index]), -int(np.floor(np.log10(abs(1/(2*np.pi*self.wdata[min_index]))))) + (5 - 1)))
                self.magicElementRealVariable = self.rdata[min_index]
            else:
                pass    #clicked outside of plot

        def magic_update():
            try:
                float(self.magicElementRVariable.get())
                float(self.magicElementTVariable.get())
                self.numVoigtSpinbox.invoke("buttonup")
                self.paramEntryVariables[len(self.paramEntryVariables)-2].set(round(float(self.magicElementRVariable.get()), -int(np.floor(np.log10(abs(float(self.magicElementRVariable.get()))))) + (5 - 1)))
                self.paramEntryVariables[len(self.paramEntryVariables)-1].set(round(float(self.magicElementTVariable.get()), -int(np.floor(np.log10(abs(float(self.magicElementTVariable.get()))))) + (5 - 1)))
                try:
                    self.magicSubplot.plot(self.magicElementRealVariable, float(self.magicElementRVariable.get()), "r+")    #Mark the point that was used
                    self.magicCanvasInput.draw()        #Necessary to refresh the plot
                    #---Change the button to say "Added" for 0.5 seconds---
                    self.magicElementEnterButton.configure(text="Added")
                    self.after(500, lambda: self.magicElementEnterButton.configure(text="Add"))
                except:
                    pass
                self.magicElementTVariable.set('0')
                self.magicElementRVariable.set('0')
                #self.magicInput.clf()
                #self.magicPlot.withdraw()
            except:
                messagebox.showwarning("Bad value(s)", "Invalid element or values", parent=self.magicPlot)
                #self.magicPlot.deiconify()
                self.magicPlot.lift()
        
        def magicFinger():
            if (self.magicPlot.state() != "withdrawn"):
                self.magicPlot.deiconify()
                self.magicPlot.lift()
            else:
                self.magicInput.set_facecolor(self.backgroundColor)
                self.magicPlot.deiconify()
                self.magicPlot.configure(background=self.backgroundColor)
                x = np.array(self.rdata)    #Doesn't plot without this
                y = np.array(self.jdata)
                dataColor = "tab:blue"
                fitColor = "orange"
                if (self.topGUI.getTheme() == "dark"):
                    dataColor = "cyan"
                    fitColor = "gold"
                else:
                    dataColor = "tab:blue"
                    fitColor = "orange"
                if (len(self.fits) > 0):
                    Zfit = np.zeros(len(self.wdata), dtype=np.complex128)
                    if (not self.capUsed):
                        for i in range(len(self.wdata)):
                            Zfit[i] = self.fits[0]
                            for k in range(1, len(self.fits), 2):
                                Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                    else:
                        for i in range(len(self.wdata)):
                            Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*self.wdata[i]*self.resultCap)
                            for k in range(1, len(self.fits), 2):
                                Zfit[i] += (self.fits[k]/(1+(1j*self.wdata[i]*2*np.pi*self.fits[k+1])))
                rightPoint = max(self.rdata)
                topPoint = max(-1*self.jdata)
                with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                    self.magicSubplot = self.magicInput.add_subplot(111)
                    self.magicSubplot.set_facecolor(self.backgroundColor)
                    self.magicSubplot.yaxis.set_ticks_position("both")
                    self.magicSubplot.yaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                    self.magicSubplot.xaxis.set_ticks_position("both")
                    self.magicSubplot.xaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                    pointsPlot, = self.magicSubplot.plot(x, -1*y, "o", color=dataColor)
                    if (len(self.fits) > 0):
                        self.magicSubplot.plot(np.array(Zfit.real), np.array(-1*Zfit.imag), color=fitColor)
                    self.magicSubplot.axis("equal")
                    self.magicSubplot.set_title("Magic Finger Nyquist Plot", color=self.foregroundColor)
                    self.magicSubplot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                    self.magicSubplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    self.magicInput.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                    self.magicCanvasInput.draw()
                self.magicCanvasInput.callbacks.connect('button_press_event', magic_click)
                self.magicElementEnterButton.configure(command=magic_update)
                enterAxes = self.magicCanvasInput.mpl_connect('axes_enter_event', graphOverMagic)
                leaveAxes = self.magicCanvasInput.mpl_connect('axes_leave_event', graphOutMagic)
                annot = self.magicSubplot.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
                annot.set_visible(False)
                def update_annot(ind):
                    x,y = pointsPlot.get_data()
                    xval = x[ind["ind"][0]]
                    yval = y[ind["ind"][0]]
                    annot.xy = (xval, yval)
                    text = "Zr=%.3g"%xval + "\nZj=-%.3g"%yval + "\nf=%.5g"%self.wdata[np.where(self.rdata == xval)][0]
                    annot.set_text(text)
                    #---Check if we're within 5% of the right or top edges, and adjust label positions accordingly
                    if (rightPoint != 0):
                        if (abs(xval - rightPoint)/rightPoint <= 0.2):
                            annot.set_position((-90, -10))
                    if (topPoint != 0):
                        if (abs(yval - topPoint)/topPoint <= 0.05):
                            annot.set_position((10, -20))
                    else:
                        annot.set_position((10, -20))
                def hover(event):
                    vis = annot.get_visible()
                    if event.inaxes == self.magicSubplot:
                        cont, ind = pointsPlot.contains(event)
                        if cont:
                            update_annot(ind)
                            annot.set_visible(True)
                            self.magicInput.canvas.draw_idle()
                        else:
                            if vis:
                                annot.set_position((10,10))
                                annot.set_visible(False)
                                self.magicInput.canvas.draw_idle()
                self.magicInput.canvas.mpl_connect("motion_notify_event", hover)
                
                
            def on_closing():   #Clear the figure before closing the popup
                self.magicInput.clf()
                self.magicPlot.withdraw()
            
            self.magicPlot.protocol("WM_DELETE_WINDOW", on_closing)
        
        def changeFreqs():            
            self.freqWindow.deiconify()
            self.freqWindow.lift()
            self.lowerSpinboxVariable.set(str(self.lowDelete))
            self.upperSpinboxVariable.set(str(self.upDelete))
            #round_to_n = lambda x, n: 0 if x == 0 else round(x, -int(np.floor(np.log10(abs(x)))) + (n - 1))   #To round to n sig figs; from Roy Hyunjin Han at https://stackoverflow.com/questions/3410976/how-to-round-a-number-to-significant-figures-in-python
            dataColor = "tab:blue"
            deletedColor = "#A9CCE3"
            if (self.topGUI.getTheme() == "dark"):
                dataColor = "cyan"
            else:
                dataColor = "tab:blue"
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                self.figFreq = Figure(figsize=(5, 5), dpi=100)
                self.canvasFreq = FigureCanvasTkAgg(self.figFreq, master=self.freqWindow)
                self.canvasFreq.get_tk_widget().grid(row=1,column=1, rowspan=5)
                self.realFreq = self.figFreq.add_subplot(211)
                self.realFreq.set_facecolor(self.backgroundColor)
                self.realFreqDeletedLow, = self.realFreq.plot(self.wdataRaw[:self.lowDelete], self.rdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                self.realFreqDeletedHigh, = self.realFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                self.realFreqPlot, = self.realFreq.plot(self.wdata, self.rdata, "o", color=dataColor)
                self.realFreq.set_xscale("log")
                self.realFreq.get_xaxis().set_visible(False)
                self.realFreq.set_title("Real Impedance", color=self.foregroundColor)
                self.imagFreq = self.figFreq.add_subplot(212)
                self.imagFreq.set_facecolor(self.backgroundColor)
                self.imagFreqDeletedLow, = self.imagFreq.plot(self.wdataRaw[:self.lowDelete], -1*self.jdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                self.imagFreqDeletedHigh, = self.imagFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], -1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                self.imagFreqPlot, = self.imagFreq.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                self.imagFreq.set_xscale("log")
                self.imagFreq.set_title("Imaginary Impedance", color=self.foregroundColor)
                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.canvasFreq.draw()
            def updateFreqs():
                try:
                    if (self.lowerSpinboxVariable.get() == ""):
                        self.lowDelete = 0
                    else:
                        self.lowDelete = int(self.lowerSpinboxVariable.get())
                    if (self.upperSpinboxVariable.get() == ""):
                        self.upDelete = 0
                    else:
                        self.upDelete = int(self.upperSpinboxVariable.get())
                except:
                    messagebox.showwarning("Value error", "The number of frequencies deleted must be a positive integer", parent=self.freqWindow)
                    return
                if (self.lowDelete >= len(self.wdataRaw) or self.upDelete >= len(self.wdataRaw) or (self.lowDelete + self.upDelete) >= len(self.wdataRaw)):
                    messagebox.showwarning("Value error", "The total number of frequencies deleted must be less than the total number of frequencies", parent=self.freqWindow)
                    return
                if (self.lowDelete < 0 or self.upDelete < 0):
                    messagebox.showwarning("Value error", "The number of frequencies deleted must be a positive integer", parent=self.freqWindow)
                    return
                if (self.upDelete == 0):
                    self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                    self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                    self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                else:
                    self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                    self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                    self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                #if (self.upperSpinboxVariable.get() == ""):
                #    self.upperSpinboxVariable.set("0")
                #if (self.lowerSpinboxVariable.get() == ""):
                #    self.lowerSpinboxVariable.set("0")
                #self.rs.setLower(np.log10(min(self.wdata)))
                #self.justUpdated = True
                #self.rs.setUpper(np.log10(max(self.wdata)))
                self.lengthOfData = len(self.wdata)
                self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata))) #%f" % round_to_n(min(self.wdata), 6)).strip("0"))
                self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata))) #%f" % round_to_n(max(self.wdata), 6)).strip("0"))
                self.realFreqPlot.set_ydata(self.rdata)
                self.realFreqPlot.set_xdata(self.wdata)
                self.realFreqDeletedHigh.set_ydata(self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:])
                self.realFreqDeletedHigh.set_xdata(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:])
                self.realFreqDeletedLow.set_ydata(self.rdataRaw[:self.lowDelete])
                self.realFreqDeletedLow.set_xdata(self.wdataRaw[:self.lowDelete])
                self.imagFreqDeletedHigh.set_ydata(-1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:])
                self.imagFreqDeletedHigh.set_xdata(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:])
                self.imagFreqDeletedLow.set_ydata(-1*self.jdataRaw[:self.lowDelete])
                self.imagFreqDeletedLow.set_xdata(self.wdataRaw[:self.lowDelete])
                self.imagFreqPlot.set_ydata(-1*self.jdata)
                self.imagFreqPlot.set_xdata(self.wdata)
                self.canvasFreq.draw()
                #self.updateFreqButton.configure(text="Updated")
                #self.after(500, lambda : self.updateFreqButton.configure(text="Update Frequencies"))
            def changeFreqSpinboxLower(e = None):
                try:
                    higherDelete = 0 if self.upperSpinboxVariable.get() == "" else int(self.upperSpinboxVariable.get())
                    lowerDelete = 0 if self.lowerSpinboxVariable.get() == "" else int(self.lowerSpinboxVariable.get())
                    #print(len(self.wdataRaw)-1-higherDelete-lowerDelete)
                    #self.lowerSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete-lowerDelete)
                    self.upperSpinbox.configure(to=len(self.wdataRaw)-1-lowerDelete)
                    #if (higherDelete == 0 and lowerDelete == 0):
                    #    self.rs.setLower(np.log10(min(self.wdataRaw.copy())))
                    #elif (higherDelete == 0):
                    #    self.rs.setLower(np.log10(min(self.wdataRaw.copy()[lowerDelete:])))
                    #elif (lowerDelete == 0):
                    #    self.rs.setLower(np.log10(min(self.wdataRaw.copy()[:-1*higherDelete])))
                    #else:
                    #    self.rs.setLower(np.log10(min(self.wdataRaw.copy()[lowerDelete:-1*higherDelete])))
                    updateFreqs()
                except:
                    pass
            
            def changeFreqSpinboxUpper(e = None):
                try:
                    higherDelete = 0 if self.upperSpinboxVariable.get() == "" else int(self.upperSpinboxVariable.get())
                    lowerDelete = 0 if self.lowerSpinboxVariable.get() == "" else int(self.lowerSpinboxVariable.get())
                    self.lowerSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete)
                    #self.upperSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete-lowerDelete)
                    #if (higherDelete == 0 and lowerDelete == 0):
                    #    self.rs.setUpper(np.log10(max(self.wdataRaw.copy())))
                    #elif (higherDelete == 0):
                    #    self.rs.setUpper(np.log10(max(self.wdataRaw.copy()[lowerDelete:])))
                    #elif (lowerDelete == 0):
                    #    self.rs.setUpper(np.log10(max(self.wdataRaw.copy()[:-1*higherDelete])))
                    #else:
                    #    self.rs.setUpper(np.log10(max(self.wdataRaw.copy()[lowerDelete:-1*higherDelete])))
                    updateFreqs()
                except:
                    pass
            
            #self.rs.setPaintTicks(True)
            #self.rs.setSnapToTicks(False) 
            #self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
            #self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
#            self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
            #self.rs.setNumberOfMajorTicks(10)
            #self.rs.showMinorTicks(False)
#            self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
            #self.rs.setLower(np.log10(min(self.wdata)))
           # self.rs.setUpper(np.log10(max(self.wdata)))
            #self.rs.setFocus()
            self.lowerLabel = tk.Label(self.freqWindow, text="Number of low frequencies\n to delete", fg=self.foregroundColor, bg=self.backgroundColor)
            #self.rangeLabel = tk.Label(self.freqWindow, text="Log of Frequency", fg=self.foregroundColor, bg=self.backgroundColor)
            self.upperLabel = tk.Label(self.freqWindow, text="Number of high frequencies\n to delete", fg=self.foregroundColor, bg=self.backgroundColor)
            self.lowerLabel.grid(column=0, row=1, pady=(80, 5), padx=3, sticky="N")
            #self.rangeLabel.grid(column=1, row=1, pady=(85, 5), sticky="N")
            self.upperLabel.grid(column=2, row=1, pady=(80, 5), padx=3, sticky="N")
            self.lowerSpinbox = tk.Spinbox(self.freqWindow, from_=0, to=(len(self.wdataRaw)-1), textvariable=self.lowerSpinboxVariable, state="normal", width=6, validate="all", validatecommand=valfreqLow2, command=changeFreqSpinboxLower)
            self.lowerSpinbox.grid(column=0, row=2, padx=(3,0), sticky="N")
            self.upperSpinbox = tk.Spinbox(self.freqWindow, from_=0, to=(len(self.wdataRaw)-1), textvariable=self.upperSpinboxVariable, state="normal", width=6, validate="all", validatecommand=valfreqHigh2, command=changeFreqSpinboxUpper)
            self.upperSpinbox.grid(column=2, row=2, padx=(0,3), sticky="N")
            self.lowerSpinbox.bind("<KeyRelease>", changeFreqSpinboxLower)
            self.upperSpinbox.bind("<KeyRelease>", changeFreqSpinboxUpper)
            self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata))) #%f" % round_to_n(min(self.wdata), 6)).strip("0"))
            self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata))) #%f" % round_to_n(max(self.wdata), 6)).strip("0"))
            self.lowestUndeleted.grid(column=0, row=3, sticky="N")
            self.highestUndeleted.grid(column=2, row=3, sticky="N")
            #self.updateFreqButton = ttk.Button(self.freqWindow, text="Update Frequencies", width=20)
            #self.updateFreqButton.grid(column=1, row=1, pady=30, sticky="S")
            #updateFreqButton_ttp = CreateToolTip(self.updateFreqButton, "Change the number of frequencies used in fitting")
            #def focusRange(event):
            #    self.lowerSpinbox.configure(state="readonly")
            #    self.upperSpinbox.configure(state="readonly")
            #def focusLower(event):
            #    self.lowerSpinbox.configure(state="normal")
            #    self.upperSpinbox.configure(state="readonly")
            #def focusUpper(event):
            #    self.upperSpinbox.configure(state="normal")
            #    self.lowerSpinbox.configure(state="readonly")
            #self.rs.bind("<FocusIn>", lambda e: focusRange(e))
            #self.lowerSpinbox.bind("<FocusIn>", lambda e: focusLower(e))
            #self.lowerSpinbox.bind("<KeyRelease>", changeFreqSpinboxLower)
            #self.upperSpinbox.bind("<FocusIn>", lambda e: focusUpper(e))
            #self.upperSpinbox.bind("<KeyRelease>", changeFreqSpinboxUpper)
            #self.justUpdated = False      
              
            #def changeSlider(event):
            #    pass
                #if (not self.justUpdated):
                #    low = 10**self.rs.getLower()
                #    high = 10**self.rs.getUpper()
                #    for i in range(len(self.wdataRaw)):
                #        if (low >= self.wdataRaw[i]):
                #            self.lowerSpinboxVariable.set(i)
                #    for i in range(len(self.wdataRaw)-1, 0, -1):
                #        if (high <= self.wdataRaw[i]):
                #            self.upperSpinboxVariable.set(len(self.wdataRaw)-1-i)
                #self.justUpdated = False
            #self.rs.subscribe(changeSlider) 
            #self.updateFreqButton.configure(command=updateFreqs)
            
            
            def on_closing():
                self.upperSpinboxVariable.set(str(self.upDelete))
                self.lowerSpinboxVariable.set(str(self.lowDelete))
                self.freqWindow.withdraw()
            self.freqWindow.protocol("WM_DELETE_WINDOW", on_closing)
        
        def simpleParams():
            simplePopup = tk.Toplevel()
            self.simplePopups.append(simplePopup)
            simplePopup.title("Evaluate Simple Parameters")
            simplePopup.iconbitmap(resource_path('img/elephant3.ico'))
            simplePopup.configure(background=self.backgroundColor)
            charFreqLabel = tk.Label(simplePopup, text="Characteristic Frequencies:", bg=self.backgroundColor, fg=self.foregroundColor)
            rFrame = tk.Frame(simplePopup, bg=self.backgroundColor)
            r0label = tk.Label(rFrame, text="R\u2080 = ", bg=self.backgroundColor, fg=self.foregroundColor)
            rplabel = tk.Label(rFrame, text="R\u209A = ", bg=self.backgroundColor, fg=self.foregroundColor)
            simpleResultsFrame = tk.Frame(simplePopup, bg=self.backgroundColor)
            simpleResultsScrollbar = ttk.Scrollbar(simpleResultsFrame, orient=tk.VERTICAL)     
            simpleResults = ttk.Treeview(simpleResultsFrame, columns=("type", "cf", "t", "c"), height=5, selectmode="browse", yscrollcommand=simpleResultsScrollbar.set)
            simpleResults.heading("type", text="Type")
            simpleResults.heading("cf", text="Char. Freq.")
            simpleResults.heading("t", text="Time constant")
            simpleResults.heading("c", text="Capacitance")
            simpleResults.column("#0", width=0)
            simpleResults.column("type", width=120, anchor=tk.CENTER)
            simpleResults.column("cf", width=120, anchor=tk.CENTER)
            simpleResults.column("t", width=120, anchor=tk.CENTER)
            simpleResults.column("c", width=120, anchor=tk.CENTER)
            simpleResultsScrollbar['command'] = simpleResults.yview
            simpleResults.grid(column=0, row=0, sticky="NSEW")
            simpleResultsScrollbar.grid(column=1, row=0, sticky="NS")
            charFreqLabel.grid(column=1, row=0, sticky="SW")
            simpleResultsFrame.grid(column=1, row=1, sticky="W")
            r0label.grid(column=0, row=0, sticky="NW")
            rplabel.grid(column=0, row=1, sticky="NW")
            rFrame.grid(column=1, row=2, sticky="NW")
            
            simpleFreqsA = np.logspace(-10, 10, 5000)
            ZfitA = np.zeros(len(simpleFreqsA), dtype=np.complex128)
            if (not self.capUsed):
                for i in range(len(simpleFreqsA)):
                    ZfitA[i] = self.fits[0]
                    for k in range(1, len(self.fits), 2):
                        ZfitA[i] += (self.fits[k]/(1+(1j*simpleFreqsA[i]*2*np.pi*self.fits[k+1])))
                phase_fit = np.arctan2(ZfitA.imag, ZfitA.real) * (180/np.pi)
            else:
                for i in range(len(simpleFreqsA)):
                    ZfitA[i] = self.fits[0] + 1/(1j*2*np.pi*simpleFreqsA[i]*self.resultCap)
                    for k in range(1, len(self.fits), 2):
                        ZfitA[i] += (self.fits[k]/(1+(1j*simpleFreqsA[i]*2*np.pi*self.fits[k+1])))
                phase_fit = np.arctan2(ZfitA.imag, ZfitA.real) * (180/np.pi)
            Zpolar = self.fits[0]
            ZpolarSigma = self.sigmas[0]
            for i in range(1, len(self.fits), 2):
                Zpolar += self.fits[i]
                ZpolarSigma += self.sigmas[i]**2
            ZpolarSigma = np.sqrt(ZpolarSigma)
            lf = -10
            hf = 10
            for i in range(len(simpleFreqsA)):
                if (abs(ZfitA[i].real - Zpolar)/ZfitA[i].real >= 0.0001):
                    lf = np.log10(simpleFreqsA[i]) if simpleFreqsA[i] < self.wdata[0] else np.log10(self.wdata[0])
                    break
            for i in range(len(simpleFreqsA)-1, 0, -1):
                if (abs(ZfitA[i].real - self.fits[0])/ZfitA[i].real >= 0.0001):
                    hf = np.log10(simpleFreqsA[i]) if simpleFreqsA[i] > self.wdata[len(self.wdata)-1] else np.log10(self.wdata[len(self.wdata)-1])
                    break
            simpleFreqs = np.logspace(lf, hf, 100000)
            Zfit = np.zeros(len(simpleFreqs), dtype=np.complex128)
            if (not self.capUsed):
                for i in range(len(simpleFreqs)):
                    Zfit[i] = self.fits[0]
                    for k in range(1, len(self.fits), 2):
                        Zfit[i] += (self.fits[k]/(1+(1j*simpleFreqs[i]*2*np.pi*self.fits[k+1])))
                phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)
            else:
                for i in range(len(simpleFreqs)):
                    Zfit[i] = self.fits[0] + 1/(1j*2*np.pi*simpleFreqs[i]*self.resultCap)
                    for k in range(1, len(self.fits), 2):
                        Zfit[i] += (self.fits[k]/(1+(1j*simpleFreqs[i]*2*np.pi*self.fits[k+1])))
                phase_fit = np.arctan2(Zfit.imag, Zfit.real) * (180/np.pi)
            #---Local minima/maxima code from https://tcoil.info/find-peaks-and-valleys-in-dataset-with-python/---
            b = (np.diff(np.sign(np.diff(-1*Zfit.imag))) > 0).nonzero()[0] + 1         # local min
            c = (np.diff(np.sign(np.diff(-1*Zfit.imag))) < 0).nonzero()[0] + 1         # local max
            for val in c:
                simpleResults.insert("", tk.END, text="", values=("Maximum", "%.5g"%simpleFreqs[val], "%.5g"%(1/(2*np.pi*simpleFreqs[val])), "%.5g"%(1/(2*np.pi*simpleFreqs[val]*Zpolar))))
            for val in b:
                simpleResults.insert("", tk.END, text="", values=("Minimum", "%.5g"%simpleFreqs[val], "%.5g"%(1/(2*np.pi*simpleFreqs[val])), "%.5g"%(1/(2*np.pi*simpleFreqs[val]*Zpolar))))
            
            r0label.configure(text="R\u2080 = %.5g"%self.fits[0] + " \u00B1 %.3g"%self.sigmas[0])
            rplabel.configure(text="R\u209A = %.5g"%Zpolar + " \u00B1 %.3g"%ZpolarSigma)
            dataColor = "tab:blue"
            fitColor = "orange"
            if (self.topGUI.getTheme() == "dark"):
                dataColor = "cyan"
                fitColor = "gold"
            else:
                dataColor = "tab:blue"
                fitColor = "orange"
            notebook = ttk.Notebook(simplePopup)
            tabNyquist = tk.Frame(notebook, background=self.backgroundColor)
            tabFreq = tk.Frame(notebook, background=self.backgroundColor)
            tabPhase = tk.Frame(notebook, background=self.backgroundColor)
            notebook.add(tabNyquist, text="Nyquist")
            notebook.add(tabFreq, text="-Zj vs. freq.")
            notebook.add(tabPhase, text="Phase angle")
            notebook.grid(column=0, row=0, rowspan=4, sticky="NSEW")
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                simpleFig = Figure()
                simplePlot = simpleFig.add_subplot(111)
                simpleFig.set_facecolor(self.backgroundColor)
                simplePlot.set_facecolor(self.backgroundColor)
                simplePlot.yaxis.set_ticks_position("both")
                simplePlot.yaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                simplePlot.xaxis.set_ticks_position("both")
                simplePlot.xaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                simplePlot.plot(self.rdata, -1*self.jdata, "o", color=dataColor, zorder=2)
                simplePlot.plot(Zfit.real, -1*Zfit.imag, color=fitColor, zorder=1)
                simplePlot.ticklabel_format(axis="y", style="sci", scilimits=(0,0))
                pointsX = [self.fits[0], Zpolar]
                pointsY = [0, 0]
                pointsFreq = ["\u221E", "0"]
                for val in c:
                    pointsX.append(Zfit[val].real)
                    pointsY.append(-1*Zfit[val].imag)
                    pointsFreq.append("%.4g"%simpleFreqs[val])
                for val in b:
                    pointsX.append(Zfit[val].real)
                    pointsY.append(-1*Zfit[val].imag)
                    pointsFreq.append("%.4g"%simpleFreqs[val])
                pointsPlot = simplePlot.scatter(pointsX[2:], pointsY[2:], s=100, c="red", marker="*", zorder=3)
                rPlot = simplePlot.scatter(pointsX[:2], pointsY[:2], s=100, c="red", marker="^", zorder=4)
                simplePlot.axis("equal")
                simplePlot.set_title("Nyquist", color=self.foregroundColor, y=1.06)
                simplePlot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                simplePlot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
            topPoint = max(-1*Zfit.imag)    
            simpleCanvas = FigureCanvasTkAgg(simpleFig, tabNyquist)
            simpleCanvas.draw()
            simpleCanvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, pady=5, padx=5, expand=True)
            toolbar = NavigationToolbar2Tk(simpleCanvas, tabNyquist)
            simpleCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            annot = simplePlot.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
            annot.set_visible(False)
            def update_annot(ind, which):
                if which == 0:
                    pos = pointsPlot.get_offsets()[ind["ind"][0]]
                else:
                   pos = rPlot.get_offsets()[ind["ind"][0]] 
                annot.xy = pos
                if which == 0:
                    text = "Zr=%.4g"%pos[0] + "\nZj=-%.4g"%pos[1] + "\nf=" + pointsFreq[pointsX.index(pos[0])]
                else:
                    text = "Zr=%.4g"%pos[0] + "\nZj=%.4g"%pos[1] + "\nf=" + pointsFreq[pointsX.index(pos[0])]
                annot.set_text(text)
                if (abs(pos[0] - Zpolar)/Zpolar <= 0.05):
                    annot.set_position((-20, 10))
                if (abs(pos[1] - topPoint)/topPoint <= 0.05):
                    annot.set_position((0, 10))
            def hover(event):
                vis = annot.get_visible()
                if event.inaxes == simplePlot:
                    cont, ind = pointsPlot.contains(event)
                    cont2, ind2 = rPlot.contains(event)
                    if cont:
                        update_annot(ind, 0)
                        annot.set_visible(True)
                        simpleFig.canvas.draw_idle()
                    elif cont2:
                        update_annot(ind2, 1)
                        annot.set_visible(True)
                        simpleFig.canvas.draw_idle()
                    else:
                        if vis:
                            annot.set_visible(False)
                            annot.set_position((10,10))
                            simpleFig.canvas.draw_idle()       
            simpleFig.canvas.mpl_connect("motion_notify_event", hover)
            
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                simpleFig3 = Figure()
                simplePlot3 = simpleFig3.add_subplot(111)
                simpleFig3.set_facecolor(self.backgroundColor)
                simplePlot3.set_facecolor(self.backgroundColor)
                simplePlot3.yaxis.set_ticks_position("both")
                simplePlot3.yaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                simplePlot3.xaxis.set_ticks_position("both")
                simplePlot3.xaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                simplePlot3.plot(self.wdata, -1*self.jdata, "o", color=dataColor, zorder=2)
                simplePlot3.plot(simpleFreqs, -1*Zfit.imag, color=fitColor, zorder=1)
                simplePlot3.ticklabel_format(axis="y", style="sci", scilimits=(0,0))
                pointsX3 = []
                pointsY3 = []
                for val in c:
                    pointsX3.append(simpleFreqs[val])
                    pointsY3.append(-1*Zfit[val].imag)
                for val in b:
                    pointsX3.append(simpleFreqs[val])
                    pointsY3.append(-1*Zfit[val].imag)
                pointsPlot3 = simplePlot3.scatter(pointsX3, pointsY3, s=100, c="red", marker="*", zorder=3)
                simplePlot3.set_xscale("log")
                simplePlot3.set_title("Imaginary impedance", color=self.foregroundColor, y=1.06)
                simplePlot3.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                simplePlot3.set_ylabel("-Zj / Ω", color=self.foregroundColor)   
            simpleCanvas3 = FigureCanvasTkAgg(simpleFig3, tabFreq)
            simpleCanvas3.draw()
            simpleCanvas3.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, pady=5, padx=5, expand=True)
            toolbar3 = NavigationToolbar2Tk(simpleCanvas3, tabFreq)
            simpleCanvas3._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            annot3 = simplePlot3.annotate("", xy=(0,0), xytext=(10,10), textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
            annot3.set_visible(False)
            topPoint3 = max(-1*Zfit.imag)
            rightPoint3 = min(simpleFreqs)
            def update_annot3(ind):
                pos = pointsPlot3.get_offsets()[ind["ind"][0]]
                annot3.xy = pos
                text = "Zj=-%.4g"%pos[1] + "\nf=%.4g"%pos[0]
                annot3.set_text(text)
                if (abs(pos[0] - rightPoint3)/rightPoint3 <= 0.05):
                    annot3.set_position((-20, 10))
                if (abs(pos[1] - topPoint3)/topPoint3 <= 0.05):
                    annot3.set_position((15, 0))
            def hover3(event):
                vis3 = annot3.get_visible()
                if event.inaxes == simplePlot3:
                    cont3, ind3 = pointsPlot3.contains(event)
                    if cont3:
                        update_annot3(ind3)
                        annot3.set_visible(True)
                        simpleFig3.canvas.draw_idle()
                    else:
                        if vis3:
                            annot3.set_visible(False)
                            annot3.set_position((10, 10))
                            simpleFig3.canvas.draw_idle()       
            simpleFig3.canvas.mpl_connect("motion_notify_event", hover3)
            
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                simpleFig2 = Figure()
                simplePlot2 = simpleFig2.add_subplot(111)
                simpleFig2.set_facecolor(self.backgroundColor)
                simplePlot2.set_facecolor(self.backgroundColor)
                simplePlot2.yaxis.set_ticks_position("both")
                simplePlot2.yaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                simplePlot2.xaxis.set_ticks_position("both")
                simplePlot2.xaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")
                simplePlot2.plot(self.wdata, np.arctan2(self.jdata, self.rdata)*(180/np.pi), "o", color=dataColor)
                simplePlot2.plot(simpleFreqs, phase_fit, color=fitColor)
                pointsX2 = []
                pointsY2 = []
                for val in c:
                    pointsX2.append(simpleFreqs[val])
                    pointsY2.append(np.arctan2(Zfit[val].imag, Zfit[val].real)*(180/np.pi))
                for val in b:
                    pointsX2.append(simpleFreqs[val])
                    pointsY2.append(np.arctan2(Zfit[val].imag, Zfit[val].real)*(180/np.pi))
                pointsPlot2 = simplePlot2.scatter(pointsX2, pointsY2, s=100, c="red", marker="*", zorder=3)
                simplePlot2.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                simplePlot2.set_ylim(bottom=0, top=-90)
                simplePlot2.set_xscale("log")
                simplePlot2.set_title("Phase Angle", color=self.foregroundColor, y=1.06)
                simplePlot2.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                simplePlot2.set_ylabel("Phase angle / degrees", color=self.foregroundColor)
            simpleCanvas2 = FigureCanvasTkAgg(simpleFig2, tabPhase)
            simpleCanvas2.draw()
            simpleCanvas2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, pady=5, padx=5, expand=True)#.grid(column=0, row=0, pady=5, padx=5, sticky="NSEW")
            topPoint2 = min(np.arctan2(Zfit.imag, Zfit.real)*(180/np.pi))
            rightPoint2 = min(simpleFreqs)
            annot2 = simplePlot2.annotate("", xy=(0,0), xytext=(10,15), textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
            annot2.set_visible(False)
            def update_annot2(ind):
                pos = pointsPlot2.get_offsets()[ind["ind"][0]]
                annot2.xy = pos
                text = "\u03D5=%.4g\u00B0"%pos[1] + "\nf=%.4g"%pos[0]
                annot2.set_text(text)
                if (abs(pos[0] - rightPoint2)/rightPoint2 <= 0.05):
                    annot2.set_position((-20, 15))
                if (abs(pos[1] - topPoint2)/topPoint2 <= 0.05):
                    annot2.set_position((10, 0))
            def hover2(event):
                vis2 = annot2.get_visible()
                if event.inaxes == simplePlot2:
                    cont2, ind2 = pointsPlot2.contains(event)
                    if cont2:
                        update_annot2(ind2)
                        annot2.set_visible(True)
                        simpleFig2.canvas.draw_idle()
                    else:
                        if vis2:
                            annot2.set_visible(False)
                            annot2.set_position((10,15))
                            simpleFig2.canvas.draw_idle()      
            simpleFig2.canvas.mpl_connect("motion_notify_event", hover2)
            toolbar2 = NavigationToolbar2Tk(simpleCanvas2, tabPhase)
            simpleCanvas2._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)#.grid(column=0, row=0, sticky="NSEW")
            tabNyquist.pack_propagate(False)
            tabFreq.pack_propagate(False)
            tabPhase.pack_propagate(False)
            simplePopup.resizable(False, False)
            
        #---The file browse field---
        self.browseFrame = tk.Frame(self, bg=self.backgroundColor)
        self.browseEntry = ttk.Entry(self.browseFrame, state="readonly", width=40)
        self.browseLabel = tk.Label(self.browseFrame, text="Input file:   ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.browseButton = ttk.Button(self.browseFrame, text="Browse...", command=browse)
        self.freqRangeButton = ttk.Button(self.browseFrame, text="Frequency Range", state="disabled", command=changeFreqs)
        self.browseLabel.grid(column=0, row=0, sticky="W")
        self.browseButton.grid(column=1, row=0, sticky="W")
        self.browseEntry.grid(column=2, row=0, sticky="W", padx=5)
        self.freqRangeButton.grid(column=3, row=0, sticky="W")
        self.browseFrame.grid(column=0, row=0, sticky="W")
        browse_ttp = CreateToolTip(self.browseButton, 'Browse for a .mmfile or .mmfitting file')
        frequencyRange_ttp = CreateToolTip(self.freqRangeButton, 'Change frequencies used in fitting')
        
        #---The number of Voigt elements changer---
        self.numVoigtFrame = tk.Frame(self, bg=self.backgroundColor)
        self.numVoigtLabel = tk.Label(self.numVoigtFrame, text="Number of Line Shapes:   ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.numVoigtMinus = tk.Label(self.numVoigtFrame, text="\u2212", bg="gray60", fg="white", justify=tk.CENTER, width=2)
        self.numVoigtTextVariable = tk.StringVar(self, "1")
        self.numVoigtNumber = ttk.Entry(self.numVoigtFrame, textvariable=self.numVoigtTextVariable, width=3, justify=tk.CENTER, state="readonly")
        self.numVoigtPlus = tk.Label(self.numVoigtFrame, text="\u002B", bg="gray60", fg="white", justify=tk.CENTER, width=2)
        self.numVoigtVariable = tk.StringVar(self, "1")
        self.numVoigtSpinbox = tk.Spinbox(self.numVoigtFrame, from_=1, to=99, textvariable=self.numVoigtVariable, exportselection=0, state="disabled", width=3, command=changeNVE)
        self.numMonteLabel = tk.Label(self.numVoigtFrame, text="     Number of Simulations:   ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.numMonteVariable = tk.StringVar(self, self.topGUI.getMC())
        vcmd3 = (self.register(validateMC), '%P')
        self.numMonteEntry = ttk.Entry(self.numVoigtFrame, textvariable=self.numMonteVariable, validate="all", validatecommand=vcmd3, width=6)
        self.numVoigtLabel.grid(column=0, row=0)
        self.numVoigtMinus.grid(column=1, row=0)
        self.numVoigtNumber.grid(column=2, row=0)
        self.numVoigtPlus.grid(column=3, row=0)
        #self.numVoigtSpinbox.grid(column=1, row=0)
        self.numMonteLabel.grid(column=4, row=0)
        self.numMonteEntry.grid(column=5, row=0)
        self.numVoigtFrame.grid(column=0, row=1, sticky="W", pady=15)
        numMonte_ttp = CreateToolTip(self.numMonteEntry, 'Number of simulations for confidence interval calculations')
        
        #---The options frame (fit type, weighting, asssumed noise)---
        self.optionsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.fitTypeLabel = tk.Label(self.optionsFrame, text="Fit type:   ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.fitTypeVariable = tk.StringVar(self, self.topGUI.getFit())
        self.fitTypeCombobox = ttk.Combobox(self.optionsFrame, textvariable=self.fitTypeVariable, value=("Real", "Imaginary", "Complex"), exportselection=0, state="readonly", width=15)
        self.weightingLabel = tk.Label(self.optionsFrame, text="   Weighting:   ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.weightingVariable = tk.StringVar(self, self.topGUI.getWeight())
        self.weightingCombobox = ttk.Combobox(self.optionsFrame, textvariable=self.weightingVariable, value=("None", "Modulus", "Proportional", "Error model"), exportselection=0, state="readonly", width=15)
        self.weightingCombobox.bind("<<ComboboxSelected>>", checkWeight)
        self.alphaLabel = tk.Label(self.optionsFrame, text="   α: ", bg=self.backgroundColor, fg=self.foregroundColor)  #Assumed noise
        self.alphaVariable = tk.StringVar(self, self.topGUI.getAlpha())
        self.alphaEntry = ttk.Entry(self.optionsFrame, textvariable=self.alphaVariable, width=5)
        self.alphaEntry.bind("<FocusIn>", lambda e: self.alphaEntry.selection_range(0, tk.END))
        self.fitTypeLabel.grid(column=0, row=0, sticky="W")
        self.fitTypeCombobox.grid(column=1, row=0)
        self.weightingLabel.grid(column=2, row=0)
        self.weightingCombobox.grid(column=3, row=0)
        self.alphaLabel.grid(column=4, row=0)
        self.alphaEntry.grid(column=5, row=0)
        self.optionsFrame.grid(column=0, row=2, sticky="W", pady=5)
        fitType_ttp = CreateToolTip(self.fitTypeCombobox, 'Which part of the data is being fit')
        weighting_ttp = CreateToolTip(self.weightingCombobox, 'How the regression is being weighted')
        alpha_ttp = CreateToolTip(self.alphaEntry, 'The assumed noise (multiplied by the weighting)')
        
        #---If error structure weighting is chosen---
        self.errorStructureFrame = tk.Frame(self.optionsFrame, bg=self.backgroundColor)
        self.errorAlphaCheckboxVariable = tk.IntVar(self, 0)
        self.errorAlphaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorAlphaCheckboxVariable, text="\u03B1 = ", command=checkErrorStructure)
        self.errorAlphaVariable = tk.StringVar(self, "0.1")
        self.errorAlphaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorAlphaVariable, state="disabled", width=6)
        self.errorBetaCheckboxVariable = tk.IntVar(self, 0)
        self.errorBetaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorBetaCheckboxVariable, text="\u03B2 = ", command=checkErrorStructure)
        self.errorBetaVariable = tk.StringVar(self, "0.1")
        self.errorBetaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorBetaVariable, state="disabled", width=6)
        self.errorBetaReCheckboxVariable = tk.IntVar(self, 0)
        self.errorBetaReVariable = tk.StringVar(self, "1")
        self.errorBetaReCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorBetaReCheckboxVariable, text="Re = ", state="disabled", command=checkErrorStructure)
        self.errorBetaReEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorBetaReVariable, state="disabled", width=6)
        self.errorGammaCheckboxVariable = tk.IntVar(self, 1)
        self.errorGammaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorGammaCheckboxVariable, text="\u03B3 = ", command=checkErrorStructure)
        self.errorGammaVariable = tk.StringVar(self, "0.1")
        self.errorGammaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorGammaVariable, width=6)
        self.errorDeltaCheckboxVariable = tk.IntVar(self, 1)
        self.errorDeltaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorDeltaCheckboxVariable, text="\u03B4 = ", command=checkErrorStructure)
        self.errorDeltaVariable = tk.StringVar(self, "0.1")
        self.errorDeltaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorDeltaVariable, width=6)
        self.errorAlphaCheckbox.grid(column=0, row=0)
        self.errorAlphaEntry.grid(column=1, row=0, padx=(2, 15))
        self.errorBetaCheckbox.grid(column=2, row=0)
        self.errorBetaEntry.grid(column=3, row=0, padx=(2, 15))
        self.errorBetaReCheckbox.grid(column=4, row=0)
        self.errorBetaReEntry.grid(column=5, row=0, padx=(2, 15))
        self.errorGammaCheckbox.grid(column=6, row=0)
        self.errorGammaEntry.grid(column=7, row=0, padx=(2, 15))
        self.errorDeltaCheckbox.grid(column=8, row=0)
        self.errorDeltaEntry.grid(column=9, row=0, padx=(2, 15))
        
        def popup_alpha(event):
            try:
                self.popup_menuA.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuA.grab_release()
        def paste_alpha():
            if (self.focus_get() != self.errorAlphaEntry):
                self.errorAlphaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.errorAlphaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_alpha():
            if (self.focus_get() != self.errorAlphaEntry):
                pyperclip.copy(self.errorAlphaEntry.get())
            else:
                try:
                    stringToCopy = self.errorAlphaEntry.selection_get()
                except:
                    stringToCopy = self.errorAlphaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuA = tk.Menu(self.errorAlphaEntry, tearoff=0)
        self.popup_menuA.add_command(label="Copy", command=copy_alpha)
        self.popup_menuA.add_command(label="Paste", command=paste_alpha)
        self.errorAlphaEntry.bind("<Button-3>", popup_alpha)
        def popup_beta(event):
            try:
                self.popup_menuB.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuB.grab_release()
        def paste_beta():
            if (self.focus_get() != self.errorBetaEntry):
                self.errorBetaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.errorBetaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_beta():
            if (self.focus_get() != self.errorBetaEntry):
                pyperclip.copy(self.errorBetaEntry.get())
            else:
                try:
                    stringToCopy = self.errorBetaEntry.selection_get()
                except:
                    stringToCopy = self.errorBetaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuB = tk.Menu(self.errorBetaEntry, tearoff=0)
        self.popup_menuB.add_command(label="Copy", command=copy_beta)
        self.popup_menuB.add_command(label="Paste", command=paste_beta)
        self.errorBetaEntry.bind("<Button-3>", popup_beta)
        def popup_re(event):
            try:
                self.popup_menuR.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuR.grab_release()
        def paste_re():
            if (self.focus_get() != self.errorBetaReEntry):
                self.errorBetaReEntry.insert(tk.END, pyperclip.paste())
            else:
                self.errorBetaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_re():
            if (self.focus_get() != self.errorBetaReEntry):
                pyperclip.copy(self.errorBetaReEntry.get())
            else:
                try:
                    stringToCopy = self.errorBetaReEntry.selection_get()
                except:
                    stringToCopy = self.errorBetaReEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuR = tk.Menu(self.errorBetaReEntry, tearoff=0)
        self.popup_menuR.add_command(label="Copy", command=copy_re)
        self.popup_menuR.add_command(label="Paste", command=paste_re)
        self.errorBetaReEntry.bind("<Button-3>", popup_re)
        def popup_gamma(event):
            try:
                self.popup_menuG.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuG.grab_release()
        def paste_gamma():
            if (self.focus_get() != self.errorBetaEntry):
                self.errorGammaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.errorGammaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_gamma():
            if (self.focus_get() != self.errorGammaEntry):
                pyperclip.copy(self.errorGammaEntry.get())
            else:
                try:
                    stringToCopy = self.errorGammaEntry.selection_get()
                except:
                    stringToCopy = self.errorGammaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuG = tk.Menu(self.errorGammaEntry, tearoff=0)
        self.popup_menuG.add_command(label="Copy", command=copy_gamma)
        self.popup_menuG.add_command(label="Paste", command=paste_gamma)
        self.errorGammaEntry.bind("<Button-3>", popup_gamma)
        def popup_delta(event):
            try:
                self.popup_menuD.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuD.grab_release()
        def paste_delta():
            if (self.focus_get() != self.errorDeltaEntry):
                self.errorDeltaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.errorDeltaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_delta():
            if (self.focus_get() != self.errorDeltaEntry):
                pyperclip.copy(self.errorDeltaEntry.get())
            else:
                try:
                    stringToCopy = self.errorDeltaEntry.selection_get()
                except:
                    stringToCopy = self.errorDeltaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuD = tk.Menu(self.errorDeltaEntry, tearoff=0)
        self.popup_menuD.add_command(label="Copy", command=copy_delta)
        self.popup_menuD.add_command(label="Paste", command=paste_delta)
        self.errorDeltaEntry.bind("<Button-3>", popup_delta)
        
        #---The fitting parameters---
        self.parametersFrame = tk.Frame(self, bg=self.backgroundColor)
        self.parametersButton = ttk.Button(self.parametersFrame, text="Edit Model Parameters", width=30, state="disabled", command=paramsPopup)
        self.parametersLoadButton = ttk.Button(self.parametersFrame, text="Load Existing Parameters...", state="disabled", command=loadParams)
        self.undoButton = ttk.Button(self.parametersFrame, text="Undo Fit", state="disabled", command=undoFit)
#        self.redoButton = ttk.Button(self.parametersFrame, text="Redo Fit", state="disabled", command=redoFit)
        self.parametersButton.grid(column=0, row=0)
        self.parametersLoadButton.grid(column=1, row=0, padx=5)
        self.undoButton.grid(column=2, row=0)
#        self.redoButton.grid(column=3, row=0, padx=5)
        self.parametersFrame.grid(column = 0, row=3, sticky="W", pady=5)
        params_ttp = CreateToolTip(self.parametersButton, 'View and modify fitting parameters')
        paramsLoad_ttp = CreateToolTip(self.parametersLoadButton, 'Load existing parameters from .mmfitting file')
        undo_ttp = CreateToolTip(self.undoButton, 'Removes current parameters and loads last fits')
        
        #---The run button---
        self.measureModelFrame = tk.Frame(self, bg=self.backgroundColor)
        self.measureModelButton = ttk.Button(self.measureModelFrame, text="Run", state="disabled", width=20, command=runMeasurementModel)
        self.autoButton = ttk.Button(self.measureModelFrame, text="Auto Fit", state="disabled", command=runAutoFitter)
        self.magicButton = ttk.Button(self.measureModelFrame, text="\"Magic Finger\"", state="disabled", command=magicFinger)
        lmButton_ttp = CreateToolTip(self.measureModelButton, 'Regress to the model with Levenberg–Marquardt algorithm')
        autoButton_ttp = CreateToolTip(self.autoButton, 'Opens popup to attempt to automatically fit data')
        magicButton_ttp = CreateToolTip(self.magicButton, 'Opens popup to select initial guess from Nyquist plot')
        
        #self.progressBar = ttk.Progressbar(self.measureModelFrame, mode="indeterminate", orient=tk.HORIZONTAL)
        self.measureModelButton.grid(column=0, row=0, sticky="W")
        self.autoButton.grid(column=1, row=0, sticky="W", padx=5)
        self.magicButton.grid(column=2, row=0, padx=5, sticky="W")
        self.measureModelFrame.grid(column=0, row=4, sticky="W", pady=5)
        
        self.sep = ttk.Separator(self, orient=tk.HORIZONTAL)
        self.sep.grid(column=0, row=5, sticky="EW")
        
        #---The results display---
        self.resultsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.resultRe = tk.Label(self.resultsFrame, text="Re (Rsol) = ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.resultRp = tk.Label(self.resultsFrame, text="Polarization Impedance = ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.resultC = tk.Label(self.resultsFrame, text="Capacitance = ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.resultAlert = tk.Label(self.resultsFrame, text="\u26A0", fg="red2", bg=self.backgroundColor)
        self.copyButton = ttk.Button(self.resultsFrame, text="Copy values and std. devs. as spreadsheet", width=40, command=copyVals)
        self.resultsParamFrame = tk.Frame(self.resultsFrame, bg=self.backgroundColor)
        self.resultsViewScrollbar = ttk.Scrollbar(self.resultsParamFrame, orient=tk.VERTICAL)     
        self.resultsView = ttk.Treeview(self.resultsParamFrame, columns=("param", "value", "sigma", "percent"), height=5, selectmode="browse", yscrollcommand=self.resultsViewScrollbar.set)
        self.resultsView.heading("param", text="Parameter")
        self.resultsView.heading("value", text="Value")
        self.resultsView.heading("sigma", text="Std. Dev.")
        self.resultsView.heading("percent", text="95% CI")
        self.resultsView.column("#0", width=0)
        self.resultsView.column("param", width=120, anchor=tk.CENTER)
        self.resultsView.column("value", width=120, anchor=tk.CENTER)
        self.resultsView.column("sigma", width=120, anchor=tk.CENTER)
        self.resultsView.column("percent", width=120, anchor=tk.CENTER)
        self.resultsViewScrollbar['command'] = self.resultsView.yview
        self.advancedResultsButton = ttk.Button(self.resultsFrame, text="Advanced results", command=advancedResults)
        #self.results= scrolledtext.ScrolledText(self.resultsFrame, width=50, height=10, bg="white", state="disabled")
        self.resultRe.grid(column=0, row=0, sticky="W")
        self.resultRp.grid(column=0, row=1, sticky="W")
        self.resultC.grid(column=0, row=2, sticky="W")
        self.copyButton.grid(column=0, row=4, sticky="E")
        self.resultsView.grid(column=0, row=0, sticky="W")
        self.resultsViewScrollbar.grid(column=1, row=0, sticky="NS")
        self.resultsParamFrame.grid(column=0, row=3, sticky="W")
        self.advancedResultsButton.grid(column=0, row=4, pady=3, sticky="W")
        self.resultsView.bind("<Button-1>", handle_click)
        self.resultsView.bind("<Motion>", handle_motion)
        #self.results.grid(column=0, row=4, sticky="W") 
        copy_ttp = CreateToolTip(self.copyButton, 'Copy fitted values and standard deviations in a format that can be opened in a spreadsheet')
        advancedResults_ttp = CreateToolTip(self.advancedResultsButton, 'Open popup with detailed output')
        resultAlert_ttp = CreateToolTip(self.resultAlert, 'At least one 95% confidence interval is nan or greater than 100%')
        
        #---The plotting options---
        self.graphFrame = tk.Frame(self, bg=self.backgroundColor)
        self.simpleButton = ttk.Button(self.graphFrame, text="Evaluate Simple Parameters", command=simpleParams)
        self.includeFrame = tk.Frame(self.graphFrame, bg=self.backgroundColor)
#        self.includeLabel = tk.Label(self.includeFrame, text="Include: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.confidenceIntervalCheckboxVariable = tk.IntVar(self, 0)
        self.confidenceIntervalCheckbox = ttk.Checkbutton(self.includeFrame, text="Include CI", variable=self.confidenceIntervalCheckboxVariable)
        self.mouseOverCheckboxVariable = tk.IntVar(self, 1)
        self.mouseOverCheckbox = ttk.Checkbutton(self.includeFrame, text="Mouseover labels", variable=self.mouseOverCheckboxVariable)
        self.editReButton = ttk.Button(self.includeFrame, text="Update R\u2091", command=updateRe)
        self.plotButton = ttk.Button(self.graphFrame, text="Plot", command=plotResults)
        self.simpleButton.grid(column=0, row=0, sticky="W", columnspan=2, pady=(8, 5))
        self.plotButton.grid(column=0, row=1, sticky="W", pady=(5,5))
#        self.includeLabel.grid(column=0, row=0, sticky="W")
        self.confidenceIntervalCheckbox.grid(column=1, row=1, sticky="W")
        self.mouseOverCheckbox.grid(column=2, row=1, sticky="W", padx=(20, 5))
#        self.editReLabel.grid(column=2, row=0, sticky="E")
#        self.editRe.grid(column=3, row=0, sticky="E")
        self.editReButton.grid(column=3, row=1, sticky="E", padx=(20,0))
        self.includeFrame.grid(column=1, row=1, sticky="W")
#        self.errorStructureCheckbox.grid(column=3, row=0)
        plot_ttp = CreateToolTip(self.plotButton, 'Opens a window with plots using data and fitted parameters')
        editRe_ttp = CreateToolTip(self.editReButton, 'Opens a window to change ohmic resistance')
        include_ttp = CreateToolTip(self.confidenceIntervalCheckbox, 'Include confidence interval ellipses and bars when plotting')
        mouseover_ttp = CreateToolTip(self.mouseOverCheckbox, 'Show data label on mouseover when viewing larger popup plots')
        simple_ttp = CreateToolTip(self.simpleButton, 'Opens a window to analyze characteristic frequencies')
        
        #---The save buttons---
        self.saveFrame = tk.Frame(self.graphFrame, bg=self.backgroundColor)
        self.saveCurrent = ttk.Button(self.saveFrame, text="Save Current Options and Parameters", width=35, command=saveCurrent)
        self.saveCurrent.grid(column=0, row=0)
        self.saveResiduals = ttk.Button(self.saveFrame, text="Save Residuals for Error Analysis", width=35, command=saveResiduals)
        self.saveResiduals.grid(column=1, row=0, padx=5)
        self.saveAll = ttk.Button(self.saveFrame, text="Export All Results", width=20, command=saveAll)
        self.saveAll.grid(column=0, row=1, sticky="W", pady=5)
        self.saveFrame.grid(column=0, row=2, columnspan=4, pady=10, sticky="W")
        saveParams_ttp = CreateToolTip(self.saveCurrent, 'Saves current options and parameters as a .mmfitting file')
        saveResiduals_ttp = CreateToolTip(self.saveResiduals, 'Saves current residuals as a .mmresiduals file for use in the Error File Preparation tab')
        saveAll_ttp = CreateToolTip(self.saveAll, 'Saves all results as a .txt file, including data, fits, confidence intervals, parameters, and standard deviatons')
    
         #---Close all popups---
        self.closeAllPopupsButton = ttk.Button(self, text="Close all popups", command=self.topGUI.closeAllPopups)
        self.closeAllPopupsButton.grid(column=0, row=8, sticky="W", pady=10)
        closeAllPopups_ttp = CreateToolTip(self.closeAllPopupsButton, 'Close all open popup windows')
    
    def plusNVE_self(self, event):
        self.numVoigtSpinbox.invoke("buttonup")
        self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white")
    def minusNVE_self(self, event):
        self.numVoigtSpinbox.invoke("buttondown")
        self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white")
    
    def setEllipseColor(self, color):
        self.ellipseColor = color
    
    def setThemeLight(self):
        self.backgroundColor = "#FFFFFF"
        self.foregroundColor = "#000000"
        self.configure(bg="#FFFFFF")
        self.browseFrame.configure(background="#FFFFFF")
        self.browseLabel.configure(background="#FFFFFF", foreground="#000000")
        self.numVoigtFrame.configure(background="#FFFFFF")
        self.numVoigtLabel.configure(background="#FFFFFF", foreground="#000000")
        self.numMonteLabel.configure(background="#FFFFFF", foreground="#000000")
        self.optionsFrame.configure(background="#FFFFFF")
        self.fitTypeLabel.configure(background="#FFFFFF", foreground="#000000")
        self.weightingLabel.configure(background="#FFFFFF", foreground="#000000")
        self.alphaLabel.configure(background="#FFFFFF", foreground="#000000")
        self.parametersFrame.configure(background="#FFFFFF")
        self.measureModelFrame.configure(background="#FFFFFF")
        self.resultsFrame.configure(background="#FFFFFF")
        self.resultRe.configure(background="#FFFFFF", foreground="#000000")
        self.resultRp.configure(background="#FFFFFF", foreground="#000000")
        self.resultC.configure(background="#FFFFFF", foreground="#000000")
        self.resultAlert.configure(background="#FFFFFF", foreground="red2")
        self.resultsParamFrame.configure(background="#FFFFFF")
        self.graphFrame.configure(background="#FFFFFF")
        self.includeFrame.configure(background="#FFFFFF")
#        self.includeLabel.configure(background="#FFFFFF", foreground="#000000")
        self.saveFrame.configure(background="#FFFFFF")
        self.errorStructureFrame.configure(background="#FFFFFF")
        self.freqWindow.configure(background="#FFFFFF")
        self.advancedParamFrame.configure(background="#FFFFFF")
        self.advancedOptionsPopup.configure(background="#FFFFFF")
        self.paramListbox.configure(background="#FFFFFF", foreground="#000000")
        self.advancedChoiceLabel.configure(background="#FFFFFF", foreground="#000000")
        #self.rs.configure(background="#FFFFFF")
        #self.rs.configure(tickColor="#000000")
        self.autoFitWindow.configure(bg="#FFFFFF")
        try:
            self.autoErrorFrame.configure(background="#FFFFFF")
            self.autoStatusLabel.configure(background="#FFFFFF", foreground="#000000")
            self.autoMaxLabel.configure(background="#FFFFFF", foreground="#000000")
            self.autoWeightingLabel.configure(background="#FFFFFF", foreground="#000000")
            self.autoNMCLabel.configure(background="#FFFFFF", foreground="#000000")
            self.autoSliderFrame.configure(background="#FFFFFF")
            self.autoRadioLabel.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        try:
            self.undeletedFrame.configure(background="#FFFFFF")
            self.lowestUndeleted.configure(background="#FFFFFF", foreground="#000000")
            self.highestUndeleted.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        #try:
        #    self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
        #    self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
        #    self.rs.setLower(np.log10(min(self.wdata)))
        #    self.rs.setUpper(np.log10(max(self.wdata)))
        #except:
        #    pass
        try:
            self.upperLabel.configure(background="#FFFFFF", foreground="#000000")
            self.lowerLabel.configure(background="#FFFFFF", foreground="#000000")
            #self.rangeLabel.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        try:
            self.progStatus.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        self.paramPopup.configure(background="#FFFFFF")
        try:
            self.buttonFrame.configure(background="#FFFFFF")
        except:
            pass
        try:
            self.reFrame.configure(background="#FFFFFF")
        except:
            pass
        try:
           self.pframe.configure(background="#FFFFFF")
        except:
            pass
        try:
            self.pcanvas.configure(background="#FFFFFF")
        except:
            pass
        try:
            self.reLabel.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        try:
            self.howMany.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        if (self.paramPopup.state() != "withdrawn"):
            for label in self.paramLabels:
                label.configure(background="#FFFFFF", foreground="#000000")
            for label in self.paramFrames:
                label.configure(background="#FFFFFF", foreground="#000000")
        
        try:
            self.advLowerLabel.configure(background="#FFFFFF", foreground="#000000")
            self.advUpperLabel.configure(background="#FFFFFF", foreground="#000000")
            self.advNumberLabel.configure(background="#FFFFFF", foreground="#000000")
            self.advSelectionLabel.configure(background="#FFFFFF", foreground="#000000")
            self.advCustomLabel.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
        
        s = ttk.Style()
        s.configure('TCheckbutton', background='#FFFFFF', foreground="#000000")
        s.configure('TRadiobutton', background='#FFFFFF', foreground='#000000')
        try:
            s.configure('Treeview', background='#FFFFFF', foreground='#000000', fieldbackground='#FFFFFF')
        except:
            pass
        
        try:
            if (tk.Toplevel.winfo_exists(self.resultsWindow)):
                self.advResults.configure(background='#FFFFFF', foreground='#000000')
        except:
            pass
                    
    def setThemeDark(self):
        self.backgroundColor = "#424242"
        self.foregroundColor = "#FFFFFF"
        self.configure(bg="#424242")
        self.browseFrame.configure(background="#424242")
        self.browseLabel.configure(background="#424242", foreground="#FFFFFF")
        self.numVoigtFrame.configure(background="#424242")
        self.numVoigtLabel.configure(background="#424242", foreground="#FFFFFF")
        self.numMonteLabel.configure(background="#424242", foreground="#FFFFFF")
        self.optionsFrame.configure(background="#424242")
        self.fitTypeLabel.configure(background="#424242", foreground="#FFFFFF")
        self.weightingLabel.configure(background="#424242", foreground="#FFFFFF")
        self.alphaLabel.configure(background="#424242", foreground="#FFFFFF")
        self.parametersFrame.configure(background="#424242")
        self.measureModelFrame.configure(background="#424242")
        self.resultsFrame.configure(background="#424242")
        self.resultRe.configure(background="#424242", foreground="#FFFFFF")
        self.resultRp.configure(background="#424242", foreground="#FFFFFF")
        self.resultC.configure(background="#424242", foreground="#FFFFFF")
        self.resultAlert.configure(background="#424242", foreground="red2")
        self.resultsParamFrame.configure(background="#424242")
        self.graphFrame.configure(background="#424242")
        self.includeFrame.configure(background="#424242")
#        self.includeLabel.configure(background="#424242", foreground="#FFFFFF")
        self.saveFrame.configure(background="#424242")
        self.errorStructureFrame.configure(background="#424242")
        self.freqWindow.configure(background="#424242")
        self.advancedParamFrame.configure(background="#424242")
        self.advancedOptionsPopup.configure(background="#424242")
        self.paramListbox.configure(background="#424242", foreground="#FFFFFF")
        self.advancedChoiceLabel.configure(background="#424242", foreground="#FFFFFF")
        #self.rs.configure(background="#424242")
        #self.rs.configure(tickColor="#FFFFFF")
        self.autoFitWindow.configure(background="#424242")
        try:
            self.autoErrorFrame.configure(background="#424242")
            self.autoStatusLabel.configure(background="#424242", foreground="#FFFFFF")
            self.autoMaxLabel.configure(background="#424242", foreground="#FFFFFF")
            self.autoWeightingLabel.configure(background="#424242", foreground="#FFFFFF")
            self.autoNMCLabel.configure(background="#424242", foreground="#FFFFFF")
            self.autoSliderFrame.configure(background="#424242")
            self.autoRadioLabel.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
        try:
            self.undeletedFrame.configure(background="#424242")
            self.lowestUndeleted.configure(background="#424242", foreground="#FFFFFF")
            self.highestUndeleted.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
#        try:
#            self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
#            self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
#            self.rs.setLower(np.log10(min(self.wdata)))
#            self.rs.setUpper(np.log10(max(self.wdata)))
#        except:
#            pass
        try:
            self.upperLabel.configure(background="#424242", foreground="#FFFFFF")
            self.lowerLabel.configure(background="#424242", foreground="#FFFFFF")
            #self.rangeLabel.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
        try:
            self.progStatus.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
        self.paramPopup.configure(background="#424242")
        try:
            self.buttonFrame.configure(background="#424242")
        except:
            pass
        try:
           self.reFrame.configure(background="#424242")
        except:
            pass
        try:
           self.pframe.configure(background="#424242")
        except:
            pass
        try:
            self.pcanvas.configure(background="#424242")
        except:
            pass
        try:
            self.reLabel.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
        try:
            self.howMany.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
        if (self.paramPopup.state() != "withdrawn"):
            for label in self.paramLabels:
                label.configure(background="#424242", foreground="#FFFFFF")
            for label in self.paramFrames:
                label.configure(background="#424242", foreground="#FFFFFF")
        
        try:
            self.advLowerLabel.configure(background="#424242", foreground="#FFFFFF")
            self.advUpperLabel.configure(background="#424242", foreground="#FFFFFF")
            self.advNumberLabel.configure(background="#424242", foreground="#FFFFFF")
            self.advSelectionLabel.configure(background="#424242", foreground="#FFFFFF")
            self.advCustomLabel.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
        
        s = ttk.Style()
        s.configure('TCheckbutton', background='#424242', foreground="#FFFFFF")
        s.configure('TRadiobutton', background='#424242', foreground='#FFFFFF')
        #s.configure('Treeview', background='#424242', foreground='#FFFFFF', fieldbackground='#424242')
        
        try:
            if (tk.Toplevel.winfo_exists(self.resultsWindow)):
                self.advResults.configure(background='#424242', foreground='#FFFFFF')
        except:
            pass
    
    def browseEnter(self, n):
#        if (self.browseEntry.get() == ""):
        try:
            data = np.loadtxt(n)
            w_in = data[:,0]
            r_in = data[:,1]
            j_in = data[:,2]
#            self.wdata = w_in
#            self.rdata = r_in
#            self.jdata = j_in
            self.wdataRaw = w_in
            self.rdataRaw = r_in
            self.jdataRaw = j_in
            doneSorting = False
            while (not doneSorting):
                doneSorting = True
                for i in range(len(self.wdataRaw)-1):
                    if (self.wdataRaw[i] > self.wdataRaw[i+1]):
                        temp = self.wdataRaw[i]
                        self.wdataRaw[i] = self.wdataRaw[i+1]
                        self.wdataRaw[i+1] = temp
                        tempR = self.rdataRaw[i]
                        self.rdataRaw[i] = self.rdataRaw[i+1]
                        self.rdataRaw[i+1] = tempR
                        tempJ = self.jdataRaw[i]
                        self.jdataRaw[i] = self.jdataRaw[i+1]
                        self.jdataRaw[i+1] = tempJ
                        doneSorting = False
            self.wdata = self.wdataRaw.copy()
            self.rdata = self.rdataRaw.copy()
            self.jdata = self.jdataRaw.copy()
            
            if (self.browseEntry.get() != "" and self.topGUI.getFreqLoad() == 1):
                try:
                    if (self.upDelete == 0):
                        self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                        self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                        self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                    else:
                        self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                        self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                        self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                    self.lowerSpinboxVariable.set(str(self.lowDelete))
                    self.upperSpinboxVariable.set(str(self.upDelete))
                except:
                    messagebox.showwarning("Frequency error", "There are more frequencies set to be deleted than data points. The number of frequencies to delete has been reset to 0.")
                    self.upDelete = 0
                    self.lowDelete = 0
                    self.lowerSpinboxVariable.set(str(self.lowDelete))
                    self.upperSpinboxVariable.set(str(self.upDelete))
                    self.wdata = self.wdataRaw.copy()
                    self.rdata = self.rdataRaw.copy()
                    self.jdata = self.jdataRaw.copy()
                #self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                #self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                #self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                #self.rs.setNumberOfMajorTicks(10)
                #self.rs.showMinorTicks(False)
                #self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                #self.rs.setLower(np.log10(min(self.wdata)))
                #self.rs.setUpper(np.log10(max(self.wdata)))
                self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata))) #%f" % round_to_n(min(self.wdata), 6)).strip("0"))
                self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata))) #%f" % round_to_n(max(self.wdata), 6)).strip("0"))
                #self.wdata = self.wdataRaw.copy()
                #self.rdata = self.rdataRaw.copy()
                #self.jdata = self.jdataRaw.copy()
            else:
                self.upDelete = 0
                self.lowDelete = 0
                self.lowerSpinboxVariable.set(str(self.lowDelete))
                self.upperSpinboxVariable.set(str(self.upDelete))
                self.wdata = self.wdataRaw.copy()
                self.rdata = self.rdataRaw.copy()
                self.jdata = self.jdataRaw.copy()
                #self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                #self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                #self.rs.setNumberOfMajorTicks(10)
                #self.rs.showMinorTicks(False)
                #self.rs.setLower(np.log10(min(self.wdata)))
                #self.rs.setUpper(np.log10(max(self.wdata)))
                self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata))) #%f" % round_to_n(min(self.wdata), 6)).strip("0"))
                self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata))) #%f" % round_to_n(max(self.wdata), 6)).strip("0"))
            try:
                self.figFreq.clear()
                dataColor = "tab:blue"
                deletedColor = "#A9CCE3"
                if (self.topGUI.getTheme() == "dark"):
                    dataColor = "cyan"
                else:
                    dataColor = "tab:blue"
                with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                    self.figFreq = Figure(figsize=(5, 5), dpi=100)
                    self.canvasFreq = FigureCanvasTkAgg(self.figFreq, master=self.freqWindow)
                    self.canvasFreq.get_tk_widget().grid(row=1,column=1, rowspan=5)
                    self.realFreq = self.figFreq.add_subplot(211)
                    self.realFreq.set_facecolor(self.backgroundColor)
                    self.realFreqDeletedLow, = self.realFreq.plot(self.wdataRaw[:self.lowDelete], self.rdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                    self.realFreqDeletedHigh, = self.realFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], self.rdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                    self.realFreqPlot, = self.realFreq.plot(self.wdata, self.rdata, "o", color=dataColor)
                    self.realFreq.set_xscale("log")
                    self.realFreq.get_xaxis().set_visible(False)
                    self.realFreq.set_title("Real Impedance", color=self.foregroundColor)
                    self.imagFreq = self.figFreq.add_subplot(212)
                    self.imagFreq.set_facecolor(self.backgroundColor)
                    self.imagFreqDeletedLow, = self.imagFreq.plot(self.wdataRaw[:self.lowDelete], -1*self.jdataRaw[:self.lowDelete], "o", markerfacecolor="None", color=deletedColor)
                    self.imagFreqDeletedHigh, = self.imagFreq.plot(self.wdataRaw[len(self.wdataRaw)-1-self.upDelete:], -1*self.jdataRaw[len(self.wdataRaw)-1-self.upDelete:], "o", markerfacecolor="None", color=deletedColor)
                    self.imagFreqPlot, = self.imagFreq.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                    self.imagFreq.set_xscale("log")
                    self.imagFreq.set_title("Imaginary Impedance", color=self.foregroundColor)
                    self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    self.canvasFreq.draw()
            except:
                pass
            self.lengthOfData = len(self.wdata)
            self.tauDefault = 1/((max(w_in)-min(w_in))/(np.log10(abs(max(w_in)/min(w_in)))))
            self.rDefault = (max(r_in)+min(r_in))/2
            self.cDefault = 1/(self.jdata[0] * -2 * np.pi * min(w_in))
            self.parametersButton.configure(state="normal")
            self.freqRangeButton.configure(state="normal")
            self.parametersLoadButton.configure(state="normal")
            self.measureModelButton.configure(state="normal")
            #self.trfButton.configure(state="normal")
            self.magicButton.configure(state="normal")
            #self.saveCurrent.configure(state="normal")
            self.numVoigtSpinbox.configure(state="readonly")
            self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
            self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtMinus.bind("<ButtonRelease-1>", self.minusNVE_self)
            self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
            self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
            self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtPlus.bind("<ButtonRelease-1>", self.plusNVE_self)
            self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
            if (self.browseEntry.get() == ""):
                self.paramComboboxVariables.append(tk.StringVar(self, "+ or -"))
                self.tauComboboxVariables.append(tk.StringVar(self, "+"))
                self.paramEntryVariables[0] = tk.StringVar(self, str(round(self.rDefault, -int(np.floor(np.log10(abs(self.rDefault)))) + (4 - 1))))
                self.paramEntryVariables.append(tk.StringVar(self, str(round(self.rDefault, -int(np.floor(np.log10(abs(self.rDefault)))) + (4 - 1)))))
                self.paramEntryVariables.append(tk.StringVar(self, str(round(self.tauDefault, -int(np.floor(np.log10(abs(self.tauDefault)))) + (4 - 1)))))
                self.capacitanceEntryVariable.set(str(round(self.cDefault, -int(np.floor(np.log10(abs(self.cDefault)))) + (4 - 1))))
            self.browseEntry.configure(state="normal")
            self.browseEntry.delete(0,tk.END)
            self.browseEntry.insert(0, n)
            self.browseEntry.configure(state="readonly")
        except:
            messagebox.showerror("File error", "Error 9:\nThere was an error loading or reading the file")
    
    def closeWindows(self):
        try:
            self.paramPopup.withdraw()
            self.pframe.destroy()
            self.pcanvas.destroy()
            self.vsb.destroy()
            self.addButton.destroy()
            self.removeButton.destroy()
            self.buttonFrame.destroy()
            self.howMany.destroy()
            self.advancedOptions.destroy()
            self.paramComboboxes.clear()
            self.tauComboboxes.clear()
            self.paramLabels.clear()
            self.paramEntries.clear()
            self.paramFrames.clear()
            self.paramDeleteButtons.clear()
            if (self.paramEntryVariables[0].get() == ''):
                self.paramEntryVariables[0].set(str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1))))
            for i in range(1, len(self.paramEntryVariables), 2):
                if (self.paramEntryVariables[i].get() == ''):
                    self.paramEntryVariables[i].set(str(round(self.rDefault, -int(np.floor(np.log10(self.rDefault))) + (4 - 1))))
                if (self.paramEntryVariables[i+1].get() == ''):
                    self.paramEntryVariables[i+1].set(str(round(self.tauDefault, -int(np.floor(np.log10(self.tauDefault))) + (4 - 1))))
        except:
            pass
        try:
            self.advancedOptionsPopup.withdraw()
            self.paramListbox.delete(0, tk.END)
            self.advCustomLabel.destroy()
            self.advCustomEntry.destroy()
            self.advCheckbox.destroy()
            self.advSelectionLabel.destroy()
            self.advSelection.destroy()
            self.advLowerLabel.destroy()
            self.advLowerEntry.destroy()
            self.advUpperLabel.destroy()
            self.advUpperEntry.destroy()
            self.advNumberLabel.destroy()
            self.advNumberEntry.destroy()
        except:
            pass
        self.autoFitWindow.withdraw()
        self.upperSpinboxVariable.set(str(self.upDelete))
        self.lowerSpinboxVariable.set(str(self.lowDelete))
        self.freqWindow.withdraw()
        self.magicInput.clf()
        self.magicPlot.withdraw()
        for resultWindow in self.resultsWindows:
            try:
                resultWindow.destroy()
            except:
                pass
        for resultPlotBig in self.resultPlotBigs:
            try:
                resultPlotBig.destroy()
            except:
                pass
        for resultPlotBigFig in self.resultPlotBigFigs:
            try:
                resultPlotBigFig.clear()
            except:
                pass
        for updateWindow in self.updateWindows:
            try:
                updateWindow.destroy()
            except:
                pass
        for updateFig in self.updateFigs:
            try:
                updateFig.clear()
            except:
                pass
        for simplePopup in self.simplePopups:
            try:
                simplePopup.destroy()
            except:
                pass
        for pltFig in self.pltFigs:
            try:
                pltFig.clear()
            except:
                pass
        for resultPlot in self.resultPlots:
            try:
                resultPlot.destroy()
            except:
                pass
        for nyCanvasButton in self.saveNyCanvasButtons:
            try:
                nyCanvasButton.destroy()
            except:
                pass
        for nyCanvasButton_ttp in self.saveNyCanvasButton_ttps:
            try:
                del nyCanvasButton_ttp
            except:
                pass
        self.alreadyPlotted = False
    
    def checkErrorStructure(self):
        if (self.errorAlphaCheckboxVariable.get() == 1):
            self.errorAlphaEntry.configure(state="normal")
        else:
            self.errorAlphaEntry.configure(state="disabled")
        if (self.errorBetaCheckboxVariable.get() == 1):
            self.errorBetaEntry.configure(state="normal")
            self.errorBetaReCheckbox.configure(state="normal")
        else:
            self.errorBetaEntry.configure(state="disabled")
            self.errorBetaReCheckboxVariable.set(0)
            self.errorBetaReEntry.configure(state="disabled")
            self.errorBetaReCheckbox.configure(state="disabled")
        if (self.errorBetaReCheckboxVariable.get() == 1):
            self.errorBetaReEntry.configure(state="normal")
        else:
            self.errorBetaReEntry.configure(state="disabled")
        if (self.errorGammaCheckboxVariable.get() == 1):
            self.errorGammaEntry.configure(state="normal")
        else:
            self.errorGammaEntry.configure(state="disabled")
        if (self.errorDeltaCheckboxVariable.get() == 1):
            self.errorDeltaEntry.configure(state="normal")
        else:
            self.errorDeltaEntry.configure(state="disabled")
    
    def cancelThreads(self):
        for t in self.currentThreads:
            t.terminate()
    
    def browseEnterFitting(self, n):
        try:
            self.paramLabels = []
            self.paramEntries = []
            self.paramEntryVariables = []
            #self.paramEntryVariables.append(tk.StringVar(self, ""))
            self.paramComboboxes = []
            self.paramComboboxVariables = []
            #self.paramComboboxVariables.append(tk.StringVar(self, "+"))
            self.tauComboboxes = []
            self.tauComboboxVariables = []
            toLoad = open(n)
            fileToLoad = toLoad.readline().split("filename: ")[1][:-1]
            numberOfVoigt = int(toLoad.readline().split("number_voigt: ")[1][:-1])
            if (numberOfVoigt < 1):
                raise ValueError
            numberOfSimulations = int(toLoad.readline().split("number_simulations: ")[1][:-1])
            if (numberOfSimulations < 1):
                raise ValueError
            fitType = toLoad.readline().split("fit_type: ")[1][:-1]
            if (fitType != "Real" and fitType != "Complex" and fitType != "Imaginary"):
                raise ValueError
            uD = int(toLoad.readline().split("upDelete: ")[1][:-1])
            if (uD < 0):
                raise ValueError
            lD = int(toLoad.readline().split("lowDelete: ")[1][:-1])
            if (lD < 0):
                raise ValueError
            weightType = toLoad.readline().split("weight_type: ")[1][:-1]
            if (weightType != "Modulus" and weightType != "None" and weightType != "Proportional" and weightType != "Error model"):
                raise ValueError
            errorAlpha = toLoad.readline().split("error_alpha: ")[1][:-1]
            errorBeta = toLoad.readline().split("error_beta: ")[1][:-1]
            errorBetaRe = toLoad.readline().split("error_betRe: ")[1][:-1]
            errorGamma = toLoad.readline().split("error_gamma: ")[1][:-1]
            errorDelta = toLoad.readline().split("error_delta: ")[1][:-1]
            useAlpha = False
            useBeta = False
            useBetaRe = False
            useGamma = False
            useDelta = False
            if (errorAlpha[0] != "n"):
                errorAlpha = float(errorAlpha)
                useAlpha = True
            else:
                errorAlpha = float(errorAlpha[1:])
            if (errorBeta[0] != "n"):
                errorBeta = float(errorBeta)
                useBeta = True
            else:
                errorBeta = float(errorBeta[1:])
            if (errorBetaRe[0] != "n" and useBeta):
                errorBetaRe = float(errorBetaRe)
                useBetaRe = True
            else:
                errorBetaRe = float(errorBetaRe[1:])
            if (errorGamma[0] != "n"):
                errorGamma = float(errorGamma)
                useGamma = True
            else:
                errorGamma = float(errorGamma[1:])
            if (errorDelta[0] != "n"):
                errorDelta = float(errorDelta)
                useDelta = True
            else:
                errorDelta = float(errorDelta[1:])
            alphaValue = float(toLoad.readline().split("alpha: ")[1][:-1])
            if (alphaValue <= 0):
                raise ValueError
            capValue = toLoad.readline().split("cap: ")[1][:-1]
            if (capValue == "N"):
                pass
            else:
                capValue = float(capValue)
            capCombo = toLoad.readline().split("cap_combo: ")[1][:-1]
            if (capCombo != "+" and capCombo != "+ or -" and capCombo != "-" and capCombo != "fixed"):
                raise ValueError
            toLoad.readline()
            paramValues = []
            while (True):
                p = toLoad.readline()[:-1]
                if ("paramComboboxes" in p):
                    break
                else:
                    paramValues.append(float(p))
            pComboboxes = []
            while (True):
                p = toLoad.readline()[:-1]
                if ("tauComboboxes" in p):
                    break
                else:
                    if (p != "+" and p != "+ or -" and p != "-" and p != "fixed"):
                        raise ValueError
                    pComboboxes.append(p)
            tComboboxes = []
            while (True):
                p = toLoad.readline()
                if not p:
                    break
                else:
                    p = p[:-1]
                    if (p != "+" and p != "fixed"):
                        raise ValueError
                    tComboboxes.append(p)
            toLoad.close()
            data = np.loadtxt(fileToLoad)
            w_in = data[:,0]
            r_in = data[:,1]
            j_in = data[:,2]
    #            self.wdata = w_in
    #            self.rdata = r_in
    #            self.jdata = j_in
            self.wdataRaw = w_in
            self.rdataRaw = r_in
            self.jdataRaw = j_in
            doneSorting = False
            while (not doneSorting):
                doneSorting = True
                for i in range(len(self.wdataRaw)-1):
                    if (self.wdataRaw[i] > self.wdataRaw[i+1]):
                        temp = self.wdataRaw[i]
                        self.wdataRaw[i] = self.wdataRaw[i+1]
                        self.wdataRaw[i+1] = temp
                        tempR = self.rdataRaw[i]
                        self.rdataRaw[i] = self.rdataRaw[i+1]
                        self.rdataRaw[i+1] = tempR
                        tempJ = self.jdataRaw[i]
                        self.jdataRaw[i] = self.jdataRaw[i+1]
                        self.jdataRaw[i+1] = tempJ
                        doneSorting = False
            self.upDelete = uD
            self.lowDelete = lD
            if (self.upDelete == 0):
                self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                self.jdata = self.jdataRaw.copy()[self.lowDelete:]
            else:
                self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
            self.lengthOfData = len(self.wdata)
            self.tauDefault = 1/((max(w_in)-min(w_in))/(np.log10(max(w_in)/min(w_in))))
            self.rDefault = (max(r_in)+min(r_in))/2
            self.cDefault = 1/(self.jdata[0] * -2 * np.pi * min(w_in))
            self.browseEntry.configure(state="normal")
            self.browseEntry.delete(0,tk.END)
            self.browseEntry.insert(0, fileToLoad)
            self.browseEntry.configure(state="readonly")
            self.parametersButton.configure(state="normal")
            self.freqRangeButton.configure(state="normal")
            self.parametersLoadButton.configure(state="normal")
            self.measureModelButton.configure(state="normal")
            #self.trfButton.configure(state="normal")
            self.magicButton.configure(state="normal")
            #self.saveCurrent.configure(state="normal")
            self.numMonteVariable.set(str(numberOfSimulations))
            self.alphaVariable.set(str(alphaValue))
            self.fitTypeVariable.set(fitType)
            self.weightingVariable.set(weightType)
            if (weightType == "Error model"):
                if (not useAlpha):
                    self.errorAlphaCheckboxVariable.set(0)
                    self.errorAlphaEntry.configure(state="disabled")
                    self.errorAlphaVariable.set(errorAlpha)
                else:
                    self.errorAlphaCheckboxVariable.set(1)
                    self.errorAlphaEntry.configure(state="normal")
                    self.errorAlphaVariable.set(errorAlpha)
                if (not useBeta):
                    self.errorBetaCheckboxVariable.set(0)
                    self.errorBetaEntry.configure(state="disabled")
                    self.errorBetaVariable.set(errorBeta)
                else:
                    self.errorBetaCheckboxVariable.set(1)
                    self.errorBetaEntry.configure(state="normal")
                    self.errorBetaVariable.set(errorBeta)
                if (not useBetaRe):
                    self.errorBetaReCheckboxVariable.set(0)
                    self.errorBetaReEntry.configure(state="disabled")
                    self.errorBetaReVariable.set(errorBetaRe)
                else:
                    self.errorBetaReCheckboxVariable.set(1)
                    self.errorBetaReCheckbox.configure(state="normal")
                    self.errorBetaReEntry.configure(state="normal")
                    self.errorBetaReVariable.set(errorBetaRe)
                if (not useGamma):
                    self.errorGammaCheckboxVariable.set(0)
                    self.errorGammaEntry.configure(state="disabled")
                    self.errorGammaVariable.set(errorGamma)
                else:
                    self.errorGammaCheckboxVariable.set(1)
                    self.errorGammaEntry.configure(state="normal")
                    self.errorGammaVariable.set(errorGamma)
                if (not useDelta):
                    self.errorDeltaCheckboxVariable.set(0)
                    self.errorDeltaEntry.configure(state="disabled")
                    self.errorDeltaVariable.set(errorDelta)
                else:
                    self.errorDeltaCheckboxVariable.set(1)
                    self.errorDeltaEntry.configure(state="normal")
                    self.errorDeltaVariable.set(errorDelta)
            if (self.weightingVariable.get() == "None"):
                self.alphaLabel.grid_remove()
                self.alphaEntry.grid_remove()
                self.errorStructureFrame.grid_remove()
            elif (self.weightingVariable.get() == "Error model"):
                self.alphaLabel.grid_remove()
                self.alphaEntry.grid_remove()
                self.errorStructureFrame.grid(column=0, row=1, pady=5, sticky="W", columnspan=5)
                self.checkErrorStructure()
            else:
                self.alphaLabel.grid(column=4, row=0)
                self.alphaEntry.grid(column=5, row=0)
                self.errorStructureFrame.grid_remove()
            self.numVoigtVariable.set(numberOfVoigt)
            self.numVoigtSpinbox.configure(state="readonly")
            self.numVoigtTextVariable.set(self.numVoigtVariable.get())
            self.numVoigtMinus.bind("<Enter>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue1", fg="white"))
            self.numVoigtMinus.bind("<Leave>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtMinus.bind("<Button-1>", lambda e: self.numVoigtMinus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtMinus.bind("<ButtonRelease-1>", self.minusNVE_self)
            self.numVoigtMinus.configure(cursor="hand2", bg="DodgerBlue3")
            self.numVoigtPlus.bind("<Enter>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue1", fg="white"))
            self.numVoigtPlus.bind("<Leave>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtPlus.bind("<Button-1>", lambda e: self.numVoigtPlus.configure(bg="DodgerBlue3", fg="white"))
            self.numVoigtPlus.bind("<ButtonRelease-1>", self.plusNVE_self)
            self.numVoigtPlus.configure(cursor="hand2", bg="DodgerBlue3")
            self.nVoigt = numberOfVoigt
            self.paramComboboxVariables.clear()
            self.tauComboboxVariables.clear()
            self.paramEntryVariables.clear()
            #---Clear existing data and load---
            for pvar in pComboboxes:
                self.paramComboboxVariables.append(tk.StringVar(self, pvar))
            for tvar in tComboboxes:
                self.tauComboboxVariables.append(tk.StringVar(self, tvar))
            self.paramEntryVariables.append(tk.StringVar(self, str(paramValues[0])))
            for pval in paramValues[1:]:
                self.paramEntryVariables.append(tk.StringVar(self, str(pval)))
            if (capValue == "N"):
                self.capacitanceCheckboxVariable.set(0)
                self.capacitanceEntryVariable.set(str(round(self.cDefault, -int(np.floor(np.log10(abs(self.cDefault)))) + (4 - 1))))
                self.capacitanceComboboxVariable.set("+ or -")
            else:
                self.capacitanceCheckboxVariable.set(1)
                self.capacitanceEntryVariable.set(str(capValue))
                self.capacitanceComboboxVariable.set(capCombo)
        except:
            messagebox.showerror("File error", "Error 9:\nThere was an error loading or reading the file")