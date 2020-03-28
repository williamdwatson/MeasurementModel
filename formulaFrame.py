# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 10:29:23 2019

@author: willdwat
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename, asksaveasfile
import os
import sys
import ctypes
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import customFitting
import threading
import queue
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
import tkinter.scrolledtext as scrolledtext
import re
from functools import partial
import pyperclip
#--------------------------------pyperclip-----------------------------------
#     Source: https://pypi.org/project/pyperclip/
#     Email: al@inventwithpython.com
#     By: Al Sweigart
#     License: BSD
#----------------------------------------------------------------------------
import matplotlib.patches as patches

class FileExtensionError(Exception):
    pass

class FileLengthError(Exception):
    pass

class FrequencyMismatchError(Exception):
    pass

class ThreadedTaskCustom(threading.Thread):
    def __init__(self, queue, fitType, numMonte, weight, noise, customFormula, w, r, j, paramNames, paramGuesses, paramLimits, errorParams):
        threading.Thread.__init__(self)
        self.queue = queue
        self.fT = fitType
        self.nM = numMonte
        self.wght = weight
        self.aN = noise
        self.formula = customFormula
        self.wdat = w
        self.rdat = r
        self.jdat = j
        self.pN = paramNames
        self.pG = paramGuesses
        self.pL = paramLimits
        self.eP = errorParams
        self.fittingObject = customFitting.customFitter()
    def run(self):
        try:
            r, s, sdr, sdi, chi, aic, rF, iF = self.fittingObject.findFit(self.fT, self.nM, self.wght, self.aN, self.formula, self.wdat, self.rdat, self.jdat, self.pN, self.pG, self.pL, self.eP)
            self.queue.put((r, s, sdr, sdi, chi, aic, rF, iF))
        except:
            self.queue.put(("@", "@", "@", "@", "@", "@", "@", "@"))
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

class CreateToolTip(object):
    """
    create a tooltip for a given widget
    Code from: https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter with answer by crxguy52
    """
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

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

class fF(tk.Frame):
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
        self.ellipseColor = self.topGUI.getEllipseColor()
        self.saveNeeded = False
        
        self.wdataRaw = []
        self.rdataRaw = []
        self.jdataRaw = []
        self.lengthOfData = 0
        self.freqWindow = tk.Toplevel(background=self.backgroundColor)
        self.freqWindow.withdraw()
        self.freqWindow.title("Change Frequency Range: Custom Formula Fitting")
        self.freqWindow.iconbitmap(resource_path('img/elephant3.ico'))
        self.lowestUndeleted = tk.Label(self.freqWindow, text="Lowest Remaining Frequency: 0", background=self.backgroundColor, foreground=self.foregroundColor)
        self.highestUndeleted = tk.Label(self.freqWindow, text="Highest Remaining Frequency: 0", background=self.backgroundColor, foreground=self.foregroundColor)
        self.rs = RangeSlider(self.freqWindow, lowerBound=-4, upperBound=7, background=self.backgroundColor, tickColor=self.foregroundColor)
        self.rs.grid(column=1, row=1)
        self.lowerSpinboxVariable = tk.StringVar(self, "0")
        self.upperSpinboxVariable = tk.StringVar(self, "0")
        self.upDelete = 0
        self.lowDelete = 0
        
        self.paramPopup = tk.Toplevel(bg=self.backgroundColor)
        self.paramPopup.withdraw()
        self.pcanvas = tk.Canvas(self.paramPopup, borderwidth=0, highlightthickness=0, background=self.backgroundColor)
        self.pframe = tk.Frame(self.pcanvas, background=self.backgroundColor)
        self.vsb = tk.Scrollbar(self.paramPopup, orient="vertical", command=self.pcanvas.yview)
        self.buttonFrame = tk.Frame(self.paramPopup, bg=self.backgroundColor)
        self.addButton = ttk.Button(self.buttonFrame, text="Add Parameter")
        self.removeButton = ttk.Button(self.buttonFrame, text="Remove Last Parameter")
        addButton_ttp = CreateToolTip(self.addButton, "Add a parameter")
        removeButton_ttp = CreateToolTip(self.removeButton, "Remove last parameter")
        self.howMany = tk.Label(self.paramPopup, text="Number of parameters: 0", bg=self.backgroundColor, fg=self.foregroundColor)
        self.addButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.removeButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.howMany.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.buttonFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=False, pady=1)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.pcanvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.pcanvas.create_window((4,4), window=self.pframe, anchor="nw")
        self.varNum = 0
        self.aR = ""
        self.currentThreads = []
        
        def _on_mousewheel_help(event):
            xpos, ypos = self.vsbh.get()
            xpos_round = round(xpos, 2)     #Avoid floating point errors
            ypos_round = round(ypos, 2)
            if (xpos_round != 0.00 or ypos_round != 1.00):
                self.hcanvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.helpPopup = tk.Toplevel(bg=self.backgroundColor)
        self.helpPopup.withdraw()
        self.hcanvas = tk.Canvas(self.helpPopup, borderwidth=0, highlightthickness=0, background=self.backgroundColor)
        self.hframe = tk.Frame(self.hcanvas, background=self.backgroundColor)
        self.helpText = "BE CAREFUL WITH THE CODE, AS IT HAS THE POWER TO MODIFY THE PROGRAM OR DAMAGE YOUR COMPUTER.\n\nAny valid Python code can be entered into the formula box. Clicking the \"Run fitting\" button will check the syntax of the code and then execute it.\n\n"\
                        "The fitting algorithm will vary the parameters listed in the \"Fitting Parameters\" window. The names given to these parameters can be used and modified in the code. Parameters cannot be named after Python reserved words.\n\n"\
                        "The frequencies from the input file can be accessed as an array named \"freq\". For instance, to access the first frequency, use \"freq[0]\" in the code.\n"\
                        "The imaginary part of the impendance is an array named Zj. For instance, to access the first imaginary impedance point, use \"Zj[0]\" in the code.\n"\
                        "The real part of the impedance is an array named Zr. For instance, to access the first real impedance point, use \"Zr[0]\" in the code.\n\n"\
                        "The imaginary number can be accessed as \"1j\". Built-in functions include: PI, SQRT, SIN, COS, TAN, RAD2DEG, DEG2RAD, SINH, COSH, TANH, ABS, LN, LOG, EXP, ARCSIN, ARCCOS, ARCTAN, ARCSINH, ARCCOSH, and ARCTANH.\n\n"\
                        "Trig functions take their arguments in radians.\n\n"\
                        "The code should set variables named \"Zreal\" and \"Zimag\" to arrays of the values for the real and imaginary parts of the impedance, respectively. Each point in the array should correspond to the frequency at that index.\n\n"\
                        "Most usual Python packages, such as numpy, can be imported. Python built-ins can also be accessed."
        self.helpTextLabel = tk.Label(self.hframe, text=self.helpText, justify=tk.LEFT, bg=self.backgroundColor, fg=self.foregroundColor)
        self.helpTextLabel.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.helpTextLabel.bind("<MouseWheel>", _on_mousewheel_help)
        self.vsbh = tk.Scrollbar(self.helpPopup, orient="vertical", command=self.hcanvas.yview)
        self.vsbh.pack(side=tk.RIGHT, fill=tk.Y)
        self.hcanvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.hcanvas.create_window((4,4), window=self.hframe, anchor="nw")
        
        self.numParams = 0
        self.paramNameEntries = []
        self.paramNameValues = []
        self.paramNameLabels = []
        self.paramComboboxes = []
        self.paramComboboxValues = []
        self.paramValueEntries = []
        self.paramValueValues = []
        self.paramValueLabels = []
        self.paramDeleteButtons = []
        
        def OpenFile():
            name = askopenfilename(initialdir=self.topGUI.getCurrentDirectory(), filetypes = [("Measurement model file", "*.mmfile *.mmcustom"), ("Measurement model data (*.mmfile)", "*.mmfile"), ("Measurement model custom fitting (*.mmcustom)", "*.mmcustom")], title = "Choose a file")
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
                        try:
                            if (self.upDelete == 0):
                                self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                                self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                                self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                            else:
                                self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                        except:
                            messagebox.showwarning("Frequency error", "There are more frequencies set to be deleted than data points. The number of frequencies to delete has been reset to 0.")
                            self.upDelete = 0
                            self.lowDelete = 0
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                            self.wdata = self.wdataRaw.copy()
                            self.rdata = self.rdataRaw.copy()
                            self.jdata = self.jdataRaw.copy()
                        self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                        self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                        #self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                        self.rs.setNumberOfMajorTicks(10)
                        self.rs.showMinorTicks(False)
                        #self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                        self.rs.setLower(np.log10(min(self.wdata)))
                        self.rs.setUpper(np.log10(max(self.wdata)))
                        #self.wdata = self.wdataRaw.copy()
                        #self.rdata = self.rdataRaw.copy()
                        #self.jdata = self.jdataRaw.copy()
                        self.lengthOfData = len(self.wdata)
                        self.freqRangeButton.configure(state="normal")
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, n)
                        self.browseEntry.configure(state="readonly")
                        self.runButton.configure(state="normal")
                    except:
                        messagebox.showerror("File error", "Error: There was an error loading or reading the file")
                elif (fext == ".mmcustom"):
                    try:
                        toLoad = open(n)
                        fileToLoad = toLoad.readline().split("filename: ")[1][:-1]
                        numberOfSimulations = int(toLoad.readline().split("number_simulations: ")[1][:-1])
                        if (numberOfSimulations < 1):
                            raise ValueError
                        fitType = toLoad.readline().split("fit_type: ")[1][:-1]
                        if (fitType != "Real" and fitType != "Complex" and fitType != "Imaginary"):
                            raise ValueError
                        alphaVal = float(toLoad.readline().split("alpha: ")[1][:-1])
                        if (alphaVal <= 0):
                            raise ValueError
                        uD = int(toLoad.readline().split("upDelete: ")[1][:-1])
                        if (uD < 0):
                            raise ValueError
                        lD = int(toLoad.readline().split("lowDelete: ")[1][:-1])
                        if (lD < 0):
                            raise ValueError
                        weightType = toLoad.readline().split("weighting: ")[1][:-1]
                        if (weightType != "None" and weightType != "Modulus" and weightType != "Proportional" and weightType != "Error model" and weightType != "Custom"):
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
                        if (errorBetaRe[0] != "n"):
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
                        toLoad.readline() #Skip the "formula:"
                        formulaIn = ""
                        while True:
                            nextLineIn = toLoad.readline()
                            if "mmparams:" in nextLineIn:
                                break
                            else:
                                formulaIn += nextLineIn
                        while True:
                            p = toLoad.readline()
                            if not p:
                                break
                            else:
                                p = p[:-1]
                                pname, pval, pcombo = p.split("~=~")
                                self.numParams += 1
                                self.paramNameValues.append(tk.StringVar(self, pname))
                                self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10))
                                self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                                self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                                self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                                self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                                self.paramValueValues.append(tk.StringVar(self, pval))
                                self.paramValueEntries.append(ttk.Entry(self.pframe, textvariable=self.paramValueValues[self.numParams-1], width=10))
                                self.paramDeleteButtons.append(ttk.Button(self.pframe, text="Delete", command=partial(deleteParam, self.numParams-1)))
                                self.paramNameLabels[len(self.paramNameLabels)-1].grid(column=0, row=self.numParams, pady=5)
                                self.paramNameEntries[len(self.paramNameEntries)-1].grid(column=1, row=self.numParams)
                                self.paramValueLabels[len(self.paramValueLabels)-1].grid(column=2, row=self.numParams, padx=(10, 3))
                                self.paramComboboxes[len(self.paramComboboxes)-1].grid(column=3, row=self.numParams)
                                self.paramValueEntries[len(self.paramValueEntries)-1].grid(column=4, row=self.numParams, padx=(3,0))
                                self.paramDeleteButtons[len(self.paramDeleteButtons)-1].grid(column=5, row=self.numParams, padx=3)
                                self.howMany.configure(text="Number of parameters: " + str(self.numParams))
                                self.paramNameLabels[len(self.paramNameLabels)-1].bind("<MouseWheel>", _on_mousewheel)
                                self.paramNameEntries[len(self.paramNameEntries)-1].bind("<MouseWheel>", _on_mousewheel)
                                self.paramValueLabels[len(self.paramValueLabels)-1].bind("<MouseWheel>", _on_mousewheel)
                                self.paramValueEntries[len(self.paramValueEntries)-1].bind("<MouseWheel>", _on_mousewheel)
                                self.paramDeleteButtons[len(self.paramDeleteButtons)-1].bind("<MouseWheel>", _on_mousewheel)
                                self.paramNameEntries[len(self.paramNameEntries)-1].bind("<KeyRelease>", keyup)
                        toLoad.close()
                        data = np.loadtxt(fileToLoad)
                        w_in = data[:,0]
                        r_in = data[:,1]
                        j_in = data[:,2]
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
                        self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
                        self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
                        #self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                        self.rs.setNumberOfMajorTicks(10)
                        self.rs.showMinorTicks(False)
                        #self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
                        self.rs.setLower(np.log10(min(self.wdata)))
                        self.rs.setUpper(np.log10(max(self.wdata)))
                        self.lengthOfData = len(self.wdata)
                        self.freqRangeButton.configure(state="normal")
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, n)
                        self.browseEntry.configure(state="readonly")
                        self.runButton.configure(state="normal")
                        self.monteCarloValue.set(str(numberOfSimulations))
                        self.fittingTypeComboboxValue.set(str(fitType))
                        self.weightingComboboxValue.set(weightType)
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
                        if (self.weightingComboboxValue.get() == "None"):
                            self.noiseFrame.grid_remove()
                            self.errorStructureFrame.grid_remove()
                        elif (self.weightingComboboxValue.get() == "Error model"):
                            self.noiseFrame.grid_remove()
                            self.errorStructureFrame.grid(column=0, row=2, pady=(5,0), columnspan=8)
                        elif (self.weightingComboboxValue.get() == "Custom"):
                            self.noiseFrame.grid_remove()
                            self.errorStructureFrame.grid_remove()
                        else:
                            self.noiseFrame.grid(column=3, row=1, pady=(5,0), sticky="W")
                            self.errorStructureFrame.grid_remove()
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, fileToLoad)
                        self.browseEntry.configure(state="readonly")
                        self.customFormula.delete("1.0", tk.END)
                        self.customFormula.insert("1.0", formulaIn)
                        keyup("")
                    except:
                        messagebox.showerror("File error", "Error: There was an error loading or reading the file")
                else:
                    messagebox.showerror("File error", "Error: The file has an unknown extension")
        
        def changeFreqs():            
            if (self.freqWindow.state() == "withdrawn"):
                self.freqWindow.deiconify()
            else:
                self.freqWindow.lift()
            self.lowerSpinboxVariable.set(str(self.lowDelete))
            self.upperSpinboxVariable.set(str(self.upDelete))
            
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
                if (self.upperSpinboxVariable.get() == ""):
                    self.upperSpinboxVariable.set("0")
                if (self.lowerSpinboxVariable.get() == ""):
                    self.lowerSpinboxVariable.set("0")
                #self.rs.setLower(np.log10(min(self.wdata)))
                #self.justUpdated = True
                #self.rs.setUpper(np.log10(max(self.wdata)))
                self.lengthOfData = len(self.wdata)
                self.lowestUndeleted.configure(text="Lowest remaining frequency: {:.4e}".format(min(self.wdata)))
                self.highestUndeleted.configure(text="Highest remaining frequency: {:.4e}".format(max(self.wdata)))
                #self.updateFreqButton.configure(text="Updated")
                #self.after(500, lambda : self.updateFreqButton.configure(text="Update Frequencies"))
            def changeFreqSpinboxLower(e = None):
                try:
                    higherDelete = 0 if self.upperSpinboxVariable.get() == "" else int(self.upperSpinboxVariable.get())
                    lowerDelete = 0 if self.lowerSpinboxVariable.get() == "" else int(self.lowerSpinboxVariable.get())
                    #print(len(self.wdataRaw)-1-higherDelete-lowerDelete)
                    #self.lowerSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete-lowerDelete)
                    self.upperSpinbox.configure(to=len(self.wdataRaw)-1-lowerDelete)
                    if (higherDelete == 0 and lowerDelete == 0):
                        self.rs.setLower(np.log10(min(self.wdataRaw.copy())))
                    elif (higherDelete == 0):
                        self.rs.setLower(np.log10(min(self.wdataRaw.copy()[lowerDelete:])))
                    elif (lowerDelete == 0):
                        self.rs.setLower(np.log10(min(self.wdataRaw.copy()[:-1*higherDelete])))
                    else:
                        self.rs.setLower(np.log10(min(self.wdataRaw.copy()[lowerDelete:-1*higherDelete])))
                    updateFreqs()
                except:
                    pass
            
            def changeFreqSpinboxUpper(e = None):
                try:
                    higherDelete = 0 if self.upperSpinboxVariable.get() == "" else int(self.upperSpinboxVariable.get())
                    lowerDelete = 0 if self.lowerSpinboxVariable.get() == "" else int(self.lowerSpinboxVariable.get())
                    self.lowerSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete)
                    #self.upperSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete-lowerDelete)
                    if (higherDelete == 0 and lowerDelete == 0):
                        self.rs.setUpper(np.log10(max(self.wdataRaw.copy())))
                    elif (higherDelete == 0):
                        self.rs.setUpper(np.log10(max(self.wdataRaw.copy()[lowerDelete:])))
                    elif (lowerDelete == 0):
                        self.rs.setUpper(np.log10(max(self.wdataRaw.copy()[:-1*higherDelete])))
                    else:
                        self.rs.setUpper(np.log10(max(self.wdataRaw.copy()[lowerDelete:-1*higherDelete])))
                    updateFreqs()
                except:
                    pass
            
            self.rs.setPaintTicks(True)
            self.rs.setSnapToTicks(False) 
            self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
            self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
#            self.rs.setMajorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
            self.rs.setNumberOfMajorTicks(10)
            self.rs.showMinorTicks(False)
#            self.rs.setMinorTickSpacing((abs(np.log10(max(self.wdata))) + abs(np.log10(min(self.wdata))))/10)
            self.rs.setLower(np.log10(min(self.wdata)))
            self.rs.setUpper(np.log10(max(self.wdata)))
            #self.rs.setFocus()
            self.lowerLabel = tk.Label(self.freqWindow, text="Number of low frequencies\n to delete", bg=self.backgroundColor, fg=self.foregroundColor)
            self.rangeLabel = tk.Label(self.freqWindow, text="Log of Frequency", bg=self.backgroundColor, fg=self.foregroundColor)
            self.upperLabel = tk.Label(self.freqWindow, text="Number of high frequencies\n to delete", bg=self.backgroundColor, fg=self.foregroundColor)
            self.lowerLabel.grid(column=0, row=1, pady=80, padx=3, sticky="N")
            self.rangeLabel.grid(column=1, row=1, pady=85, sticky="N")
            self.upperLabel.grid(column=2, row=1, pady=80, padx=3, sticky="N")
            self.lowerSpinbox = tk.Spinbox(self.freqWindow, from_=0, to=(len(self.wdataRaw)-1), textvariable=self.lowerSpinboxVariable, state="readonly", width=3, command=changeFreqSpinboxLower)
            self.lowerSpinbox.grid(column=0, row=1, padx=(3,0))
            self.upperSpinbox = tk.Spinbox(self.freqWindow, from_=0, to=(len(self.wdataRaw)-1), textvariable=self.upperSpinboxVariable, state="readonly", width=3, command=changeFreqSpinboxUpper)
            self.upperSpinbox.grid(column=2, row=1, padx=(0,3))
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
        """
        def keyup_return(event):
            textToSearch = self.customFormula.get("1.0", tk.INSERT)
            if (len(textToSearch) > 3):
                if (textToSearch[-2] == ")"):
                    self.customFormula.insert(tk.INSERT + " - 1c", ":")
                    self.customFormula.insert(tk.INSERT, "\t")
        """
        def keyup(event):
            self.customFormula.tag_remove("tagZreal", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords2", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords3", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords4", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords5", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords6", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords7", "1.0", tk.END)
            self.customFormula.tag_remove("tagReservedWords8", "1.0", tk.END)
            self.customFormula.tag_remove("tagComment", "1.0", tk.END)
            self.customFormula.tag_remove("tagUserVariables", "1.0", tk.END)
            self.customFormula.tag_remove("tagBuiltIns", "1.0", tk.END)
            self.customFormula.tag_remove("tagPrebuiltFormulas", "1.0", tk.END)
            self.customFormula.tag_remove("tagQuotations", "1.0", tk.END)
            self.customFormula.tag_remove("tagApostrophes", "1.0", tk.END)
            self.customFormula.tag_remove("tagFreq", "1.0", tk.END)
            self.customFormula.tag_remove("tagZ", "1.0", tk.END)
            self.customFormula.tag_remove("tagWeight", "1.0", tk.END)
            textToSearch = self.customFormula.get("1.0", tk.END)
            if (textToSearch[:-1] == self.defaultFormula):
                self.saveNeeded = False
            else:
                self.saveNeeded = True
            whereStr = [m.start() for m in re.finditer(r'\bZreal\b', textToSearch)]
            whereStr.extend([m.start() for m in re.finditer(r'\bZimag\b', textToSearch)])
            whereFreq = [m.start() for m in re.finditer(r'\bfreq\b', textToSearch)]
            whereZ = [m.start() for m in re.finditer(r'\bZr\b', textToSearch)]
            whereZ.extend([m.start() for m in re.finditer(r'\bZj\b', textToSearch)])
            reservedWords2 = [m.start() for m in re.finditer(r'\bis\b', textToSearch)]
            reservedWords2.extend([m.start() for m in re.finditer(r'\bas\b', textToSearch)])
            reservedWords2.extend([m.start() for m in re.finditer(r'\bif\b', textToSearch)])
            reservedWords2.extend([m.start() for m in re.finditer(r'\bor\b', textToSearch)])
            reservedWords2.extend([m.start() for m in re.finditer(r'\bin\b', textToSearch)])
            reservedWords3 = [m.start() for m in re.finditer(r'\bfor\b', textToSearch)]
            reservedWords3.extend([m.start() for m in re.finditer(r'\btry\b', textToSearch)])
            reservedWords3.extend([m.start() for m in re.finditer(r'\bdef\b', textToSearch)])
            reservedWords3.extend([m.start() for m in re.finditer(r'\band\b', textToSearch)])
            reservedWords3.extend([m.start() for m in re.finditer(r'\bdel\b', textToSearch)])
            reservedWords3.extend([m.start() for m in re.finditer(r'\bnot\b', textToSearch)])
            reservedWords4 = [m.start() for m in re.finditer(r'\bNone\b', textToSearch)]
            reservedWords4.extend([m.start() for m in re.finditer(r'\bTrue\b', textToSearch)])
            reservedWords4.extend([m.start() for m in re.finditer(r'\bfrom\b', textToSearch)])
            reservedWords4.extend([m.start() for m in re.finditer(r'\bwith\b', textToSearch)])
            reservedWords4.extend([m.start() for m in re.finditer(r'\belif\b', textToSearch)])
            reservedWords4.extend([m.start() for m in re.finditer(r'\belse\b', textToSearch)])
            reservedWords4.extend([m.start() for m in re.finditer(r'\bpass\b', textToSearch)])
            reservedWords5 = [m.start() for m in re.finditer(r'\bFalse\b', textToSearch)]
            reservedWords5.extend([m.start() for m in re.finditer(r'\bclass\b', textToSearch)])
            reservedWords5.extend([m.start() for m in re.finditer(r'\bwhile\b', textToSearch)])
            reservedWords5.extend([m.start() for m in re.finditer(r'\byield\b', textToSearch)])
            reservedWords5.extend([m.start() for m in re.finditer(r'\bbreak\b', textToSearch)])
            reservedWords5.extend([m.start() for m in re.finditer(r'\braise\b', textToSearch)])
            reservedWords6 = [m.start() for m in re.finditer(r'\breturn\b', textToSearch)]
            reservedWords6.extend([m.start() for m in re.finditer(r'\blambda\b', textToSearch)])
            reservedWords6.extend([m.start() for m in re.finditer(r'\bglobal\b', textToSearch)])
            reservedWords6.extend([m.start() for m in re.finditer(r'\bassert\b', textToSearch)])
            reservedWords6.extend([m.start() for m in re.finditer(r'\bimport\b', textToSearch)])
            reservedWords6.extend([m.start() for m in re.finditer(r'\bexcept\b', textToSearch)])
            reservedWords7 = [m.start() for m in re.finditer(r'\bfinally\b', textToSearch)]
            reservedWords8 = [m.start() for m in re.finditer(r'\bcontinue\b', textToSearch)]
            reservedWords8.extend([m.start() for m in re.finditer(r'\bnonlocal\b', textToSearch)])
            whereWeight = [m.start() for m in re.finditer(r'\bweighting\b', textToSearch)]
            whereComment = []
            userVariablesBegin = []
            userVariablesEnd = []
            for i in range(len(self.paramNameValues)):
                tempList = [m.start() for m in re.finditer(r'\b' + self.paramNameValues[i].get() + r'\b', textToSearch)]
                userVariablesBegin.extend(tempList)
                if (len(tempList) > 0):
                    for j in range(len(tempList)):
                        userVariablesEnd.append(tempList[j] + len(self.paramNameValues[i].get()))
            
            builtIns = ['id', 'abs', 'set', 'all', 'min', 'any', 'dir', 'hex', 'bin', 'oct', 'int', 'str', 'ord', 'sum', 'pow', 'len', 'chr', 'zip', 'map', 'max', 'hash', 'dict', 'help', 'next', 'bool', 'eval', 'open', 'exec', 'iter', 'type', 'list', 'vars', 'repr'\
                        'slice', 'ascii', 'input', 'super', 'bytes', 'float', 'print', 'tuple', 'range', 'round', 'divmod', 'object', 'sorted', 'filter', 'format', 'locals', 'delattr', 'setattr', 'getattr', 'compile', 'globals', 'complex', 'hasattr', 'callable'\
                        'property', 'reversed', 'enumerate', 'bytearray', 'frozenset', 'memoryview', 'breakpoint', 'isinstance', 'issubclass', '__import__', 'classmethod', 'staticmethod']
            builtInsBegin = []
            builtInsEnd = []
            for i in range(len(builtIns)):
                tempList = [m.start() for m in re.finditer(r'\b' + builtIns[i] + r'\(', textToSearch)]
                builtInsBegin.extend(tempList)
                if (len(tempList) > 0):
                    for j in range(len(tempList)):
                        builtInsEnd.append(tempList[j] + len(builtIns[i]))
            
            prebuiltFormulas = ['1j', 'PI', 'SIN', 'COS', 'TAN', 'COSH', 'SINH', 'TANH', 'ARCSIN', 'ARCCOS', 'ARCTAN', 'ARCSINH', 'ARCCOSH', 'ARCTANH', 'SQRT', 'LN', 'LOG', 'EXP', 'ABS', 'DEG2RAD', 'RAD2DEG']
            prebuiltFormulasBegin = []
            prebuiltFormulasEnd = []
            for i in range(len(prebuiltFormulas)):
                tempList = [m.start() for m in re.finditer(r'\b' + prebuiltFormulas[i] + r'\b', textToSearch)]
                prebuiltFormulasBegin.extend(tempList)
                if (len(tempList) > 0):
                    for j in range(len(tempList)):
                        prebuiltFormulasEnd.append(tempList[j] + len(prebuiltFormulas[i]))
                        
            quotationBegin = []
            quotationEnd = []
            apostropheBegin = []
            apostropheEnd = []
            isComment = False
            isQuote = False
            isApostrophe = False
            for i in range(len(textToSearch)):
                if (textToSearch[i] == "#" and not isQuote and not isApostrophe):
                    whereComment.append(i)
                    isComment = True
                elif (textToSearch[i] == "\n"):
                    isComment = False
                elif (not isComment and not isApostrophe and textToSearch[i] == "\""):
                    if (len(quotationBegin) > len(quotationEnd)):
                        quotationEnd.append(i)
                        isQuote = False
                    else:
                        quotationBegin.append(i)
                        isQuote = True
                elif (not isComment and not isQuote and textToSearch[i] == "\'"):
                    if (len(apostropheBegin) > len(apostropheEnd)):
                        apostropheEnd.append(i)
                        isApostrophe = False
                    else:
                        apostropheBegin.append(i)
                        isApostrophe = True

            for i in range(len(whereStr)):
                self.customFormula.tag_add("tagZreal", "1.0 + " + str(whereStr[i]) + "c", "1.0 + " + str(whereStr[i] + 5) + "c")
                self.customFormula.tag_config("tagZreal", foreground="red")
            for i in range(len(whereFreq)):
                self.customFormula.tag_add("tagFreq", "1.0 + " + str(whereFreq[i]) + "c", "1.0 + " + str(whereFreq[i] + 4) + "c")
                self.customFormula.tag_config("tagFreq", foreground="red")
            for i in range(len(whereZ)):
                self.customFormula.tag_add("tagZ", "1.0 + " + str(whereZ[i]) + "c", "1.0 + " + str(whereZ[i] + 2) + "c")
                self.customFormula.tag_config("tagZ", foreground="red")
            if (self.weightingComboboxValue.get() == "Custom"):
                for i in range(len(whereWeight)):
                    self.customFormula.tag_add("tagWeight", "1.0 + " + str(whereWeight[i]) + "c", "1.0 + " + str(whereWeight[i] + 9) + "c")
                    self.customFormula.tag_config("tagWeight", foreground="red")
            for i in range(len(reservedWords2)):
                self.customFormula.tag_add("tagReservedWords2", "1.0 + " + str(reservedWords2[i]) + "c", "1.0 + " + str(reservedWords2[i] + 2) + "c")
                self.customFormula.tag_config("tagReservedWords2", foreground="DodgerBlue3")
            for i in range(len(reservedWords3)):
                self.customFormula.tag_add("tagReservedWords3", "1.0 + " + str(reservedWords3[i]) + "c", "1.0 + " + str(reservedWords3[i] + 3) + "c")
                self.customFormula.tag_config("tagReservedWords3", foreground="DodgerBlue3")
            for i in range(len(reservedWords4)):
                self.customFormula.tag_add("tagReservedWords4", "1.0 + " + str(reservedWords4[i]) + "c", "1.0 + " + str(reservedWords4[i] + 4) + "c")
                self.customFormula.tag_config("tagReservedWords4", foreground="DodgerBlue3")
            for i in range(len(reservedWords5)):
                self.customFormula.tag_add("tagReservedWords5", "1.0 + " + str(reservedWords5[i]) + "c", "1.0 + " + str(reservedWords5[i] + 5) + "c")
                self.customFormula.tag_config("tagReservedWords5", foreground="DodgerBlue3")
            for i in range(len(reservedWords6)):
                self.customFormula.tag_add("tagReservedWords6", "1.0 + " + str(reservedWords6[i]) + "c", "1.0 + " + str(reservedWords6[i] + 6) + "c")
                self.customFormula.tag_config("tagReservedWords6", foreground="DodgerBlue3")
            for i in range(len(reservedWords7)):
                self.customFormula.tag_add("tagReservedWords7", "1.0 + " + str(reservedWords7[i]) + "c", "1.0 + " + str(reservedWords7[i] + 7) + "c")
                self.customFormula.tag_config("tagReservedWords7", foreground="DodgerBlue3")
            for i in range(len(reservedWords8)):
                self.customFormula.tag_add("tagReservedWords8", "1.0 + " + str(reservedWords8[i]) + "c", "1.0 + " + str(reservedWords8[i] + 8) + "c")
                self.customFormula.tag_config("tagReservedWords8", foreground="DodgerBlue3")
            for i in range(len(builtInsBegin)):
                self.customFormula.tag_add("tagBuiltIns", "1.0 + " + str(builtInsBegin[i]) + "c", "1.0 + " + str(builtInsEnd[i]) + "c")
                self.customFormula.tag_config("tagBuiltIns", foreground="violet red")
            for i in range(len(prebuiltFormulasBegin)):
                self.customFormula.tag_add("tagPrebuiltFormulas", "1.0 + " + str(prebuiltFormulasBegin[i]) + "c", "1.0 + " + str(prebuiltFormulasEnd[i]) + "c")
                self.customFormula.tag_config("tagPrebuiltFormulas", foreground="DarkOrchid2")
            for i in range(len(userVariablesBegin)):
                self.customFormula.tag_add("tagUserVariables", "1.0 + " + str(userVariablesBegin[i]) + "c", "1.0 + " + str(userVariablesEnd[i]) + "c")
                self.customFormula.tag_config("tagUserVariables", foreground="red")
            for i in range(len(whereComment)):
                self.customFormula.tag_remove("tagZreal", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords2", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords3", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords4", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords5", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords6", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords7", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagReservedWords8", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagUserVariables", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagBuiltIns", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagPrebuiltFormulas", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagQuotations", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagApostrophes", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagFreq", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagZ", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_remove("tagWeight", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_add("tagComment", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
                self.customFormula.tag_config("tagComment", foreground="gray60")
            for i in range(len(quotationBegin)):
                if (len(quotationBegin) > len(quotationEnd) and i == len(quotationBegin)-1):
                    self.customFormula.tag_add("tagQuotations", "1.0 + " + str(quotationBegin[i]) + "c", "1.0 + " + str(len(textToSearch)) + "c")
                else:
                    self.customFormula.tag_add("tagQuotations", "1.0 + " + str(quotationBegin[i]) + "c", "1.0 + " + str(quotationEnd[i]+1) + "c")
                self.customFormula.tag_config("tagQuotations", foreground="green")
            for i in range(len(apostropheBegin)):
                if (len(apostropheBegin) > len(apostropheEnd) and i == len(apostropheBegin)-1):
                    self.customFormula.tag_add("tagApostrophes", "1.0 + " + str(apostropheBegin[i]) + "c", "1.0 + " + str(len(textToSearch)) + "c")
                else:
                    self.customFormula.tag_add("tagApostrophes", "1.0 + " + str(apostropheBegin[i]) + "c", "1.0 + " + str(apostropheEnd[i]+1) + "c")
                self.customFormula.tag_config("tagApostrophes", foreground="green")
        
        def deleteParam(paramNumber):
            if paramNumber == 0 and self.numParams == 1:
                removeParam()
            else:
                for i in range(paramNumber, self.numParams-1):
                    self.paramNameValues[i].set(self.paramNameValues[i+1].get())
                    self.paramComboboxValues[i].set(self.paramComboboxValues[i+1].get())
                    self.paramValueValues[i].set(self.paramValueValues[i+1].get())
                removeParam()
        
        def addParam():
            self.numParams += 1
            self.paramNameValues.append(tk.StringVar(self, "var" + str(self.varNum)))
            self.varNum += 1
            self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10))
            self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
            self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
            self.paramComboboxValues.append(tk.StringVar(self, "+ or -"))
            self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
            self.paramValueValues.append(tk.StringVar(self, "0"))
            self.paramValueEntries.append(ttk.Entry(self.pframe, textvariable=self.paramValueValues[self.numParams-1], width=10))
            self.paramDeleteButtons.append(ttk.Button(self.pframe, text="Delete", command=partial(deleteParam, self.numParams-1)))
            self.paramNameLabels[len(self.paramNameLabels)-1].grid(column=0, row=self.numParams, pady=5)
            self.paramNameEntries[len(self.paramNameEntries)-1].grid(column=1, row=self.numParams)
            self.paramValueLabels[len(self.paramValueLabels)-1].grid(column=2, row=self.numParams, padx=(10, 3))
            self.paramComboboxes[len(self.paramComboboxes)-1].grid(column=3, row=self.numParams)
            self.paramValueEntries[len(self.paramValueEntries)-1].grid(column=4, row=self.numParams, padx=(3,0))
            self.paramDeleteButtons[len(self.paramDeleteButtons)-1].grid(column=5, row=self.numParams, padx=3)
            self.howMany.configure(text="Number of parameters: " + str(self.numParams))
            self.paramNameLabels[len(self.paramNameLabels)-1].bind("<MouseWheel>", _on_mousewheel)
            self.paramNameEntries[len(self.paramNameEntries)-1].bind("<MouseWheel>", _on_mousewheel)
            self.paramValueLabels[len(self.paramValueLabels)-1].bind("<MouseWheel>", _on_mousewheel)
            self.paramValueEntries[len(self.paramValueEntries)-1].bind("<MouseWheel>", _on_mousewheel)
            self.paramDeleteButtons[len(self.paramDeleteButtons)-1].bind("<MouseWheel>", _on_mousewheel)
            self.paramNameEntries[len(self.paramNameEntries)-1].bind("<KeyRelease>", keyup)
            keyup("")
        
        def removeParam():
            if (self.numParams != 0):
                self.paramNameValues.pop()
                self.paramComboboxValues.pop()
                self.paramValueValues.pop()
                self.paramNameEntries.pop().grid_remove()
                self.paramComboboxes.pop().grid_remove()
                self.paramValueEntries.pop().grid_remove()
                self.paramNameLabels.pop().grid_remove()
                self.paramValueLabels.pop().grid_remove()
                self.paramDeleteButtons.pop().grid_remove()
                self.numParams -= 1
                self.howMany.configure(text="Number of parameters: " + str(self.numParams))
                keyup("")

        def _on_mousewheel(event):
            xpos, ypos = self.vsb.get()
            xpos_round = round(xpos, 2)     #Avoid floating point errors
            ypos_round = round(ypos, 2)
            if (xpos_round != 0.00 or ypos_round != 1.00):
                self.pcanvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def parameterPopup():
            self.pcanvas.configure(yscrollcommand=self.vsb.set)
            self.pcanvas.bind("<MouseWheel>", _on_mousewheel)
            self.pframe.bind("<MouseWheel>", _on_mousewheel)
            def onFrameConfigure(canvas):
                #Reset the scroll region to encompass the inner frame
                canvas.configure(scrollregion=canvas.bbox("all"))
            
            def onClose():
                self.paramPopup.withdraw()
                #addButton.destroy()
                #removeButton.destroy()
                #buttonFrame.destroy()
                #self.howMany.destroy()
            
            self.addButton.configure(command=addParam)
            self.removeButton.configure(command=removeParam)
            self.paramPopup.title("Custom fitting parameters")
            self.paramPopup.iconbitmap(resource_path("img/elephant3.ico"))
            self.paramPopup.geometry("500x400")
            self.paramPopup.deiconify()
            self.paramPopup.protocol("WM_DELETE_WINDOW", onClose)
            
            self.pframe.bind("<Configure>", lambda event, canvas=self.pcanvas: onFrameConfigure(canvas))
        
        def helpPopup(event):
            self.hcanvas.configure(yscrollcommand=self.vsbh.set)
            self.hcanvas.bind("<MouseWheel>", _on_mousewheel)
            self.hframe.bind("<MouseWheel>", _on_mousewheel)
            def onHelpFrameConfigure(canvas):
                #Reset the scroll region to encompass the inner frame
                canvas.configure(scrollregion=canvas.bbox("all"))
            def onClose():
                self.helpPopup.withdraw()
            def helpConfigure(configEvent):
                self.helpTextLabel.configure(wraplength=configEvent.width)
            self.helpPopup.title("Custom formula help")
            self.helpPopup.iconbitmap(resource_path("img/elephant3.ico"))
            self.helpPopup.geometry("500x400")
            self.helpPopup.deiconify()
            self.helpPopup.protocol("WM_DELETE_WINDOW", onClose)
            
            self.hframe.bind("<Configure>", lambda event, canvas=self.hcanvas: onHelpFrameConfigure(canvas))
            self.helpPopup.bind("<Configure>", helpConfigure)
        
        def process_queue_custom():
            try:
                r,s,sdR,sdI,chi,aic,realF,imagF = self.queue.get(0)
                self.runButton.configure(state="normal")
                self.browseButton.configure(state="normal")
                self.customFormula.configure(state="normal")
                self.freqRangeButton.configure(state="normal")
                self.parametersButton.configure(state="normal")
                self.fittingTypeCombobox.configure(state="readonly")
                self.monteCarloEntry.configure(state="normal")
                self.saveFormulaButton.configure(state="normal")
                self.weightingCombobox.configure(state="readonly")
                self.noiseEntry.configure(state="normal")
                self.cancelButton.grid_remove()
                self.prog_bar.stop()
                self.prog_bar.destroy()
                if (len(r) == 1):
                    if (r == "b"):
                        messagebox.showerror("Fitting error", "There was an error in the fitting.")
                    elif (r == "^"):
                        messagebox.showerror("Fitting error", "The fitting failed.")
                    elif (r == "@"):
                        pass    #The fitting was cancelled
                else:
                    if len(s) == 1:
                        if (s == "-"):
                            messagebox.showwarning("Variance error", "A minimization was found, but the standard deviations and confidence intervals could not be calculated.")
                    for i in range(len(r)):
                        self.paramValueValues[i].set("%.5g"%r[i])
                    self.aR = ""
                    self.aR += "File name: " + self.browseEntry.get() + "\n"
                    self.aR += "Number of data: " + str(self.lengthOfData) + "\n"
                    self.aR += "Number of parameters: " + str(len(r)) + "\n\n"
                    for i in range(len(r)):
                        self.aR += self.paramNameValues[i].get() + " = " + str(float(r[i]))
                        if (len(s) == 1):
                            self.aR += "\n"
                        else:
                            self.aR += " ± " + str(float(s[i])) + "\n"
                    self.aR += "\nChi-squared statistic = %.4g"%chi + "\n"
                    self.aR += "Chi-squared/Degrees of freedom = %.4g"%(chi/(self.lengthOfData*2-len(r))) + "\n"
                    try:
                        self.aR += "Akaike Information Criterion = %.4g"%aic
                    except:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    self.resultsView.delete(*self.resultsView.get_children())
                    self.resultAlert.grid_remove()
                    if (len(s) > 1):
                        for i in range(len(r)):
                            if (abs(s[i]*2*100/r[i]) > 100 or np.isnan(s[i])):
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=2, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"))
                    else:
                        for i in range(len(r)):
                            self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "nan", "nan"))
                            self.resultAlert.grid(column=0, row=2, sticky="E")
                    self.resultsView.tag_configure("bad", background="yellow")
                    self.resultsFrame.grid(column=0, row=5, sticky="W", pady=5)
                    self.graphFrame.grid(column=0, row=6, sticky="W", pady=5)
                    self.saveFrame.grid(column=0, row=7, sticky="W", pady=5)
                    self.realFit = realF
                    self.imagFit = imagF
                    self.sdrReal = sdR
                    self.sdrImag = sdI
                    self.fits = r
                    self.sigmas = s
                    self.chiSquared = chi
            except:
                self.after(100, process_queue_custom)
        
        def checkFitting():
            
            #---Check if a formula has been entered----
            formula = self.customFormula.get("1.0", tk.END)
            if ("".join(formula.split()) == ""):
                messagebox.showwarning("Error", "Formula is empty")
                return
            
            #---Check that none of the variable names are repeated, and that none are Python reserved words or variables used in this code
            for i in range(len(self.paramNameValues)):
                name = self.paramNameValues[i].get()
                if (name == "False" or name == "None" or name == "True" or name == "and" or name == "as" or name == "assert" or name == "break" or name == "class" or name == "continue" or name == "def" or name == "del" or name == "elif" or name == "else" or name == "except"\
                     or name == "finally" or name == "for" or name == "from" or name == "global" or name == "if" or name == "import" or name == "in" or name == "is" or name == "lambda" or name == "nonlocal" or name == "not" or name == "or" or name == "pass" or name == "raise"\
                      or name == "return" or name == "try" or name == "while" or name == "with" or name == "yield"):
                    messagebox.showwarning("Error", "The variable name \"" + name + "\" is a Python reserved word. Change the variable name.")
                    return
                elif ("self." in name):
                    messagebox.showwarning("Error", "The variable name \"" + name + "\" contains \"self.\"; change the variable name.")
                    return
                elif (name == "freq" or name == "Zr" or name == "Zj" or name == "Zreal" or name == "Zimag" or name == "weighting"):
                    messagebox.showwarning("Error", "The variable name \"" + name + "\" is used by the fitting program; change the variable name.")
                    return
                for j in range(i+1, len(self.paramNameValues)):
                    if (name == self.paramNameValues[j].get()):
                        messagebox.showwarning("Error", "Two or more variables have the same name.")
                        return
            
            #---Replace the functions with np.<function>, and then attempt to compile the code to look for syntax errors---        
            try:
                prebuiltFormulas = ['PI', 'SIN', 'COS', 'TAN', 'COSH', 'SINH', 'TANH', 'ARCSIN', 'ARCCOS', 'ARCTAN', 'ARCSINH', 'ARCCOSH', 'ARCTANH', 'SQRT', 'LN', 'LOG', 'EXP', 'ABS', 'DEG2RAD', 'RAD2DEG']
                formula = formula.replace("^", "**")    #Replace ^ with ** for exponentiation (this could prevent some features like regex from being used effectively)
                for pf in prebuiltFormulas:
                    toReplace = "np." + pf.lower()
                    formula = formula.replace(pf, toReplace)
                compile(formula, 'user_generated_formula', 'exec')
            except:
                messagebox.showwarning("Compile error", "There was an issue compiling the code")
                return
            
            #---Check if the variable "freq" appears in the code, as it's likely a mistake if it doesn't---
            textToSearch = self.customFormula.get("1.0", tk.END)
            whereFreq = [m.start() for m in re.finditer(r'\bfreq\b', textToSearch)]
            if (len(whereFreq) == 0):
                messagebox.showwarning("No \"freq\"", "The variable \"freq\" does not appear in the code. This may cause an error.")
            elif (len(whereFreq) == 1 and whereFreq[0] == 5):
                messagebox.showwarning("No \"freq\"", "The variable \"freq\" does not seem to be in the code. This may cause an error.")
            
            self.queue = queue.Queue()
            fit_type = 3
            if (self.fittingTypeComboboxValue.get() == "Real"):
                fit_type = 1
            elif (self.fittingTypeCombobox.get() == "Imaginary"):
                fit_type = 2
            num_monte = 1000
            try:
                num_monte = int(self.monteCarloValue.get())
            except:
                messagebox.showwarning("Bad number of simulations", "The number of simulations must be an integer")
                return
            param_names = []
            for pNV in self.paramNameValues:
                param_names.append(pNV.get())
            param_guesses = []
            for pNG in self.paramValueValues:
                try:
                    param_guesses.append(float(pNG.get()))
                except:
                    messagebox.showwarning("Bad parameter value", "The parameter values must be real numbes")
                    return
            param_limits = []
            for pL in self.paramComboboxValues:
                if (pL.get() == "+"):
                    param_limits.append("+")
                elif (pL.get() == "-"):
                    param_limits.append("-")
                elif (pL.get() == "fixed"):
                    param_limits.append("f")
                else:
                    param_limits.append("n")
            weight = 0
            errorModelParams = np.zeros(5)
            if (self.weightingComboboxValue.get() == "Modulus"):
                weight = 1
            elif (self.weightingComboboxValue.get() == "Proportional"):
                weight = 2
            elif (self.weightingComboboxValue.get() == "Error model"):
                if (self.errorAlphaCheckboxVariable.get() != 1 and self.errorBetaCheckboxVariable.get() != 1 and self.errorGammaCheckboxVariable.get() != 1 and self.errorDeltaCheckboxVariable.get() != 1):
                    messagebox.showwarning("Bad error structure", "At least one error structure value must be checkd")
                    return
                try:
                    if (self.errorAlphaCheckboxVariable.get() == 1):
                        if (self.errorAlphaVariable.get() == "" or self.errorAlphaVariable.get() == " " or self.errorAlphaVariable.get() == "."):
                            errorModelParams[0] = 0
                        else:
                            errorModelParams[0] = float(self.errorAlphaVariable.get())
                    if (self.errorBetaCheckboxVariable.get() == 1):
                        if (self.errorBetaVariable.get() == "" or self.errorBetaVariable.get() == " " or self.errorBetaVariable.get() == "."):
                            errorModelParams[1] = 0
                        else:
                            errorModelParams[1] = float(self.errorBetaVariable.get())
                    if (self.errorBetaReCheckboxVariable.get() == 1):
                        if (self.errorBetaReVariable.get() == "" or self.errorBetaReVariable.get() == " " or self.errorBetaReVariable.get() == "."):
                            errorModelParams[2] = 0
                        else:
                            errorModelParams[2] = float(self.errorBetaReVariable.get())
                    if (self.errorGammaCheckboxVariable.get() == 1):
                        if (self.errorGammaVariable.get() == "" or self.errorGammaVariable.get() == " " or self.errorGammaVariable.get() == "."):
                            errorModelParams[3] = 0
                        else:
                            errorModelParams[3] = float(self.errorGammaVariable.get())
                    if (self.errorDeltaCheckboxVariable.get() == 1):
                        if (self.errorDeltaVariable.get() == "" or self.errorDeltaVariable.get() == " " or self.errorDeltaVariable.get() == "."):
                            errorModelParams[4] = 0
                        else:
                            errorModelParams[4] = float(self.errorDeltaVariable.get())
                except:
                    messagebox.showerror("Value error", "One of the error structure parameters has an invalid value")
                    return
                if (errorModelParams[0] == 0 and errorModelParams[1] == 0 and errorModelParams[3] == 0 and errorModelParams[4] == 0):
                    messagebox.showerror("Value error", "At least one of the error structure parameters must be nonzero")
                    return
                weight = 3
            elif (self.weightingComboboxValue.get() == "Custom"):
                weight = 4
            assumed_noise = float(self.noiseEntryValue.get())
            self.runButton.configure(state="disabled")
            self.browseButton.configure(state="disabled")
            self.customFormula.configure(state="disabled")
            self.freqRangeButton.configure(state="disabled")
            self.parametersButton.configure(state="disabled")
            self.fittingTypeCombobox.configure(state="disabled")
            self.monteCarloEntry.configure(state="disabled")
            self.saveFormulaButton.configure(state="disabled")
            self.weightingCombobox.configure(state="disabled")
            self.noiseEntry.configure(state="disabled")
            self.prog_bar = ttk.Progressbar(self.runFrame, orient="horizontal", length=150, mode="indeterminate")
            self.prog_bar.grid(column=3, row=0, padx=5)
            self.prog_bar.start(40)
            
            self.currentThreads.append(ThreadedTaskCustom(self.queue, fit_type, num_monte, weight, assumed_noise, formula, self.wdata, self.jdata, self.rdata, param_names, param_guesses, param_limits, errorModelParams))
            self.currentThreads[len(self.currentThreads) - 1].start()
            self.cancelButton.configure(command=lambda: self.currentThreads[len(self.currentThreads) - 1].terminate())
            self.cancelButton.grid(column=2, row=0, sticky="W", padx=15)
            self.after(100, process_queue_custom)
        
        def popup_formula(event):
            try:
                self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu.grab_release()
        
        def copyFormula():
            try:
                pyperclip.copy(self.customFormula.get(tk.SEL_FIRST, tk.SEL_LAST))
            except:
                pass
        
        def cutFormula():
            try:
                pyperclip.copy(self.customFormula.get(tk.SEL_FIRST, tk.SEL_LAST))
                self.customFormula.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except:
                pass
        
        def pasteFormula():
            try:
                toPaste = pyperclip.paste()
                self.customFormula.insert(tk.INSERT, toPaste)
                keyup(None)
            except:
                pass
            
        def undoFormula():
            try:
                self.customFormula.edit_undo()
            except:
                pass
            
        def redoFormula():
            try:
                self.customFormula.edit_redo()
            except:
                pass
        
        def handle_click(event):
            if self.resultsView.identify_region(event.x, event.y) == "separator":
                if (self.resultsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def handle_motion(event):
            if self.resultsView.identify_region(event.x, event.y) == "separator":
                if (self.resultsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def copyVals():
            self.copyButton.configure(text="Copied")
            self.clipboard_clear()
            stringToCopy = "Param\tValue\tStd. Dev.\n"
            for child in self.resultsView.get_children():
                resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                stringToCopy += resultName + "\t" + resultValue + "\t" + resultStdDev + "\n"
            pyperclip.copy(stringToCopy[:-1])
            self.after(500, lambda : self.copyButton.configure(text="Copy values and std. devs. as spreadsheet"))
        
        def advancedResults():
            self.resultsWindow = tk.Toplevel()
            self.resultsWindow.title("Advanced results")
            self.resultsWindow.iconbitmap(resource_path("img/elephant3.ico"))
            self.resultsWindow.geometry("1000x550")
            self.advResults = scrolledtext.ScrolledText(self.resultsWindow, width=50, height=10, bg=self.backgroundColor, fg=self.foregroundColor, state="normal")
            self.advResults.insert(tk.INSERT, self.aR)
            self.advResults.configure(state="disabled")
            self.advResults.pack(fill=tk.BOTH, expand=True)
        
        def onclick(event):
            try:
                if (not event.inaxes):      #If a subplot isn't clicked
                    return
                resultPlotBig = tk.Toplevel()
                resultPlotBig.iconbitmap(resource_path('img/elephant3.ico'))
                w, h = self.winfo_screenwidth(), self.winfo_screenheight()
                resultPlotBig.geometry("%dx%d+0+0" % (w/2, h/2))
                with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                    fig = Figure()
                    fig.set_facecolor(self.backgroundColor)
                    RealFit = np.array(self.realFit)
                    ImagFit = np.array(self.imagFit)
                    phase_fit = np.arctan2(ImagFit, RealFit) * (180/np.pi)
                    larger = fig.add_subplot(111)
                    larger.set_facecolor(self.backgroundColor)
                    larger.yaxis.set_ticks_position("both")
                    larger.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                    larger.xaxis.set_ticks_position("both")
                    larger.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
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
                        larger.plot(self.rdata, -1*self.jdata, "o", color=dataColor)
                        larger.plot(RealFit, -1*ImagFit, color=fitColor, marker="x", markeredgecolor="black")
                        if (self.confInt):
                            for i in range(len(self.wdata)):
                                ellipse = patches.Ellipse(xy=(RealFit[i], -1*ImagFit[i]), width=4*self.sdrReal[i], height=4*self.sdrImag[i])
                                larger.add_artist(ellipse)
                                ellipse.set_alpha(0.1)
                                ellipse.set_facecolor(self.ellipseColor)
                        larger.axis("equal")
                        larger.set_title("Nyquist Plot", color=self.foregroundColor)
                        larger.set_xlabel("Zr / Ω", color=self.foregroundColor)
                        larger.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    elif (event.inaxes == self.bplot):
                        resultPlotBig.title("Real Impedance Plot")
                        larger.plot(self.wdata, self.rdata, "o", color=dataColor)
                        larger.plot(self.wdata, RealFit, color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(RealFit[i] + 2*self.sdrReal[i], RealFit[i] - 2*self.sdrReal[i])
                                error_below[i] = min(RealFit[i] + 2*self.sdrReal[i], RealFit[i] - 2*self.sdrReal[i])
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        larger.set_xscale("log")
                        larger.set_title("Real Impedance Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Zr / Ω", color=self.foregroundColor)
                    elif (event.inaxes == self.cplot):
                        resultPlotBig.title("Imaginary Impedance Plot")
                        larger.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                        larger.plot(self.wdata, -1*ImagFit, color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(ImagFit[i] + 2*self.sdrImag[i], ImagFit[i] - 2*self.sdrImag[i])
                                error_below[i] = min(ImagFit[i] + 2*self.sdrImag[i], ImagFit[i] - 2*self.sdrImag[i])
                            larger.plot(self.wdata, -1*error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, -1*error_below, "--", color=self.ellipseColor)
                        larger.set_xscale("log")
                        larger.set_title("Imaginary Impedance Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    elif (event.inaxes == self.dplot):
                        resultPlotBig.title("|Z| Bode Plot")
                        larger.plot(self.wdata, np.sqrt(self.jdata**2 + self.rdata**2), "o", color=dataColor)
                        larger.plot(self.wdata, np.sqrt(ImagFit**2 + RealFit**2), color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = np.sqrt(max((RealFit[i]+2*self.sdrReal[i])**2, (RealFit[i]-2*self.sdrReal[i])**2) + max((ImagFit[i]+2*self.sdrImag[i])**2, (ImagFit[i]-2*self.sdrImag[i])**2))
                                error_below[i] = np.sqrt(min((RealFit[i]+2*self.sdrReal[i])**2, (RealFit[i]-2*self.sdrReal[i])**2) + min((ImagFit[i]+2*self.sdrImag[i])**2, (ImagFit[i]-2*self.sdrImag[i])**2))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        larger.set_xscale("log")
                        larger.set_yscale("log")
                        larger.set_title("|Z| Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                    elif (event.inaxes == self.eplot):
                        phase_fit = np.arctan2(ImagFit, RealFit) * (180/np.pi)
                        actual_phase = np.arctan2(self.jdata, self.rdata) * (180/np.pi)
                        resultPlotBig.title("Phase Angle Bode Plot")
                        larger.plot(self.wdata, actual_phase, "o", color=dataColor)
                        larger.plot(self.wdata, phase_fit, color=fitColor)
                        if (self.confInt):
                            error_above = np.arctan2((ImagFit+2*self.sdrImag), (RealFit+2*self.sdrReal)) * (180/np.pi)
                            error_below = np.arctan2((ImagFit-2*self.sdrImag), (RealFit-2*self.sdrReal)) * (180/np.pi)
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
    #                    larger.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                        larger.set_ylim(bottom=-90, top=0)
                        larger.set_xscale("log")
                        larger.set_title("Phase Angle Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Phase Angle / Degrees", color=self.foregroundColor)
                    elif (event.inaxes == self.hplot):
                        resultPlotBig.title("Log(Zr) vs Log|ω|")
                        larger.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                        larger.plot(self.wdata, np.log10(abs(RealFit)), color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i]))) #np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                                error_below[i] = min(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i])))#np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
    #                            error_above[i] = np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
    #                            error_below[i] = np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        larger.set_xscale("log")
                        larger.set_title("Log|Zr|", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Log|Zr|", color=self.foregroundColor)
                    elif (event.inaxes == self.iplot):
                         resultPlotBig.title("Log|Zj| vs Log|ω|")
                         larger.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                         larger.plot(self.wdata, np.log10(abs(ImagFit)), color=fitColor)
                         if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))#np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                                error_below[i] = min(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))#np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
    #                            error_above[i] = np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
    #                            error_below[i] = np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                         larger.set_xscale("log")
                         larger.set_title("Log|Zj|", color=self.foregroundColor)
                         larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                         larger.set_ylabel("Log|Zj|", color=self.foregroundColor)
                    elif (event.inaxes == self.kplot):
                        normalized_residuals_real = np.zeros(len(self.wdata))
                        normalized_error_real_below = np.zeros(len(self.wdata))
                        normalized_error_real_above = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            normalized_residuals_real[i] = (self.rdata[i] - RealFit[i])/RealFit[i]
                            normalized_error_real_below[i] = 2*self.sdrReal[i]/RealFit[i]
                            normalized_error_real_above[i] = -2*self.sdrReal[i]/RealFit[i]
                        resultPlotBig.title("Real Normalized Residuals")
                        larger.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                        if (self.confInt):
                            larger.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                        larger.axhline(0, color="black", linewidth=1.0)
                        larger.set_xscale("log")
                        larger.set_title("Real Normalized Residuals", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("(Zr-Z\u0302rmodel)/Zr", color=self.foregroundColor)
                    elif (event.inaxes == self.lplot):
                        normalized_residuals_imag = np.zeros(len(self.wdata))
                        normalized_error_imag_below = np.zeros(len(self.wdata))
                        normalized_error_imag_above = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            normalized_residuals_imag[i] = (self.jdata[i] - ImagFit[i])/ImagFit[i]
                            normalized_error_imag_below[i] = 2*self.sdrImag[i]/ImagFit[i]
                            normalized_error_imag_above[i] = -2*self.sdrImag[i]/ImagFit[i]
                        resultPlotBig.title("Imaginary Normalized Residuals")
                        larger.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                        if (self.confInt):
                            larger.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                        larger.axhline(0, color="black", linewidth=1.0)
                        larger.set_xscale("log")
                        larger.set_title("Imaginary Normalized Residuals", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("(Zj-Z\u0302jmodel)/Zj", color=self.foregroundColor)
    
                    larger.xaxis.label.set_fontsize(20)
                    larger.yaxis.label.set_fontsize(20)
                    larger.title.set_fontsize(30)    
                largerCanvas = FigureCanvasTkAgg(fig, resultPlotBig)
                largerCanvas.draw()
                largerCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                toolbar = NavigationToolbar2Tk(largerCanvas, resultPlotBig)
                toolbar.update()
                largerCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                def on_closing():   #Clear the figure before closing the popup
                    fig.clear()
                    resultPlotBig.destroy()
                    #self.resultPlot.lift(self.parent)       #Keep other popups open if a lower one is closed

                #self.resultPlot.lift(self.parent)
                resultPlotBig.protocol("WM_DELETE_WINDOW", on_closing)

            except:
                pass
        
        def graphOut(event):
            self.nyCanvas._tkcanvas.config(cursor="arrow")
        
        def graphOver(event):
            #axes = event.inaxes
            #autoAxis = event.inaxes
            whichCan = self.nyCanvas._tkcanvas
            #he = int(whichCan.winfo_height())
            #wi = int(whichCan.winfo_width())
            whichCan.config(cursor="hand2")
        
        def plotResults():
            self.resultPlot = tk.Toplevel(background=self.backgroundColor)
            self.resultPlot.title("Measurement Model Plots")
            self.resultPlot.iconbitmap(resource_path('img/elephant3.ico'))
            self.resultPlot.state("zoomed")
            self.confInt = False if self.confidenceIntervalCheckboxVariable.get() == 0 else True
            RealFit = np.array(self.realFit)
            ImagFit = np.array(self.imagFit)
            phase_fit = np.arctan2(self.imagFit, self.realFit) * (180/np.pi)
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                self.f = Figure()#figsize=((5,4), dpi=100)
                self.f.set_facecolor(self.backgroundColor)
                dataColor = "tab:blue"
                fitColor = "orange"
                if (self.topGUI.getTheme() == "dark"):
                    dataColor = "cyan"
                    fitColor = "gold"
                else:
                    dataColor = "tab:blue"
                    fitColor = "orange"
                
                self.aplot = self.f.add_subplot(331)
                self.aplot.set_facecolor(self.backgroundColor)
                self.aplot.yaxis.set_ticks_position("both")
                self.aplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.aplot.xaxis.set_ticks_position("both")
                self.aplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.aplot.plot(self.rdata, -1*self.jdata, "o", color=dataColor)
                self.aplot.plot(RealFit, -1*ImagFit, color=fitColor, marker="x", markeredgecolor="black")
                if (self.confInt):
                    for i in range(len(self.wdata)):
                        ellipse = patches.Ellipse(xy=(RealFit[i], -1*ImagFit[i]), width=4*self.sdrReal[i], height=4*self.sdrImag[i])
                        self.aplot.add_artist(ellipse)
                        ellipse.set_alpha(0.1)
                        ellipse.set_facecolor(self.ellipseColor)
                self.aplot.axis("equal")
                #self.aplot.set_title("Nyquist")
                extra = patches.Rectangle((0, 0), 1, 1, fc="w", fill=False, edgecolor='none', linewidth=0)
                legA = self.aplot.legend([extra], ["Nyquist"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legA.get_texts():
                    text.set_color(self.foregroundColor)
                self.aplot.set_xlabel("Zr / Ω", color=self.foregroundColor)
                self.aplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                
                self.bplot = self.f.add_subplot(332)
                self.bplot.set_facecolor(self.backgroundColor)
                self.bplot.yaxis.set_ticks_position("both")
                self.bplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.bplot.xaxis.set_ticks_position("both")
                self.bplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.bplot.plot(self.wdata, self.rdata, "o", color=dataColor)
                self.bplot.plot(self.wdata, RealFit, color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(RealFit[i] + 2*self.sdrReal[i], RealFit[i] - 2*self.sdrReal[i])
                        error_below[i] = min(RealFit[i] + 2*self.sdrReal[i], RealFit[i] - 2*self.sdrReal[i])
                    self.bplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.bplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.bplot.set_xscale('log')
                #self.bplot.set_title("Real Impedance")
                legB = self.bplot.legend([extra], ["Real Impedance"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legB.get_texts():
                    text.set_color(self.foregroundColor)
                self.bplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.bplot.set_ylabel("Zr / Ω", color=self.foregroundColor)
                
                self.cplot = self.f.add_subplot(333)
                self.cplot.set_facecolor(self.backgroundColor)
                self.cplot.yaxis.set_ticks_position("both")
                self.cplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.cplot.xaxis.set_ticks_position("both")
                self.cplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.cplot.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                self.cplot.plot(self.wdata, -1*ImagFit, color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(ImagFit[i] + 2*self.sdrImag[i], ImagFit[i] - 2*self.sdrImag[i])
                        error_below[i] = min(ImagFit[i] + 2*self.sdrImag[i], ImagFit[i] - 2*self.sdrImag[i])
                    self.cplot.plot(self.wdata, -1*error_above, "--", color=self.ellipseColor)
                    self.cplot.plot(self.wdata, -1*error_below, "--", color=self.ellipseColor)
                self.cplot.set_xscale('log')
                #self.cplot.set_title("Imaginary Impedance")
                legC = self.cplot.legend([extra], ["Imaginary Impedance"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legC.get_texts():
                    text.set_color(self.foregroundColor)
                self.cplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.cplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                
                self.dplot = self.f.add_subplot(334)
                self.dplot.set_facecolor(self.backgroundColor)
                self.dplot.yaxis.set_ticks_position("both")
                self.dplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.dplot.xaxis.set_ticks_position("both")
                self.dplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.dplot.plot(self.wdata, np.sqrt(self.rdata**2 + self.jdata**2), "o", color=dataColor)
                self.dplot.plot(self.wdata, np.sqrt(RealFit**2 + ImagFit**2), color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = np.sqrt(max((RealFit[i]+2*self.sdrReal[i])**2, (RealFit[i]-2*self.sdrReal[i])**2) + max((ImagFit[i]+2*self.sdrImag[i])**2, (ImagFit[i]-2*self.sdrImag[i])**2))
                        error_below[i] = np.sqrt(min((RealFit[i]+2*self.sdrReal[i])**2, (RealFit[i]-2*self.sdrReal[i])**2) + min((ImagFit[i]+2*self.sdrImag[i])**2, (ImagFit[i]-2*self.sdrImag[i])**2))
                    self.dplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.dplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.dplot.set_xscale('log')
                self.dplot.set_yscale('log')
                #self.dplot.set_title("|Z| Bode")
                legD = self.dplot.legend([extra], ["|Z| Bode"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legD.get_texts():
                    text.set_color(self.foregroundColor)
                self.dplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.dplot.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                
                self.eplot = self.f.add_subplot(335)
                self.eplot.set_facecolor(self.backgroundColor)
                self.eplot.yaxis.set_ticks_position("both")
                self.eplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.eplot.xaxis.set_ticks_position("both")
                self.eplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.eplot.plot(self.wdata, np.arctan2(self.jdata, self.rdata)*(180/np.pi), "o", color=dataColor)
                self.eplot.plot(self.wdata, phase_fit, color=fitColor)
                if (self.confInt):
    #                error_above = np.zeros(len(self.wdata))
    #                error_below = np.zeros(len(self.wdata))
    #                for i in range(len(self.wdata)):
    #                    error_above[i] = np.arctan2(max((Zfit.imag[i]+2*self.sdrImag[i]), (Zfit.imag[i]-2*self.sdrImag[i])), min((Zfit.real[i]+2*self.sdrReal[i]), (Zfit.real[i]-2*self.sdrReal[i])))
    #                    error_below[i] = np.arctan2(min((Zfit.imag[i]+2*self.sdrImag[i]), (Zfit.imag[i]-2*self.sdrImag[i])), max((Zfit.real[i]+2*self.sdrReal[i]), (Zfit.real[i]-2*self.sdrReal[i])))
                    error_above = np.arctan2((ImagFit+2*self.sdrImag) , (RealFit+2*self.sdrReal)) * (180/np.pi)
                    error_below = np.arctan2((ImagFit-2*self.sdrImag) , (RealFit-2*self.sdrReal)) * (180/np.pi)
                    self.eplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.eplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
    #            self.eplot.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                self.eplot.set_ylim(bottom=-90, top=0)
                self.eplot.set_xscale('log')
                #self.eplot.set_title("Phase Angle Bode")
                legE = self.eplot.legend([extra], ["Phase Angle Bode"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legE.get_texts():
                    text.set_color(self.foregroundColor)
                self.eplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.eplot.set_ylabel("Phase angle / Deg.", color=self.foregroundColor)
                
                self.hplot = self.f.add_subplot(336)
                self.hplot.set_facecolor(self.backgroundColor)
                self.hplot.yaxis.set_ticks_position("both")
                self.hplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.hplot.xaxis.set_ticks_position("both")
                self.hplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.hplot.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                self.hplot.plot(self.wdata, np.log10(abs(RealFit)), color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i]))) #np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                        error_below[i] = min(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i])))#np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                    self.hplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.hplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.hplot.set_xscale('log')
                #self.hplot.set_title("Log|Zr|")
                legH = self.hplot.legend([extra], ["Log|Zr|"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legH.get_texts():
                    text.set_color(self.foregroundColor)
                self.hplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.hplot.set_ylabel("Log|Zr|", color=self.foregroundColor)
                
                self.iplot = self.f.add_subplot(337)
                self.iplot.set_facecolor(self.backgroundColor)
                self.iplot.yaxis.set_ticks_position("both")
                self.iplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.iplot.xaxis.set_ticks_position("both")
                self.iplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.iplot.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                self.iplot.plot(self.wdata, np.log10(abs(ImagFit)), color=fitColor)
                if (self.confInt):
                    error_above = np.zeros(len(self.wdata))
                    error_below = np.zeros(len(self.wdata))
                    for i in range(len(self.wdata)):
                        error_above[i] = max(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))#np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                        error_below[i] = min(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))#np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                    self.iplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.iplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.iplot.set_xscale('log')
                #self.iplot.set_title("Log|Zj|")
                legI = self.iplot.legend([extra], ["Log|Zj|"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legI.get_texts():
                    text.set_color(self.foregroundColor)
                self.iplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.iplot.set_ylabel("Log|Zj|", color=self.foregroundColor)
                
                normalized_residuals_real = np.zeros(len(self.wdata))
                normalized_error_real_above = np.zeros(len(self.wdata))
                normalized_error_real_below = np.zeros(len(self.wdata))
                normalized_residuals_imag = np.zeros(len(self.wdata))
                normalized_error_imag_above = np.zeros(len(self.wdata))
                normalized_error_imag_below = np.zeros(len(self.wdata))
                for i in range(len(self.wdata)):
                    normalized_residuals_real[i] = (self.rdata[i] - RealFit[i])/RealFit[i]
                    normalized_error_real_above[i] = 2*self.sdrReal[i]/RealFit[i]
                    normalized_error_real_below[i] = -2*self.sdrReal[i]/RealFit[i]
                    normalized_residuals_imag[i] = (self.jdata[i] - ImagFit[i])/ImagFit[i]
                    normalized_error_imag_below[i] = 2*self.sdrImag[i]/ImagFit[i]
                    normalized_error_imag_above[i] = -2*self.sdrImag[i]/ImagFit[i]
                
                self.kplot = self.f.add_subplot(338)
                self.kplot.set_facecolor(self.backgroundColor)
                self.kplot.yaxis.set_ticks_position("both")
                self.kplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.kplot.xaxis.set_ticks_position("both")
                self.kplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.kplot.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                if (self.confInt):
                    self.kplot.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                    self.kplot.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                #fplot.plot(self.wdata, phase_fit, color="orange")
                self.kplot.axhline(0, color="black", linewidth=1.0)
                self.kplot.set_xscale('log')
                #self.kplot.set_title("Real Residuals")
                legK = self.kplot.legend([extra], ["Real Residuals"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                for text in legK.get_texts():
                    text.set_color(self.foregroundColor)
                self.kplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.kplot.set_ylabel("(Zr-Zrmodel)/Zr", color=self.foregroundColor)
                
                self.lplot = self.f.add_subplot(339)
                self.lplot.set_facecolor(self.backgroundColor)
                self.lplot.yaxis.set_ticks_position("both")
                self.lplot.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.lplot.xaxis.set_ticks_position("both")
                self.lplot.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                self.lplot.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                if (self.confInt):
                    self.lplot.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                    self.lplot.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                #fplot.plot(self.wdata, phase_fit, color="orange")
                self.lplot.axhline(0, color="black", linewidth=1.0)
                self.lplot.set_xscale('log')
                #self.lplot.set_title("Imaginary Residuals")
                legL = self.lplot.legend([extra], ["Imaginary Residuals"], loc=9, frameon=False, bbox_to_anchor=(0., 1.01, 1., .101), borderaxespad=0., prop={'size': 17})
                self.lplot.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.lplot.set_ylabel("(Zj-Zjmodel)/Zj", color=self.foregroundColor)
                for text in legL.get_texts():
                    text.set_color(self.foregroundColor)
                self.nyCanvas = FigureCanvasTkAgg(self.f, self.resultPlot)
                self.f.subplots_adjust(top=0.95, bottom=0.1, right=0.95, left=0.05)
            self.nyCanvas.draw()
            self.nyCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
            self.nyCanvas.mpl_connect('button_press_event', onclick)
            self.nyCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            enterAxes = self.f.canvas.mpl_connect('axes_enter_event', graphOver)
            leaveAxes = self.f.canvas.mpl_connect('axes_leave_event', graphOut)
            def on_closing():   #Clear the figure before closing the popup
                self.f.clear()
                self.resultPlot.destroy()
                self.alreadyPlotted = False
            
            self.resultPlot.protocol("WM_DELETE_WINDOW", on_closing)
            
        def saveFitting():
            valuesMatcher = []
            a = True
            for child in self.resultsView.get_children():
                resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                valuesMatcher.append(resultValue)
            if (len(valuesMatcher) != 0):
                if (len(valuesMatcher) != len(self.paramValueValues)):
                    a = messagebox.askokcancel("Values do not match", "The values under \"Edit Parameters\" do not match the current fitted values. Only values under \"Edit Parameters\" will be saved. Continue?")
                else:
                    dontMatch = False
                    for i in range(len(valuesMatcher)):
                        if (valuesMatcher[i] != self.paramValueValues[i].get()):
                            dontMatch = True
                            break
                    if (dontMatch):
                        a = messagebox.askokcancel("Values do not match", "The values under \"Edit Parameters\" do not match the current fitted values. Only values under \"Edit Parameters\" will be saved. Continue?")
            if (a):
                if (self.browseEntry.get() == ""):
                    stringToSave = "filename: NONE\n"
                else:
                    stringToSave = "filename: " + self.browseEntry.get() + "\n"
                stringToSave += "number_simulations: " + self.monteCarloValue.get() + "\n"
                stringToSave += "fit_type: " + self.fittingTypeComboboxValue.get() + "\n"
                stringToSave += "alpha: " + self.noiseEntryValue.get() + "\n"
                stringToSave += "upDelete: " + str(self.upDelete) + "\n"
                stringToSave += "lowDelete: " + str(self.lowDelete) + "\n"
                stringToSave += "weighting: " + self.weightingComboboxValue.get() + "\n"
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
                stringToSave += "formula: \n"
                stringToSave += self.customFormula.get("1.0", tk.END)
                stringToSave += "\nmmparams: \n"
                for i in range(len(self.paramNameValues)):
                    stringToSave += self.paramNameValues[i].get() + "~=~" + self.paramValueValues[i].get() + "~=~" + self.paramComboboxValues[i].get() + "\n"
                defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
                saveName = asksaveasfile(mode='w', defaultextension=".mmcustom", initialfile=defaultSaveName, filetypes=[("Measurement model custom fitting", ".mmcustom")])
                directory = os.path.dirname(str(saveName))
                self.topGUI.setCurrentDirectory(directory)
                if saveName is None:     #If save is cancelled
                    return
                saveName.write(stringToSave)
                saveName.close()
                self.saveNeeded = False
                self.saveFormulaButton.configure(text="Saved")
                self.after(1000, lambda : self.saveFormulaButton.configure(text="Save Fitting"))
        
        def saveResiduals():
            messagebox.showinfo("Warning", "Note: The \"ohmic resistance\" for custom data sets is saved as 0.")
            stringToSave = "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\n"
            for i in range(len(self.wdata)):
                stringToSave += str(self.wdata[i]) + "\t" + str(self.rdata[i] - self.realFit[i]) + "\t" + str(self.jdata[i] - self.imagFit[i]) + "\t" + str(self.rdata[i]) + "\t" + str(self.jdata[i]) + "\n"
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
            saveName = asksaveasfile(mode='w', defaultextension=".mmresiduals", initialfile=defaultSaveName, filetypes=[("Measurement model residuals", ".mmresiduals")])
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if saveName is None:     #If save is cancelled
                return
            saveName.write(stringToSave)
            saveName.close()
            self.saveResiduals.configure(text="Saved")
            self.after(1000, lambda : self.saveResiduals.configure(text="Save Residuals"))
        
        def saveAll():
            stringToSave = "freq\tZrdata\tZjdata\tZrmodel\tZjmodel\tSigmaZrconf\tSigmaZjconf\n"
            for i in range(len(self.wdata)):
                stringToSave += str(self.wdata[i]) + "\t" + str(self.rdata[i]) + "\t" + str(self.jdata[i]) + "\t" + str(self.realFit[i]) + "\t" + str(self.imagFit[i]) + "\t"  + str(self.sdrReal[i]) + "\t" + str(self.sdrImag[i]) + "\n"
            stringToSave += "----------------------------------------------------------------------------------\n"
            for i in range(len(self.paramNameValues)):
                 stringToSave += self.paramNameValues[i].get() + " = " + str(self.fits[i]) + "\tStd. Dev. = " + str(self.sigmas[i]) + "\n"
            stringToSave += "Chi-squared = " + str(self.chiSquared)
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))                
            saveName = asksaveasfile(title="Save All Results", mode='w', defaultextension=".txt", initialfile=defaultSaveName, filetypes=[("Text file (*.txt)", ".txt")])
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if saveName is None:
                return
            saveName.write(stringToSave)
            saveName.close()
            self.saveAll.configure(text="Saved")
            self.after(1000, lambda : self.saveAll.configure(text="Export All Results"))
        
        def checkWeight(e):
            if (self.weightingComboboxValue.get() == "None"):
                self.noiseFrame.grid_remove()
                self.errorStructureFrame.grid_remove()
            elif (self.weightingComboboxValue.get() == "Error model"):
                self.noiseFrame.grid_remove()
                self.errorStructureFrame.grid(column=0, row=2, pady=(5,0), columnspan=8)
            elif (self.weightingComboboxValue.get() == "Custom"):
                self.noiseFrame.grid_remove()
                self.errorStructureFrame.grid_remove()
            else:
                self.noiseFrame.grid(column=3, row=1, pady=(5,0), sticky="W")
                self.errorStructureFrame.grid_remove()
            keyup("e")
        
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
        
        self.browseFrame = tk.Frame(self, bg=self.backgroundColor)
        self.browseEntry = ttk.Entry(self.browseFrame, state="readonly", width=40)
        self.browseLabel = tk.Label(self.browseFrame, text="File to fit:   ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.browseButton = ttk.Button(self.browseFrame, text="Browse...", command=browse)
        self.freqRangeButton = ttk.Button(self.browseFrame, text="Frequency Range", state="disabled", command=changeFreqs)
        self.browseLabel.grid(column=0, row=0, sticky="W")
        self.browseButton.grid(column=1, row=0, sticky="W")
        self.browseEntry.grid(column=2, row=0, sticky="W", padx=5)
        self.freqRangeButton.grid(column=3, row=0, sticky="W")
        self.browseFrame.grid(column=0, row=0, sticky="W")
        browse_ttp = CreateToolTip(self.browseButton, 'Browse for a .mmfile or .mmcustom file')
        frequencyRange_ttp = CreateToolTip(self.freqRangeButton, 'Change frequencies used in fitting')
        
        self.fittingButtonFrame = tk.Frame(self, bg=self.backgroundColor)
        self.parametersButton = ttk.Button(self.fittingButtonFrame, text="Fitting parameters", command=parameterPopup)
        self.fittingTypeComboboxValue = tk.StringVar(self, "Complex")
        self.fittingTypeLabel = tk.Label(self.fittingButtonFrame, text="Fit type: ", background=self.backgroundColor, foreground=self.foregroundColor)
        self.fittingTypeCombobox = ttk.Combobox(self.fittingButtonFrame, textvariable=self.fittingTypeComboboxValue, value=("Real", "Imaginary", "Complex"), state="readonly", exportselection=0, width=10)
        self.monteCarloValue = tk.StringVar(self, "1000")
        self.monteCarloLabel = tk.Label(self.fittingButtonFrame, text="Number of Simulations: ", background=self.backgroundColor, foreground=self.foregroundColor)
        self.monteCarloEntry = ttk.Entry(self.fittingButtonFrame, textvariable=self.monteCarloValue, width=8, exportselection=0)
        self.codeLabel = tk.Label(self.fittingButtonFrame, text="Code:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.helpLabel = tk.Label(self.fittingButtonFrame, text="Help", bg=self.backgroundColor, fg="blue", cursor="hand2")
        if (self.topGUI.getTheme() == "dark"):
            self.helpLabel.configure(fg="skyblue")
        self.helpLabel.bind("<Button-1>", helpPopup)
        self.weightingLabel = tk.Label(self.fittingButtonFrame, text="Weighting:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.weightingComboboxValue = tk.StringVar(self, "Modulus")
        self.weightingCombobox = ttk.Combobox(self.fittingButtonFrame, textvariable=self.weightingComboboxValue, value=("None", "Modulus", "Proportional", "Error model", "Custom"), state="readonly", exportselection=0, width=13)
        self.weightingCombobox.bind("<<ComboboxSelected>>", checkWeight)
        self.noiseFrame = tk.Frame(self.fittingButtonFrame, bg=self.backgroundColor)
        self.noiseLabel = tk.Label(self.noiseFrame, text="α: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.noiseEntryValue = tk.StringVar(self, "1")
        self.noiseEntry = ttk.Entry(self.noiseFrame, textvariable=self.noiseEntryValue, width=8)
        self.parametersButton.grid(column=0, row=0, sticky="W")
        self.fittingTypeLabel.grid(column=1, row=0, sticky="W", padx=(10, 4))
        self.fittingTypeCombobox.grid(column=2, row=0, sticky="W", padx=(0, 10))
        self.monteCarloLabel.grid(column=3, row=0, sticky="W", padx=(0, 4))
        self.monteCarloEntry.grid(column=4, row=0, sticky="W")
        self.codeLabel.grid(column=0, row=3, pady=(5, 0), sticky="W")
        self.weightingLabel.grid(column=1, row=1, pady=(5,0), sticky="W")
        self.weightingCombobox.grid(column=2, row=1, pady=(5,0), sticky="W")
        self.noiseFrame.grid(column=3, row=1, pady=(5,0), sticky="W")
        self.noiseLabel.grid(column=0, row=0, sticky="E")
        self.noiseEntry.grid(column=1, row=0, sticky="W")
        self.helpLabel.grid(column=3, row=3, pady=(5, 0), sticky="E")
        self.fittingButtonFrame.grid(column=0, row=1, pady=(10, 0), sticky="W")
        paramButton_ttp = CreateToolTip(self.parametersButton, 'Add, remove, or edit fitting parameters')
        fittingType_ttp = CreateToolTip(self.fittingTypeCombobox, 'Change which part of the data is fitted')
        monteCarlo_ttp = CreateToolTip(self.monteCarloEntry, 'Number of Monte Carlo simulations')
        weighting_ttp = CreateToolTip(self.weightingCombobox, 'Weighting used in objective function')
        noise_ttp = CreateToolTip(self.noiseEntry, 'Assumed noise (multiplied by weighting)')
        
        #---If error structure weighting is chosen---
        self.errorStructureFrame = tk.Frame(self.fittingButtonFrame, bg=self.backgroundColor)
        self.errorAlphaCheckboxVariable = tk.IntVar(self, 0)
        self.errorAlphaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorAlphaCheckboxVariable, text="\u03B1 = ", command=checkErrorStructure)
        self.errorAlphaVariable = tk.StringVar(self, "0.1")
        self.errorAlphaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorAlphaVariable, state="disabled", exportselection=0, width=6)
        self.errorBetaCheckboxVariable = tk.IntVar(self, 0)
        self.errorBetaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorBetaCheckboxVariable, text="\u03B2 = ", command=checkErrorStructure)
        self.errorBetaVariable = tk.StringVar(self, "0.1")
        self.errorBetaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorBetaVariable, state="disabled", exportselection=0, width=6)
        self.errorBetaReCheckboxVariable = tk.IntVar(self, 0)
        self.errorBetaReCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorBetaReCheckboxVariable, state="disabled", text="Re = ", command=checkErrorStructure)
        self.errorBetaReVariable = tk.StringVar(self, "0.1")
        self.errorBetaReEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorBetaReVariable, state="disabled", exportselection=0, width=6)
        self.errorGammaCheckboxVariable = tk.IntVar(self, 1)
        self.errorGammaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorGammaCheckboxVariable, text="\u03B3 = ", command=checkErrorStructure)
        self.errorGammaVariable = tk.StringVar(self, "0.1")
        self.errorGammaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorGammaVariable, exportselection=0, width=6)
        self.errorDeltaCheckboxVariable = tk.IntVar(self, 1)
        self.errorDeltaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorDeltaCheckboxVariable, text="\u03B4 = ", command=checkErrorStructure)
        self.errorDeltaVariable = tk.StringVar(self, "0.1")
        self.errorDeltaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorDeltaVariable, exportselection=0, width=6)
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
        
        self.customFunctionFrame = tk.Frame(self, bg=self.backgroundColor)
        self.customFunctionContainer = tk.Frame(self.customFunctionFrame, borderwidth=1, relief="sunken")
        self.customFormula = tk.Text(self.customFunctionContainer, width=60, height=10, wrap="none", borderwidth=0, undo=True)
        self.defaultFormula = "#Use freq for the independent variable and 1j for the complex number\nZ = \nZreal = Z.real\nZimag = Z.imag"
        self.customFormula.insert(tk.END, self.defaultFormula)
        customFormulaVertical = tk.Scrollbar(self.customFunctionContainer, orient="vertical", command=self.customFormula.yview)
        customFormulaHorizontal = tk.Scrollbar(self.customFunctionContainer, orient="horizontal", command=self.customFormula.xview)
        self.customFormula.configure(yscrollcommand=customFormulaVertical.set, xscrollcommand=customFormulaHorizontal.set)
        formulaFont = tkfont.Font(font=self.customFormula['font'])
        tab_width = formulaFont.measure(' ' * 4)
        self.customFormula.configure(tabs=(tab_width,))
        self.customFormula.bind("<KeyRelease>", keyup)
        #self.customFormula.bind("<KeyRelease-Return>", keyup_return)
        self.customFormula.grid(row=0, column=0, sticky="nsew")
        customFormulaVertical.grid(row=0, column=1, sticky="ns")
        customFormulaHorizontal.grid(row=1, column=0, sticky="ew")
        self.customFunctionContainer.grid_rowconfigure(0, weight=1)
        self.customFunctionContainer.grid_columnconfigure(0, weight=1)
        self.customFunctionContainer.grid(column=0, row=0)
        self.customFunctionFrame.grid(column=0, row=2, pady=(0, 10), sticky="W")
        keyup(None)
        
        self.popup_menu = tk.Menu(self.customFormula, tearoff=0)
        self.popup_menu.add_command(label="Copy", command=copyFormula)
        self.popup_menu.add_command(label="Cut", command=cutFormula)
        self.popup_menu.add_command(label="Paste", command=pasteFormula)
        self.popup_menu.add_separator()
        self.popup_menu.add_command(label="Undo", command=undoFormula)
        self.popup_menu.add_command(label="Redo", command=redoFormula)
        self.customFormula.bind("<Button-3>", popup_formula)
        
        self.runFrame = tk.Frame(self, bg=self.backgroundColor)
        self.runButton = ttk.Button(self.runFrame, text="Run fitting", state="disabled", command=checkFitting)
        self.saveFormulaButton = ttk.Button(self.runFrame, text="Save fitting", command=saveFitting)
        self.cancelButton = ttk.Button(self.runFrame, text="Cancel fitting")
        self.runButton.grid(column=0, row=0, sticky="W")
        self.saveFormulaButton.grid(column=1, row=0, sticky="W", padx=5)
        self.runFrame.grid(column=0, row=3, sticky="W")
        runButton_ttp = CreateToolTip(self.runButton, 'Fit formula using Levenberg-Marquardt algorithm')
        saveFormualButton_ttp = CreateToolTip(self.saveFormulaButton, 'Save the current formula and parameter values as a .mmcustom file')
        cancelButton_ttp = CreateToolTip(self.cancelButton, 'Cancel fitting')
        
        self.sep = ttk.Separator(self, orient=tk.HORIZONTAL)
        self.sep.grid(column=0, row=4, sticky="EW", pady=5)
        
        #---The results display---
        self.resultsFrame = tk.Frame(self, bg=self.backgroundColor)
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
        self.includeFrame = tk.Frame(self.graphFrame, bg=self.backgroundColor)
        self.confidenceIntervalCheckboxVariable = tk.IntVar(self, 0)
        self.confidenceIntervalCheckbox = ttk.Checkbutton(self.includeFrame, text="Include Confidence Interval", variable=self.confidenceIntervalCheckboxVariable)
        self.plotButton = ttk.Button(self.graphFrame, text="Plot", command=plotResults)
        self.plotButton.grid(column=0, row=0, sticky="W", pady=(8,5))
        self.confidenceIntervalCheckbox.grid(column=1, row=0, sticky="W", padx=5)
        self.includeFrame.grid(column=1, row=0, sticky="W")
        plot_ttp = CreateToolTip(self.plotButton, 'Opens a window with plots using data and fitted parameters')
        include_ttp = CreateToolTip(self.confidenceIntervalCheckbox, 'Include confidence interval ellipses and bars when plotting')
        
        #---The save and export buttons---
        self.saveFrame = tk.Frame(self, bg=self.backgroundColor)
        self.saveResiduals = ttk.Button(self.saveFrame, text="Save Residuals", command=saveResiduals)
        self.saveAll = ttk.Button(self.saveFrame, text="Export All Results", command=saveAll)
        self.saveResiduals.grid(column=0, row=0, sticky="W", padx=(0, 5))
        self.saveAll.grid(column=1, row=0)
        saveResiduals_ttp = CreateToolTip(self.saveResiduals, 'Save residuals for error analysis')
        saveAll_ttp = CreateToolTip(self.saveAll, 'Export all results, including data, fits, formula, and parameters')
    
    def _on_mousewheel(self, event):
            xpos, ypos = self.vsb.get()
            xpos_round = round(xpos, 2)     #Avoid floating point errors
            ypos_round = round(ypos, 2)
            if (xpos_round != 0.00 or ypos_round != 1.00):
                self.pcanvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def keyup(self, event):
        self.customFormula.tag_remove("tagZreal", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords2", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords3", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords4", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords5", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords6", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords7", "1.0", tk.END)
        self.customFormula.tag_remove("tagReservedWords8", "1.0", tk.END)
        self.customFormula.tag_remove("tagComment", "1.0", tk.END)
        self.customFormula.tag_remove("tagUserVariables", "1.0", tk.END)
        self.customFormula.tag_remove("tagBuiltIns", "1.0", tk.END)
        self.customFormula.tag_remove("tagQuotations", "1.0", tk.END)
        self.customFormula.tag_remove("tagApostrophes", "1.0", tk.END)
        self.customFormula.tag_remove("tagFreq", "1.0", tk.END)
        self.customFormula.tag_remove("tagZ", "1.0", tk.END)
        self.customFormula.tag_remove("tagWeight", "1.0", tk.END)
        textToSearch = self.customFormula.get("1.0", tk.END)
        if (textToSearch[:-1] == self.defaultFormula):
            self.saveNeeded = False
        else:
            self.saveNeeded = True
        whereStr = [m.start() for m in re.finditer(r'\bZreal\b', textToSearch)]
        whereStr.extend([m.start() for m in re.finditer(r'\bZimag\b', textToSearch)])
        whereFreq = [m.start() for m in re.finditer(r'\bfreq\b', textToSearch)]
        whereZ = [m.start() for m in re.finditer(r'\bZr\b', textToSearch)]
        whereZ.extend([m.start() for m in re.finditer(r'\bZj\b', textToSearch)])
        reservedWords2 = [m.start() for m in re.finditer(r'\bis\b', textToSearch)]
        reservedWords2.extend([m.start() for m in re.finditer(r'\bas\b', textToSearch)])
        reservedWords2.extend([m.start() for m in re.finditer(r'\bif\b', textToSearch)])
        reservedWords2.extend([m.start() for m in re.finditer(r'\bor\b', textToSearch)])
        reservedWords2.extend([m.start() for m in re.finditer(r'\bin\b', textToSearch)])
        reservedWords3 = [m.start() for m in re.finditer(r'\bfor\b', textToSearch)]
        reservedWords3.extend([m.start() for m in re.finditer(r'\btry\b', textToSearch)])
        reservedWords3.extend([m.start() for m in re.finditer(r'\bdef\b', textToSearch)])
        reservedWords3.extend([m.start() for m in re.finditer(r'\band\b', textToSearch)])
        reservedWords3.extend([m.start() for m in re.finditer(r'\bdel\b', textToSearch)])
        reservedWords3.extend([m.start() for m in re.finditer(r'\bnot\b', textToSearch)])
        reservedWords4 = [m.start() for m in re.finditer(r'\bNone\b', textToSearch)]
        reservedWords4.extend([m.start() for m in re.finditer(r'\bTrue\b', textToSearch)])
        reservedWords4.extend([m.start() for m in re.finditer(r'\bfrom\b', textToSearch)])
        reservedWords4.extend([m.start() for m in re.finditer(r'\bwith\b', textToSearch)])
        reservedWords4.extend([m.start() for m in re.finditer(r'\belif\b', textToSearch)])
        reservedWords4.extend([m.start() for m in re.finditer(r'\belse\b', textToSearch)])
        reservedWords4.extend([m.start() for m in re.finditer(r'\bpass\b', textToSearch)])
        reservedWords5 = [m.start() for m in re.finditer(r'\bFalse\b', textToSearch)]
        reservedWords5.extend([m.start() for m in re.finditer(r'\bclass\b', textToSearch)])
        reservedWords5.extend([m.start() for m in re.finditer(r'\bwhile\b', textToSearch)])
        reservedWords5.extend([m.start() for m in re.finditer(r'\byield\b', textToSearch)])
        reservedWords5.extend([m.start() for m in re.finditer(r'\bbreak\b', textToSearch)])
        reservedWords5.extend([m.start() for m in re.finditer(r'\braise\b', textToSearch)])
        reservedWords6 = [m.start() for m in re.finditer(r'\breturn\b', textToSearch)]
        reservedWords6.extend([m.start() for m in re.finditer(r'\blambda\b', textToSearch)])
        reservedWords6.extend([m.start() for m in re.finditer(r'\bglobal\b', textToSearch)])
        reservedWords6.extend([m.start() for m in re.finditer(r'\bassert\b', textToSearch)])
        reservedWords6.extend([m.start() for m in re.finditer(r'\bimport\b', textToSearch)])
        reservedWords6.extend([m.start() for m in re.finditer(r'\bexcept\b', textToSearch)])
        reservedWords7 = [m.start() for m in re.finditer(r'\bfinally\b', textToSearch)]
        reservedWords8 = [m.start() for m in re.finditer(r'\bcontinue\b', textToSearch)]
        reservedWords8.extend([m.start() for m in re.finditer(r'\bnonlocal\b', textToSearch)])
        whereWeight = [m.start() for m in re.finditer(r'\bweighting\b', textToSearch)]
        whereComment = []
        userVariablesBegin = []
        userVariablesEnd = []
        for i in range(len(self.paramNameValues)):
            tempList = [m.start() for m in re.finditer(r'\b' + self.paramNameValues[i].get() + r'\b', textToSearch)]
            userVariablesBegin.extend(tempList)
            if (len(tempList) > 0):
                for j in range(len(tempList)):
                    userVariablesEnd.append(tempList[j] + len(self.paramNameValues[i].get()))
        
        builtIns = ['id', 'abs', 'set', 'all', 'min', 'any', 'dir', 'hex', 'bin', 'oct', 'int', 'str', 'ord', 'sum', 'pow', 'len', 'chr', 'zip', 'map', 'max', 'hash', 'dict', 'help', 'next', 'bool', 'eval', 'open', 'exec', 'iter', 'type', 'list', 'vars', 'repr'\
                    'slice', 'ascii', 'input', 'super', 'bytes', 'float', 'print', 'tuple', 'range', 'round', 'divmod', 'object', 'sorted', 'filter', 'format', 'locals', 'delattr', 'setattr', 'getattr', 'compile', 'globals', 'complex', 'hasattr', 'callable'\
                    'property', 'reversed', 'enumerate', 'bytearray', 'frozenset', 'memoryview', 'breakpoint', 'isinstance', 'issubclass', '__import__', 'classmethod', 'staticmethod']
        builtInsBegin = []
        builtInsEnd = []
        for i in range(len(builtIns)):
            tempList = [m.start() for m in re.finditer(r'\b' + builtIns[i] + r'\(', textToSearch)]
            builtInsBegin.extend(tempList)
            if (len(tempList) > 0):
                for j in range(len(tempList)):
                    builtInsEnd.append(tempList[j] + len(builtIns[i]))
        
        quotationBegin = []
        quotationEnd = []
        apostropheBegin = []
        apostropheEnd = []
        isComment = False
        isQuote = False
        isApostrophe = False
        for i in range(len(textToSearch)):
            if (textToSearch[i] == "#" and not isQuote and not isApostrophe):
                whereComment.append(i)
                isComment = True
            elif (textToSearch[i] == "\n"):
                isComment = False
            elif (not isComment and not isApostrophe and textToSearch[i] == "\""):
                if (len(quotationBegin) > len(quotationEnd)):
                    quotationEnd.append(i)
                    isQuote = False
                else:
                    quotationBegin.append(i)
                    isQuote = True
            elif (not isComment and not isQuote and textToSearch[i] == "\'"):
                if (len(apostropheBegin) > len(apostropheEnd)):
                    apostropheEnd.append(i)
                    isApostrophe = False
                else:
                    apostropheBegin.append(i)
                    isApostrophe = True

        for i in range(len(whereStr)):
            self.customFormula.tag_add("tagZreal", "1.0 + " + str(whereStr[i]) + "c", "1.0 + " + str(whereStr[i] + 5) + "c")
            self.customFormula.tag_config("tagZreal", foreground="red")
        for i in range(len(whereFreq)):
            self.customFormula.tag_add("tagFreq", "1.0 + " + str(whereFreq[i]) + "c", "1.0 + " + str(whereFreq[i] + 4) + "c")
            self.customFormula.tag_config("tagFreq", foreground="red")
        for i in range(len(whereZ)):
            self.customFormula.tag_add("tagZ", "1.0 + " + str(whereZ[i]) + "c", "1.0 + " + str(whereZ[i] + 2) + "c")
            self.customFormula.tag_config("tagZ", foreground="red")
        if (self.weightingComboboxValue.get() == "Custom"):
            for i in range(len(whereWeight)):
                self.customFormula.tag_add("tagWeight", "1.0 + " + str(whereWeight[i]) + "c", "1.0 + " + str(whereWeight[i] + 9) + "c")
                self.customFormula.tag_config("tagWeight", foreground="red")
        for i in range(len(reservedWords2)):
            self.customFormula.tag_add("tagReservedWords2", "1.0 + " + str(reservedWords2[i]) + "c", "1.0 + " + str(reservedWords2[i] + 2) + "c")
            self.customFormula.tag_config("tagReservedWords2", foreground="DodgerBlue3")
        for i in range(len(reservedWords3)):
            self.customFormula.tag_add("tagReservedWords3", "1.0 + " + str(reservedWords3[i]) + "c", "1.0 + " + str(reservedWords3[i] + 3) + "c")
            self.customFormula.tag_config("tagReservedWords3", foreground="DodgerBlue3")
        for i in range(len(reservedWords4)):
            self.customFormula.tag_add("tagReservedWords4", "1.0 + " + str(reservedWords4[i]) + "c", "1.0 + " + str(reservedWords4[i] + 4) + "c")
            self.customFormula.tag_config("tagReservedWords4", foreground="DodgerBlue3")
        for i in range(len(reservedWords5)):
            self.customFormula.tag_add("tagReservedWords5", "1.0 + " + str(reservedWords5[i]) + "c", "1.0 + " + str(reservedWords5[i] + 5) + "c")
            self.customFormula.tag_config("tagReservedWords5", foreground="DodgerBlue3")
        for i in range(len(reservedWords6)):
            self.customFormula.tag_add("tagReservedWords6", "1.0 + " + str(reservedWords6[i]) + "c", "1.0 + " + str(reservedWords6[i] + 6) + "c")
            self.customFormula.tag_config("tagReservedWords6", foreground="DodgerBlue3")
        for i in range(len(reservedWords7)):
            self.customFormula.tag_add("tagReservedWords7", "1.0 + " + str(reservedWords7[i]) + "c", "1.0 + " + str(reservedWords7[i] + 7) + "c")
            self.customFormula.tag_config("tagReservedWords7", foreground="DodgerBlue3")
        for i in range(len(reservedWords8)):
            self.customFormula.tag_add("tagReservedWords8", "1.0 + " + str(reservedWords8[i]) + "c", "1.0 + " + str(reservedWords8[i] + 8) + "c")
            self.customFormula.tag_config("tagReservedWords8", foreground="DodgerBlue3")
        for i in range(len(builtInsBegin)):
            self.customFormula.tag_add("tagBuiltIns", "1.0 + " + str(builtInsBegin[i]) + "c", "1.0 + " + str(builtInsEnd[i]) + "c")
            self.customFormula.tag_config("tagBuiltIns", foreground="violet red")
        for i in range(len(userVariablesBegin)):
            self.customFormula.tag_add("tagUserVariables", "1.0 + " + str(userVariablesBegin[i]) + "c", "1.0 + " + str(userVariablesEnd[i]) + "c")
            self.customFormula.tag_config("tagUserVariables", foreground="red")
        for i in range(len(whereComment)):
            self.customFormula.tag_remove("tagZreal", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords2", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords3", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords4", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords5", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords6", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords7", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagReservedWords8", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagUserVariables", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagBuiltIns", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagQuotations", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagApostrophes", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagFreq", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagZ", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_remove("tagWeight", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_add("tagComment", "1.0 + " + str(whereComment[i]) + "c", "1.0 + " + str(whereComment[i]) + "c lineend")
            self.customFormula.tag_config("tagComment", foreground="gray60")
        for i in range(len(quotationBegin)):
            if (len(quotationBegin) > len(quotationEnd) and i == len(quotationBegin)-1):
                self.customFormula.tag_add("tagQuotations", "1.0 + " + str(quotationBegin[i]) + "c", "1.0 + " + str(len(textToSearch)) + "c")
            else:
                self.customFormula.tag_add("tagQuotations", "1.0 + " + str(quotationBegin[i]) + "c", "1.0 + " + str(quotationEnd[i]+1) + "c")
            self.customFormula.tag_config("tagQuotations", foreground="green")
        for i in range(len(apostropheBegin)):
            if (len(apostropheBegin) > len(apostropheEnd) and i == len(apostropheBegin)-1):
                self.customFormula.tag_add("tagApostrophes", "1.0 + " + str(apostropheBegin[i]) + "c", "1.0 + " + str(len(textToSearch)) + "c")
            else:
                self.customFormula.tag_add("tagApostrophes", "1.0 + " + str(apostropheBegin[i]) + "c", "1.0 + " + str(apostropheEnd[i]+1) + "c")
            self.customFormula.tag_config("tagApostrophes", foreground="green")
    
    def removeParam(self):
        if (self.numParams != 0):
            self.paramNameValues.pop()
            self.paramComboboxValues.pop()
            self.paramValueValues.pop()
            self.paramNameEntries.pop().grid_remove()
            self.paramComboboxes.pop().grid_remove()
            self.paramValueEntries.pop().grid_remove()
            self.paramNameLabels.pop().grid_remove()
            self.paramValueLabels.pop().grid_remove()
            self.paramDeleteButtons.pop().grid_remove()
            self.numParams -= 1
            self.howMany.configure(text="Number of parameters: " + str(self.numParams))
            self.keyup("")
    
    def deleteParam(self, paramNumber):
        if paramNumber == 0 and self.numParams == 1:
            self.removeParam()
        else:
            for i in range(paramNumber, self.numParams-1):
                self.paramNameValues[i].set(self.paramNameValues[i+1].get())
                self.paramComboboxValues[i].set(self.paramComboboxValues[i+1].get())
                self.paramValueValues[i].set(self.paramValueValues[i+1].get())
            self.removeParam()
    
    def formulaEnter(self, n):
        fname, fext = os.path.splitext(n)
        directory = os.path.dirname(str(n))
        self.topGUI.setCurrentDirectory(directory)
        if (fext == ".mmfile"):
            try:
                data = np.loadtxt(n)
                w_in = data[:,0]
                r_in = data[:,1]
                j_in = data[:,2]
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
                self.lengthOfData = len(self.wdata)
                self.freqRangeButton.configure(state="normal")
                self.browseEntry.configure(state="normal")
                self.browseEntry.delete(0,tk.END)
                self.browseEntry.insert(0, n)
                self.browseEntry.configure(state="readonly")
                self.runButton.configure(state="normal")
            except:
                messagebox.showerror("File error", "Error: There was an error loading or reading the file")
        elif (fext == ".mmcustom"):
            try:
                toLoad = open(n)
                fileToLoad = toLoad.readline().split("filename: ")[1][:-1]
                numberOfSimulations = int(toLoad.readline().split("number_simulations: ")[1][:-1])
                if (numberOfSimulations < 1):
                    raise ValueError
                fitType = toLoad.readline().split("fit_type: ")[1][:-1]
                if (fitType != "Real" and fitType != "Complex" and fitType != "Imaginary"):
                    raise ValueError
                alphaVal = float(toLoad.readline().split("alpha: ")[1][:-1])
                if (alphaVal <= 0):
                    raise ValueError
                uD = int(toLoad.readline().split("upDelete: ")[1][:-1])
                if (uD < 0):
                    raise ValueError
                lD = int(toLoad.readline().split("lowDelete: ")[1][:-1])
                if (lD < 0):
                    raise ValueError
                weightType = toLoad.readline().split("weighting: ")[1][:-1]
                if (weightType != "None" and weightType != "Modulus" and weightType != "Proportional" and weightType != "Error model" and weightType != "Custom"):
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
                if (errorBetaRe[0] != "n"):
                    errorBetaRe = float(errorBetaRe)
                    useBetaRe = True
                else:
                    errorBetaRe = float(errorBeta[1:])
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
                toLoad.readline() #Skip the "formula:"
                formulaIn = ""
                while True:
                    nextLineIn = toLoad.readline()
                    if "mmparams:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
                while True:
                    p = toLoad.readline()
                    if not p:
                        break
                    else:
                        p = p[:-1]
                        pname, pval, pcombo = p.split("~=~")
                        self.numParams += 1
                        self.paramNameValues.append(tk.StringVar(self, pname))
                        self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10))
                        self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                        self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                        self.paramValueValues.append(tk.StringVar(self, pval))
                        self.paramValueEntries.append(ttk.Entry(self.pframe, textvariable=self.paramValueValues[self.numParams-1], width=10))
                        self.paramDeleteButtons.append(ttk.Button(self.pframe, text="Delete", command=partial(self.deleteParam, self.numParams-1)))
                        self.paramNameLabels[len(self.paramNameLabels)-1].grid(column=0, row=self.numParams, pady=5)
                        self.paramNameEntries[len(self.paramNameEntries)-1].grid(column=1, row=self.numParams)
                        self.paramValueLabels[len(self.paramValueLabels)-1].grid(column=2, row=self.numParams, padx=(10, 3))
                        self.paramComboboxes[len(self.paramComboboxes)-1].grid(column=3, row=self.numParams)
                        self.paramValueEntries[len(self.paramValueEntries)-1].grid(column=4, row=self.numParams, padx=(3,0))
                        self.paramDeleteButtons[len(self.paramDeleteButtons)-1].grid(column=5, row=self.numParams, padx=3)
                        self.howMany.configure(text="Number of parameters: " + str(self.numParams))
                        self.paramNameLabels[len(self.paramNameLabels)-1].bind("<MouseWheel>", self._on_mousewheel)
                        self.paramNameEntries[len(self.paramNameEntries)-1].bind("<MouseWheel>", self._on_mousewheel)
                        self.paramValueLabels[len(self.paramValueLabels)-1].bind("<MouseWheel>", self._on_mousewheel)
                        self.paramValueEntries[len(self.paramValueEntries)-1].bind("<MouseWheel>", self._on_mousewheel)
                        self.paramDeleteButtons[len(self.paramDeleteButtons)-1].bind("<MouseWheel>", self._on_mousewheel)
                        self.paramNameEntries[len(self.paramNameEntries)-1].bind("<KeyRelease>", self.keyup)
                toLoad.close()
                data = np.loadtxt(fileToLoad)
                w_in = data[:,0]
                r_in = data[:,1]
                j_in = data[:,2]
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
                self.freqRangeButton.configure(state="normal")
                self.browseEntry.configure(state="normal")
                self.browseEntry.delete(0,tk.END)
                self.browseEntry.insert(0, n)
                self.browseEntry.configure(state="readonly")
                self.runButton.configure(state="normal")
                self.monteCarloValue.set(str(numberOfSimulations))
                self.fittingTypeComboboxValue.set(str(fitType))
                self.weightingComboboxValue.set(weightType)
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
                if (self.weightingComboboxValue.get() == "None"):
                    self.noiseFrame.grid_remove()
                    self.errorStructureFrame.grid_remove()
                elif (self.weightingComboboxValue.get() == "Error model"):
                    self.noiseFrame.grid_remove()
                    self.errorStructureFrame.grid(column=0, row=2, pady=(5,0), columnspan=8)
                elif (self.weightingComboboxValue.get() == "Custom"):
                    self.noiseFrame.grid_remove()
                    self.errorStructureFrame.grid_remove()
                else:
                    self.noiseFrame.grid(column=3, row=1, pady=(5,0), sticky="W")
                    self.errorStructureFrame.grid_remove()
                self.browseEntry.configure(state="normal")
                self.browseEntry.delete(0,tk.END)
                self.browseEntry.insert(0, fileToLoad)
                self.browseEntry.configure(state="readonly")
                self.customFormula.delete("1.0", tk.END)
                self.customFormula.insert("1.0", formulaIn[:-2])    #Insert custom formula into code box and remove last two new line characters (they come from the saving format)
                self.keyup("")
            except:
                messagebox.showerror("File error", "Error: There was an error loading or reading the file")
        else:
            messagebox.showerror("File error", "Error: The file has an unknown extension")
    
    def setThemeLight(self):
        self.backgroundColor = "#FFFFFF"
        self.foregroundColor = "#000000"
        self.configure(background="#FFFFFF")
        self.customFunctionFrame.configure(background="#FFFFFF")
        self.fittingButtonFrame.configure(background="#FFFFFF")
        self.codeLabel.configure(background="#FFFFFF", foreground="#000000")
        self.helpLabel.configure(background="#FFFFFF", foreground="blue")
        self.runFrame.configure(background="#FFFFFF")
        self.pframe.configure(background="#FFFFFF")
        self.pcanvas.configure(background="#FFFFFF")
        self.helpPopup.configure(background="#FFFFFF")
        self.hframe.configure(background="#FFFFFF")
        self.hcanvas.configure(background="#FFFFFF")
        self.helpTextLabel.configure(background="#FFFFFF", foreground="#000000")
        self.fittingTypeLabel.configure(background="#FFFFFF", foreground="#000000")
        self.monteCarloLabel.configure(background="#FFFFFF", foreground="#000000")
        self.resultsFrame.configure(background="#FFFFFF")
        self.graphFrame.configure(background="#FFFFFF")
        self.includeFrame.configure(background="#FFFFFF")
        self.browseFrame.configure(background="#FFFFFF")
        self.browseLabel.configure(background="#FFFFFF", foreground="#000000")
        self.weightingLabel.configure(background="#FFFFFF", foreground="#000000")
        self.noiseLabel.configure(background="#FFFFFF", foreground="#000000")
        self.noiseFrame.configure(background="#FFFFFF")
        self.freqWindow.configure(background="#FFFFFF")
        self.rs.configure(background="#FFFFFF")
        self.rs.configure(tickColor="#000000")
        self.howMany.configure(background="#FFFFFF", foreground="#000000")
        self.paramPopup.configure(background="#FFFFFF")
        self.saveFrame.configure(background="#FFFFFF")
        self.lowestUndeleted.configure(background="#FFFFFF", foreground="#000000")
        self.highestUndeleted.configure(background="#FFFFFF", foreground="#000000")
        for paramName in self.paramNameLabels:
            paramName.configure(background="#FFFFFF", foreground="#000000")
        for paramVal in self.paramValueLabels:
            paramVal.configure(background="#FFFFFF", foreground="#000000")
        try:
            self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
            self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
            self.rs.setLower(np.log10(min(self.wdata)))
            self.rs.setUpper(np.log10(max(self.wdata)))
        except:
            pass
        try:
            self.upperLabel.configure(background="#FFFFFF", foreground="#000000")
            self.lowerLabel.configure(background="#FFFFFF", foreground="#000000")
            self.rangeLabel.configure(background="#FFFFFF", foreground="#000000")
        except:
            pass
                                 
    def setThemeDark(self):
        self.backgroundColor = "#424242"
        self.foregroundColor = "#FFFFFF"
        self.configure(background="#424242")
        self.customFunctionFrame.configure(background="#424242")
        self.fittingButtonFrame.configure(background="#424242")
        self.codeLabel.configure(background="#424242", foreground="#FFFFFF")
        self.helpLabel.configure(background="#424242", foreground="skyblue")
        self.runFrame.configure(background="#424242")
        self.pframe.configure(background="#424242")
        self.pcanvas.configure(background="#424242")
        self.helpPopup.configure(background="#424242")
        self.hframe.configure(background="#424242")
        self.hcanvas.configure(background="#424242")
        self.helpTextLabel.configure(background="#424242", foreground="#FFFFFF")
        self.fittingTypeLabel.configure(background="#424242", foreground="#FFFFFF")
        self.monteCarloLabel.configure(background="#424242", foreground="#FFFFFF")
        self.resultsFrame.configure(background="#424242")
        self.graphFrame.configure(background="#424242")
        self.includeFrame.configure(background="#424242")
        self.browseFrame.configure(background="#424242")
        self.browseLabel.configure(background="#424242", foreground="#FFFFFF")
        self.weightingLabel.configure(background="#424242", foreground="#FFFFFF")
        self.noiseLabel.configure(background="#424242", foreground="#FFFFFF")
        self.noiseFrame.configure(background="#424242")
        self.freqWindow.configure(background="#424242")
        self.lowestUndeleted.configure(background="#424242", foreground="#FFFFFF")
        self.highestUndeleted.configure(background="#424242", foreground="#FFFFFF")
        self.howMany.configure(background="#424242", foreground="#FFFFFF")
        self.paramPopup.configure(background="#424242")
        for paramName in self.paramNameLabels:
            paramName.configure(background="#424242", foreground="#FFFFFF")
        for paramVal in self.paramValueLabels:
            paramVal.configure(background="#424242", foreground="#FFFFFF")
        self.rs.configure(background="#424242")
        self.rs.configure(tickColor="#FFFFFF")
        try:
            self.rs.setLowerBound((np.log10(min(self.wdataRaw))))
            self.rs.setUpperBound((np.log10(max(self.wdataRaw))))
            self.rs.setLower(np.log10(min(self.wdata)))
            self.rs.setUpper(np.log10(max(self.wdata)))
        except:
            pass
        try:
            self.upperLabel.configure(background="#424242", foreground="#FFFFFF")
            self.lowerLabel.configure(background="#424242", foreground="#FFFFFF")
            self.rangeLabel.configure(background="#424242", foreground="#FFFFFF")
        except:
            pass
    
    def setEllipseColor(self, color):
        self.ellipseColor = color
    
    def needToSave(self):
        return self.saveNeeded
    
    def cancelThreads(self):
        for t in self.currentThreads:
            t.terminate()