# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 10:29:23 2019

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
import tkinter.font as tkfont
from tkinter import messagebox
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename, asksaveasfile, askdirectory
from tkinter import simpledialog
import os
import sys
import ctypes
import comtypes.client as cc
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import customFitting
import bootstrapFitting
import simplexFitting
import dataSimulation
import threading
import queue
#import time
#from rangeSlider import RangeSlider
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
import webbrowser
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
    def __init__(self, queue, eI, lp, fitType, numMonte, weight, noise, customFormula, w, r, j, paramNames, paramGuesses, paramLimits, errorParams):
        threading.Thread.__init__(self)
        self.queue = queue
        self.extraI = eI
        self.listPer = lp
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
            r, s, sdr, sdi, chi, aic, rF, iF, bP, cW = self.fittingObject.findFit(self.extraI, self.listPer, self.fT, self.nM, self.wght, self.aN, self.formula, self.wdat, self.rdat, self.jdat, self.pN, self.pG, self.pL, self.eP)
            self.queue.put((r, s, sdr, sdi, chi, aic, rF, iF, bP, cW))
        except Exception as inst:
            print(inst)
            self.queue.put(("@", "@", "@", "@", "@", "@", "@", "@", "@"))
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

class ThreadedTaskBootstrap(threading.Thread):
    def __init__(self, queue, eI, lp, nBS, r_in, s_in, sdR_in, sdI_in, chi_in, aic_in, realF_in, imagF_in, fitType, numMonte, weight, noise, customFormula, w, r, j, paramNames, paramGuesses, paramLimits, errorParams):
        threading.Thread.__init__(self)
        self.queue = queue
        self.extraI = eI
        self.percent_list = lp
        self.nBootStrap = nBS
        self.in_r = r_in
        self.in_s = s_in
        self.in_sdR = sdR_in
        self.in_sdI = sdI_in
        self.in_chi = chi_in
        self.in_aic = aic_in
        self.in_realF = realF_in
        self.in_imagF = imagF_in
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
        self.fittingObject = bootstrapFitting.bootstrapFitter()
    def run(self):
        try:
            r, s, sdr, sdi, chi, aic, rF, iF, cW = self.fittingObject.findFit(self.extraI, self.percent_list, self.nBootStrap, self.in_r, self.in_s, self.in_sdR, self.in_sdI, self.in_chi, self.in_aic, self.in_realF, self.in_imagF, self.fT, self.nM, self.wght, self.aN, self.formula, self.wdat, self.rdat, self.jdat, self.pN, self.pG, self.pL, self.eP)
            self.queue.put((r, s, sdr, sdi, chi, aic, rF, iF, cW))
        except Exception:
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
        # Creates a toplevel window
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
        
        #From David Hasselberg, https://stackoverflow.com/questions/36219613/make-tkinter-prompt-inherit-parent-windows-icon
        class StringDialog(simpledialog._QueryString):  #Modification of simpledialog to include an icon
            def body(self, master):
                super().body(master)
                self.iconbitmap(resource_path('img/elephant3.ico'))
        
            def ask_string(title, prompt, **kargs):
                d = StringDialog(title, prompt, **kargs)
                return d.result
        
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
        self.lowestUndeleted = tk.Text(self.freqWindow, height=1, width=40, borderwidth=0, bg=self.backgroundColor, fg=self.foregroundColor) #tk.Label(self.freqWindow, text="Lowest Remaining Frequency: 0", background=self.backgroundColor, foreground=self.foregroundColor)
        self.lowestUndeleted.configure(inactiveselectbackground=self.lowestUndeleted.cget("selectbackground"))
        self.lowestUndeleted.insert(1.0, "Lowest Remaining Frequency: 0")
        self.lowestUndeleted.configure(state="disabled")
        self.highestUndeleted = tk.Text(self.freqWindow, height=1, width=40, borderwidth=0, bg=self.backgroundColor, fg=self.foregroundColor)
        self.highestUndeleted.configure(inactiveselectbackground=self.highestUndeleted.cget("selectbackground"))
        self.highestUndeleted.insert(1.0, "Lowest Remaining Frequency: 0")
        self.highestUndeleted.configure(state="disabled")
        self.lowerSpinboxVariable = tk.StringVar(self, "0")
        self.upperSpinboxVariable = tk.StringVar(self, "0")
        self.upDelete = 0
        self.lowDelete = 0
        
        def popup_formula_description(event):
            try:
                self.popup_menu_description.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu_description.grab_release()
        
        def copyFormula_description():
            try:
                pyperclip.copy(self.formulaDescriptionEntry.get(tk.SEL_FIRST, tk.SEL_LAST))
            except Exception:
                pass
        
        def cutFormula_description():
            try:
                pyperclip.copy(self.formulaDescriptionEntry.get(tk.SEL_FIRST, tk.SEL_LAST))
                self.formulaDescriptionEntry.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except Exception:
                pass
        
        def pasteFormula_description():
            try:
                toPaste = pyperclip.paste()
                self.formulaDescriptionEntry.insert(tk.INSERT, toPaste)
                keyup(None)
            except Exception:
                pass
            
        def undoFormula_description():
            try:
                self.formulaDescriptionEntry.edit_undo()
            except Exception:
                pass
            
        def redoFormula_description():
            try:
                self.formulaDescriptionEntry.edit_redo()
            except Exception:
                pass
        
        self.importPathWindow = tk.Toplevel(bg=self.backgroundColor)
        self.importPathWindow.withdraw()
        self.importPathWindow.geometry("750x400")
        self.importPathWindow.title("Python import paths")
        self.importPathWindow.iconbitmap(resource_path('img/elephant3.ico'))
        self.importPathLabel = tk.Label(self.importPathWindow, text="Paths to search: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.importPathListboxFrame = tk.Frame(self.importPathWindow, bg=self.backgroundColor)
        self.importPathListboxScrollbar = ttk.Scrollbar(self.importPathListboxFrame, orient=tk.VERTICAL)
        self.importPathListboxScrollbarHorizontal = ttk.Scrollbar(self.importPathListboxFrame, orient=tk.HORIZONTAL)
        self.importPathListbox = tk.Listbox(self.importPathListboxFrame, height=3, width=65, selectmode=tk.BROWSE, activestyle='none', yscrollcommand=self.importPathListboxScrollbar.set, xscrollcommand=self.importPathListboxScrollbarHorizontal.set)
        self.importPathListboxScrollbar['command'] = self.importPathListbox.yview
        self.importPathListboxScrollbarHorizontal['command'] = self.importPathListbox.xview
        self.importPathButtonFrame = tk.Frame(self.importPathWindow, bg=self.backgroundColor)
        self.importPathButton = ttk.Button(self.importPathButtonFrame, text="Browse...")
        self.importPathRemoveButton = ttk.Button(self.importPathButtonFrame, text="Remove")
        self.importPathButton.pack(side=tk.TOP, fill=tk.NONE, expand=False)
        self.importPathRemoveButton.pack(side=tk.TOP, fill=tk.NONE, expand=False, pady=5)
        self.importPathListboxScrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=False)
        self.importPathListboxScrollbarHorizontal.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.importPathListbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.importPathLabel.pack(side=tk.TOP, anchor=tk.W, fill=tk.NONE, pady=3, expand=False)
        self.importPathListboxFrame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.importPathButtonFrame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5, expand=False)
        importPathButton_ttp = CreateToolTip(self.importPathButton, "Browse for directories to add to the Python module search path")
        importPathRemove_ttp = CreateToolTip(self.importPathRemoveButton, "Remove selected path")
        
        def removeFile_importPath(e=None):
            if (self.importPathListbox.size() > 0):
                 items = list(map(int, self.importPathListbox.curselection()))
                 if (len(items) != 0):
                     self.importPathListbox.delete(tk.ANCHOR)
        self.importPathRemoveButton.configure(command=removeFile_importPath)
        self.popup_menu_importPath = tk.Menu(self.importPathListbox, tearoff=0)
        self.popup_menu_importPath.add_command(label="Remove directory", command=removeFile_importPath)
        def popup_inputFiles(event):
            try:
                self.popup_menu_importPath.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu_importPath.grab_release()
        self.importPathListbox.bind("<Button-3>", popup_inputFiles)
        for val in self.topGUI.getDefaultImports():
            if (len(val) > 0):
                self.importPathListbox.insert(tk.END, val)
        
        self.simulationWindow = tk.Toplevel(bg=self.backgroundColor)
        self.simulationWindow.withdraw()
        self.simulationWindow.title("Synthetic Data")
        self.simulationWindow.iconbitmap(resource_path('img/elephant3.ico'))
        self.simulationFreqFrame = tk.Frame(self.simulationWindow, bg=self.backgroundColor)
        self.simulationLowerFreqLabel = tk.Label(self.simulationFreqFrame, text="Lower frequency: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.simulationLowerFreq = ttk.Entry(self.simulationFreqFrame)
        self.simulationLowerFreq.insert(0, "1E-5")
        self.simulationUpperFreqLabel = tk.Label(self.simulationFreqFrame, text="Upper frequency: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.simulationUpperFreq = ttk.Entry(self.simulationFreqFrame)
        self.simulationUpperFreq.insert(0, "1E5")
        self.simulationNumFreqLabel = tk.Label(self.simulationFreqFrame, text="Points per decade: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.simulationNumFreq = ttk.Entry(self.simulationFreqFrame)
        self.simulationNumFreq.insert(0, "10")
        self.numberOfPointsLabel = tk.Label(self.simulationWindow, text="Total number of points: 100", bg=self.backgroundColor, fg=self.foregroundColor)
        self.runSimulationButton = ttk.Button(self.simulationWindow, width=40, text="Generate synthetic data")
        self.simViewFrame = tk.Frame(self.simulationWindow, bg=self.backgroundColor)
        self.simViewScrollbar = ttk.Scrollbar(self.simViewFrame, orient=tk.VERTICAL)
        self.simView = ttk.Treeview(self.simViewFrame, columns=("freq", "real", "imag"), height=10, selectmode="browse", yscrollcommand=self.simViewScrollbar.set)
        self.simView.heading("freq", text="Frequency")
        self.simView.heading("real", text="Real part")
        self.simView.heading("imag", text="Imaginary part")
        self.simView.column("#0", width=50)
        self.simView.column("freq", width=150)
        self.simView.column("imag", width=150)
        self.simView.column("real", width=150)
        self.simViewScrollbar['command'] = self.simView.yview
        self.saveSimulationFrame = tk.Frame(self.simulationWindow, bg=self.backgroundColor)
        self.saveSimulationButton = ttk.Button(self.saveSimulationFrame, text="Save")
        self.copySimulationButton = ttk.Button(self.saveSimulationFrame, text="Copy values as spreadsheet")
        self.plotSimulationButton = ttk.Button(self.saveSimulationFrame, text="Plot")
        self.plotSimulationCheckboxVariable = tk.IntVar(self, 1)
        self.plotSimulationCheckbox = ttk.Checkbutton(self.saveSimulationFrame, text="Mouseover labels", variable=self.plotSimulationCheckboxVariable)
        self.simView.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.simViewScrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=False)
        self.simulationLowerFreqLabel.pack(side=tk.LEFT, expand=False)
        self.simulationLowerFreq.pack(side=tk.LEFT, fill=tk.X, padx=5, expand=True)
        self.simulationUpperFreqLabel.pack(side=tk.LEFT, expand=False)
        self.simulationUpperFreq.pack(side=tk.LEFT, fill=tk.X, padx=5, expand=True)
        self.simulationNumFreqLabel.pack(side=tk.LEFT, expand=False)
        self.simulationNumFreq.pack(side=tk.LEFT, fill=tk.X, padx=(5, 0), expand=True)
        self.simulationFreqFrame.pack(side=tk.TOP, fill=tk.X, pady=5, padx=3)
        self.numberOfPointsLabel.pack(side=tk.TOP, fill=tk.NONE, anchor=tk.W, pady=5)
        self.runSimulationButton.pack(side=tk.TOP, fill=tk.NONE, pady=5, padx=3, expand=True)
        self.saveSimulationButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.copySimulationButton.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3)
        self.plotSimulationButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.plotSimulationCheckbox.pack(side=tk.LEFT, fill=tk.X, padx=(3, 0), expand=True)
        runSimulationButton_ttp = CreateToolTip(self.runSimulationButton, 'Generate synthetic data')
        saveSimulationButton_ttp = CreateToolTip(self.saveSimulationButton, 'Save generated data')
        copySimulationButton_ttp = CreateToolTip(self.copySimulationButton, 'Copy values in a format to transfer to a spreadsheet')
        plotSimulationButton_ttp = CreateToolTip(self.plotSimulationButton, 'Plot the synthetic data')
        plotSimulationCheckbox_ttp = CreateToolTip(self.plotSimulationCheckbox, 'Include mouseover labels when plotting')
        
        self.formulaDescriptionWindow = tk.Toplevel(bg=self.backgroundColor)
        self.formulaDescriptionWindow.withdraw()
        self.formulaDescriptionWindow.geometry("1000x700")
        self.formulaDescriptionWindow.title("Custom formula description")
        self.formulaDescriptionWindow.iconbitmap(resource_path('img/elephant3.ico'))
        self.formulaDescriptionLabel = tk.Label(self.formulaDescriptionWindow, text="Description: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.formulaDescriptionFrame = tk.Frame(self.formulaDescriptionWindow, bg=self.backgroundColor)
        self.formulaDescriptionContainer = tk.Frame(self.formulaDescriptionFrame, borderwidth=1, relief="sunken")
        self.formulaDescriptionEntry = tk.Text(self.formulaDescriptionContainer, width=60, height=4, wrap="none", borderwidth=0, undo=True, fg=self.foregroundColor)
        if (self.topGUI.getTheme() == "dark"):
            self.formulaDescriptionEntry.configure(background="#6B6B6B")
        else:
            self.formulaDescriptionEntry.configure(background="#FFFFFF")
        self.formulaDescriptionVertical = tk.Scrollbar(self.formulaDescriptionContainer, orient="vertical", command=self.formulaDescriptionEntry.yview)
        self.formulaDescriptionHorizontal = tk.Scrollbar(self.formulaDescriptionContainer, orient="horizontal", command=self.formulaDescriptionEntry.xview)
        self.formulaDescriptionEntry.configure(yscrollcommand=self.formulaDescriptionVertical.set, xscrollcommand=self.formulaDescriptionHorizontal.set)
        self.formulaLatexFrame = tk.Frame(self.formulaDescriptionWindow, bg=self.backgroundColor)
        self.formulaDescriptionLatexLabel = tk.Label(self.formulaLatexFrame, text="Equation (Latex math): ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.formulaDescriptionLatexEntry = scrolledtext.ScrolledText(self.formulaLatexFrame, height=4)
        self.formulaDescriptionVertical.pack(side=tk.RIGHT, fill=tk.Y, expand=False)
        self.formulaDescriptionHorizontal.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.formulaDescriptionEntry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.formulaDescriptionLabel.pack(side=tk.TOP, anchor=tk.W)
        self.formulaDescriptionContainer.pack(side=tk.TOP, fill=tk.BOTH, expand=False)
        self.formulaDescriptionFrame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=5)
        self.formulaDescriptionLatexLabel.pack(side=tk.LEFT)
        self.formulaDescriptionLatexEntry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.latexFig = Figure(figsize=(6, 2), dpi=100)
        self.latexFig.set_facecolor(self.backgroundColor)
        self.latexAx = self.latexFig.add_subplot(111)
        self.latexCanvas = FigureCanvasTkAgg(self.latexFig, master=self.formulaDescriptionWindow)
        self.latexCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.X, expand=True)
        self.latexCanvas._tkcanvas.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        self.latexAx.get_xaxis().set_visible(False)
        self.latexAx.get_yaxis().set_visible(False)
        self.latexFig.tight_layout()
        self.latexAx.axis("off")
        self.formulaLatexFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=False, pady=15, padx=5)
        self.popup_menu_description = tk.Menu(self.formulaDescriptionEntry, tearoff=0)
        self.popup_menu_description.add_command(label="Copy", command=copyFormula_description)
        self.popup_menu_description.add_command(label="Cut", command=cutFormula_description)
        self.popup_menu_description.add_command(label="Paste", command=pasteFormula_description)
        self.popup_menu_description.add_separator()
        self.popup_menu_description.add_command(label="Undo", command=undoFormula_description)
        self.popup_menu_description.add_command(label="Redo", command=redoFormula_description)
        self.formulaDescriptionEntry.bind("<Button-3>", popup_formula_description)
        
        self.loadFormulaWindow = tk.Toplevel(bg=self.backgroundColor)
        self.loadFormulaWindow.withdraw()
        self.loadFormulaWindow.geometry("1350x750")
        self.loadFormulaWindow.title("Load formula")
        self.loadFormulaWindow.iconbitmap(resource_path('img/elephant3.ico'))
        self.loadFormulaFrame = tk.Frame(self.loadFormulaWindow, bg=self.backgroundColor)
        self.loadFormulaDirectoryFrame = tk.Frame(self.loadFormulaFrame, bg=self.backgroundColor)
        self.loadFormulaDirectoryEntry = ttk.Entry(self.loadFormulaDirectoryFrame, width=40, state="normal")
        self.loadFormulaDirectoryEntry.insert(0, self.topGUI.getDefaultFormulaDirectory())
        self.loadFormulaDirectoryEntry.configure(state="readonly")
        self.loadFormulaDirectoryLabel = tk.Label(self.loadFormulaDirectoryFrame, text="Directory: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.loadFormulaDirectoryButton = ttk.Button(self.loadFormulaDirectoryFrame, text="Browse..")
        self.loadFormulaDirectoryLabel.pack(side=tk.LEFT, expand=False)
        self.loadFormulaDirectoryEntry.pack(side=tk.LEFT, fill=tk.X, padx=3, expand=True)
        self.loadFormulaDirectoryButton.pack(side=tk.RIGHT, expand=False)
        self.loadFormulaNodes = dict()
        self.loadFormulaFrameTwo = tk.Frame(self.loadFormulaFrame)
        self.loadFormulaTree = ttk.Treeview(self.loadFormulaFrameTwo, selectmode="browse")
        self.loadFormulaYSB = ttk.Scrollbar(self.loadFormulaFrameTwo, orient='vertical', command=self.loadFormulaTree.yview)
        self.loadFormulaXSB = ttk.Scrollbar(self.loadFormulaFrameTwo, orient='horizontal', command=self.loadFormulaTree.xview)
        self.loadFormulaTree.configure(yscroll=self.loadFormulaYSB.set, xscroll=self.loadFormulaXSB.set)
        self.loadFormulaXSB.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.loadFormulaTree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.loadFormulaYSB.pack(side=tk.RIGHT, fill=tk.Y, expand=False)
        self.loadFormulaDirectoryFrame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5), expand=False)
        self.loadFormulaFrameTwo.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        abspath = os.path.abspath(self.topGUI.getDefaultFormulaDirectory())
        self.insert_node('', abspath, abspath)
        self.loadFormulaTree.bind('<<TreeviewOpen>>', self.open_node)
        self.loadFormulaDescriptionFrame = tk.Frame(self.loadFormulaWindow, bg=self.backgroundColor)
        self.loadFormulaDescriptionLabel = tk.Label(self.loadFormulaDescriptionFrame, text="Description: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.loadFormulaDescription = scrolledtext.ScrolledText(self.loadFormulaDescriptionFrame, height=3, fg=self.foregroundColor, state="disabled")
        if (self.topGUI.getTheme() == "dark"):
            self.loadFormulaDescription.configure(background="#6B6B6B")
        else:
            self.loadFormulaDescription.configure(background="#FFFFFF")
        self.loadFormulaCodeLabel = tk.Label(self.loadFormulaDescriptionFrame, text="Code: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.loadFormulaCode = scrolledtext.ScrolledText(self.loadFormulaDescriptionFrame, height=10, fg=self.foregroundColor, state="disabled")
        if (self.topGUI.getTheme() == "dark"):
            self.loadFormulaCode.configure(background="#6B6B6B")
        else:
            self.loadFormulaCode.configure(background="#FFFFFF")
        self.loadFormulaEquationLabel = tk.Label(self.loadFormulaDescriptionFrame, text="Equation: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.load_latexFig = Figure(figsize=(6, 2), dpi=100)
        self.load_latexFig.set_facecolor(self.backgroundColor)
        self.load_latexAx = self.load_latexFig.add_subplot(111)
        self.load_latexCanvas = FigureCanvasTkAgg(self.load_latexFig, master=self.loadFormulaDescriptionFrame)
        self.load_latexCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.X, expand=True)
        self.load_latexCanvas._tkcanvas.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        self.loadFormulaDescriptionLabel.pack(side=tk.TOP, expand=False, anchor=tk.W)
        self.loadFormulaDescription.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.loadFormulaCodeLabel.pack(side=tk.TOP, expand=False, anchor=tk.W)
        self.loadFormulaCode.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.loadFormulaEquationLabel.pack(side=tk.TOP, expand=False, anchor=tk.W)
        self.load_latexAx.get_xaxis().set_visible(False)
        self.load_latexAx.get_yaxis().set_visible(False)
        self.load_latexFig.tight_layout()
        self.load_latexAx.axis("off")
        self.loadFormulaButtonFrame = tk.Frame(self.loadFormulaWindow, bg=self.backgroundColor)
        self.loadFormula = ttk.Button(self.loadFormulaButtonFrame, width=20, text="Load Formula", state="disabled", command=self.loadFormulaCommand)
        self.loadFormulaIncludeLabel = tk.Label(self.loadFormulaButtonFrame, text="Include: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.loadFormulaParamsVariable = tk.IntVar(self, 1)
        self.loadFormulaParams = ttk.Checkbutton(self.loadFormulaButtonFrame, variable=self.loadFormulaParamsVariable, text="Parameters")
        self.loadFormulaCodeVariable = tk.IntVar(self, 1)
        self.loadFormulaCodeCheck = ttk.Checkbutton(self.loadFormulaButtonFrame, variable=self.loadFormulaCodeVariable, text="Code")
        self.loadFormulaOtherVariable = tk.IntVar(self, 0)
        self.loadFormulaOther = ttk.Checkbutton(self.loadFormulaButtonFrame, variable=self.loadFormulaOtherVariable, text="Other fitting settings")
        self.loadFormulaOther.grid(column=5, row=0)
        self.loadFormulaCodeCheck.grid(column=4, row=0)
        self.loadFormulaParams.grid(column=2, row=0)
        self.loadFormulaIncludeLabel.grid(column=1, row=0, padx=5)
        self.loadFormula.grid(column=0, row=0)
        self.loadFormulaButtonFrame.pack(side=tk.BOTTOM, anchor=tk.CENTER, fill=tk.NONE, pady=5, expand=False)
        self.loadFormulaFrame.pack(side=tk.LEFT, fill=tk.BOTH, pady=5, padx=(0,5), expand=True)
        self.loadFormulaDescriptionFrame.pack(side=tk.RIGHT, fill=tk.BOTH, pady=5, expand=True)
        self.loadFormulaTree.bind("<ButtonRelease-1>", self.clickFormula)
        loadFormulaDirectoryButton_ttp = CreateToolTip(self.loadFormulaDirectoryButton, 'Look for directory containing .mmformula files')
        loadFormula_ttp = CreateToolTip(self.loadFormula, 'Load formula to main window')
        loadFormulaParams_ttp = CreateToolTip(self.loadFormulaParams, 'Include parameters (names and values)')
        loadFormulaCode_ttp = CreateToolTip(self.loadFormulaCodeCheck, 'Include formula')
        loadFormulaOther_ttp = CreateToolTip(self.loadFormulaOther, 'Include other fitting settings (number of simulations, fit type, and weighting)')
        
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
        self.bottomButtonsFrame = tk.Frame(self.paramPopup, bg=self.backgroundColor)
        self.advancedOptionsButton = ttk.Button(self.bottomButtonsFrame, text="Advanced options")
        advancedOptions_ttp = CreateToolTip(self.advancedOptionsButton, 'Opens a popup containing advanced parameter settings')
        self.simplexButton = ttk.Button(self.bottomButtonsFrame, text="Step-by-step simplex", state="disabled")
        simplexButton_ttp = CreateToolTip(self.simplexButton, 'Perform one interation of a simplex algorithm (will replace current parameter values)')
        self.addButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.removeButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.advancedOptionsButton.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.simplexButton.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bottomButtonsFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.howMany.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.buttonFrame.pack(side=tk.BOTTOM, fill=tk.X, expand=False, pady=1)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.pcanvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.pcanvas.create_window((4,4), window=self.pframe, anchor="nw")
        self.advancedOptionsWindow = tk.Toplevel(bg=self.backgroundColor)
        self.advancedOptionsWindow.withdraw()
        self.advancedOptionsWindow.title("Advanced parameter options")
        self.advancedOptionsWindow.iconbitmap(resource_path("img/elephant3.ico"))
        self.advancedOptionsWindow.geometry("500x400")
        self.paramListbox = tk.Listbox(self.advancedOptionsWindow, selectmode=tk.BROWSE, activestyle='none', background=self.backgroundColor, foreground=self.foregroundColor, exportselection=False)
        self.paramListboxScrollbar = ttk.Scrollbar(self.advancedOptionsWindow, orient=tk.VERTICAL)
        self.paramListboxScrollbar['command'] = self.paramListbox.yview
        self.paramListbox.configure(yscrollcommand=self.paramListboxScrollbar.set)
        self.advancedOptionsLabel = tk.Label(self.advancedOptionsWindow, text="", bg=self.backgroundColor, fg=self.foregroundColor, font=('Helvetica', 10, 'bold'))
        self.advancedUpperLimitFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.advancedUpperLimitLabel = tk.Label(self.advancedUpperLimitFrame, text="Upper limit: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.advancedUpperLimitEntry = ttk.Entry(self.advancedUpperLimitFrame)
        self.advancedUpperLimitLabel.pack(side=tk.LEFT)
        self.advancedUpperLimitEntry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.advancedLowerLimitFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.advancedLowerLimitLabel = tk.Label(self.advancedLowerLimitFrame, text="Lower limit: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.advancedLowerLimitEntry = ttk.Entry(self.advancedLowerLimitFrame)
        self.advancedLowerLimitLabel.pack(side=tk.LEFT)
        self.advancedLowerLimitEntry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.multistartCheckbox = ttk.Checkbutton(self.advancedOptionsWindow, text="Enable multistart")
        self.multistartSpacingFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.multistartSpacing = ttk.Combobox(self.multistartSpacingFrame, values=("Logarithmic", "Linear", "Random", "Custom"), state="disabled", exportselection=False)
        self.multistartSpacingLabel = tk.Label(self.multistartSpacingFrame, text="Point spacing: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.multistartSpacingLabel.pack(side=tk.LEFT)
        self.multistartSpacing.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.multistartLowerFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.multistartLowerLabel = tk.Label(self.multistartLowerFrame, text="Lower bound: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.multistartLowerEntry = ttk.Entry(self.multistartLowerFrame, state="disabled")
        self.multistartLowerLabel.pack(side=tk.LEFT)
        self.multistartLowerEntry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.multistartUpperFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.multistartUpperLabel = tk.Label(self.multistartUpperFrame, text="Upper bound: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.multistartUpperEntry = ttk.Entry(self.multistartUpperFrame, state="disabled")
        self.multistartUpperLabel.pack(side=tk.LEFT)
        self.multistartUpperEntry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.multistartNumberFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.multistartNumberLabel = tk.Label(self.multistartNumberFrame, text="Number of points: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.multistartNumberEntry = ttk.Entry(self.multistartNumberFrame, state="disabled")
        self.multistartNumberLabel.pack(side=tk.LEFT)
        self.multistartNumberEntry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.multistartCustomFrame = tk.Frame(self.advancedOptionsWindow, bg=self.backgroundColor)
        self.multistartCustomLabel = tk.Label(self.multistartCustomFrame, text="Enter values separated by commas:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.multistartCustomEntry = ttk.Entry(self.multistartCustomFrame)
        self.multistartCustomLabel.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.multistartCustomEntry.pack(side=tk.TOP, fill=tk.X, expand=True)
        self.paramListbox.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        self.paramListboxScrollbar.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        self.advancedOptionsLabel.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.upperLimits = []
        self.lowerLimits = []
        self.multistartCheckboxVariables = []
        self.multistartSpacingVariables = []
        self.multistartLowerVariables = []
        self.multistartUpperVariables = []
        self.multistartNumberVariables = []
        self.multistartCustomVariables = []
        self.paramListboxSelection = 0
        self.varNum = 0
        self.whatFit = "C"
        self.aR = ""
        self.bootstrapEntryVariable = tk.StringVar(self, "500")
        self.bootstrapRunning = False
        self.currentThreads = []
        self.listPercent = []
        self.bestParams = []
        self.latexIn = ""
        self.resultsWindows = []
        self.resultPlotBigs = []
        self.resultPlotBigFigs = []
        self.resultPlots = []
        self.resultPlotFigs = []
        self.saveNyCanvasButtons = []
        self.saveNyCanvasButton_ttps = []
        self.simPlots = []
        self.simPlotFigs = []
        self.simPlotBigs = []
        self.simPlotBigFigs = []
        self.simSaveNyCanvasButtons = []
        self.simSaveNyCanvasButton_ttps = []
        self.loadedFormula = ""
        
        def checkUpperAndLower(e=None):
            try:
                if (self.upperLimits[self.paramListboxSelection].get() == "0" and self.lowerLimits[self.paramListboxSelection].get() == "-inf"):
                    self.paramComboboxValues[self.paramListboxSelection].set("-")
                elif (self.upperLimits[self.paramListboxSelection].get() == "inf" and self.lowerLimits[self.paramListboxSelection].get() == "-inf"):
                    self.paramComboboxValues[self.paramListboxSelection].set("+ or -")
                elif (self.upperLimits[self.paramListboxSelection].get() == "inf" and self.lowerLimits[self.paramListboxSelection].get() == "0"):
                    self.paramComboboxValues[self.paramListboxSelection].set("+")
                else:
                    self.paramComboboxValues[self.paramListboxSelection].set("custom")
            except Exception:
                pass    #If it fails, it isn't a big deal
        self.advancedUpperLimitEntry.bind("<KeyRelease>", checkUpperAndLower)
        self.advancedLowerLimitEntry.bind("<KeyRelease>", checkUpperAndLower)
        
        def _on_mousewheel_help(event):
            xpos, ypos = self.vsbh.get()
            xpos_round = round(xpos, 2)     #Avoid floating point errors
            ypos_round = round(ypos, 2)
            if (xpos_round != 0.00 or ypos_round != 1.00):
                self.hcanvas.yview_scroll(int(-1*(event.delta/120)), "units")

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
        
        def validateIt(P):
            if P == "":
                return True
            else:
                return P.isalnum()
        self.valcom = (self.register(validateIt), '%P')
        
        def validateFreqHigh(P):
            if (P == ""):
                return True
            try:
                val = int(P)
            except Exception:
                return False
            if (val < 0):
                return False
            elif (val + self.lowDelete > len(self.wdataRaw)):
                return False
            return True
        valfreqHigh = (self.register(validateFreqHigh), '%P')
        
        def validateFreqLow(P):
            if (P == ""):
                return True
            try:
                val = int(P)
            except Exception:
                return False
            if (val < 0):
                return False
            elif (val + self.upDelete > len(self.wdataRaw)):
                return False
            return True
        valfreqLow = (self.register(validateFreqLow), '%P')
        
        def OpenFile():
            name = askopenfilename(initialdir=self.topGUI.getCurrentDirectory(), filetypes = [("Measurement model file", "*.mmfile *.mmcustom"), ("Measurement model data (*.mmfile)", "*.mmfile"), ("Measurement model custom fitting (*.mmcustom)", "*.mmcustom")], title = "Choose a file")
            try:
                with open(name,'r') as UseFile:
                    return name
            except Exception:
                return "+"
        
        def limitChanged(num, e):
            if (self.paramComboboxValues[num].get() == "+"):
                self.upperLimits[num].set("inf")
                self.lowerLimits[num].set("0")
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
            elif (self.paramComboboxValues[num].get() == "+ or -"):
                self.upperLimits[num].set("inf")
                self.lowerLimits[num].set("-inf")
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
            elif (self.paramComboboxValues[num].get() == "-"):
                self.upperLimits[num].set("0")
                self.lowerLimits[num].set("-inf")
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
            elif (self.paramComboboxValues[num].get() == "fixed" and self.paramListboxSelection == num):
                self.advancedLowerLimitEntry.configure(state="disabled")
                self.advancedUpperLimitEntry.configure(state="disabled")
            elif (self.paramComboboxValues[num].get() == "custom"):
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
                advancedOptionsPopup()
                self.paramListbox.selection_clear(0, tk.END)
                self.paramListbox.select_set(num)
                self.paramListboxSelection = num
                onSelect(None, n=num)
        
        def browse():
            n = OpenFile()
            if (n != '+'):
                fname, fext = os.path.splitext(n)
                directory = os.path.dirname(str(n))
                self.topGUI.setCurrentDirectory(directory)
                if (fext == ".mmfile"):
                    try:
                        with open(n,'r') as UseFile:
                            filetext = UseFile.read()
                            lines = filetext.splitlines()
                        if ("frequency" in lines[0].lower()):
                            data = np.loadtxt(n, skiprows=1)
                        else:
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
                        if (self.topGUI.getFreqLoadCustom() == 1):
                            try:
                                if (self.upDelete == 0):
                                    self.wdata = self.wdataRaw.copy()[self.lowDelete:]
                                    self.rdata = self.rdataRaw.copy()[self.lowDelete:]
                                    self.jdata = self.jdataRaw.copy()[self.lowDelete:]
                                else:
                                    self.wdata = self.wdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                    self.rdata = self.rdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                                    self.jdata = self.jdataRaw.copy()[self.lowDelete:-1*self.upDelete]
                            except Exception:
                                messagebox.showwarning("Frequency error", "There are more frequencies set to be deleted than data points. The number of frequencies to delete has been reset to 0.")
                                self.upDelete = 0
                                self.lowDelete = 0
                                self.lowerSpinboxVariable.set(str(self.lowDelete))
                                self.upperSpinboxVariable.set(str(self.upDelete))
                                self.wdata = self.wdataRaw.copy()
                                self.rdata = self.rdataRaw.copy()
                                self.jdata = self.jdataRaw.copy()

                            self.lowestUndeleted.configure(state="normal")
                            self.lowestUndeleted.delete(1.0, tk.END)
                            self.lowestUndeleted.insert(1.0, "Lowest remaining frequency: {:.4e}".format(min(self.wdata)))
                            self.lowestUndeleted.configure(state="disabled")
                            self.highestUndeleted.configure(state="normal")
                            self.highestUndeleted.delete(1.0, tk.END)
                            self.highestUndeleted.insert(1.0, "Highest remaining frequency: {:.4e}".format(max(self.wdata)))
                            self.highestUndeleted.configure(state="disabled")
                        else:
                            self.upDelete = 0
                            self.lowDelete = 0
                            self.lowerSpinboxVariable.set(str(self.lowDelete))
                            self.upperSpinboxVariable.set(str(self.upDelete))
                            self.wdata = self.wdataRaw.copy()
                            self.rdata = self.rdataRaw.copy()
                            self.jdata = self.jdataRaw.copy()
                            self.lowestUndeleted.configure(state="normal")
                            self.lowestUndeleted.delete(1.0, tk.END)
                            self.lowestUndeleted.insert(1.0, "Lowest remaining frequency: {:.4e}".format(min(self.wdata)))
                            self.lowestUndeleted.configure(state="disabled")
                            self.highestUndeleted.configure(state="normal")
                            self.highestUndeleted.delete(1.0, tk.END)
                            self.highestUndeleted.insert(1.0, "Highest remaining frequency: {:.4e}".format(max(self.wdata)))
                            self.highestUndeleted.configure(state="disabled")
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
                                self.realFreq.set_ylabel("Zr / Ω", color=self.foregroundColor)
                                self.imagFreq.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                                self.canvasFreq.draw()
                        except Exception:
                            pass
                        self.lengthOfData = len(self.wdata)
                        self.freqRangeButton.configure(state="normal")
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, n)
                        self.browseEntry.configure(state="readonly")
                        self.runButton.configure(state="normal")
                        self.simplexButton.configure(state="normal")
                    except Exception:
                        messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
                elif (fext == ".mmcustom"):
                    try:
                        toLoad = open(n)
                        firstLine = toLoad.readline()
                        fileVersion = 0
                        if "version: 1.1" in firstLine:
                            fileVersion = 1
                        elif "version: 1.2" in firstLine:
                            fileVersion = 2
                        if (fileVersion == 0):
                            fileToLoad = firstLine.split("filename: ")[1][:-1]
                        else:
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
                            if (fileVersion == 0):
                                if "mmparams:" in nextLineIn:
                                    break
                                else:
                                    formulaIn += nextLineIn
                            elif (fileVersion == 1):
                                if "description:" in nextLineIn:
                                    break
                                else:
                                    formulaIn += nextLineIn
                            elif (fileVersion == 2):
                                if "imports:" in nextLineIn:
                                    break
                                else:
                                    formulaIn += nextLineIn
                        for i in range(len(self.paramNameEntries)):
                            self.paramNameEntries[i].grid_remove()
                            self.paramNameLabels[i].grid_remove()
                            self.paramValueEntries[i].grid_remove()
                            self.paramValueLabels[i].grid_remove()
                            self.paramComboboxes[i].grid_remove()
                            self.paramDeleteButtons[i].grid_remove()
                        self.numParams = 0
                        self.paramNameValues.clear()
                        self.paramNameEntries.clear()
                        self.paramNameLabels.clear()
                        self.paramValueLabels.clear()
                        self.paramComboboxValues.clear()
                        self.paramComboboxes.clear()
                        self.paramValueValues.clear()
                        self.paramValueEntries.clear()
                        self.paramDeleteButtons.clear()
                        self.upperLimits.clear()
                        self.lowerLimits.clear()
                        self.paramListbox.delete(0, tk.END)
                        self.multistartCheckboxVariables.clear()
                        self.multistartSpacingVariables.clear()
                        self.multistartLowerVariables.clear()
                        self.multistartUpperVariables.clear()
                        self.multistartNumberVariables.clear()
                        self.multistartCustomVariables.clear()
                        self.advancedLowerLimitEntry.configure(state="normal")
                        self.advancedUpperLimitEntry.configure(state="normal")
                        imports = ""
                        if (fileVersion == 2):
                            while True:
                                nextLineIn = toLoad.readline()
                                if "description:" in nextLineIn:
                                    break
                                else:
                                    imports += nextLineIn
                        toImport = imports.rstrip().split('\n')
                        for tI in toImport:
                            if (len(tI) > 0):
                                self.importPathListbox.insert(tk.END, tI)
                        if fileVersion != 0:
                            self.formulaDescriptionEntry.delete("1.0", tk.END)
                            descriptionIn = ""
                            while True:
                                nextLineIn = toLoad.readline()
                                if "latex:" in nextLineIn:
                                    break
                                else:
                                    descriptionIn += nextLineIn
                            latexIn = toLoad.readline()
                            self.formulaDescriptionEntry.insert("1.0", descriptionIn.rstrip())
                            self.formulaDescriptionLatexEntry.delete("1.0", tk.END)
                            self.formulaDescriptionLatexEntry.insert("1.0", latexIn.rstrip())
                            self.latexAx.clear()
                            self.latexAx.axis("off")
                            self.latexCanvas.draw()
                            self.graphLatex()
                            toLoad.readline()
                        while True:
                            p = toLoad.readline()
                            if not p:
                                break
                            else:
                                p = p[:-1]
                                try:
                                    pname, pval, pcombo, plimup, plimlow, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                                    self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                                    self.multistartLowerVariables.append(tk.StringVar(self, plower))
                                    self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                                    self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                                    self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                                    self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                                    self.upperLimits.append(tk.StringVar(self, plimup))
                                    self.lowerLimits.append(tk.StringVar(self, plimlow))
                                except Exception:
                                    try:
                                        pname, pval, pcombo, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                                        self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                                        self.multistartLowerVariables.append(tk.StringVar(self, plower))
                                        self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                                        self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                                        self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                                        self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                                        self.upperLimits.append(tk.StringVar(self, "inf"))
                                        self.lowerLimits.append(tk.StringVar(self, "-inf"))
                                    except Exception:
                                        pname, pval, pcombo = p.split("~=~")
                                        self.multistartCheckboxVariables.append(tk.IntVar(self, 0))
                                        self.multistartLowerVariables.append(tk.StringVar(self, "1E-5"))
                                        self.multistartUpperVariables.append(tk.StringVar(self, "1E5"))
                                        self.multistartSpacingVariables.append(tk.StringVar(self, "Logarithmic"))
                                        self.multistartCustomVariables.append(tk.StringVar(self, "0.1,1,10,100,1000"))
                                        self.multistartNumberVariables.append(tk.StringVar(self, "10"))
                                        self.upperLimits.append(tk.StringVar(self, "inf"))
                                        self.lowerLimits.append(tk.StringVar(self, "-inf"))
                                self.numParams += 1
                                if (self.numParams == 1):
                                    self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[0])
                                    self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[0])
                                    self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[0])
                                    self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[0])
                                    self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[0])
                                    self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[0])
                                    self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                    self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                    self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[0])
                                    self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[0])
                                    self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                    self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                    self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                    self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                    self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                    if (self.multistartCheckboxVariables[0].get() == 0):
                                        self.multistartSpacing.configure(state="disabled")
                                        self.multistartUpperEntry.configure(state="disabled")
                                        self.multistartLowerEntry.configure(state="disabled")
                                        self.multistartNumberEntry.configure(state="disabled")
                                        self.multistartCustomEntry.configure(state="disabled")
                                    else:
                                        if (self.multistartSpacingVariables[0].get() == "Custom"):
                                            self.multistartSpacing.configure(state="readonly")
                                            self.multistartUpperEntry.configure(state="normal")
                                            self.multistartLowerEntry.configure(state="normal")
                                            self.multistartNumberEntry.configure(state="normal")
                                            self.multistartCustomEntry.configure(state="normal")
                                            self.multistartLowerFrame.pack_forget()
                                            self.multistartUpperFrame.pack_forget()
                                            self.multistartNumberFrame.pack_forget()
                                            self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                        else:
                                            self.multistartSpacing.configure(state="readonly")
                                            self.multistartUpperEntry.configure(state="normal")
                                            self.multistartLowerEntry.configure(state="normal")
                                            self.multistartNumberEntry.configure(state="normal")
                                            self.multistartCustomEntry.configure(state="normal")
                                self.paramNameValues.append(tk.StringVar(self, pname))
                                self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10, validate="all", validatecommand=self.valcom))
                                self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                                self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                                self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                                self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed", "custom"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                                self.paramComboboxes[-1].bind("<<ComboboxSelected>>", partial(limitChanged, self.numParams-1))
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
                                self.paramListbox.insert(tk.END, pname)
                                if (self.paramComboboxValues[0].get() == "fixed"):
                                    self.advancedLowerLimitEntry.configure(state="disabled")
                                    self.advancedUpperLimitEntry.configure(state="disabled")
                        toLoad.close()
                        self.paramListbox.selection_clear(0, tk.END)
                        self.paramListbox.select_set(0)
                        self.paramListboxSelection = 0
                        onSelect(None, n=0)
                        try:
                            with open(fileToLoad,'r') as UseFile:
                                filetext = UseFile.read()
                                lines = filetext.splitlines()
                            if ("frequency" in lines[0].lower()):
                                data = np.loadtxt(fileToLoad, skiprows=1)
                            else:
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
                                    self.realFreq.set_ylabel("Zr / Ω", color=self.foregroundColor)
                                    self.imagFreq.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                                    self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                                    self.canvasFreq.draw()
                            except Exception:
                                pass
                            self.lengthOfData = len(self.wdata)
                            self.freqRangeButton.configure(state="normal")
                            self.runButton.configure(state="normal")
                            self.simplexButton.configure(state="normal")
                        except Exception:
                            messagebox.showerror("File not found", "Error 53: \nThe linked .mmfile could not be found")
                            fileToLoad = ""
                        self.browseEntry.configure(state="normal")
                        self.browseEntry.delete(0,tk.END)
                        self.browseEntry.insert(0, n)
                        self.browseEntry.configure(state="readonly")
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
                        self.noiseEntryValue.set(alphaVal)
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
                        self.customFormula.insert("1.0", formulaIn.rstrip())
                        keyup("")
                    except Exception:
                        messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
                else:
                    messagebox.showerror("File error", "Error 43:\n The file has an unknown extension")
        
        def changeFreqs():            
            if (self.freqWindow.state() == "withdrawn"):
                self.freqWindow.deiconify()
            else:
                self.freqWindow.lift()
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
                self.realFreq.set_ylabel("Zr / Ω", color=self.foregroundColor)
                self.imagFreq.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.canvasFreq.draw()
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
                except Exception:
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

                self.lengthOfData = len(self.wdata)
                self.lowestUndeleted.configure(state="normal")
                self.lowestUndeleted.delete(1.0, tk.END)
                self.lowestUndeleted.insert(1.0, "Lowest remaining frequency: {:.4e}".format(min(self.wdata)))
                self.lowestUndeleted.configure(state="disabled")
                self.highestUndeleted.configure(state="normal")
                self.highestUndeleted.delete(1.0, tk.END)
                self.highestUndeleted.insert(1.0, "Highest remaining frequency: {:.4e}".format(max(self.wdata)))
                self.highestUndeleted.configure(state="disabled")
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

            def changeFreqSpinboxLower(e = None):
                try:
                    higherDelete = 0 if self.upperSpinboxVariable.get() == "" else int(self.upperSpinboxVariable.get())
                    lowerDelete = 0 if self.lowerSpinboxVariable.get() == "" else int(self.lowerSpinboxVariable.get())
                    self.upperSpinbox.configure(to=len(self.wdataRaw)-1-lowerDelete)
                    updateFreqs()
                except Exception:
                    pass
            
            def changeFreqSpinboxUpper(e = None):
                try:
                    higherDelete = 0 if self.upperSpinboxVariable.get() == "" else int(self.upperSpinboxVariable.get())
                    lowerDelete = 0 if self.lowerSpinboxVariable.get() == "" else int(self.lowerSpinboxVariable.get())
                    self.lowerSpinbox.configure(to=len(self.wdataRaw)-1-higherDelete)
                    updateFreqs()
                except Exception:
                    pass
            
            self.lowerLabel = tk.Label(self.freqWindow, text="Number of low frequencies\n to delete", bg=self.backgroundColor, fg=self.foregroundColor)
            self.rangeLabel = tk.Label(self.freqWindow, text="Log of Frequency", bg=self.backgroundColor, fg=self.foregroundColor)
            self.upperLabel = tk.Label(self.freqWindow, text="Number of high frequencies\n to delete", bg=self.backgroundColor, fg=self.foregroundColor)
            self.lowerLabel.grid(column=0, row=1, pady=(80, 0), padx=3)
            self.upperLabel.grid(column=2, row=1, pady=(80, 0), padx=3)
            self.lowerSpinbox = tk.Spinbox(self.freqWindow, from_=0, to=(len(self.wdataRaw)-1), textvariable=self.lowerSpinboxVariable, state="normal", width=6, validate="all", validatecommand=valfreqLow, command=changeFreqSpinboxLower)
            self.lowerSpinbox.grid(column=0, row=2, padx=(3,0))
            self.upperSpinbox = tk.Spinbox(self.freqWindow, from_=0, to=(len(self.wdataRaw)-1), textvariable=self.upperSpinboxVariable, state="normal", width=6, validate="all", validatecommand=valfreqHigh, command=changeFreqSpinboxUpper)
            self.lowerSpinbox.bind("<KeyRelease>", changeFreqSpinboxLower)
            self.upperSpinbox.bind("<KeyRelease>", changeFreqSpinboxUpper)
            self.upperSpinbox.grid(column=2, row=2, padx=(0,3))
            self.lowestUndeleted.configure(state="normal")
            self.lowestUndeleted.delete(1.0, tk.END)
            self.lowestUndeleted.insert(1.0, "Lowest remaining frequency: {:.4e}".format(min(self.wdata)))
            self.lowestUndeleted.configure(state="disabled")
            self.highestUndeleted.configure(state="normal")
            self.highestUndeleted.delete(1.0, tk.END)
            self.highestUndeleted.insert(1.0, "Highest remaining frequency: {:.4e}".format(max(self.wdata)))
            self.highestUndeleted.configure(state="disabled")
            self.lowestUndeleted.grid(column=0, row=3, sticky="N")
            self.highestUndeleted.grid(column=2, row=3, sticky="N")            
            
            def on_closing():
                self.upperSpinboxVariable.set(str(self.upDelete))
                self.lowerSpinboxVariable.set(str(self.lowDelete))
                self.freqWindow.withdraw()
            self.freqWindow.protocol("WM_DELETE_WINDOW", on_closing)

        def keyup(event):
            self.paramListbox.delete(0, tk.END)
            for n in self.paramNameValues:
                self.paramListbox.insert(tk.END, n.get())
            try:
                self.advancedOptionsLabel.configure(text=self.paramNameValues[self.paramListboxSelection].get())
            except Exception:
                pass
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
                    self.multistartCheckboxVariables[i].set(self.multistartCheckboxVariables[i+1].get())
                    self.multistartLowerVariables[i].set(self.multistartLowerVariables[i+1].get())
                    self.multistartUpperVariables[i].set(self.multistartUpperVariables[i+1].get())
                    self.multistartSpacingVariables[i].set(self.multistartSpacingVariables[i+1].get())
                    self.multistartCustomVariables[i].set(self.multistartCustomVariables[i+1].get())
                    self.multistartNumberVariables[i].set(self.multistartNumberVariables[i+1].get())
                    self.upperLimits[i].set(self.upperLimits[i+1].get())
                    self.lowerLimits[i].set(self.lowerLimits[i+1].get())
                removeParam()
        
        def addParam():
            self.numParams += 1
            self.paramNameValues.append(tk.StringVar(self, "var" + str(self.varNum)))
            self.varNum += 1
            self.multistartCheckboxVariables.append(tk.IntVar(self, 0))
            self.multistartLowerVariables.append(tk.StringVar(self, "1E-5"))
            self.multistartUpperVariables.append(tk.StringVar(self, "1E5"))
            self.multistartSpacingVariables.append(tk.StringVar(self, "Logarithmic"))
            self.multistartCustomVariables.append(tk.StringVar(self, "0.1,1,10,100,1000"))
            self.multistartNumberVariables.append(tk.StringVar(self, "10"))
            self.upperLimits.append(tk.StringVar(self, "inf"))
            self.lowerLimits.append(tk.StringVar(self, "-inf"))
            self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10, validate="all", validatecommand=self.valcom))
            self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
            self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
            self.paramComboboxValues.append(tk.StringVar(self, "+ or -"))
            self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed", "custom"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
            self.paramComboboxes[-1].bind("<<ComboboxSelected>>", partial(limitChanged, self.numParams-1))
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
            self.paramListbox.insert(tk.END, self.paramNameValues[-1].get())
            if (self.numParams == 1):
                self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[0])
                self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[0])
                self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[0])
                self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[0])
                self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[0])
                self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[0])
                self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[0])
                self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[0])
                self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                self.multistartSpacing.configure(state="disabled")
                self.multistartUpperEntry.configure(state="disabled")
                self.multistartLowerEntry.configure(state="disabled")
                self.multistartNumberEntry.configure(state="disabled")
                self.multistartCustomEntry.configure(state="disabled")
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
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
                self.paramListbox.delete(tk.END)
                self.paramListbox.selection_clear(0, tk.END)
                self.paramListbox.select_set(0)
                self.paramListboxSelection = 0
                onSelect(None, n=0)
                self.multistartCheckboxVariables.pop()
                self.multistartLowerVariables.pop()
                self.multistartUpperVariables.pop()
                self.multistartSpacingVariables.pop()
                self.multistartCustomVariables.pop()
                self.multistartNumberVariables.pop()
                self.upperLimits.pop()
                self.lowerLimits.pop()
                keyup("")
        
        def advCheck():
            if (self.multistartCheckboxVariables[self.paramListboxSelection].get()):
                self.multistartSpacing.configure(state="readonly")
                self.multistartUpperEntry.configure(state="normal")
                self.multistartLowerEntry.configure(state="normal")
                self.multistartNumberEntry.configure(state="normal")
                self.multistartCustomEntry.configure(state="normal")
            else:
                self.multistartSpacing.configure(state="disabled")
                self.multistartUpperEntry.configure(state="disabled")
                self.multistartLowerEntry.configure(state="disabled")
                self.multistartNumberEntry.configure(state="disabled")
                self.multistartCustomEntry.configure(state="disabled")
        
        def advSelectionChange(e):
            if (self.multistartSpacingVariables[self.paramListboxSelection].get() == "Custom"):
                self.multistartLowerFrame.pack_forget()
                self.multistartUpperFrame.pack_forget()
                self.multistartNumberFrame.pack_forget()
                self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
            else:
                self.multistartCustomFrame.pack_forget()
                self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
        
        def onSelect(e, n=-1):
            try:
                if (n == -1):
                    num_selected = self.paramListbox.curselection()[0]
                else:
                    num_selected = n
                self.paramListboxSelection = num_selected
                self.advancedOptionsLabel.configure(text=self.paramNameValues[num_selected].get())
                self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                if (self.paramComboboxValues[num_selected].get() == "fixed"):
                    self.advancedLowerLimitEntry.configure(state="disabled")
                    self.advancedUpperLimitEntry.configure(state="disabled")
                else:
                    self.advancedLowerLimitEntry.configure(state="normal")
                    self.advancedUpperLimitEntry.configure(state="normal")
                if (self.multistartSpacingVariables[num_selected].get() == "Custom"):
                    self.multistartLowerFrame.pack_forget()
                    self.multistartUpperFrame.pack_forget()
                    self.multistartNumberFrame.pack_forget()
                    self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                else:
                    self.multistartCustomFrame.pack_forget()
                    self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                    self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                    self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                if (self.multistartCheckboxVariables[num_selected].get()):
                    self.multistartSpacing.configure(state="readonly")
                    self.multistartUpperEntry.configure(state="normal")
                    self.multistartLowerEntry.configure(state="normal")
                    self.multistartNumberEntry.configure(state="normal")
                    self.multistartCustomEntry.configure(state="normal")
                else:
                    self.multistartSpacing.configure(state="disabled")
                    self.multistartUpperEntry.configure(state="disabled")
                    self.multistartLowerEntry.configure(state="disabled")
                    self.multistartNumberEntry.configure(state="disabled")
                    self.multistartCustomEntry.configure(state="disabled")
                self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[num_selected])
                self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[num_selected])
                self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[num_selected])
                self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[num_selected])
                self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[num_selected])
                self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[num_selected])
                self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[num_selected])
                self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[num_selected])
            except Exception:     #No parameters left
                self.paramListboxSelection = 0
                self.advancedOptionsLabel.configure(text="")
                self.advancedLowerLimitFrame.pack_forget()
                self.advancedUpperLimitFrame.pack_forget()
                self.multistartCheckbox.pack_forget()
                self.multistartSpacingFrame.pack_forget()
                self.multistartLowerFrame.pack_forget()
                self.multistartUpperFrame.pack_forget()
                self.multistartNumberFrame.pack_forget()
                self.multistartCustomFrame.pack_forget()
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
        
        self.multistartCheckbox.configure(command=advCheck)
        self.multistartSpacing.bind("<<ComboboxSelected>>", advSelectionChange)
        self.paramListbox.bind("<<ListboxSelect>>", onSelect)
        
        def advancedOptionsPopup():
            self.advancedOptionsWindow.deiconify()
            self.advancedOptionsWindow.lift()
            def onClose():
                self.advancedOptionsWindow.withdraw()
            
            self.advancedOptionsWindow.protocol("WM_DELETE_WINDOW", onClose)
        
        def loadParams():
            a = True
            if (len(self.paramNameEntries) > 0):
                a = messagebox.askokcancel("Remove parameters", "Loading parameters will remove existing parameters. Continue?", parent=self.paramPopup)
            if a:
                def OpenFileParam():
                    name = askopenfilename(initialdir=self.topGUI.getCurrentDirectory(), filetypes = [("Custom fitting file", "*.mmcustom"), ("Measurement model custom fitting (*.mmcustom)", "*.mmcustom")], title = "Choose a file", parent=self.paramPopup)
                    try:
                        with open(name,'r') as UseFile:
                            return name
                    except Exception:
                        return "+"
                n = OpenFile()
                if (n != '+'):
                    fname, fext = os.path.splitext(n)
                    directory = os.path.dirname(str(n))
                    self.topGUI.setCurrentDirectory(directory)
                    if (fext == ".mmcustom"):
                        try:
                            toLoad = open(n)
                            for i in range(13):
                                toLoad.readline()
                            while True:
                                nextLineIn = toLoad.readline()
                                if "mmparams:" in nextLineIn:
                                    break
                            for i in range(len(self.paramNameEntries)):
                                self.paramNameEntries[i].grid_remove()
                                self.paramNameLabels[i].grid_remove()
                                self.paramValueEntries[i].grid_remove()
                                self.paramValueLabels[i].grid_remove()
                                self.paramComboboxes[i].grid_remove()
                                self.paramDeleteButtons[i].grid_remove()
                            self.numParams = 0
                            self.paramNameValues.clear()
                            self.paramNameEntries.clear()
                            self.paramNameLabels.clear()
                            self.paramValueLabels.clear()
                            self.paramComboboxValues.clear()
                            self.paramComboboxes.clear()
                            self.paramValueValues.clear()
                            self.paramValueEntries.clear()
                            self.paramDeleteButtons.clear()
                            self.paramListbox.delete(0, tk.END)
                            self.multistartCheckboxVariables.clear()
                            self.multistartSpacingVariables.clear()
                            self.multistartLowerVariables.clear()
                            self.multistartUpperVariables.clear()
                            self.multistartNumberVariables.clear()
                            self.multistartCustomVariables.clear()
                            self.upperLimits.clear()
                            self.lowerLimits.clear()
                            while True:
                                p = toLoad.readline()
                                if not p:
                                    break
                                else:
                                    p = p[:-1]
                                    try:
                                        pname, pval, pcombo, plimup, plimlow, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                                        self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                                        self.multistartLowerVariables.append(tk.StringVar(self, plower))
                                        self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                                        self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                                        self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                                        self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                                        self.upperLimits.append(tk.StringVar(self, plimup))
                                        self.lowerLimits.append(tk.StringVar(self, plimlow))
                                    except Exception:
                                        pname, pval, pcombo = p.split("~=~")
                                        self.multistartCheckboxVariables.append(tk.IntVar(self, 0))
                                        self.multistartLowerVariables.append(tk.StringVar(self, "1E-5"))
                                        self.multistartUpperVariables.append(tk.StringVar(self, "1E5"))
                                        self.multistartSpacingVariables.append(tk.StringVar(self, "Logarithmic"))
                                        self.multistartCustomVariables.append(tk.StringVar(self, "0.1,1,10,100,1000"))
                                        self.multistartNumberVariables.append(tk.StringVar(self, "10"))
                                        self.upperLimits.append(tk.StringVar(self, "inf"))
                                        self.lowerLimits.append(tk.StringVar(self, "-inf"))
                                    self.numParams += 1
                                    if (self.numParams == 1):
                                        self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[0])
                                        self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[0])
                                        self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[0])
                                        self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[0])
                                        self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[0])
                                        self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[0])
                                        self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                        self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                        self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                        self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                        self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                        self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                                        self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                        if (self.paramComboboxValues[0].get() == "fixed"):
                                            self.advancedLowerLimitFrame.configure(state="disabled")
                                            self.advancedUpperLimitFrame.configure(state="disabled")
                                        if (self.multistartCheckboxVariables[0].get() == 0):
                                            self.multistartSpacing.configure(state="disabled")
                                            self.multistartUpperEntry.configure(state="disabled")
                                            self.multistartLowerEntry.configure(state="disabled")
                                            self.multistartNumberEntry.configure(state="disabled")
                                            self.multistartCustomEntry.configure(state="disabled")
                                        else:
                                            if (self.multistartSpacingVariables[0].get() == "Custom"):
                                                self.multistartSpacing.configure(state="readonly")
                                                self.multistartUpperEntry.configure(state="normal")
                                                self.multistartLowerEntry.configure(state="normal")
                                                self.multistartNumberEntry.configure(state="normal")
                                                self.multistartCustomEntry.configure(state="normal")
                                                self.multistartLowerFrame.pack_forget()
                                                self.multistartUpperFrame.pack_forget()
                                                self.multistartNumberFrame.pack_forget()
                                                self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                            else:
                                                self.multistartSpacing.configure(state="readonly")
                                                self.multistartUpperEntry.configure(state="normal")
                                                self.multistartLowerEntry.configure(state="normal")
                                                self.multistartNumberEntry.configure(state="normal")
                                                self.multistartCustomEntry.configure(state="normal")
                                    self.paramNameValues.append(tk.StringVar(self, pname))
                                    self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10, validate="all", validatecommand=self.valcom))
                                    self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                                    self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                                    self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                                    self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed", "custom"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                                    self.paramComboboxes[-1].bind("<<ComboboxSelected>>", partial(limitChanged, self.numParams-1))
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
                                    self.paramListbox.insert(tk.END, pname)
                            toLoad.close()
                            self.paramListbox.selection_clear(0, tk.END)
                            self.paramListbox.select_set(0)
                            self.paramListboxSelection = 0
                            self.paramListbox.event_generate("<<ListboxSelect>>")
                            keyup("")
                            parameterPopup()
                        except Exception:
                            messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
                    else:
                        messagebox.showerror("File error", "Error 43:\n The file has an unknown extension")
            else:   #Cancel is clicked
                pass
        
        def simplexFit():
            if (len(self.paramValueValues) < 1):    #If there are no parameters to fit
                pass
            else:
                formula = self.customFormula.get("1.0", tk.END)
                if ("".join(formula.split()) == ""):
                    messagebox.showwarning("Error", "Error 44: \nFormula is empty", parent=self.paramPopup)
                    return
                
                #---Check that none of the variable names are repeated, and that none are Python reserved words or variables used in this code
                for i in range(len(self.paramNameValues)):
                    name = self.paramNameValues[i].get()
                    if (name == "False" or name == "None" or name == "True" or name == "and" or name == "as" or name == "assert" or name == "break" or name == "class" or name == "continue" or name == "def" or name == "del" or name == "elif" or name == "else" or name == "except"\
                         or name == "finally" or name == "for" or name == "from" or name == "global" or name == "if" or name == "import" or name == "in" or name == "is" or name == "lambda" or name == "nonlocal" or name == "not" or name == "or" or name == "pass" or name == "raise"\
                          or name == "return" or name == "try" or name == "while" or name == "with" or name == "yield"):
                        messagebox.showwarning("Error", "Error 45: \nThe variable name \"" + name + "\" is a Python reserved word. Change the variable name.", parent=self.paramPopup)
                        return
                    elif (name == "freq" or name == "Zr" or name == "Zj" or name == "Zreal" or name == "Zimag" or name == "weighting"):
                        messagebox.showwarning("Error", "Error 47: \nThe variable name \"" + name + "\" is used by the fitting program; change the variable name.", parent=self.paramPopup)
                        return
                    for j in range(i+1, len(self.paramNameValues)):
                        if (name == self.paramNameValues[j].get()):
                            messagebox.showwarning("Error", "Error 48: \nTwo or more variables have the same name.", parent=self.paramPopup)
                            return
                
                #---Replace the functions with np.<function>, and then attempt to compile the code to look for syntax errors---        
                try:
                    prebuiltFormulas = ['PI', 'ARCSINH', 'ARCCOSH', 'ARCTANH', 'ARCSIN', 'ARCCOS', 'ARCTAN', 'COSH', 'SINH', 'TANH', 'SIN', 'COS', 'TAN', 'SQRT', 'EXP', 'ABS', 'DEG2RAD', 'RAD2DEG']
                    formula = formula.replace("^", "**")    #Replace ^ with ** for exponentiation (this could prevent some features like regex from being used effectively)
                    formula = formula.replace("LN", "np.emath.log")     #Replace LN with the natural log (the emath version returns complex numbers for negative arguments)
                    formula = formula.replace("LOG", "np.emath.log10")  #Replace LOG with the base-10 log (the emath version returns complex numbers for negative arguments)
                    for pf in prebuiltFormulas:
                        toReplace = "np." + pf.lower()
                        formula = formula.replace(pf, toReplace)
                    formula = formula.replace("\n", "\n\t")
                    formula = "try:\n\t" + formula
                    formula = formula.rstrip()
                    formula += "\n\tif any(np.isnan(Zreal)) or any(np.isnan(Zimag)):\n\t\traise Exception\nexcept Exception:\n\tZreal = np.full(len(freq), 1E300)\n\tZimag = np.full(len(freq), 1E300)"
                    compile(formula, 'user_generated_formula', 'exec')
                except Exception:
                    messagebox.showwarning("Compile error", "There was an issue compiling the code", parent=self.paramPopup)
                    return
                
                #---Check if the variable "freq" appears in the code, as it's likely a mistake if it doesn't---
                textToSearch = self.customFormula.get("1.0", tk.END)
                whereFreq = [m.start() for m in re.finditer(r'\bfreq\b', textToSearch)]
                if (len(whereFreq) == 0):
                    messagebox.showwarning("No \"freq\"", "The variable \"freq\" does not appear in the code. This may cause an error.", parent=self.paramPopup)
                elif (len(whereFreq) == 1 and whereFreq[0] == 5):
                    messagebox.showwarning("No \"freq\"", "The variable \"freq\" does not seem to be in the code. This may cause an error.", parent=self.paramPopup)
                
                self.queue = queue.Queue()
                fit_type = 3
                if (self.fittingTypeComboboxValue.get() == "Real"):
                    fit_type = 1
                elif (self.fittingTypeCombobox.get() == "Imaginary"):
                    fit_type = 2
                param_names = []
                for pNV in self.paramNameValues:
                    param_names.append(pNV.get())
                param_guesses = []
                try:
                    for pG in self.paramValueValues:
                        param_guesses.append(float(pG.get()))
                except Exception:
                    messagebox.showwarning("Value error", "Error 54: \nThe parameter guesses must be real numbers", parent=self.paramPopup)
                    return
                weight = 0
                errorModelParams = np.zeros(5)
                if (self.weightingComboboxValue.get() == "Modulus"):
                    weight = 1
                elif (self.weightingComboboxValue.get() == "Proportional"):
                    weight = 2
                elif (self.weightingComboboxValue.get() == "Error model"):
                    if (self.errorAlphaCheckboxVariable.get() != 1 and self.errorBetaCheckboxVariable.get() != 1 and self.errorGammaCheckboxVariable.get() != 1 and self.errorDeltaCheckboxVariable.get() != 1):
                        messagebox.showwarning("Bad error structure", "At least one error structure value must be checkd", parent=self.paramPopup)
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
                    except Exception:
                        messagebox.showerror("Value error", "One of the error structure parameters has an invalid value", parent=self.paramPopup)
                        return
                    if (errorModelParams[0] == 0 and errorModelParams[1] == 0 and errorModelParams[3] == 0 and errorModelParams[4] == 0):
                        messagebox.showerror("Value error", "At least one of the error structure parameters must be nonzero", parent=self.paramPopup)
                        return
                    weight = 3
                elif (self.weightingComboboxValue.get() == "Custom"):
                    weight = 4
                assumed_noise = float(self.noiseEntryValue.get())
                param_limits = []
                for i in range(len(self.paramComboboxValues)):
                    pL = self.paramComboboxValues[i]
                    if (pL.get() == "+"):
                        param_limits.append("+")
                    elif (pL.get() == "-"):
                        param_limits.append("-")
                    elif (pL.get() == "fixed"):
                        param_limits.append("f")
                    elif (pL.get() == "custom"):
                        try:
                            float(self.upperLimits[i].get())
                            float(self.lowerLimits[i].get())
                        except Exception:
                            messagebox.showwarning("Value error", "Error 61: \nThe upper and lower limits must be real numbers", parent=self.paramPopup)
                            return
                        param_limits.append(str(self.upperLimits[i].get()) + ";" + str(self.lowerLimits[i].get()))
                    else:
                        param_limits.append("n")
                result = simplexFitting.findFit(fit_type, weight, assumed_noise, formula, self.wdata, self.rdata, self.jdata, param_names, param_guesses, param_limits, errorModelParams)
                for i in range(len(self.paramValueValues)):
                    self.paramValueValues[i].set(str(round(result[i], -int(np.floor(np.log10(abs(result[i])))) + (4 - 1))))
        
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
            
            self.addButton.configure(command=addParam)
            self.removeButton.configure(command=removeParam)
            self.advancedOptionsButton.configure(command=advancedOptionsPopup)
            self.simplexButton.configure(command=simplexFit)
            self.paramPopup.title("Custom fitting parameters")
            self.paramPopup.iconbitmap(resource_path("img/elephant3.ico"))
            self.paramPopup.geometry("500x400")
            self.paramPopup.deiconify()
            self.paramPopup.protocol("WM_DELETE_WINDOW", onClose)
            
            self.pframe.bind("<Configure>", lambda event, canvas=self.pcanvas: onFrameConfigure(canvas))
        
        def helpPopup(event):
            webbrowser.open_new(r'file://' + os.path.dirname(os.path.realpath(__file__)) + '\Formula_guide.pdf')
        
        def graphLatex(e=None):
            try:
                tmptext = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '')   #self.formulaDescriptionLatexVariable.get()
                tmptext2 = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
                split_text = tmptext.split(r"\\")
                tmptext = ""
                first = True
                for t in split_text:
                    if first:
                        tmptext = "$"+t+"$"
                        first = False
                    else:
                        tmptext += "\n" + "$"+t+"$"
                if tmptext2 == "":
                   with plt.rc_context({'mathtext.fontset': "cm"}):
                        self.latexAx.clear()
                        self.latexAx.axis("off")
                        self.latexCanvas.draw()
                elif tmptext2 != '\\' and tmptext2[-1] != '\\':
                    windowWidth = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).width*self.latexFig.dpi
                    windowHeight = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).height*self.latexFig.dpi
                    with plt.rc_context({'mathtext.fontset': "cm"}):
                        self.latexAx.clear()
                        rLatex = self.latexFig.canvas.get_renderer()
                        latexSize = 30
                        latexHeight = 0.5
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                        bb = tLatex.get_window_extent(renderer=rLatex)
                        while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                            self.latexAx.clear()
                            if (latexHeight <= 0.05):
                                latexHeight = 0
                                while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                    self.latexAx.clear()
                                    latexSize -=1
                                    if (latexSize <= 0):
                                        latexSize = 1
                                        break
                                    tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                    bb = tLatex.get_window_extent(renderer=rLatex)
                                break
                            latexHeight -= 0.05
                            tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)
                            bb = tLatex.get_window_extent(renderer=rLatex)
                        while (bb.width > windowWidth):
                            self.latexAx.clear()
                            latexSize -= 1
                            if (latexSize <= 0):
                                latexSize = 1
                                break
                            tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                            bb = tLatex.get_window_extent(renderer=rLatex)
                        self.latexAx.axis("off")
                        self.latexCanvas.draw()
            except Exception:
                pass
        
        def formulaDescriptionPopup(event):
            self.formulaDescriptionWindow.deiconify()
            self.formulaDescriptionWindow.lift()
            self.formulaDescriptionLatexEntry.bind('<KeyRelease>', graphLatex)
            self.formulaDescriptionWindow.bind('<Configure>', graphLatex)
            def on_closing_description():
                self.formulaDescriptionWindow.withdraw()
            
            self.formulaDescriptionWindow.protocol("WM_DELETE_WINDOW", on_closing_description)
        
        def importDirectoryBrowse():
            folder = askdirectory(parent=self.importPathWindow, initialdir=self.topGUI.getCurrentDirectory())
            folder_str = str(folder)
            if (len(folder_str) > 1):
                self.importPathListbox.insert(tk.END, folder_str)
        
        def importPathPopup(event):
            self.importPathWindow.deiconify()
            self.importPathWindow.lift()
            self.importPathButton.configure(command=importDirectoryBrowse)
            def on_closing_importPath():
                self.importPathWindow.withdraw()
            
            self.importPathWindow.protocol("WM_DELETE_WINDOW", on_closing_importPath)
        
        def process_queue_bootstrap():
            try:
                r,s,sdR,sdI,chi,aic,realF,imagF,custom_weight = self.queueBootstrap.get(0)
                self.progStatus.grid_remove()
                self.progStatus.destroy()
                self.simplexButton.configure(state="normal")
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
                self.fitWeightR = np.ones(len(self.wdata))
                self.fitWeightJ = np.ones(len(self.wdata))
                if (self.weightingComboboxValue.get() == "Proportional"):
                    for i in range(len(self.wdata)):
                        self.fitWeightR[i] = float(self.noiseEntryValue.get()) * self.rdata[i]
                    for i in range(len(self.wdata)):
                        self.fitWeightJ[i] = float(self.noiseEntryValue.get()) * self.jdata[i]
                elif (self.weightingComboboxValue.get() == "Modulus"):
                    for i in range(len(self.wdata)):
                        self.fitWeightR[i] = float(self.noiseEntryValue.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                        self.fitWeightJ[i] = float(self.noiseEntryValue.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                elif (self.weightingComboboxValue.get() == "Error model"):
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
                elif (self.weightingComboboxValue.get() == "Custom"):
                    for i in range(len(self.wdata)):
                        self.fitWeightR[i] = custom_weight[i]
                        self.fitWeightJ[i] = custom_weight[i]
                try:
                    self.taskbar.SetProgressState(self.masterWindowId, 0x0)
                except Exception:
                    pass
                if (len(r) == 1):
                    if (r == "b"):
                        messagebox.showerror("Fitting error", "There was an error in the simulations.")
                        return
                    elif (r == "^"):
                        messagebox.showerror("Fitting error", "A solution could not be found.")
                        return
                    elif (r == "@"):
                        return    #The fitting was cancelled
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
                except Exception:
                    self.aR += "Akaike Information Criterion could not be calculated"
                
                self.resultsView.delete(*self.resultsView.get_children())
                self.resultAlert.grid_remove()
                if (len(s) != 1):
                    for i in range(len(r)):
                        if (r[i] == 0):
                            self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "0%"))
                        else:
                            if (abs(s[i]*2*100/r[i]) > 100 or np.isnan(s[i])):
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=3, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"))
                else:
                    for i in range(len(r)):
                        self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "nan", "nan"))
                        self.resultAlert.grid(column=0, row=3, sticky="E")
                self.resultsView.tag_configure("bad", background="yellow")
                self.resultsFrame.grid(column=0, row=5, sticky="W", pady=5)
                self.graphFrame.grid(column=0, row=6, sticky="W", pady=5)
                self.saveFrame.grid(column=0, row=7, sticky="W", pady=5)
                self.realFit = realF
                self.imagFit = imagF
                self.sdrReal = sdR
                self.sdrImag = sdI
                if (len(sdR) == 1):
                    if (sdR == "-"):
                        self.sdrReal = np.zeros(len(self.wdata))
                        self.sdrImag = np.zeros(len(self.wdata))
                self.fits = r
                self.sigmas = s
                self.chiSquared = chi
            except Exception:
                percent_to_step = (len(self.listPercent) - 1)/self.listPercent[0]
                #print(percent_to_step*100)
                self.prog_bar.config(value=percent_to_step*100)
                try:
                    self.taskbar.SetProgressValue(self.masterWindowId, (len(self.listPercent) - 1), self.listPercent[0])
                    self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                except Exception:
                    pass
                if (self.listPercent[len(self.listPercent)-1] == 2):
                    self.progStatus.configure(text="Starting processes...")
                elif (percent_to_step >= 1):
                    self.progStatus.configure(text="Calculating CIs...")
                elif (percent_to_step > 0.001):
                    self.progStatus.configure(text="{:.1%}".format(percent_to_step))
                self.after(100, process_queue_bootstrap)
        
        def bootstrapStdDevs(r,s,sdR,sdI,chi,aic,realF,imagF):
            self.bootstrapLabel.destroy()
            self.bootstrapEntry.destroy()
            self.bootstrapButton.destroy()
            self.bootstrapCancel.destroy()
            self.bootstrapPopup.destroy()
            try:
                bootstrapNum = int(self.bootstrapEntryVariable.get())
                if (bootstrapNum <= 0):
                    raise Exception
            except Exception:
                messagebox.showwarning("Invalid value", "Error 53: \nThe number of simulations must be a positive integer")
                return
            self.queueBootstrap = queue.Queue()
            #---Check if a formula has been entered----
            formula = self.customFormula.get("1.0", tk.END)
            if ("".join(formula.split()) == ""):
                messagebox.showwarning("Error", "Error 44: \nFormula is empty")
                return
            
            #---Check that none of the variable names are repeated, and that none are Python reserved words or variables used in this code
            for i in range(len(self.paramNameValues)):
                name = self.paramNameValues[i].get()
                if (name == "False" or name == "None" or name == "True" or name == "and" or name == "as" or name == "assert" or name == "break" or name == "class" or name == "continue" or name == "def" or name == "del" or name == "elif" or name == "else" or name == "except"\
                     or name == "finally" or name == "for" or name == "from" or name == "global" or name == "if" or name == "import" or name == "in" or name == "is" or name == "lambda" or name == "nonlocal" or name == "not" or name == "or" or name == "pass" or name == "raise"\
                      or name == "return" or name == "try" or name == "while" or name == "with" or name == "yield"):
                    messagebox.showwarning("Error", "Error 45: \nThe variable name \"" + name + "\" is a Python reserved word. Change the variable name.")
                    return
                elif (name == "freq" or name == "Zr" or name == "Zj" or name == "Zreal" or name == "Zimag" or name == "weighting"):
                    messagebox.showwarning("Error", "Error 47: \nThe variable name \"" + name + "\" is used by the fitting program; change the variable name.")
                    return
                for j in range(i+1, len(self.paramNameValues)):
                    if (name == self.paramNameValues[j].get()):
                        messagebox.showwarning("Error", "Error 48: \nTwo or more variables have the same name.")
                        return
            
            #---Replace the functions with np.<function>, and then attempt to compile the code to look for syntax errors---        
            try:
                prebuiltFormulas = ['PI', 'ARCSINH', 'ARCCOSH', 'ARCTANH', 'ARCSIN', 'ARCCOS', 'ARCTAN', 'COSH', 'SINH', 'TANH', 'SIN', 'COS', 'TAN', 'SQRT', 'EXP', 'ABS', 'DEG2RAD', 'RAD2DEG']
                formula = formula.replace("^", "**")    #Replace ^ with ** for exponentiation (this could prevent some features like regex from being used effectively)
                formula = formula.replace("LN", "np.emath.log")#, "cmath.log")
                formula = formula.replace("LOG", "np.emath.log10")
                for pf in prebuiltFormulas:
                    toReplace = "np." + pf.lower()
                    formula = formula.replace(pf, toReplace)
                formula = formula.replace("\n", "\n\t")
                formula = "try:\n\t" + formula
                formula = formula.rstrip()
                formula += "\n\tif any(np.isnan(Zreal)) or any(np.isnan(Zimag)):\n\t\traise Exception\nexcept Exception:\n\tZreal = np.full(len(freq), 1E300)\n\tZimag = np.full(len(freq), 1E300)"
                compile(formula, 'user_generated_formula', 'exec')
            except Exception:
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
            except Exception:
                messagebox.showwarning("Bad number of simulations", "The number of simulations must be an integer")
                return
            param_names = []
            for pNV in self.paramNameValues:
                param_names.append(pNV.get())
            param_limits = []
            for pL in self.paramComboboxValues:
                if (pL.get() == "+"):
                    param_limits.append("+")
                elif (pL.get() == "-"):
                    param_limits.append("-")
                elif (pL.get() == "fixed"):
                    param_limits.append("f")
                elif (pL.get() == "custom"):
                    try:
                        float(self.upperLimits[i].get())
                        float(self.lowerLimits[i].get())
                    except Exception:
                        messagebox.showwarning("Value error", "Error 61: \nThe upper and lower limits must be real numbers")
                        return
                    param_limits.append(str(self.upperLimits[i].get()) + ";" + str(self.lowerLimits[i].get()))
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
                except Exception:
                    messagebox.showerror("Value error", "One of the error structure parameters has an invalid value")
                    return
                if (errorModelParams[0] == 0 and errorModelParams[1] == 0 and errorModelParams[3] == 0 and errorModelParams[4] == 0):
                    messagebox.showerror("Value error", "At least one of the error structure parameters must be nonzero")
                    return
                weight = 3
            elif (self.weightingComboboxValue.get() == "Custom"):
                weight = 4
            assumed_noise = float(self.noiseEntryValue.get())
            self.simplexButton.configure(state="disabled")
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
            self.prog_bar = ttk.Progressbar(self.runFrame, orient="horizontal", length=150, mode="determinate")
            self.prog_bar.grid(column=4, row=0, padx=5)
            self.progStatus = tk.Label(self.runFrame, text="Initializing...", bg=self.backgroundColor, fg=self.foregroundColor)
            self.progStatus.grid(column=5, row=0)
            try:
                cc.GetModule('tl.tlb')
                import comtypes.gen.TaskbarLib as tbl
                self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                self.taskbar.HrInit()
                self.taskbar.ActivateTab(self.masterWindowId)
                self.taskbar.SetProgressState(self.masterWindowId, 0x2)
            except Exception:
                pass
            self.listPercent = [bootstrapNum]
            extra_imports = self.importPathListbox.get(0, tk.END)
            for eI in extra_imports:
                if not os.path.isdir(eI):
                    messagebox.showerror("Path error", "Error 63 :\nThe additional import path \"" + eI + "\" does not exist.")
                    self.simplexButton.configure(state="normal")
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
                    return
            self.currentThreads.append(ThreadedTaskBootstrap(self.queueBootstrap, extra_imports, self.listPercent, bootstrapNum, r,s,sdR,sdI,chi,aic,realF,imagF,fit_type, num_monte, weight, assumed_noise, formula, self.wdata, self.jdata, self.rdata, param_names, self.bestParams, param_limits, errorModelParams))
            self.currentThreads[len(self.currentThreads) - 1].start()
            self.bootstrapRunning = True
            self.cancelButton.configure(command=lambda: self.currentThreads[len(self.currentThreads)-1].terminate())
            self.cancelButton.grid(column=3, row=0, sticky="W", padx=15)
            self.after(100, process_queue_bootstrap)
        
        def process_queue_custom():
            try:
                r,s,sdR,sdI,chi,aic,realF,imagF, bestP, custom_weight = self.queue.get(0)
                self.simplexButton.configure(state="normal")
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
                self.fitWeightR = np.ones(len(self.wdata))
                self.fitWeightJ = np.ones(len(self.wdata))
                if (self.weightingComboboxValue.get() == "Proportional"):
                    for i in range(len(self.wdata)):
                        self.fitWeightR[i] = float(self.noiseEntryValue.get()) * self.rdata[i]
                    for i in range(len(self.wdata)):
                        self.fitWeightJ[i] = float(self.noiseEntryValue.get()) * self.jdata[i]
                elif (self.weightingComboboxValue.get() == "Modulus"):
                    for i in range(len(self.wdata)):
                        self.fitWeightR[i] = float(self.noiseEntryValue.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                        self.fitWeightJ[i] = float(self.noiseEntryValue.get()) * np.sqrt(self.rdata[i]**2 + self.jdata[i]**2)
                elif (self.weightingComboboxValue.get() == "Error model"):
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
                elif (self.weightingComboboxValue.get() == "Custom"):
                    if (len(custom_weight) == 1):
                        messagebox.showwarning("Error", "The error structure could not be calculated")
                        return
                    else:
                        for i in range(len(custom_weight)):
                            self.fitWeightR[i] = custom_weight[i]
                            self.fitWeightJ[i] = custom_weight[i]
                try:
                    self.progStatus.grid_remove()
                    self.progStatus.destroy()
                except Exception:
                    pass
                try:
                    self.taskbar.SetProgressState(self.masterWindowId, 0x0)
                except Exception:
                    pass
                continueResults = True
                if (len(r) == 1):
                    if (r == "b"):
                        messagebox.showerror("Fitting error", "There was an error in the fitting.")
                        continueResults = False
                    elif (r == "^"):
                        messagebox.showerror("Fitting error", "The fitting did not find a solution.")
                        continueResults = False
                    elif (r == "@"):
                        continueResults = False #The fitting was cancelled
                if continueResults:
                    self.bestParams = list(bestP[0])
                    doBootstrap = False
                    if len(s) == 1:
                        if (s == "-"):
                            doBootstrap = messagebox.askyesno("Variance error", "A minimization was found, but the standard deviations and confidence intervals could not be calculated. Would you like to calculate these with a Monte Carlo simulation?")
                    if (doBootstrap):
                        self.bootstrapPopup = tk.Toplevel(bg=self.backgroundColor)
                        self.bootstrapPopup.geometry("400x100")
                        self.bootstrapLabel = tk.Label(self.bootstrapPopup, text="Number of simulations: ", fg=self.foregroundColor, bg=self.backgroundColor)
                        self.bootstrapEntry = ttk.Entry(self.bootstrapPopup, textvariable=self.bootstrapEntryVariable)
                        self.bootstrapButton = ttk.Button(self.bootstrapPopup, text="Run", command=lambda: bootstrapStdDevs(r,s,sdR,sdI,chi,aic,realF,imagF))
                        self.bootstrapCancel = ttk.Button(self.bootstrapPopup, text="Cancel")
                        self.bootstrapLabel.pack(side=tk.TOP, fill=tk.X)
                        self.bootstrapEntry.pack(side=tk.TOP, fill=tk.X)
                        self.bootstrapButton.pack(side=tk.LEFT, expand=True, fill=tk.X)
                        self.bootstrapCancel.pack(side=tk.RIGHT, expand=True, fill=tk.X)
                        self.bootstrapPopup.title("Monte Carlo standard deviations")
                        self.bootstrapPopup.iconbitmap(resource_path("img/elephant3.ico"))
                        self.bootstrapPopup.lift()
                        self.bootstrapPopup.grab_set()
                        self.bootstrapPopup.bind("<Return>", lambda e: bootstrapStdDevs(r,s,sdR,sdI,chi,aic,realF,imagF))
                        def bootstrap_closing():
                            self.bootstrapLabel.destroy()
                            self.bootstrapEntry.destroy()
                            self.bootstrapButton.destroy()
                            self.bootstrapCancel.destroy()
                            self.bootstrapPopup.destroy()
                        self.bootstrapCancel['command'] = bootstrap_closing
                        self.bootstrapPopup.protocol("WM_DELETE_WINDOW", bootstrap_closing)
                        
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
                    except Exception:
                        self.aR += "Akaike Information Criterion could not be calculated"
                    self.resultsView.delete(*self.resultsView.get_children())
                    self.resultAlert.grid_remove()
                    if (s != "-"):
                        for i in range(len(r)):
                            if (r[i] == 0):
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "0%"))
                            if (abs(s[i]*2*100/r[i]) > 100 or np.isnan(s[i])):
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"), tags=("bad",))
                                self.resultAlert.grid(column=0, row=3, sticky="E")
                            else:
                                self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "%.3g"%s[i], "%.3g"%(s[i]*2*100/r[i])+"%"))
                    else:
                        for i in range(len(r)):
                            self.resultsView.insert("", tk.END, text="", values=(self.paramNameValues[i].get(), "%.5g"%r[i], "nan", "nan"))
                            self.resultAlert.grid(column=0, row=3, sticky="E")
                    self.resultsView.tag_configure("bad", background="yellow")
                    if (self.fittingTypeComboboxValue.get() == "Complex"):
                        self.whatFit = "C"
                    elif (self.fittingTypeComboboxValue.get() == "Imaginary"):
                        self.whatFit = "J"
                    elif (self.fittingTypeComboboxValue.get() == "Real"):
                        self.whatFit = "R"
                    self.resultsFrame.grid(column=0, row=5, sticky="W", pady=5)
                    self.graphFrame.grid(column=0, row=6, sticky="W", pady=5)
                    self.saveFrame.grid(column=0, row=7, sticky="W", pady=5)
                    self.realFit = realF
                    self.imagFit = imagF
                    self.sdrReal = sdR
                    self.sdrImag = sdI
                    if (len(sdR) == 1):
                        if (sdR == "-"):
                            self.sdrReal = np.zeros(len(self.wdata))
                            self.sdrImag = np.zeros(len(self.wdata))
                    self.fits = r
                    self.sigmas = s
                    self.chiSquared = chi
            except Exception:
                if (self.listPercent[0] > 1):
                    percent_to_step = (len(self.listPercent) - 1)/self.listPercent[0]
                    self.prog_bar.config(value=percent_to_step*100)
                    try:
                        self.taskbar.SetProgressValue(self.masterWindowId, (len(self.listPercent) - 1), self.listPercent[0])
                        self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                    except Exception:
                        pass
                    if (self.listPercent[len(self.listPercent)-1] == 2):
                        self.progStatus.configure(text="Starting processes...")
                    elif (percent_to_step >= 1):
                        self.progStatus.configure(text="Calculating CIs...")
                    elif (percent_to_step > 0.001):
                        self.progStatus.configure(text="{:.1%}".format(percent_to_step))
                self.after(100, process_queue_custom)
        
        def checkFitting():
            try:
                self.masterWindowId = ctypes.windll.user32.GetForegroundWindow()
            except Exception:
                pass
            #---Check if a formula has been entered----
            formula = self.customFormula.get("1.0", tk.END)
            if ("".join(formula.split()) == ""):
                messagebox.showwarning("Error", "Error 44: \nFormula is empty")
                return
            
            #---Check that none of the variable names are repeated, and that none are Python reserved words or variables used in this code
            for i in range(len(self.paramNameValues)):
                name = self.paramNameValues[i].get()
                if (name == "False" or name == "None" or name == "True" or name == "and" or name == "as" or name == "assert" or name == "break" or name == "class" or name == "continue" or name == "def" or name == "del" or name == "elif" or name == "else" or name == "except"\
                     or name == "finally" or name == "for" or name == "from" or name == "global" or name == "if" or name == "import" or name == "in" or name == "is" or name == "lambda" or name == "nonlocal" or name == "not" or name == "or" or name == "pass" or name == "raise"\
                      or name == "return" or name == "try" or name == "while" or name == "with" or name == "yield"):
                    messagebox.showwarning("Error", "Error 45: \nThe variable name \"" + name + "\" is a Python reserved word. Change the variable name.")
                    return
                elif (name == "freq" or name == "Zr" or name == "Zj" or name == "Zreal" or name == "Zimag" or name == "weighting"):
                    messagebox.showwarning("Error", "Error 47: \nThe variable name \"" + name + "\" is used by the fitting program; change the variable name.")
                    return
                for j in range(i+1, len(self.paramNameValues)):
                    if (name == self.paramNameValues[j].get()):
                        messagebox.showwarning("Error", "Error 48: \nTwo or more variables have the same name.")
                        return
            
            #---Replace the functions with np.<function>, and then attempt to compile the code to look for syntax errors---        
            try:
                prebuiltFormulas = ['PI', 'ARCSINH', 'ARCCOSH', 'ARCTANH', 'ARCSIN', 'ARCCOS', 'ARCTAN', 'COSH', 'SINH', 'TANH', 'SIN', 'COS', 'TAN', 'SQRT', 'EXP', 'ABS', 'DEG2RAD', 'RAD2DEG']
                formula = formula.replace("^", "**")    #Replace ^ with ** for exponentiation (this could prevent some features like regex from being used effectively)
                formula = formula.replace("LN", "np.emath.log")#, "cmath.log")
                formula = formula.replace("LOG", "np.emath.log10")
                for pf in prebuiltFormulas:
                    toReplace = "np." + pf.lower()
                    formula = formula.replace(pf, toReplace)
                formula = formula.replace("\n", "\n\t")
                formula = "try:\n\t" + formula
                formula = formula.rstrip()
                formula += "\n\tif any(np.isnan(Zreal)) or any(np.isnan(Zimag)):\n\t\traise Exception\nexcept Exception:\n\tZreal = np.full(len(freq), 1E300)\n\tZimag = np.full(len(freq), 1E300)"
                compile(formula, 'user_generated_formula', 'exec')
            except Exception:
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
            except Exception:
                messagebox.showwarning("Bad number of simulations", "The number of simulations must be an integer")
                return
            param_names = []
            for pNV in self.paramNameValues:
                param_names.append(pNV.get())
            param_guesses = []
            for i in range(len(self.paramValueValues)):
                param_guesses.append([])
            for i in range(len(self.paramValueValues)):
                try:
                    param_guesses[i].append(float(self.paramValueValues[i].get()))
                except Exception:
                    messagebox.showwarning("Bad parameter value", "Error 54: \nThe parameter values must be real numbers")
                    return
                if (self.multistartCheckboxVariables[i].get()):
                        if (self.multistartSpacingVariables[i].get() == "Linear"):
                            try:
                                float(self.multistartLowerVariables[i].get())
                            except Exception:
                                messagebox.showwarning("Bad lower bound", "Error 55: \nThe lower bound on a multistart must be a real number")
                                return
                            try:
                                float(self.multistartUpperVariables[i].get())
                            except Exception:
                                messagebox.showwarning("Bad upper bound", "Error 56: \nThe upper bound on a multistart must be a real number")
                                return
                            try:
                                if (int(self.multistartNumberVariables[i].get()) <= 0):
                                    raise Exception
                            except Exception:
                                messagebox.showwarning("Bad number of multistarts", "Error 57: \nThe number of multistarts must be a positive integer")
                                return
                            param_guesses[i].extend(np.linspace(float(self.multistartLowerVariables[i].get()), float(self.multistartUpperVariables[i].get()), int(self.multistartNumberVariables[i].get())))
                        elif (self.multistartSpacingVariables[i].get() == "Logarithmic"):
                            try:
                                float(self.multistartLowerVariables[i].get())
                            except Exception:
                                messagebox.showwarning("Bad lower bound", "Error 55: \nThe lower bound on a multistart must be a real number")
                                return
                            try:
                                float(self.multistartUpperVariables[i].get())
                            except Exception:
                                messagebox.showwarning("Bad upper bound", "Error 56: \nThe upper bound on a multistart must be a real number")
                                return
                            try:
                                if (int(self.multistartNumberVariables[i].get()) <= 0):
                                    raise Exception
                            except Exception:
                                messagebox.showwarning("Bad number of multistarts", "Error 57: \nThe number of multistarts must be a positive integer")
                                return
                            if (np.sign(float(self.multistartLowerVariables[i].get())) != np.sign(float(self.multistartUpperVariables[i].get()))):
                                messagebox.showwarning("Value error", "Error 58: \nFor logarithmic spacing, the sign of the upper and lower bounds must be the same")
                                return
                            if (float(self.multistartLowerVariables[i].get()) == 0 or float(self.multistartUpperVariables[i].get()) == 0):
                                messagebox.showwarning("Value error", "Error 59: \nFor logarithmic spacing, neither the upper nor lower bound can be 0")
                                return
                            param_guesses[i].extend(np.geomspace(float(self.multistartLowerVariables[i].get()), float(self.multistartUpperVariables[i].get()), int(self.multistartNumberVariables[i].get())))
                        elif (self.multistartSpacingVariables[i].get() == "Random"):
                            try:
                                float(self.multistartLowerVariables[i].get())
                            except Exception:
                                messagebox.showwarning("Bad lower bound", "Error 55: \nThe lower bound on a multistart must be a real number")
                                return
                            try:
                                float(self.multistartUpperVariables[i].get())
                            except Exception:
                                messagebox.showwarning("Bad upper bound", "Error 56: \nThe upper bound on a multistart must be a real number")
                                return
                            try:
                                if (int(self.multistartNumberVariables[i].get()) <= 0):
                                    raise Exception
                            except Exception:
                                messagebox.showwarning("Bad number of multistarts", "Error 57: \nThe number of multistarts must be a positive integer")
                                return
                            param_guesses[i].extend((float(self.multistartUpperVariables[i].get())-float(self.multistartLowerVariables[i].get()))*np.random.rand(int(self.multistartNumberVariables[i].get()))+float(self.multistartLowerVariables[i].get()))
                        elif (self.multistartSpacingVariables[i].get() == "Custom"):
                            try:
                                param_guesses[i].extend(float(val) for val in [x.strip() for x in self.multistartCustomVariables[i].get().split(',')])
                            except Exception:
                                messagebox.showwarning("Value error", "Error 60: \nThere was a problem with the custom multistart choices")
                                return
            param_limits = []
            for i in range(len(self.paramComboboxValues)): #pL in self.paramComboboxValues:
                pL = self.paramComboboxValues[i]
                if (pL.get() == "+"):
                    param_limits.append("+")
                elif (pL.get() == "-"):
                    param_limits.append("-")
                elif (pL.get() == "fixed"):
                    param_limits.append("f")
                elif (pL.get() == "custom"):
                    try:
                        float(self.upperLimits[i].get())
                        float(self.lowerLimits[i].get())
                    except Exception:
                        messagebox.showwarning("Value error", "Error 61: \nThe upper and lower limits must be real numbers")
                        return
                    param_limits.append(str(self.upperLimits[i].get()) + ";" + str(self.lowerLimits[i].get()))
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
                except Exception:
                    messagebox.showerror("Value error", "One of the error structure parameters has an invalid value")
                    return
                if (errorModelParams[0] == 0 and errorModelParams[1] == 0 and errorModelParams[3] == 0 and errorModelParams[4] == 0):
                    messagebox.showerror("Value error", "At least one of the error structure parameters must be nonzero")
                    return
                weight = 3
            elif (self.weightingComboboxValue.get() == "Custom"):
                weight = 4
            assumed_noise = float(self.noiseEntryValue.get())
            self.simplexButton.configure(state="disabled")
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
            extra_imports = self.importPathListbox.get(0, tk.END)
            for eI in extra_imports:
                if not os.path.isdir(eI):
                    messagebox.showerror("Path error", "Error 63: \nThe additional import path \"" + eI + "\" does not exist.")
                    self.simplexButton.configure(state="normal")
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
                    return
            num_guesses = 1
            for a in param_guesses:
                num_guesses *= len(a)
            self.listPercent = [num_guesses]
            if (num_guesses > 1):
                self.prog_bar = ttk.Progressbar(self.runFrame, orient="horizontal", length=150, mode="determinate")
                self.prog_bar.grid(column=4, row=0, padx=5)
                self.progStatus = tk.Label(self.runFrame, text="Initializing...", bg=self.backgroundColor, fg=self.foregroundColor)
                self.progStatus.grid(column=5, row=0)
                try:
                    cc.GetModule('tl.tlb')
                    import comtypes.gen.TaskbarLib as tbl
                    self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                    self.taskbar.HrInit()
                    self.taskbar.ActivateTab(self.masterWindowId)
                    self.taskbar.SetProgressState(self.masterWindowId, 0x2)
                except Exception:
                    pass
            else:
                self.prog_bar = ttk.Progressbar(self.runFrame, orient="horizontal", length=150, mode="indeterminate")
                self.prog_bar.grid(column=4, row=0, padx=5)
                self.prog_bar.start(40)
                try:
                    cc.GetModule('tl.tlb')
                    import comtypes.gen.TaskbarLib as tbl
                    self.taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
                    self.taskbar.HrInit()
                    self.taskbar.ActivateTab(self.masterWindowId)
                    self.taskbar.SetProgressState(self.masterWindowId, 0x1)
                except Exception:
                    pass
            self.currentThreads.append(ThreadedTaskCustom(self.queue, extra_imports, self.listPercent, fit_type, num_monte, weight, assumed_noise, formula, self.wdata, self.rdata, self.jdata, param_names, param_guesses, param_limits, errorModelParams))
            self.currentThreads[len(self.currentThreads) - 1].start()
            self.cancelButton.configure(command=lambda: self.currentThreads[len(self.currentThreads) - 1].terminate())
            self.cancelButton.grid(column=3, row=0, sticky="W", padx=15)
            self.after(100, process_queue_custom)
        
        def popup_formula(event):
            try:
                self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu.grab_release()
        
        def copyFormula():
            try:
                pyperclip.copy(self.customFormula.get(tk.SEL_FIRST, tk.SEL_LAST))
            except Exception:
                pass
        
        def cutFormula():
            try:
                pyperclip.copy(self.customFormula.get(tk.SEL_FIRST, tk.SEL_LAST))
                self.customFormula.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except Exception:
                pass
        
        def pasteFormula():
            try:
                toPaste = pyperclip.paste()
                self.customFormula.insert(tk.INSERT, toPaste)
                keyup(None)
            except Exception:
                pass
            
        def undoFormula():
            try:
                self.customFormula.edit_undo()
            except Exception:
                pass
            
        def redoFormula():
            try:
                self.customFormula.edit_redo()
            except Exception:
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
            stringToCopy = str(self.browseEntry.get()) + "\n"
            if (self.loadedFormula != ""):
                stringToCopy += str(self.loadFormulaDirectoryEntry.get()) + "/" + str(self.loadedFormula) + ".mmformula\n"
            else:
                stringToCopy += "\n"
            stringToCopy += "Param\tValue\tStd. Dev.\n"
            for child in self.resultsView.get_children():
                resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                stringToCopy += resultName + "\t" + resultValue + "\t" + resultStdDev + "\n"
            chi_squared_over_nu = str(self.chiSquared/(self.lengthOfData*2-len(self.fits)))
            stringToCopy += "Chi^2/nu\t" + chi_squared_over_nu
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
                    whichPlot = "a"
                    if (event.inaxes == self.aplot):
                        resultPlotBig.title("Nyquist Plot")
                        pointsPlot, = larger.plot(self.rdata, -1*self.jdata, "o", color=dataColor)
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
                        whichPlot = "a"
                        rightPoint = max(self.rdata)
                        topPoint = max(-1*self.jdata)
                    elif (event.inaxes == self.bplot):
                        resultPlotBig.title("Real Impedance Plot")
                        pointsPlot, = larger.plot(self.wdata, self.rdata, "o", color=dataColor)
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
                        whichPlot = "b"
                        rightPoint = max(self.wdata)
                        topPoint = max(self.rdata)
                    elif (event.inaxes == self.cplot):
                        resultPlotBig.title("Imaginary Impedance Plot")
                        pointsPlot, = larger.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
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
                        whichPlot = "c"
                        rightPoint = max(self.wdata)
                        topPoint = max(-1*self.jdata)
                    elif (event.inaxes == self.dplot):
                        resultPlotBig.title("|Z| Bode Plot")
                        pointsPlot, = larger.plot(self.wdata, np.sqrt(self.jdata**2 + self.rdata**2), "o", color=dataColor)
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
                        whichPlot = "d"
                        rightPoint = max(self.wdata)
                        topPoint = max(np.sqrt(self.jdata**2 + self.rdata**2))
                    elif (event.inaxes == self.eplot):
                        phase_fit = np.arctan2(ImagFit, RealFit) * (180/np.pi)
                        actual_phase = np.arctan2(self.jdata, self.rdata) * (180/np.pi)
                        resultPlotBig.title("Phase Angle Bode Plot")
                        pointsPlot, = larger.plot(self.wdata, actual_phase, "o", color=dataColor)
                        larger.plot(self.wdata, phase_fit, color=fitColor)
                        if (self.confInt):
                            error_above = np.arctan2((ImagFit+2*self.sdrImag), (RealFit+2*self.sdrReal)) * (180/np.pi)
                            error_below = np.arctan2((ImagFit-2*self.sdrImag), (RealFit-2*self.sdrReal)) * (180/np.pi)
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        larger.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                        larger.set_ylim(bottom=0, top=-90)
                        larger.set_xscale("log")
                        larger.set_title("Phase Angle Bode Plot", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Phase Angle / Degrees", color=self.foregroundColor)
                        whichPlot = "e"
                        rightPoint = max(self.wdata)
                        topPoint = -90
                    elif (event.inaxes == self.hplot):
                        resultPlotBig.title("Log(Zr) vs Log|f|")
                        pointsPlot, = larger.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                        larger.plot(self.wdata, np.log10(abs(RealFit)), color=fitColor)
                        if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i])))
                                error_below[i] = min(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i])))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                        larger.set_xscale("log")
                        larger.set_title("Log|Zr|", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("Log|Zr|", color=self.foregroundColor)
                        whichPlot = "h"
                        rightPoint = max(self.wdata)
                        topPoint = max(np.log10(abs(self.rdata)))
                    elif (event.inaxes == self.iplot):
                         resultPlotBig.title("Log|Zj| vs Log|f|")
                         pointsPlot, = larger.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                         larger.plot(self.wdata, np.log10(abs(ImagFit)), color=fitColor)
                         if (self.confInt):
                            error_above = np.zeros(len(self.wdata))
                            error_below = np.zeros(len(self.wdata))
                            for i in range(len(self.wdata)):
                                error_above[i] = max(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))
                                error_below[i] = min(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))
                            larger.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                         larger.set_xscale("log")
                         larger.set_title("Log|Zj|", color=self.foregroundColor)
                         larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                         larger.set_ylabel("Log|Zj|", color=self.foregroundColor)
                         whichPlot = "i"
                         rightPoint = max(self.wdata)
                         topPoint = max(np.log10(abs(self.jdata)))
                    elif (event.inaxes == self.kplot):
                        normalized_residuals_real = np.zeros(len(self.wdata))
                        normalized_error_real_below = np.zeros(len(self.wdata))
                        normalized_error_real_above = np.zeros(len(self.wdata))
                        errStruct_real_above = np.zeros(len(self.wdata))
                        errStruct_real_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            normalized_residuals_real[i] = (self.rdata[i] - RealFit[i])/RealFit[i]
                            normalized_error_real_below[i] = 2*self.sdrReal[i]/RealFit[i]
                            normalized_error_real_above[i] = -2*self.sdrReal[i]/RealFit[i]
                            errStruct_real_above[i] = 2*self.fitWeightR[i]/self.rdata[i]
                            errStruct_real_below[i] = -2*self.fitWeightR[i]/self.rdata[i]
                        resultPlotBig.title("Real Normalized Residuals")
                        pointsPlot, = larger.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                        if (self.confInt):
                            larger.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                        if (self.errStruct):
                            larger.plot(self.wdata, errStruct_real_above, "--", color="black")
                            larger.plot(self.wdata, errStruct_real_below, "--", color="black")
                        larger.axhline(0, color="black", linewidth=1.0)
                        larger.set_xscale("log")
                        larger.set_title("Real Normalized Residuals", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("(Zr-Z\u0302rmodel)/Zr", color=self.foregroundColor)
                        whichPlot = "k"
                        rightPoint = max(self.wdata)
                        topPoint = max(normalized_residuals_real)
                    elif (event.inaxes == self.lplot):
                        normalized_residuals_imag = np.zeros(len(self.wdata))
                        normalized_error_imag_below = np.zeros(len(self.wdata))
                        normalized_error_imag_above = np.zeros(len(self.wdata))
                        errStruct_imag_above = np.zeros(len(self.wdata))
                        errStruct_imag_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            normalized_residuals_imag[i] = (self.jdata[i] - ImagFit[i])/ImagFit[i]
                            normalized_error_imag_below[i] = 2*self.sdrImag[i]/ImagFit[i]
                            normalized_error_imag_above[i] = -2*self.sdrImag[i]/ImagFit[i]
                            errStruct_imag_above[i] = 2*self.fitWeightJ[i]/self.jdata[i]
                            errStruct_imag_below[i] = -2*self.fitWeightJ[i]/self.jdata[i]
                        resultPlotBig.title("Imaginary Normalized Residuals")
                        pointsPlot, = larger.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                        if (self.confInt):
                            larger.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                            larger.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                        if (self.errStruct):
                            larger.plot(self.wdata, errStruct_imag_above, "--", color="black")
                            larger.plot(self.wdata, errStruct_imag_below, "--", color="black")
                        larger.axhline(0, color="black", linewidth=1.0)
                        larger.set_xscale("log")
                        larger.set_title("Imaginary Normalized Residuals", color=self.foregroundColor)
                        larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                        larger.set_ylabel("(Zj-Z\u0302jmodel)/Zj", color=self.foregroundColor)
                        whichPlot = "l"
                        rightPoint = max(self.wdata)
                        topPoint = max(normalized_residuals_imag)
                    larger.xaxis.label.set_fontsize(20)
                    larger.yaxis.label.set_fontsize(20)
                    larger.title.set_fontsize(30)    
                largerCanvas = FigureCanvasTkAgg(fig, resultPlotBig)
                largerCanvas.draw()
                largerCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                toolbar = NavigationToolbar2Tk(largerCanvas, resultPlotBig)
                toolbar.update()
                largerCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                if (self.mouseoverCheckboxVariable.get() == 1):
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

            except Exception:
                pass
        
        def graphOut(event):
            self.nyCanvas._tkcanvas.config(cursor="arrow")
        
        def graphOver(event):
            whichCan = self.nyCanvas._tkcanvas
            whichCan.config(cursor="hand2")
        
        def plotResults():
            self.resultPlot = tk.Toplevel(background=self.backgroundColor)
            self.resultPlots.append(self.resultPlot)
            self.resultPlot.title("Measurement Model Plots")
            self.resultPlot.iconbitmap(resource_path('img/elephant3.ico'))
            self.resultPlot.state("zoomed")
            self.confInt = False if self.confidenceIntervalCheckboxVariable.get() == 0 else True
            self.errStruct = False if self.errorStructureCheckboxVariable.get() == 0 else True
            RealFit = np.array(self.realFit)
            ImagFit = np.array(self.imagFit)
            phase_fit = np.arctan2(self.imagFit, self.realFit) * (180/np.pi)
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                self.f = Figure()
                self.resultPlotFigs.append(self.f)
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
                self.aplot.set_title("Nyquist", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                self.bplot.set_title("Real Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                self.cplot.set_title("Imaginary Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                self.dplot.set_title("|Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                    error_above = np.arctan2((ImagFit+2*self.sdrImag) , (RealFit+2*self.sdrReal)) * (180/np.pi)
                    error_below = np.arctan2((ImagFit-2*self.sdrImag) , (RealFit-2*self.sdrReal)) * (180/np.pi)
                    self.eplot.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                    self.eplot.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                self.eplot.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                self.eplot.set_ylim(bottom=0, top=-90)
                self.eplot.set_xscale('log')
                self.eplot.set_title("Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                self.hplot.set_title("Log|Zr|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                self.iplot.set_title("Log|Zj|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                self.iplot.set_ylabel("Log|Zj|", color=self.foregroundColor)
                
                normalized_residuals_real = np.zeros(len(self.wdata))
                normalized_error_real_above = np.zeros(len(self.wdata))
                normalized_error_real_below = np.zeros(len(self.wdata))
                normalized_residuals_imag = np.zeros(len(self.wdata))
                normalized_error_imag_above = np.zeros(len(self.wdata))
                normalized_error_imag_below = np.zeros(len(self.wdata))
                errStruct_real_above = np.zeros(len(self.wdata))
                errStruct_real_below = np.zeros(len(self.wdata))
                errStruct_imag_above = np.zeros(len(self.wdata))
                errStruct_imag_below = np.zeros(len(self.wdata))
                for i in range(len(self.wdata)):
                    normalized_residuals_real[i] = (self.rdata[i] - RealFit[i])/RealFit[i]
                    normalized_error_real_above[i] = 2*self.sdrReal[i]/RealFit[i]
                    normalized_error_real_below[i] = -2*self.sdrReal[i]/RealFit[i]
                    normalized_residuals_imag[i] = (self.jdata[i] - ImagFit[i])/ImagFit[i]
                    normalized_error_imag_below[i] = 2*self.sdrImag[i]/ImagFit[i]
                    normalized_error_imag_above[i] = -2*self.sdrImag[i]/ImagFit[i]
                    errStruct_real_above[i] = 2*self.fitWeightR[i]/self.rdata[i]
                    errStruct_real_below[i] = -2*self.fitWeightR[i]/self.rdata[i]
                    errStruct_imag_above[i] = 2*self.fitWeightJ[i]/self.jdata[i]
                    errStruct_imag_below[i] = -2*self.fitWeightJ[i]/self.jdata[i]
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
                if (self.errStruct):
                    self.kplot.plot(self.wdata, errStruct_real_above, "--", color="black")
                    self.kplot.plot(self.wdata, errStruct_real_below, "--", color="black")
                self.kplot.axhline(0, color="black", linewidth=1.0)
                self.kplot.set_xscale('log')
                self.kplot.set_title("Real Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
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
                if (self.errStruct):
                    self.lplot.plot(self.wdata, errStruct_imag_above, "--", color="black")
                    self.lplot.plot(self.wdata, errStruct_imag_below, "--", color="black")
                self.lplot.axhline(0, color="black", linewidth=1.0)
                self.lplot.set_xscale('log')
                self.lplot.set_title("Imaginary Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                self.lplot.set_ylabel("(Zj-Zjmodel)/Zj", color=self.foregroundColor)
                self.nyCanvas = FigureCanvasTkAgg(self.f, self.resultPlot)
                self.f.subplots_adjust(top=0.95, bottom=0.1, right=0.95, left=0.05)
            self.nyCanvas.draw()
            self.nyCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
            self.nyCanvas.mpl_connect('button_press_event', onclick)
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
                    ax_save.plot(RealFit, -1*ImagFit, color=fitColor, marker="x", markeredgecolor="black")
                    if (self.confInt):
                        for i in range(len(self.wdata)):
                            ellipse = patches.Ellipse(xy=(RealFit[i], ImagFit[i]), width=4*self.sdrReal[i], height=4*self.sdrImag[i])
                            ax_save.add_artist(ellipse)
                            ellipse.set_alpha(0.1)
                            ellipse.set_facecolor(self.ellipseColor)
                    ax_save.axis("equal")
                    ax_save.set_title("Nyquist", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Zr / Ω", color=self.foregroundColor)
                    ax_save.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_nyquist_custom.png", dpi=300)
                    ax_save.clear()
                    ax_save.axis("auto")
                    #---bplot---
                    ax_save.set_facecolor(self.backgroundColor)
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.plot(self.wdata, self.rdata, "o", color=dataColor)
                    ax_save.plot(self.wdata, RealFit, color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(RealFit[i] + 2*self.sdrReal[i], RealFit[i] - 2*self.sdrReal[i])
                            error_below[i] = min(RealFit[i] + 2*self.sdrReal[i], RealFit[i] - 2*self.sdrReal[i])
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Real Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_ylabel("Zr / Ω", color=self.foregroundColor)
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_real_impedance_custom.png", dpi=300)
                    ax_save.clear()
                    #---cplot---
                    ax_save.plot(self.wdata, -1*self.jdata, "o", color=dataColor)
                    ax_save.plot(self.wdata, -1*ImagFit, color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(ImagFit[i] + 2*self.sdrImag[i], ImagFit[i] - 2*self.sdrImag[i])
                            error_below[i] = min(ImagFit[i] + 2*self.sdrImag[i], ImagFit[i] - 2*self.sdrImag[i])
                        ax_save.plot(self.wdata, -1*error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, -1*error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Imaginary Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_imag_impedance_custom.png", dpi=300)
                    ax_save.clear()
                    #---dplot---
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.plot(self.wdata, np.sqrt(self.rdata**2 + self.jdata**2), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.sqrt(RealFit**2 + ImagFit**2), color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = np.sqrt(max((RealFit[i]+2*self.sdrReal[i])**2, (RealFit[i]-2*self.sdrReal[i])**2) + max((ImagFit[i]+2*self.sdrImag[i])**2, (ImagFit[i]-2*self.sdrImag[i])**2))
                            error_below[i] = np.sqrt(min((RealFit[i]+2*self.sdrReal[i])**2, (RealFit[i]-2*self.sdrReal[i])**2) + min((ImagFit[i]+2*self.sdrImag[i])**2, (ImagFit[i]-2*self.sdrImag[i])**2))
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_yscale('log')
                    ax_save.set_title("|Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_Z_bode_custom.png", dpi=300)
                    ax_save.clear()
                    #---eplot---
                    ax_save.yaxis.set_ticks_position("both")
                    ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    ax_save.xaxis.set_ticks_position("both")
                    ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    ax_save.plot(self.wdata, np.arctan2(self.jdata, self.rdata)*(180/np.pi), "o", color=dataColor)
                    ax_save.plot(self.wdata, phase_fit, color=fitColor)
                    if (self.confInt):
                        error_above = np.arctan2((ImagFit+2*self.sdrImag) , (RealFit+2*self.sdrReal)) * (180/np.pi)
                        error_below = np.arctan2((ImagFit-2*self.sdrImag) , (RealFit-2*self.sdrReal)) * (180/np.pi)
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                    ax_save.set_ylim(bottom=0, top=-90)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Phase angle / Deg.", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_phase_angle_bode_custom.png", dpi=300)
                    ax_save.clear()
                    #---hplot---
                    ax_save.plot(self.wdata, np.log10(abs(self.rdata)), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.log10(abs(RealFit)), color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i]))) #np.log10(max(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                            error_below[i] = min(np.log10(abs(RealFit[i]+2*self.sdrReal[i])), np.log10(abs(RealFit[i]-2*self.sdrReal[i])))#np.log10(min(abs(Zfit.real[i]+2*self.sdrReal[i]), abs(Zfit.real[i]-2*self.sdrReal[i])))
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Log|Zr|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Log|Zr|", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_log_zr_custom.png", dpi=300)
                    ax_save.clear()
                    #---iplot---
                    ax_save.plot(self.wdata, np.log10(abs(self.jdata)), "o", color=dataColor)
                    ax_save.plot(self.wdata, np.log10(abs(ImagFit)), color=fitColor)
                    if (self.confInt):
                        error_above = np.zeros(len(self.wdata))
                        error_below = np.zeros(len(self.wdata))
                        for i in range(len(self.wdata)):
                            error_above[i] = max(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))#np.log10(max(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                            error_below[i] = min(np.log10(abs(ImagFit[i]+2*self.sdrImag[i])), np.log10(abs(ImagFit[i]-2*self.sdrImag[i])))#np.log10(min(abs(Zfit.imag[i]+2*self.sdrImag[i]), abs(Zfit.imag[i]-2*self.sdrImag[i])))
                        ax_save.plot(self.wdata, error_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, error_below, "--", color=self.ellipseColor)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Log|Zj|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("Log|Zj|", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_log_zj_custom.png", dpi=300)
                    ax_save.clear()
                    #---kplot---
                    ax_save.plot(self.wdata, normalized_residuals_real, "o", markerfacecolor="None", color=dataColor)
                    if (self.confInt):
                        ax_save.plot(self.wdata, normalized_error_real_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, normalized_error_real_below, "--", color=self.ellipseColor)
                    if (self.errStruct):
                        ax_save.plot(self.wdata, errStruct_real_above, "--", color="black")
                        ax_save.plot(self.wdata, errStruct_real_below, "--", color="black")
                    ax_save.axhline(0, color="black", linewidth=1.0)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Real Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("(Zr-Zrmodel)/Zr", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_real_residuals_custom.png", dpi=300)
                    ax_save.clear()
                    #---lplot---
                    ax_save.plot(self.wdata, normalized_residuals_imag, "o", markerfacecolor="None", color=dataColor)
                    if (self.confInt):
                        ax_save.plot(self.wdata, normalized_error_imag_above, "--", color=self.ellipseColor)
                        ax_save.plot(self.wdata, normalized_error_imag_below, "--", color=self.ellipseColor)
                    if (self.errStruct):
                        ax_save.plot(self.wdata, errStruct_imag_above, "--", color="black")
                        ax_save.plot(self.wdata, errStruct_imag_below, "--", color="black")
                    ax_save.axhline(0, color="black", linewidth=1.0)
                    ax_save.set_xscale('log')
                    ax_save.set_title("Imaginary Residuals", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    ax_save.set_ylabel("(Zj-Zjmodel)/Zj", color=self.foregroundColor)
                    pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_imag_residuals_custom.png", dpi=300)
                    ax_save.clear()
                    self.saveNyCanvasButton.configure(text="Saved")
                    self.after(500, lambda: self.saveNyCanvasButton.configure(text="Save All"))
            self.saveNyCanvasButton = ttk.Button(self.resultPlot, text="Save All", command=saveAllPlots)
            self.saveNyCanvasButtons.append(self.saveNyCanvasButton)
            saveNyCanvasButton_ttp = CreateToolTip(self.saveNyCanvasButton, "Save all plots")
            self.saveNyCanvasButton_ttps.append(saveNyCanvasButton_ttp)
            self.saveNyCanvasButton.pack(side=tk.BOTTOM, pady=5, expand=False)
            enterAxes = self.f.canvas.mpl_connect('axes_enter_event', graphOver)
            leaveAxes = self.f.canvas.mpl_connect('axes_leave_event', graphOut)
            def on_closing():   #Clear the figure before closing the popup
                self.f.clear()
                self.resultPlot.destroy()
                self.saveNyCanvasButton.destroy()
                nonlocal saveNyCanvasButton_ttp
                del saveNyCanvasButton_ttp
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
                stringToSave = "version: 1.2\n"
                if (self.browseEntry.get() == ""):
                    stringToSave += "filename: NONE\n"
                else:
                    stringToSave += "filename: " + self.browseEntry.get() + "\n"
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
                extra_imports = self.importPathListbox.get(0, tk.END)
                stringToSave += "imports: \n"
                for eI in extra_imports:
                    stringToSave += eI + "\n"
                stringToSave += "description: \n"
                stringToSave += self.formulaDescriptionEntry.get("1.0", tk.END)
                stringToSave += "latex: \n"
                stringToSave += self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
                stringToSave += "\nmmparams: \n"
                for i in range(len(self.paramNameValues)):
                    stringToSave += self.paramNameValues[i].get() + "~=~" + self.paramValueValues[i].get() + "~=~" + self.paramComboboxValues[i].get() + "~=~" + self.upperLimits[i].get() + "~=~" + self.lowerLimits[i].get() + "~=~" + str(self.multistartCheckboxVariables[i].get()) + "~=~" + self.multistartSpacingVariables[i].get() + "~=~" + self.multistartUpperVariables[i].get() + "~=~" + self.multistartLowerVariables[i].get() + "~=~" + self.multistartNumberVariables[i].get() + "~=~" + self.multistartCustomVariables[i].get() + "\n"
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
        
        def saveFormula():
            stringToSave = "version: 1.2\n"
            if (self.browseEntry.get() == ""):
                stringToSave += "filename: NONE\n"
            else:
                stringToSave += "filename: " + self.browseEntry.get() + "\n"
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
            extra_imports = self.importPathListbox.get(0, tk.END)
            stringToSave += "imports: \n"
            for eI in extra_imports:
                stringToSave += eI + "\n"
            stringToSave += "description: \n"
            stringToSave += self.formulaDescriptionEntry.get("1.0", tk.END)
            stringToSave += "latex: \n"
            stringToSave += self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            stringToSave += "\nmmparams: \n"
            for i in range(len(self.paramNameValues)):
                stringToSave += self.paramNameValues[i].get() + "~=~" + self.paramValueValues[i].get() + "~=~" + self.paramComboboxValues[i].get() + "~=~" + self.upperLimits[i].get() + "~=~" + self.lowerLimits[i].get() + "~=~" + str(self.multistartCheckboxVariables[i].get()) + "~=~" + self.multistartSpacingVariables[i].get() + "~=~" + self.multistartUpperVariables[i].get() + "~=~" + self.multistartLowerVariables[i].get() + "~=~" + self.multistartNumberVariables[i].get() + "~=~" + self.multistartCustomVariables[i].get() + "\n"
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
            saveName = asksaveasfile(mode='w', defaultextension=".mmformula", initialdir=self.loadFormulaDirectoryEntry.get(), initialfile=self.loadedFormula, filetypes=[("Measurement model custom formula", ".mmformula")])
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if saveName is None:     #If save is cancelled
                return
            saveName.write(stringToSave)
            saveName.close()
            self.loadFormulaTree.delete(*self.loadFormulaTree.get_children())
            self.insert_node('', self.loadFormulaDirectoryEntry.get(), self.loadFormulaDirectoryEntry.get())
            self.formulaSaveButton.configure(text="Saved")
            self.after(1000, lambda : self.formulaSaveButton.configure(text="Save formula"))
        
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
                stringToSave += str(self.wdata[i]) + "\t" + str(self.rdata[i]) + "\t" + str(self.jdata[i]) + "\t" + str(self.realFit[i]) + "\t" + str(self.imagFit[i])
                try:
                    stringToSave += "\t"  + str(self.sdrReal[i]) + "\t" + str(self.sdrImag[i]) + "\n"
                except Exception:
                    stringToSave += "\n"
            stringToSave += "----------------------------------------------------------------------------------\n"
            stringToSave += "File name: " + str(self.browseEntry.get()) + "\n"
            if (self.loadedFormula != ""):
                stringToSave += "Formula name: " + str(self.loadFormulaDirectoryEntry.get()) + "/" + str(self.loadedFormula) + ".mmformula\n"
            else:
                stringToSave += "No loaded formula\n"
            for i in range(len(self.paramNameValues)):
                 stringToSave += self.paramNameValues[i].get() + " = " + str(self.fits[i]) 
                 try:
                     if (self.sigmas[i] == "-"):
                         stringToSave += "\tStd. Dev. = nan\n"
                     else:
                         stringToSave += "\tStd. Dev. = " + str(self.sigmas[i]) + "\n"
                 except Exception:
                     stringToSave += "\tStd. Dev. = nan\n"
            stringToSave += "Chi-squared = " + str(self.chiSquared) + "\n"
            stringToSave += "Chi-squared/Degrees of freedom = " + str(self.chiSquared/(self.lengthOfData*2-len(self.fits)))
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.browseEntry.get()))
            defaultSaveName += "_custom"            
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
        
        def loadFormulaDirectory():
            folder = askdirectory(initialdir=self.loadFormulaDirectoryEntry.get(), parent=self.loadFormulaWindow)
            folder_str = str(folder)
            if len(folder_str) == 0:
                pass
            else:
                self.loadFormulaDirectoryEntry.configure(state="normal")
                self.loadFormulaDirectoryEntry.delete(0,tk.END)
                self.loadFormulaDirectoryEntry.insert(0, folder_str)
                self.loadFormulaDirectoryEntry.configure(state="readonly")
                self.loadFormulaTree.delete(*self.loadFormulaTree.get_children())
                self.insert_node('', folder_str, folder_str)
        
        def loadCode():
            self.loadFormulaWindow.deiconify()
            self.loadFormulaWindow.lift()
            self.loadFormulaDirectoryButton.configure(command=loadFormulaDirectory)
            self.loadFormulaWindow.bind('<Configure>', self.load_graphLatex)
            def on_closing_loadCode():
                self.loadFormulaWindow.withdraw()
            self.loadFormulaWindow.protocol("WM_DELETE_WINDOW", on_closing_loadCode)
        
        def calculateNumPoints(e=None):
            try:
                with np.errstate(divide='raise', invalid='raise'):
                    num_decades = np.log10(float(self.simulationUpperFreq.get())) - np.log10(float(self.simulationLowerFreq.get()))
                num_points = int(num_decades) * int(self.simulationNumFreq.get())
                if (num_points < 0):
                    self.numberOfPointsLabel.configure(text="Total number of points: ")
                else:
                    self.numberOfPointsLabel.configure(text="Total number of points: " + str(num_points))
            except Exception:
                self.numberOfPointsLabel.configure(text="Total number of points: ")
        
        def runSimulations(e=None):
            try:
                with np.errstate(divide='raise', invalid='raise'):
                    num_decades = np.log10(float(self.simulationUpperFreq.get())) - np.log10(float(self.simulationLowerFreq.get()))
                num_points = int(num_decades) * int(self.simulationNumFreq.get())
            except Exception:
                messagebox.showwarning("Value error", "Error 64: \nThe upper and lower frequencies must be real numbers, and the points per decade must be an integer", parent=self.simulationWindow)
                return
            if (num_points <= 0):
                messagebox.showwarning("Value error", "Error 65: \nThe total number of points must be positive (i.e. the points per decade must be positive and the upper frequency must be greater than the lower frequency)", parent = self.simulationWindow)
                return
            self.sim_freqs = np.logspace(np.log10(float(self.simulationLowerFreq.get())), np.log10(float(self.simulationUpperFreq.get())), num_points)
            
            #---Check if a formula has been entered----
            formula = self.customFormula.get("1.0", tk.END)
            if ("".join(formula.split()) == ""):
                messagebox.showwarning("Error", "Error 44: \nFormula is empty", parent=self.simulationWindow)
                return
            
            #---Check that none of the variable names are repeated, and that none are Python reserved words or variables used in this code
            for i in range(len(self.paramNameValues)):
                name = self.paramNameValues[i].get()
                if (name == "False" or name == "None" or name == "True" or name == "and" or name == "as" or name == "assert" or name == "break" or name == "class" or name == "continue" or name == "def" or name == "del" or name == "elif" or name == "else" or name == "except"\
                     or name == "finally" or name == "for" or name == "from" or name == "global" or name == "if" or name == "import" or name == "in" or name == "is" or name == "lambda" or name == "nonlocal" or name == "not" or name == "or" or name == "pass" or name == "raise"\
                      or name == "return" or name == "try" or name == "while" or name == "with" or name == "yield"):
                    messagebox.showwarning("Error", "Error 45: \nThe variable name \"" + name + "\" is a Python reserved word. Change the variable name.")
                    return
                elif (name == "freq"  or name == "Zreal" or name == "Zimag" or name == "weighting"):
                    messagebox.showwarning("Error", "Error 47: \nThe variable name \"" + name + "\" is used by the fitting program; change the variable name.")
                    return
                for j in range(i+1, len(self.paramNameValues)):
                    if (name == self.paramNameValues[j].get()):
                        messagebox.showwarning("Error", "Error 48: \nTwo or more variables have the same name.")
                        return
            #---Replace the functions with np.<function>, and then attempt to compile the code to look for syntax errors---        
            try:
                prebuiltFormulas = ['PI', 'ARCSINH', 'ARCCOSH', 'ARCTANH', 'ARCSIN', 'ARCCOS', 'ARCTAN', 'COSH', 'SINH', 'TANH', 'SIN', 'COS', 'TAN', 'SQRT', 'EXP', 'ABS', 'DEG2RAD', 'RAD2DEG']
                formula = formula.replace("^", "**")    #Replace ^ with ** for exponentiation (this could prevent some features like regex from being used effectively)
                formula = formula.replace("LN", "np.emath.log")
                formula = formula.replace("LOG", "np.emath.log10")
                for pf in prebuiltFormulas:
                    toReplace = "np." + pf.lower()
                    formula = formula.replace(pf, toReplace)
                formula = formula.replace("\n", "\n\t")
                formula = "try:\n\t" + formula
                formula = formula.rstrip()
                formula += "\n\tif any(np.isnan(Zreal)) or any(np.isnan(Zimag)):\n\t\traise Exception\nexcept Exception:\n\tZreal = np.full(len(freq), 1E300)\n\tZimag = np.full(len(freq), 1E300)"
                compile(formula, 'user_generated_formula', 'exec')
            except Exception:
                messagebox.showwarning("Compile error", "There was an issue compiling the code", parent=self.simulationWindow)
                return
            
            #---Check if the variable "freq" appears in the code, as it's likely a mistake if it doesn't---
            textToSearch = self.customFormula.get("1.0", tk.END)
            whereFreq = [m.start() for m in re.finditer(r'\bfreq\b', textToSearch)]
            if (len(whereFreq) == 0):
                messagebox.showwarning("No \"freq\"", "The variable \"freq\" does not appear in the code. This may cause an error.", parent=self.simulationWindow)
            elif (len(whereFreq) == 1 and whereFreq[0] == 5):
                messagebox.showwarning("No \"freq\"", "The variable \"freq\" does not seem to be in the code. This may cause an error.", parent=self.simulationWindow)
            whereZr = [m.start() for m in re.finditer(r'\bZr\b', textToSearch)]
            whereZj = [m.start() for m in re.finditer(r'\bZj\b', textToSearch)]
            if (len(whereZr) != 0 or len(whereZj) != 0):
                messagebox.showwarning("Impedance referenced", "The code references Zr and/or Zj; as simulations are being performed at arbitrary frequencies, the built-in variables holding the data values are not available", parent=self.simulationWindow)
            param_names = []
            for pNV in self.paramNameValues:
                param_names.append(pNV.get())
            param_values = []
            for i in range(len(self.paramValueValues)):
                try:
                    param_values.append(float(self.paramValueValues[i].get()))
                except Exception:
                    messagebox.showwarning("Bad parameter value", "Error 54: \nThe parameter values must be real numbers", parent=self.simulationWindow)
                    return
            self.sim_real, self.sim_imag = dataSimulation.run_simulations(self.sim_freqs, formula, param_names, param_values)
            if (len(self.sim_real) == 1):
                messagebox.showerror("Simulation error", "There was an error with the simulations", parent=self.simulationWindow)
                return
            else:
                self.simView.delete(*self.simView.get_children())
                for i in range(len(self.sim_freqs)):
                    self.simView.insert("", tk.END, text=str(i+1), values=(str(self.sim_freqs[i]), str(self.sim_real[i]), str(self.sim_imag[i]))) #"%E"%self.sim_freqs[i], "%G"%self.sim_real[i], "%G"%self.sim_imag[i]))
                self.simViewFrame.pack(side=tk.TOP, fill=tk.X, pady=(0,5), padx=3, expand=True)
                self.saveSimulationFrame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5), padx=3, expand=True)
        
        def copySimulations(e=None):
            self.copySimulationButton.configure(text="Copied")
            toCopy = "Frequency\tReal\tImaginary"
            for i in range(len(self.sim_freqs)):
                toCopy += "\n" + str(self.sim_freqs[i]) + "\t" + str(self.sim_real[i]) + "\t" + str(self.sim_imag[i])
            pyperclip.copy(toCopy)
            self.after(500, lambda: self.copySimulationButton.configure(text="Copy values as spreadsheet"))
        
        def saveSimulations(e=None):
            if (self.loadedFormula != ""):
                defaultSaveName = self.loadedFormula + "_synthetic"
            else:
                defaultSaveName = "synthetic_data"
            saveName = asksaveasfile(title="Save Synthetic Data", mode='w', defaultextension=".mmfile", parent=self.simulationWindow, initialfile=defaultSaveName, filetypes=[("Measurement model file (*.mmfile)", ".mmfile"), ("Tab delimited (*.txt)", ".txt"), ("Comma separated (*.csv)", ".csv")])
            if saveName is None:     #If save is cancelled
                return
            filename, file_extension = os.path.splitext(str(saveName))
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if "txt" in file_extension:
                stringToSave = "Frequency\tReal\tImaginary"
                for i in range(len(self.sim_freqs)):
                    stringToSave += "\n" + str(self.sim_freqs[i]) + "\t" + str(self.sim_real[i]) + "\t" + str(self.sim_imag[i])
            elif "mmfile" in file_extension:
                stringToSave = ""
                for i in range(len(self.sim_freqs)):
                    stringToSave += str(self.sim_freqs[i]) + "\t" + str(self.sim_real[i]) + "\t" + str(self.sim_imag[i]) + "\n"
            elif "csv" in file_extension:
                stringToSave = "Frequency,Real,Imaginary"
                for i in range(len(self.sim_freqs)):
                    stringToSave += "\n" + str(self.sim_freqs[i]) + "," + str(self.sim_real[i]) + "," + str(self.sim_imag[i])
            saveName.write(stringToSave)
            saveName.close()
            self.saveSimulationButton.configure(text="Saved")
            self.after(1000, lambda : self.saveSimulationButton.configure(text="Save"))
        
        def plotSimulations(e=None):
            for s in self.simPlots:
                try:
                    s.destroy()
                except Exception:
                    pass
            for s in self.simPlotFigs:
                try:
                    s.clear()
                    plt.close(s)
                except Exception:
                    pass
            for s in self.simPlotBigs:
                try:
                    s.destroy()
                except Exception:
                    pass
            for s in self.simPlotBigFigs:
                try:
                    s.clear()
                    plt.close(s)
                except Exception:
                    pass
            for s in self.simSaveNyCanvasButtons:
                try:
                    s.destroy()
                except Exception:
                    pass
            for s in self.simSaveNyCanvasButton_ttps:
                try:
                    del s
                except Exception:
                    pass
            simPlot = tk.Toplevel(background=self.backgroundColor)
            self.simPlots.append(simPlot)
            simPlot.title("Synthetic Data Plots")
            simPlot.iconbitmap(resource_path('img/elephant3.ico'))
            simPlot.state("zoomed")
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                sim_pltFig = Figure()
                self.simPlotFigs.append(sim_pltFig)
                sim_pltFig.set_facecolor(self.backgroundColor)
                dataColor = "tab:blue"
                if (self.topGUI.getTheme() == "dark"):
                    dataColor = "cyan"
                else:
                    dataColor = "tab:blue"
                aplot = sim_pltFig.add_subplot(331)
                aplot.set_facecolor(self.backgroundColor)
                aplot.yaxis.set_ticks_position("both")
                aplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                aplot.xaxis.set_ticks_position("both")
                aplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in")
                aplot.plot(self.sim_real, -1*self.sim_imag, "o",  color=dataColor)
                aplot.axis("equal")
                aplot.set_title("Nyquist", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                aplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                
                bplot = sim_pltFig.add_subplot(332)
                bplot.set_facecolor(self.backgroundColor)
                bplot.yaxis.set_ticks_position("both")
                bplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                bplot.xaxis.set_ticks_position("both")
                bplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                bplot.plot(self.sim_freqs, self.sim_real, "o", color=dataColor)
                bplot.set_xscale('log')
                bplot.set_title("Real Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                bplot.set_ylabel("Zr / Ω", color=self.foregroundColor)
                
                cplot = sim_pltFig.add_subplot(333)
                cplot.set_facecolor(self.backgroundColor)
                cplot.yaxis.set_ticks_position("both")
                cplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                cplot.xaxis.set_ticks_position("both")
                cplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                cplot.plot(self.sim_freqs, -1*self.sim_imag, "o", color=dataColor)
                cplot.set_xscale('log')
                cplot.set_title("Imaginary Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                cplot.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                
                dplot = sim_pltFig.add_subplot(323)
                dplot.set_facecolor(self.backgroundColor)
                dplot.yaxis.set_ticks_position("both")
                dplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                dplot.xaxis.set_ticks_position("both")
                dplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                dplot.plot(self.sim_freqs, np.sqrt(self.sim_real**2 + self.sim_imag**2), "o", color=dataColor)
                dplot.set_xscale('log')
                dplot.set_yscale('log')
                dplot.set_title("|Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                dplot.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                
                eplot = sim_pltFig.add_subplot(324)
                eplot.set_facecolor(self.backgroundColor)
                eplot.yaxis.set_ticks_position("both")
                eplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                eplot.xaxis.set_ticks_position("both")
                eplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                eplot.plot(self.sim_freqs, np.arctan2(self.sim_imag, self.sim_real)*(180/np.pi), "o", color=dataColor)
                eplot.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                eplot.set_ylim(bottom=0, top=-90)
                eplot.set_xscale('log')
                eplot.set_title("Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                eplot.set_ylabel("Phase angle / Deg.", color=self.foregroundColor)
                
                hplot = sim_pltFig.add_subplot(325)
                hplot.set_facecolor(self.backgroundColor)
                hplot.yaxis.set_ticks_position("both")
                hplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                hplot.xaxis.set_ticks_position("both")
                hplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                hplot.plot(self.sim_freqs, np.log10(abs(self.sim_real)), "o", color=dataColor)
                hplot.set_xscale('log')
                hplot.set_title("Log|Zr|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                hplot.set_ylabel("Log|Zr|", color=self.foregroundColor)
                
                iplot = sim_pltFig.add_subplot(326)
                iplot.set_facecolor(self.backgroundColor)
                iplot.yaxis.set_ticks_position("both")
                iplot.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                iplot.xaxis.set_ticks_position("both")
                iplot.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                iplot.plot(self.sim_freqs, np.log10(abs(self.sim_imag)), "o", color=dataColor)
                iplot.set_xscale('log')
                iplot.set_title("Log|Zj|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                iplot.set_ylabel("Log|Zj|", color=self.foregroundColor)
            
            def sim_onclick(event):
                try:
                    if (not event.inaxes):      #If a subplot isn't clicked
                        return
                    sim_resultPlotBig = tk.Toplevel(background=self.backgroundColor)
                    self.simPlotBigs.append(sim_resultPlotBig)
                    sim_resultPlotBig.iconbitmap(resource_path('img/elephant3.ico'))
                    w, h = self.winfo_screenwidth(), self.winfo_screenheight()
                    sim_resultPlotBig.geometry("%dx%d+0+0" % (w/2, h/2))
                    with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                        sim_fig = Figure()
                        self.simPlotBigFigs.append(sim_fig)
                        sim_fig.set_facecolor(self.backgroundColor)
                        larger = sim_fig.add_subplot(111)
                        larger.set_facecolor(self.backgroundColor)
                        larger.yaxis.set_ticks_position("both")
                        larger.yaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                        larger.xaxis.set_ticks_position("both")
                        larger.xaxis.set_tick_params(direction="in", which="both", color=self.foregroundColor)
                        whichPlot = "a"
                        dataColor = "tab:blue"
                        if (self.topGUI.getTheme() == "dark"):
                            dataColor = "cyan"
                        else:
                            dataColor = "tab:blue"
                        if (event.inaxes == aplot):
                            sim_resultPlotBig.title("Nyquist Plot")
                            pointsPlot, = larger.plot(self.sim_real, -1*self.sim_imag, "o", color=dataColor)
                            rightPoint = max(self.sim_real)
                            topPoint = max(-1*self.sim_imag)
                            larger.axis("equal")
                            larger.set_title("Nyquist Plot", color=self.foregroundColor)
                            larger.set_xlabel("Zr / Ω", color=self.foregroundColor)
                            larger.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                            whichPlot = "a"
                        elif (event.inaxes == bplot):
                            sim_resultPlotBig.title("Real Impedance Plot")
                            pointsPlot, = larger.plot(self.sim_freqs, self.sim_real, "o", color=dataColor, zorder=2)
                            rightPoint = max(self.sim_freqs)
                            topPoint = max(self.sim_real)
                            larger.set_xscale("log")
                            larger.set_title("Real Impedance Plot", color=self.foregroundColor)
                            larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                            larger.set_ylabel("Zr / Ω", color=self.foregroundColor)
                            whichPlot = "b"
                        elif (event.inaxes == cplot):
                            sim_resultPlotBig.title("Imaginary Impedance Plot")
                            pointsPlot, = larger.plot(self.sim_freqs, -1*self.sim_imag, "o", color=dataColor)
                            rightPoint = max(self.sim_freqs)
                            topPoint = max(-1*self.sim_imag)
                            larger.set_xscale("log")
                            larger.set_title("Imaginary Impedance Plot", color=self.foregroundColor)
                            larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                            larger.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                            whichPlot = "c"
                        elif (event.inaxes == dplot):
                            sim_resultPlotBig.title("|Z| Bode Plot")
                            pointsPlot, = larger.plot(self.sim_freqs, np.sqrt(self.sim_imag**2 + self.sim_real**2), "o", color=dataColor)
                            rightPoint = max(self.sim_freqs)
                            topPoint = max(np.sqrt(self.sim_imag**2 + self.sim_real**2))
                            larger.set_xscale("log")
                            larger.set_yscale("log")
                            larger.set_title("|Z| Bode Plot", color=self.foregroundColor)
                            larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                            larger.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                            whichPlot = "d"
                        elif (event.inaxes == eplot):
                            actual_phase = np.arctan2(self.sim_imag, self.sim_real) * (180/np.pi)
                            sim_resultPlotBig.title("Phase Angle Bode Plot")
                            pointsPlot, = larger.plot(self.sim_freqs, actual_phase, "o", color=dataColor)
                            larger.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                            larger.set_ylim(bottom=0, top=-90)
                            rightPoint = max(self.sim_freqs)
                            topPoint = -90
                            larger.set_xscale("log")
                            larger.set_title("Phase Angle Bode Plot", color=self.foregroundColor)
                            larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                            larger.set_ylabel("Phase Angle / Degrees", color=self.foregroundColor)
                            whichPlot = "e"
                        elif (event.inaxes == hplot):
                            sim_resultPlotBig.title("Log(Zr) vs f")
                            pointsPlot, = larger.plot(self.sim_freqs, np.log10(abs(self.sim_real)), "o", color=dataColor)
                            rightPoint = max(self.sim_freqs)
                            topPoint = max(np.log10(abs(self.sim_real)))
                            larger.set_xscale("log")
                            larger.set_title("Log|Zr|", color=self.foregroundColor)
                            larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                            larger.set_ylabel("Log|Zr|", color=self.foregroundColor)
                            whichPlot = "h"
                        elif (event.inaxes == iplot):
                             sim_resultPlotBig.title("Log|Zj| vs f")
                             pointsPlot, = larger.plot(self.sim_freqs, np.log10(abs(self.sim_imag)), "o", color=dataColor)
                             rightPoint = max(self.sim_freqs)
                             topPoint = max(np.log10(abs(self.sim_imag)))
                             larger.set_xscale("log")
                             larger.set_title("Log|Zj|", color=self.foregroundColor)
                             larger.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                             larger.set_ylabel("Log|Zj|", color=self.foregroundColor)
                             whichPlot = "i"
                        larger.xaxis.label.set_fontsize(20)
                        larger.yaxis.label.set_fontsize(20)
                        larger.title.set_fontsize(30)
                        largerCanvas = FigureCanvasTkAgg(sim_fig, sim_resultPlotBig)
                        largerCanvas.draw()
                        largerCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                        toolbar = NavigationToolbar2Tk(largerCanvas, sim_resultPlotBig)
                        toolbar.update()
                        largerCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                        if (self.plotSimulationCheckboxVariable.get() == 1):
                            annot = larger.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=1), arrowprops=dict(arrowstyle="-"))
                            annot.set_visible(False)
                            def update_annot(ind):
                                x,y = pointsPlot.get_data()
                                xval = x[ind["ind"][0]]
                                yval = y[ind["ind"][0]]
                                annot.xy = (xval, yval)
                                if (whichPlot == "a"):
                                    text = "Zr=%.5g"%xval + "\nZj=-%.5g"%yval + "\nf=%.5g"%self.sim_freqs[np.where(self.sim_real == xval)][0]
                                elif (whichPlot == "b"):
                                    text = "Zr=%.5g"%yval + "\nZj=-%.5g"%self.sim_imag[np.where(self.sim_real == yval)][0] + "\nf=%.5g"%xval
                                elif (whichPlot == "c"):
                                    text = "Zr=%.5g"%self.sim_real[np.where(-1*self.sim_imag == yval)][0] + "\nZj=-%.5g"%yval + "\nf=%.5g"%xval
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
                                #---Check if we're within 5% of the right or top edges, and adjust label positions accordingly---
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
                                        sim_fig.canvas.draw_idle()
                                    else:
                                        if vis:
                                            annot.set_position((10,10))
                                            annot.set_visible(False)
                                            sim_fig.canvas.draw_idle()
                            sim_fig.canvas.mpl_connect("motion_notify_event", hover)
                        def sim_big_on_closing():   #Clear the figure before closing the popup
                            sim_fig.clear()
                            sim_resultPlotBig.destroy()
                            try:
                                plt.close(sim_fig)
                            except Exception:
                                pass
                        sim_resultPlotBig.protocol("WM_DELETE_WINDOW", sim_big_on_closing)
                except Exception:
                    pass
            
            sim_nyCanvas = FigureCanvasTkAgg(sim_pltFig, simPlot)
            sim_pltFig.subplots_adjust(top=0.95, bottom=0.1, right=0.95, left=0.05)
            sim_nyCanvas.draw()
            sim_nyCanvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
            sim_nyCanvas.mpl_connect('button_press_event', sim_onclick)
            sim_nyCanvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            
            def sim_saveAllPlots(e=None):
                folder = askdirectory(parent=simPlot, initialdir=self.topGUI.getCurrentDirectory())
                folder_str = str(folder)
                if (len(folder_str) == 0):
                    pass
                else:
                    defaultSaveName = StringDialog.ask_string("Save name", "File name: ", parent=simPlot)
                    if defaultSaveName is None:
                        return
                    sim_saveNyCanvasButton.configure(text="Saving")
                    dataColor = "tab:blue"
                    if (self.topGUI.getTheme() == "dark"):
                        dataColor = "cyan"
                    else:
                        dataColor = "tab:blue"
                    sim_pltSaveFig = plt.Figure()
                    sim_pltSaveFig.set_facecolor(self.backgroundColor)
                    sim_save_canvas = FigureCanvas(sim_pltSaveFig)
                    sim_ax_save = sim_pltSaveFig.add_subplot(111)
                    sim_ax_save.set_facecolor(self.backgroundColor)
                    #---aplot---
                    sim_ax_save.yaxis.set_ticks_position("both")
                    sim_ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    sim_ax_save.xaxis.set_ticks_position("both")
                    sim_ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    sim_ax_save.plot(self.sim_real, -1*self.sim_imag, "o",  color=dataColor)
                    sim_ax_save.axis("equal")
                    sim_ax_save.set_title("Nyquist", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_xlabel("Zr / Ω", color=self.foregroundColor)
                    sim_ax_save.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_nyquist.png", dpi=300)
                    sim_ax_save.clear()
                    sim_ax_save.axis("auto")
                    #---bplot---
                    sim_ax_save.set_facecolor(self.backgroundColor)
                    sim_ax_save.yaxis.set_ticks_position("both")
                    sim_ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    sim_ax_save.xaxis.set_ticks_position("both")
                    sim_ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    sim_ax_save.plot(self.sim_freqs, self.sim_real, "o", color=dataColor)
                    sim_ax_save.set_xscale('log')
                    sim_ax_save.set_title("Real Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_ylabel("Zr / Ω", color=self.foregroundColor)
                    sim_ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_real_impedance.png", dpi=300)
                    sim_ax_save.clear()
                    #---cplot---
                    sim_ax_save.plot(self.sim_freqs, -1*self.sim_imag, "o", color=dataColor)
                    sim_ax_save.set_xscale('log')
                    sim_ax_save.set_title("Imaginary Impedance", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    sim_ax_save.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_imag_impedance.png", dpi=300)
                    sim_ax_save.clear()
                    #---dplot---
                    sim_ax_save.yaxis.set_ticks_position("both")
                    sim_ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    sim_ax_save.xaxis.set_ticks_position("both")
                    sim_ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    sim_ax_save.plot(self.sim_freqs, np.sqrt(self.sim_real**2 + self.sim_imag**2), "o", color=dataColor)
                    sim_ax_save.set_xscale('log')
                    sim_ax_save.set_yscale('log')
                    sim_ax_save.set_title("|Z| Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    sim_ax_save.set_ylabel("|Z| / Ω", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_Z_bode.png", dpi=300)
                    sim_ax_save.clear()
                    #---eplot---
                    sim_ax_save.yaxis.set_ticks_position("both")
                    sim_ax_save.yaxis.set_tick_params(color=self.foregroundColor, direction="in")
                    sim_ax_save.xaxis.set_ticks_position("both")
                    sim_ax_save.xaxis.set_tick_params(color=self.foregroundColor, direction="in", which="both")
                    sim_ax_save.plot(self.sim_freqs, np.arctan2(self.sim_imag, self.sim_real)*(180/np.pi), "o", color=dataColor)
                    sim_ax_save.yaxis.set_ticks([-90, -75, -60, -45, -30, -15, 0])
                    sim_ax_save.set_ylim(bottom=0, top=-90)
                    sim_ax_save.set_xscale('log')
                    sim_ax_save.set_title("Phase Angle Bode", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    sim_ax_save.set_ylabel("Phase angle / Deg.", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_phase_angle_bode.png", dpi=300)
                    sim_ax_save.clear()
                    #---hplot---
                    sim_ax_save.plot(self.sim_freqs, np.log10(abs(self.sim_imag)), "o", color=dataColor)
                    sim_ax_save.set_xscale('log')
                    sim_ax_save.set_title("Log|Zr|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    sim_ax_save.set_ylabel("Log|Zr|", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_log_zr.png", dpi=300)
                    sim_ax_save.clear()
                    #---iplot---
                    sim_ax_save.plot(self.sim_freqs, np.log10(abs(self.sim_imag)), "o", color=dataColor)
                    sim_ax_save.set_xscale('log')
                    sim_ax_save.set_title("Log|Zj|", fontdict={'fontsize' : 17, 'color' : self.foregroundColor})
                    sim_ax_save.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    sim_ax_save.set_ylabel("Log|Zj|", color=self.foregroundColor)
                    sim_pltSaveFig.savefig(str(os.path.join(folder, defaultSaveName)) + "_log_zj.png", dpi=300)
                    sim_ax_save.clear()
                    sim_saveNyCanvasButton.configure(text="Saved")
                    self.after(1000, lambda: sim_saveNyCanvasButton.configure(text="Save All"))
                
            def sim_graphOut(event):
                sim_nyCanvas._tkcanvas.config(cursor="arrow")
    
            def sim_graphOver(event):
                whichCan = sim_nyCanvas._tkcanvas
                he = int(whichCan.winfo_height())
                wi = int(whichCan.winfo_width())
                whichCan.config(cursor="hand2")
            
            sim_saveNyCanvasButton = ttk.Button(simPlot, text="Save All", command=sim_saveAllPlots)
            self.simSaveNyCanvasButtons.append(sim_saveNyCanvasButton)
            sim_saveNyCanvasButton_ttp = CreateToolTip(sim_saveNyCanvasButton, "Save all plots")
            self.simSaveNyCanvasButton_ttps .append(sim_saveNyCanvasButton_ttp)
            sim_saveNyCanvasButton.pack(side=tk.BOTTOM, pady=5, expand=False)
            enterAxes = sim_pltFig.canvas.mpl_connect('axes_enter_event', sim_graphOver)
            leaveAxes = sim_pltFig.canvas.mpl_connect('axes_leave_event', sim_graphOut)
            def sim_on_closing():   #Clear the figure before closing the popup
                sim_pltFig.clear()
                simPlot.destroy()
                sim_saveNyCanvasButton.destroy()
                nonlocal sim_saveNyCanvasButton_ttp
                del sim_saveNyCanvasButton_ttp
            
            simPlot.protocol("WM_DELETE_WINDOW", sim_on_closing)
        
        def performSimulations(e=None):
            self.simulationWindow.deiconify()
            self.simulationWindow.lift()
            self.simulationUpperFreq.bind("<KeyRelease>", calculateNumPoints)
            self.simulationLowerFreq.bind("<KeyRelease>", calculateNumPoints)
            self.simulationNumFreq.bind("<KeyRelease>", calculateNumPoints)
            self.runSimulationButton.configure(command=runSimulations)
            self.copySimulationButton.configure(command=copySimulations)
            self.saveSimulationButton.configure(command=saveSimulations)
            self.plotSimulationButton.configure(command=plotSimulations)
            def on_closing_simulation():
                self.simulationWindow.withdraw()
            self.simulationWindow.protocol("WM_DELETE_WINDOW", on_closing_simulation)
        
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
        self.monteCarloEntry = ttk.Entry(self.fittingButtonFrame, textvariable=self.monteCarloValue, width=8)
        self.codeLabel = tk.Label(self.fittingButtonFrame, text="Code:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.formulaDescriptionLink = tk.Label(self.fittingButtonFrame, text="Formula description", bg=self.backgroundColor, fg="blue", cursor="hand2")
        self.importPathLink = tk.Label(self.fittingButtonFrame, text="Import paths", bg=self.backgroundColor, fg="blue", cursor="hand2")
        self.helpLabel = tk.Label(self.fittingButtonFrame, text="Help", bg=self.backgroundColor, fg="blue", cursor="hand2")
        if (self.topGUI.getTheme() == "dark"):
            self.helpLabel.configure(fg="skyblue")
            self.importPathLink.configure(fg="skyblue")
        self.helpLabel.bind("<Button-1>", helpPopup)
        self.formulaDescriptionLink.bind("<Button-1>", formulaDescriptionPopup)
        self.importPathLink.bind("<Button-1>", importPathPopup)
        if (self.topGUI.getTheme() == "dark"):
            self.formulaDescriptionLink.configure(fg="skyblue")
        self.loadCodeButton = ttk.Button(self.fittingButtonFrame, text="Load Formula", command=loadCode)
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
        self.formulaDescriptionLink.grid(column=2, row=3, pady=(5,0), sticky="W")
        self.importPathLink.grid(column=3, row=3, pady=(5, 0), sticky="E")
        self.weightingLabel.grid(column=1, row=1, pady=(5,0), sticky="W")
        self.weightingCombobox.grid(column=2, row=1, pady=(5,0), sticky="W")
        self.noiseFrame.grid(column=3, row=1, pady=(5,0), sticky="W")
        self.noiseLabel.grid(column=0, row=0, sticky="E")
        self.noiseEntry.grid(column=1, row=0, sticky="W")
        self.helpLabel.grid(column=4, row=3, pady=(5, 0), sticky="E")
        self.loadCodeButton.grid(column=4, row=1, sticky="E")
        self.fittingButtonFrame.grid(column=0, row=1, pady=(10, 0), sticky="W")
        paramButton_ttp = CreateToolTip(self.parametersButton, 'Add, remove, or edit fitting parameters')
        fittingType_ttp = CreateToolTip(self.fittingTypeCombobox, 'Change which part of the data is fitted')
        monteCarlo_ttp = CreateToolTip(self.monteCarloEntry, 'Number of Monte Carlo simulations')
        weighting_ttp = CreateToolTip(self.weightingCombobox, 'Weighting used in objective function')
        noise_ttp = CreateToolTip(self.noiseEntry, 'Assumed noise (multiplied by weighting)')
        help_ttp = CreateToolTip(self.helpLabel, 'Opens a custom formula guide PDF')
        loadCode_ttp = CreateToolTip(self.loadCodeButton, 'Load formula from a .mmformula file')
        formulaDescription_ttp = CreateToolTip(self.formulaDescriptionLink, 'Opens a popup to edit the formula description')
        importPaths_ttp = CreateToolTip(self.importPathLink, 'Opens a popup to choose directories to search for imports while running the code')
        
        #---If error structure weighting is chosen---
        self.errorStructureFrame = tk.Frame(self.fittingButtonFrame, bg=self.backgroundColor)
        self.errorAlphaCheckboxVariable = tk.IntVar(self, 0)
        self.errorAlphaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorAlphaCheckboxVariable, text="\u03B1 = ", command=checkErrorStructure)
        self.errorAlphaVariable = tk.StringVar(self, "0.1")
        self.errorAlphaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorAlphaVariable, state="disabled", width=6)
        self.errorBetaCheckboxVariable = tk.IntVar(self, 0)
        self.errorBetaCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorBetaCheckboxVariable, text="\u03B2 = ", command=checkErrorStructure)
        self.errorBetaVariable = tk.StringVar(self, "0.1")
        self.errorBetaEntry = ttk.Entry(self.errorStructureFrame, textvariable=self.errorBetaVariable, state="disabled", width=6)
        self.errorBetaReCheckboxVariable = tk.IntVar(self, 0)
        self.errorBetaReCheckbox = ttk.Checkbutton(self.errorStructureFrame, variable=self.errorBetaReCheckboxVariable, state="disabled", text="Re = ", command=checkErrorStructure)
        self.errorBetaReVariable = tk.StringVar(self, "0.1")
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
                except Exception:
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
                except Exception:
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
                except Exception:
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
                except Exception:
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
                except Exception:
                    stringToCopy = self.errorDeltaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuD = tk.Menu(self.errorDeltaEntry, tearoff=0)
        self.popup_menuD.add_command(label="Copy", command=copy_delta)
        self.popup_menuD.add_command(label="Paste", command=paste_delta)
        self.errorDeltaEntry.bind("<Button-3>", popup_delta)
        
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
        self.formulaSaveButton = ttk.Button(self.runFrame, text="Save formula", command=saveFormula)
        self.cancelButton = ttk.Button(self.runFrame, text="Cancel fitting")
        self.runSimulationsButton = ttk.Button(self.runFrame, text="Synthetic data", command=performSimulations)
        self.runButton.grid(column=0, row=0, sticky="W")
        self.formulaSaveButton.grid(column=2, row=0, sticky="W")
        self.runSimulationsButton.grid(column=1, row=0, sticky="W", padx=5)
        self.runFrame.grid(column=0, row=3, sticky="W")
        formulaSaveButton_ttp = CreateToolTip(self.formulaSaveButton, 'Save the current formula and parameters as a .mmformula file')
        runButton_ttp = CreateToolTip(self.runButton, 'Fit formula using Levenberg-Marquardt algorithm')
        cancelButton_ttp = CreateToolTip(self.cancelButton, 'Cancel fitting')
        simulationsButton_ttp = CreateToolTip(self.runSimulationsButton, 'Generate synthetic data at specified frequencies using the current code and parameters')
        
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
        self.copyButton.grid(column=0, row=5, sticky="E")
        self.resultsView.grid(column=0, row=0, sticky="W")
        self.resultsViewScrollbar.grid(column=1, row=0, sticky="NS")
        self.resultsParamFrame.grid(column=0, row=4, sticky="W")
        self.advancedResultsButton.grid(column=0, row=5, pady=3, sticky="W")
        self.resultsView.bind("<Button-1>", handle_click)
        self.resultsView.bind("<Motion>", handle_motion)
        copy_ttp = CreateToolTip(self.copyButton, 'Copy fitted values and standard deviations in a format that can be opened in a spreadsheet')
        advancedResults_ttp = CreateToolTip(self.advancedResultsButton, 'Open popup with detailed output')
        resultAlert_ttp = CreateToolTip(self.resultAlert, 'At least one 95% confidence interval is nan or greater than 100%')
        
        #---The plotting options---
        self.graphFrame = tk.Frame(self, bg=self.backgroundColor)
        self.includeFrame = tk.Frame(self.graphFrame, bg=self.backgroundColor)
        self.confidenceIntervalCheckboxVariable = tk.IntVar(self, 0)
        self.confidenceIntervalCheckbox = ttk.Checkbutton(self.includeFrame, text="Include CI", variable=self.confidenceIntervalCheckboxVariable)
        self.mouseoverCheckboxVariable = tk.IntVar(self, 1)
        self.mouseoverCheckbox = ttk.Checkbutton(self.includeFrame, text="Mouseover labels", variable=self.mouseoverCheckboxVariable)
        self.errorStructureCheckboxVariable = tk.IntVar(self, 0)
        self.errorStructureCheckbox = ttk.Checkbutton(self.includeFrame, text="Error Structure", variable=self.errorStructureCheckboxVariable)
        self.plotButton = ttk.Button(self.graphFrame, text="Plot", command=plotResults)
        self.plotButton.grid(column=0, row=0, sticky="W", pady=(8,5))
        self.confidenceIntervalCheckbox.grid(column=1, row=0, sticky="W", padx=5)
        self.errorStructureCheckbox.grid(column=2, row=0, sticky="W")
        self.mouseoverCheckbox.grid(column=3, row=0, sticky="W", padx=5)
        self.includeFrame.grid(column=1, row=0, sticky="W")
        plot_ttp = CreateToolTip(self.plotButton, 'Opens a window with plots using data and fitted parameters')
        include_ttp = CreateToolTip(self.confidenceIntervalCheckbox, 'Include confidence interval ellipses and bars when plotting')
        mouseover_ttp = CreateToolTip(self.mouseoverCheckbox, 'Show data label on mouseover when viewing larger popup plots')
        errorStructure_ttp = CreateToolTip(self.errorStructureCheckbox, 'Include error structure lines on the residual error plots')
        
        #---The save and export buttons---
        self.saveFrame = tk.Frame(self, bg=self.backgroundColor)
        self.saveResiduals = ttk.Button(self.saveFrame, text="Save Residuals", command=saveResiduals)
        self.saveAll = ttk.Button(self.saveFrame, text="Export All Results", command=saveAll)
        self.saveFormulaButton = ttk.Button(self.saveFrame, text="Save Fitting", command=saveFitting)
        saveFormualButton_ttp = CreateToolTip(self.saveFormulaButton, 'Save the current formula and parameter values as a .mmcustom file')
        self.saveResiduals.grid(column=0, row=0, sticky="W")
        self.saveFormulaButton.grid(column=1, row=0, sticky="W", padx=5)
        self.saveAll.grid(column=2, row=0)
        saveResiduals_ttp = CreateToolTip(self.saveResiduals, 'Save residuals for error analysis')
        saveAll_ttp = CreateToolTip(self.saveAll, 'Export all results, including data, fits, formula, and parameters')
        
         #---Close all popups---
        self.closeAllPopupsButton = ttk.Button(self, text="Close all popups", command=self.topGUI.closeAllPopups)
        self.closeAllPopupsButton.grid(column=0, row=8, sticky="W", pady=10)
        closeAllPopups_ttp = CreateToolTip(self.closeAllPopupsButton, 'Close all open popup windows')
        
    def _on_mousewheel(self, event):
            xpos, ypos = self.vsb.get()
            xpos_round = round(xpos, 2)     #Avoid floating point errors
            ypos_round = round(ypos, 2)
            if (xpos_round != 0.00 or ypos_round != 1.00):
                self.pcanvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def keyup(self, event=None):
        self.paramListbox.delete(0, tk.END)
        for n in self.paramNameValues:
            self.paramListbox.insert(tk.END, n.get())
        try:
            self.advancedOptionsLabel.configure(text=self.paramNameValues[self.paramListboxSelection].get())
        except Exception:
            pass
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
    
    def advancedOptionsPopup(self):
        self.advancedOptionsWindow.deiconify()
        self.advancedOptionsWindow.lift()
        def onClose():
            self.advancedOptionsWindow.withdraw()
        self.advancedOptionsWindow.protocol("WM_DELETE_WINDOW", onClose) 
    
    def onSelect(self, e, n=-1):
        try:
            if (n == -1):
                num_selected = self.paramListbox.curselection()[0]
            else:
                num_selected = n
            self.paramListboxSelection = num_selected
            self.advancedOptionsLabel.configure(text=self.paramNameValues[num_selected].get())
            self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
            self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
            self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
            self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
            if (self.paramComboboxValues[num_selected].get() == "fixed"):
                self.advancedLowerLimitEntry.configure(state="disabled")
                self.advancedUpperLimitEntry.configure(state="disabled")
            else:
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
            if (self.multistartSpacingVariables[num_selected].get() == "Custom"):
                self.multistartLowerFrame.pack_forget()
                self.multistartUpperFrame.pack_forget()
                self.multistartNumberFrame.pack_forget()
                self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
            else:
                self.multistartCustomFrame.pack_forget()
                self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
            if (self.multistartCheckboxVariables[num_selected].get()):
                self.multistartSpacing.configure(state="readonly")
                self.multistartUpperEntry.configure(state="normal")
                self.multistartLowerEntry.configure(state="normal")
                self.multistartNumberEntry.configure(state="normal")
                self.multistartCustomEntry.configure(state="normal")
            else:
                self.multistartSpacing.configure(state="disabled")
                self.multistartUpperEntry.configure(state="disabled")
                self.multistartLowerEntry.configure(state="disabled")
                self.multistartNumberEntry.configure(state="disabled")
                self.multistartCustomEntry.configure(state="disabled")
            self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[num_selected])
            self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[num_selected])
            self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[num_selected])
            self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[num_selected])
            self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[num_selected])
            self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[num_selected])
            self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[num_selected])
            self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[num_selected])
        except Exception:     #No parameters left
            self.paramListboxSelection = 0
            self.advancedOptionsLabel.configure(text="")
            self.advancedLowerLimitFrame.pack_forget()
            self.advancedUpperLimitFrame.pack_forget()
            self.multistartCheckbox.pack_forget()
            self.multistartSpacingFrame.pack_forget()
            self.multistartLowerFrame.pack_forget()
            self.multistartUpperFrame.pack_forget()
            self.multistartNumberFrame.pack_forget()
            self.multistartCustomFrame.pack_forget()
            self.advancedLowerLimitEntry.configure(state="normal")
            self.advancedUpperLimitEntry.configure(state="normal")
    
    def limitChanged(self, num, e):
        if (self.paramComboboxValues[num].get() == "+"):
            self.upperLimits[num].set("inf")
            self.lowerLimits[num].set("0")
            self.advancedLowerLimitEntry.configure(state="normal")
            self.advancedUpperLimitEntry.configure(state="normal")
        elif (self.paramComboboxValues[num].get() == "+ or -"):
            self.upperLimits[num].set("inf")
            self.lowerLimits[num].set("-inf")
            self.advancedLowerLimitEntry.configure(state="normal")
            self.advancedUpperLimitEntry.configure(state="normal")
        elif (self.paramComboboxValues[num].get() == "-"):
            self.upperLimits[num].set("0")
            self.lowerLimits[num].set("-inf")
            self.advancedLowerLimitEntry.configure(state="normal")
            self.advancedUpperLimitEntry.configure(state="normal")
        elif (self.paramComboboxValues[num].get() == "fixed" and self.paramListboxSelection == num):
            self.advancedLowerLimitEntry.configure(state="disabled")
            self.advancedUpperLimitEntry.configure(state="disabled")
        elif (self.paramComboboxValues[num].get() == "custom"):
            self.advancedLowerLimitEntry.configure(state="normal")
            self.advancedUpperLimitEntry.configure(state="normal")
            self.advancedOptionsPopup()
            self.paramListbox.selection_clear(0, tk.END)
            self.paramListbox.select_set(num)
            self.paramListboxSelection = num
            self.onSelect(None, n=num)
    
    def graphLatex(self, e=None):
        try:
            tmptext = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            tmptext2 = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            split_text = tmptext.split(r"\\")
            tmptext = ""
            first = True
            for t in split_text:
                if first:
                    tmptext = "$"+t+"$"
                    first = False
                else:
                    tmptext += "\n" + "$"+t+"$"
            if tmptext2 == "":
               with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.latexAx.clear()
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
            elif tmptext2 != '\\' and tmptext2[-1] != '\\':
                windowWidth = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).width*self.latexFig.dpi
                windowHeight = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).height*self.latexFig.dpi
                with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.latexAx.clear()
                    rLatex = self.latexFig.canvas.get_renderer()
                    latexSize = 30
                    latexHeight = 0.5
                    tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                    bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                        self.latexAx.clear()
                        if (latexHeight <= 0.05):
                            latexHeight = 0
                            while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                self.latexAx.clear()
                                latexSize -=1
                                if (latexSize <= 0):
                                    latexSize = 1
                                    break
                                tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                bb = tLatex.get_window_extent(renderer=rLatex)
                            break
                        latexHeight -= 0.05
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.width > windowWidth):
                        self.latexAx.clear()
                        latexSize -= 1
                        if (latexSize <= 0):
                            latexSize = 1
                            break
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
        except Exception:
            pass
    
    def formulaEnter(self, n):
        fname, fext = os.path.splitext(n)
        directory = os.path.dirname(str(n))
        self.topGUI.setCurrentDirectory(directory)
        if (fext == ".mmfile"):
            try:
                with open(n,'r') as UseFile:
                    filetext = UseFile.read()
                    lines = filetext.splitlines()
                if ("frequency" in lines[0].lower()):
                    data = np.loadtxt(n, skiprows=1)
                else:
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
                self.upDelete = 0
                self.lowDelete = 0
                self.lowerSpinboxVariable.set(str(self.lowDelete))
                self.upperSpinboxVariable.set(str(self.upDelete))
                self.wdata = self.wdataRaw.copy()
                self.rdata = self.rdataRaw.copy()
                self.jdata = self.jdataRaw.copy()
                self.lowestUndeleted.configure(state="normal")
                self.lowestUndeleted.delete(1.0, tk.END)
                self.lowestUndeleted.insert(1.0, "Lowest remaining frequency: {:.4e}".format(min(self.wdata)))
                self.lowestUndeleted.configure(state="disabled")
                self.highestUndeleted.configure(state="normal")
                self.highestUndeleted.delete(1.0, tk.END)
                self.highestUndeleted.insert(1.0, "Highest remaining frequency: {:.4e}".format(max(self.wdata)))
                self.highestUndeleted.configure(state="disabled")
                self.lengthOfData = len(self.wdata)
                self.freqRangeButton.configure(state="normal")
                self.browseEntry.configure(state="normal")
                self.browseEntry.delete(0,tk.END)
                self.browseEntry.insert(0, n)
                self.browseEntry.configure(state="readonly")
                self.runButton.configure(state="normal")
                self.simplexButton.configure(state="normal")
            except Exception:
                messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
        elif (fext == ".mmcustom"):
            try:
                toLoad = open(n)
                firstLine = toLoad.readline()
                fileVersion = 0
                if "version: 1.1" in firstLine:
                    fileVersion = 1
                elif "version: 1.2" in firstLine:
                    fileVersion = 2
                if (fileVersion == 0):
                    fileToLoad = firstLine.split("filename: ")[1][:-1]
                else:
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
                    if (fileVersion == 0):
                        if "mmparams:" in nextLineIn:
                            break
                        else:
                            formulaIn += nextLineIn
                    elif (fileVersion == 1):
                        if "description:" in nextLineIn:
                            break
                        else:
                            formulaIn += nextLineIn
                    elif (fileVersion == 2):
                        if "imports:" in nextLineIn:
                            break
                        else:
                            formulaIn += nextLineIn
                for i in range(len(self.paramNameEntries)):
                    self.paramNameEntries[i].grid_remove()
                    self.paramNameLabels[i].grid_remove()
                    self.paramValueEntries[i].grid_remove()
                    self.paramValueLabels[i].grid_remove()
                    self.paramComboboxes[i].grid_remove()
                    self.paramDeleteButtons[i].grid_remove()
                self.numParams = 0
                self.paramNameValues.clear()
                self.paramNameEntries.clear()
                self.paramNameLabels.clear()
                self.paramValueLabels.clear()
                self.paramComboboxValues.clear()
                self.paramComboboxes.clear()
                self.paramValueValues.clear()
                self.paramValueEntries.clear()
                self.paramDeleteButtons.clear()
                self.paramListbox.delete(0, tk.END)
                self.multistartCheckboxVariables.clear()
                self.multistartSpacingVariables.clear()
                self.multistartLowerVariables.clear()
                self.multistartUpperVariables.clear()
                self.multistartNumberVariables.clear()
                self.multistartCustomVariables.clear()
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
                imports = ""
                if (fileVersion == 2):
                    while True:
                        nextLineIn = toLoad.readline()
                        if "description:" in nextLineIn:
                            break
                        else:
                            imports += nextLineIn
                toImport = imports.rstrip().split('\n')
                for tI in toImport:
                    if (len(tI) > 0):
                        self.importPathListbox.insert(tk.END, tI)
                if (fileVersion != 0):
                    self.formulaDescriptionEntry.delete("1.0", tk.END)
                    descriptionIn = ""
                    while True:
                        nextLineIn = toLoad.readline()
                        if "latex:" in nextLineIn:
                            break
                        else:
                            descriptionIn += nextLineIn
                    latexIn = toLoad.readline()
                    self.formulaDescriptionEntry.insert("1.0", descriptionIn.rstrip())
                    self.formulaDescriptionLatexEntry.delete("1.0", tk.END)
                    self.formulaDescriptionLatexEntry.insert("1.0", latexIn.rstrip())
                    self.latexAx.clear()
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
                    self.graphLatex()
                    toLoad.readline()
                while True:
                    p = toLoad.readline()
                    if not p:
                        break
                    else:
                        p = p[:-1]
                        try:
                            pname, pval, pcombo, plimup, plimlow, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                            self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                            self.multistartLowerVariables.append(tk.StringVar(self, plower))
                            self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                            self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                            self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                            self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                            self.upperLimits.append(tk.StringVar(self, plimup))
                            self.lowerLimits.append(tk.StringVar(self, plimlow))
                        except Exception:
                            try:
                                pname, pval, pcombo, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                                self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                                self.multistartLowerVariables.append(tk.StringVar(self, plower))
                                self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                                self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                                self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                                self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                                self.upperLimits.append(tk.StringVar(self, "inf"))
                                self.lowerLimits.append(tk.StringVar(self, "-inf"))
                            except Exception:
                                pname, pval, pcombo = p.split("~=~")
                                self.multistartCheckboxVariables.append(tk.IntVar(self, 0))
                                self.multistartLowerVariables.append(tk.StringVar(self, "1E-5"))
                                self.multistartUpperVariables.append(tk.StringVar(self, "1E5"))
                                self.multistartSpacingVariables.append(tk.StringVar(self, "Logarithmic"))
                                self.multistartCustomVariables.append(tk.StringVar(self, "0.1,1,10,100,1000"))
                                self.multistartNumberVariables.append(tk.StringVar(self, "10"))
                                self.upperLimits.append(tk.StringVar(self, "inf"))
                                self.lowerLimits.append(tk.StringVar(self, "-inf"))
                        self.numParams += 1
                        if (self.numParams == 1):
                            self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[0])
                            self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[0])
                            self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[0])
                            self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[0])
                            self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[0])
                            self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[0])
                            self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[0])
                            self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[0])
                            if (self.multistartCheckboxVariables[0].get() == 0):
                                self.multistartSpacing.configure(state="disabled")
                                self.multistartUpperEntry.configure(state="disabled")
                                self.multistartLowerEntry.configure(state="disabled")
                                self.multistartNumberEntry.configure(state="disabled")
                                self.multistartCustomEntry.configure(state="disabled")
                            else:
                                if (self.multistartSpacingVariables[0].get() == "Custom"):
                                    self.multistartSpacing.configure(state="readonly")
                                    self.multistartUpperEntry.configure(state="normal")
                                    self.multistartLowerEntry.configure(state="normal")
                                    self.multistartNumberEntry.configure(state="normal")
                                    self.multistartCustomEntry.configure(state="normal")
                                    self.multistartLowerFrame.pack_forget()
                                    self.multistartUpperFrame.pack_forget()
                                    self.multistartNumberFrame.pack_forget()
                                    self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                else:
                                    self.multistartSpacing.configure(state="readonly")
                                    self.multistartUpperEntry.configure(state="normal")
                                    self.multistartLowerEntry.configure(state="normal")
                                    self.multistartNumberEntry.configure(state="normal")
                                    self.multistartCustomEntry.configure(state="normal")
                        self.paramNameValues.append(tk.StringVar(self, pname))
                        self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10, validate="all", validatecommand=self.valcom))
                        self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                        self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed", "custom"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                        self.paramComboboxes[-1].bind("<<ComboboxSelected>>", partial(self.limitChanged, self.numParams-1))
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
                        self.paramListbox.insert(tk.END, pname)
                        if (self.paramComboboxValues[0].get() == "fixed"):
                            self.advancedLowerLimitEntry.configure(state="disabled")
                            self.advancedUpperLimitEntry.configure(state="disabled")
                toLoad.close()
                self.paramListbox.selection_clear(0, tk.END)
                self.paramListbox.select_set(0)
                self.paramListboxSelection = 0
                self.paramListbox.event_generate("<<ListboxSelect>>")
                try:
                    with open(fileToLoad,'r') as UseFile:
                        filetext = UseFile.read()
                        lines = filetext.splitlines()
                    if ("frequency" in lines[0].lower()):
                        data = np.loadtxt(fileToLoad, skiprows=1)
                    else:
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
                    self.runButton.configure(state="normal")
                    self.simplexButton.configure(state="normal")
                except Exception:
                    messagebox.showerror("File not found", "Error 53: \nThe linked .mmfile could not be found")
                    fileToLoad = ""
                self.browseEntry.configure(state="normal")
                self.browseEntry.delete(0,tk.END)
                self.browseEntry.insert(0, n)
                self.browseEntry.configure(state="readonly")
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
                self.noiseEntryValue.set(alphaVal)
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
                self.customFormula.insert("1.0", formulaIn.rstrip())    #Insert custom formula into code box and remove trailing whitespace
                self.keyup("")
            except Exception:
                messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
        else:
            messagebox.showerror("File error", "Error 43: \nThe file has an unknown extension")
    
    def setThemeLight(self):
        self.backgroundColor = "#FFFFFF"
        self.foregroundColor = "#000000"
        self.configure(background="#FFFFFF")
        self.customFunctionFrame.configure(background="#FFFFFF")
        self.fittingButtonFrame.configure(background="#FFFFFF")
        self.codeLabel.configure(background="#FFFFFF", foreground="#000000")
        self.helpLabel.configure(background="#FFFFFF", foreground="blue")
        self.formulaDescriptionLink.configure(background="#FFFFFF", foreground="blue")
        self.importPathLink.configure(background="#FFFFFF", foreground="blue")
        self.runFrame.configure(background="#FFFFFF")
        self.pframe.configure(background="#FFFFFF")
        self.pcanvas.configure(background="#FFFFFF")
        self.simulationWindow.configure(background="#FFFFFF")
        self.simulationFreqFrame.configure(background="#FFFFFF")
        self.simulationLowerFreqLabel.configure(background="#FFFFFF", foreground="#000000")
        self.simulationUpperFreqLabel.configure(background="#FFFFFF", foreground="#000000")
        self.simulationNumFreqLabel.configure(background="#FFFFFF", foreground="#000000")
        self.numberOfPointsLabel.configure(background="#FFFFFF", foreground="#000000")
        self.simViewFrame.configure(background="#FFFFFF")
        self.saveSimulationFrame.configure(background="#FFFFFF")
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
        self.advancedOptionsWindow.configure(background="#FFFFFF")
        self.advancedOptionsLabel.configure(background="#FFFFFF", foreground="#000000")
        self.advancedUpperLimitLabel.configure(background="#FFFFFF", foreground="#000000")
        self.advancedLowerLimitLabel.configure(background="#FFFFFF", foreground="#000000")
        self.multistartSpacingLabel.configure(background="#FFFFFF", foreground="#000000")
        self.multistartLowerLabel.configure(background="#FFFFFF", foreground="#000000")
        self.multistartUpperLabel.configure(background="#FFFFFF", foreground="#000000")
        self.multistartNumberLabel.configure(background="#FFFFFF", foreground="#000000")
        self.multistartCustomLabel.configure(background="#FFFFFF", foreground="#000000")
        self.paramListbox.configure(background="#FFFFFF", foreground="#000000")
        self.formulaDescriptionWindow.configure(background="#FFFFFF")
        self.formulaDescriptionLabel.configure(background="#FFFFFF", foreground="#000000")
        self.formulaDescriptionFrame.configure(background="#FFFFFF")
        self.formulaDescriptionContainer.configure(background="#FFFFFF")
        self.formulaLatexFrame.configure(background="#FFFFFF")
        self.formulaDescriptionLatexLabel.configure(background="#FFFFFF", foreground="#000000")
        self.formulaDescriptionEntry.configure(background="#FFFFFF", foreground="#000000")
        self.latexFig.set_facecolor("#FFFFFF")
        self.graphLatex()
        self.load_latexFig.set_facecolor("#FFFFFF")
        try:
            tmptext = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            tmptext2 = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            split_text = tmptext.split(r"\\")
            tmptext = ""
            first = True
            for t in split_text:
                if first:
                    tmptext = "$"+t+"$"
                    first = False
                else:
                    tmptext += "\n" + "$"+t+"$"
            if tmptext2 == "":
               with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.latexAx.clear()
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
            elif tmptext2 != '\\' and tmptext2[-1] != '\\':
                windowWidth = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).width*self.latexFig.dpi
                windowHeight = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).height*self.latexFig.dpi
                with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.latexAx.clear()
                    rLatex = self.latexFig.canvas.get_renderer()
                    latexSize = 30
                    latexHeight = 0.5
                    tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                    bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                        self.latexAx.clear()
                        if (latexHeight <= 0.05):
                            latexHeight = 0
                            while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                self.latexAx.clear()
                                latexSize -=1
                                if (latexSize <= 0):
                                    latexSize = 1
                                    break
                                tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                bb = tLatex.get_window_extent(renderer=rLatex)
                            break
                        latexHeight -= 0.05
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.width > windowWidth):
                        self.latexAx.clear()
                        latexSize -= 1
                        if (latexSize <= 0):
                            latexSize = 1
                            break
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
        except Exception:
            pass
        self.loadFormulaWindow.configure(background="#FFFFFF")
        self.loadFormulaFrame.configure(background="#FFFFFF")
        self.loadFormulaDirectoryFrame.configure(background="#FFFFFF")
        self.loadFormulaDirectoryLabel.configure(background="#FFFFFF", foreground="#000000")
        self.loadFormulaFrameTwo.configure(background="#FFFFFF")
        self.loadFormulaDescriptionFrame.configure(background="#FFFFFF")
        self.loadFormulaDescriptionLabel.configure(background="#FFFFFF", foreground="#000000")
        self.loadFormulaDescription.configure(background="#FFFFFF", foreground="#000000")
        self.loadFormulaCode.configure(background="#FFFFFF", foreground="#000000")
        self.loadFormulaCodeLabel.configure(background="#FFFFFF", foreground="#000000")
        self.loadFormulaEquationLabel.configure(background="#FFFFFF", foreground="#000000")
        self.loadFormulaButtonFrame.configure(background="#FFFFFF")
        self.loadFormulaIncludeLabel.configure(background="#FFFFFF", foreground="#000000")
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
                self.realFreq.set_ylabel("Zr / Ω", color=self.foregroundColor)
                self.imagFreq.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.canvasFreq.draw()
        except Exception:
            pass
        self.howMany.configure(background="#FFFFFF", foreground="#000000")
        self.paramPopup.configure(background="#FFFFFF")
        self.bottomButtonsFrame.configure(background="#FFFFFF")
        self.saveFrame.configure(background="#FFFFFF")
        self.lowestUndeleted.configure(background="#FFFFFF", foreground="#000000")
        self.highestUndeleted.configure(background="#FFFFFF", foreground="#000000")
        for paramName in self.paramNameLabels:
            paramName.configure(background="#FFFFFF", foreground="#000000")
        for paramVal in self.paramValueLabels:
            paramVal.configure(background="#FFFFFF", foreground="#000000")
        try:
            self.upperLabel.configure(background="#FFFFFF", foreground="#000000")
            self.lowerLabel.configure(background="#FFFFFF", foreground="#000000")
            self.rangeLabel.configure(background="#FFFFFF", foreground="#000000")
        except Exception:
            pass
                                 
    def setThemeDark(self):
        self.backgroundColor = "#424242"
        self.foregroundColor = "#FFFFFF"
        self.configure(background="#424242")
        self.customFunctionFrame.configure(background="#424242")
        self.fittingButtonFrame.configure(background="#424242")
        self.codeLabel.configure(background="#424242", foreground="#FFFFFF")
        self.helpLabel.configure(background="#424242", foreground="skyblue")
        self.formulaDescriptionLink.configure(background="#424242", foreground="skyblue")
        self.importPathLink.configure(background="#424242", foreground="skyblue")
        self.runFrame.configure(background="#424242")
        self.pframe.configure(background="#424242")
        self.pcanvas.configure(background="#424242")
        self.simulationWindow.configure(background="#424242")
        self.simulationFreqFrame.configure(background="#424242")
        self.simulationLowerFreqLabel.configure(background="#424242", foreground="#FFFFFF")
        self.simulationUpperFreqLabel.configure(background="#424242", foreground="#FFFFFF")
        self.simulationNumFreqLabel.configure(background="#424242", foreground="#FFFFFF")
        self.numberOfPointsLabel.configure(background="#424242", foreground="#FFFFFF")
        self.simViewFrame.configure(background="#424242")
        self.saveSimulationFrame.configure(background="#424242")
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
        self.bottomButtonsFrame.configure(background="#424242")
        for paramName in self.paramNameLabels:
            paramName.configure(background="#424242", foreground="#FFFFFF")
        for paramVal in self.paramValueLabels:
            paramVal.configure(background="#424242", foreground="#FFFFFF")
        self.advancedOptionsWindow.configure(background="#424242")
        self.advancedOptionsLabel.configure(background="#424242", foreground="#FFFFFF")
        self.advancedUpperLimitLabel.configure(background="#424242", foreground="#FFFFFF")
        self.advancedLowerLimitLabel.configure(background="#424242", foreground="#FFFFFF")
        self.multistartSpacingLabel.configure(background="#424242", foreground="#FFFFFF")
        self.multistartLowerLabel.configure(background="#424242", foreground="#FFFFFF")
        self.multistartUpperLabel.configure(background="#424242", foreground="#FFFFFF")
        self.multistartNumberLabel.configure(background="#424242", foreground="#FFFFFF")
        self.multistartCustomLabel.configure(background="#424242", foreground="#FFFFFF")
        self.paramListbox.configure(background="#424242", foreground="#FFFFFF")
        self.formulaDescriptionWindow.configure(background="#424242")
        self.formulaDescriptionLabel.configure(background="#424242", foreground="#FFFFFF")
        self.formulaDescriptionFrame.configure(background="#424242")
        self.formulaDescriptionContainer.configure(background="#424242")
        self.formulaLatexFrame.configure(background="#424242")
        self.formulaDescriptionLatexLabel.configure(background="#424242", foreground="#FFFFFF")
        self.formulaDescriptionEntry.configure(background="#6B6B6B", foreground="#FFFFFF")
        self.latexFig.set_facecolor("#424242")
        self.graphLatex()
        self.load_latexFig.set_facecolor("#424242")
        try:
            tmptext = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            tmptext2 = self.formulaDescriptionLatexEntry.get("1.0", tk.END).rstrip().replace('\n', '').replace('\r', '') #self.formulaDescriptionLatexVariable.get()
            split_text = tmptext.split(r"\\")
            tmptext = ""
            first = True
            for t in split_text:
                if first:
                    tmptext = "$"+t+"$"
                    first = False
                else:
                    tmptext += "\n" + "$"+t+"$"
            if tmptext2 == "":
               with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.latexAx.clear()
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
            elif tmptext2 != '\\' and tmptext2[-1] != '\\':
                windowWidth = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).width*self.latexFig.dpi
                windowHeight = self.latexAx.get_window_extent().transformed(self.latexFig.dpi_scale_trans.inverted()).height*self.latexFig.dpi
                with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.latexAx.clear()
                    rLatex = self.latexFig.canvas.get_renderer()
                    latexSize = 30
                    latexHeight = 0.5
                    tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                    bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                        self.latexAx.clear()
                        if (latexHeight <= 0.05):
                            latexHeight = 0
                            while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                self.latexAx.clear()
                                latexSize -=1
                                if (latexSize <= 0):
                                    latexSize = 1
                                    break
                                tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                bb = tLatex.get_window_extent(renderer=rLatex)
                            break
                        latexHeight -= 0.05
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.width > windowWidth):
                        self.latexAx.clear()
                        latexSize -= 1
                        if (latexSize <= 0):
                            latexSize = 1
                            break
                        tLatex = self.latexAx.text(0.01, latexHeight, tmptext, transform=self.latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    self.latexAx.axis("off")
                    self.latexCanvas.draw()
        except Exception:
            pass
        self.loadFormulaWindow.configure(background="#424242")
        self.loadFormulaFrame.configure(background="#424242")
        self.loadFormulaDirectoryFrame.configure(background="#424242")
        self.loadFormulaDirectoryLabel.configure(background="#424242", foreground="#FFFFFF")
        self.loadFormulaFrameTwo.configure(background="#424242")
        self.loadFormulaDescriptionFrame.configure(background="#424242")
        self.loadFormulaDescriptionLabel.configure(background="#424242", foreground="#FFFFFF")
        self.loadFormulaDescription.configure(background="#6B6B6B", foreground="#FFFFFF")
        self.loadFormulaCode.configure(background="#6B6B6B", foreground="#FFFFFF")
        self.loadFormulaCodeLabel.configure(background="#424242", foreground="#FFFFFF")
        self.loadFormulaEquationLabel.configure(background="#424242", foreground="#FFFFFF")
        self.loadFormulaButtonFrame.configure(background="#424242")
        self.loadFormulaIncludeLabel.configure(background="#424242", foreground="#FFFFFF")
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
                self.realFreq.set_ylabel("Zr / Ω", color=self.foregroundColor)
                self.imagFreq.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                self.imagFreq.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                self.canvasFreq.draw()
        except Exception:
            pass
        self.saveFrame.configure(background="#424242")
        try:
            self.upperLabel.configure(background="#424242", foreground="#FFFFFF")
            self.lowerLabel.configure(background="#424242", foreground="#FFFFFF")
            self.rangeLabel.configure(background="#424242", foreground="#FFFFFF")
        except Exception:
            pass
    
    def setEllipseColor(self, color):
        self.ellipseColor = color
    
    def needToSave(self):
        return self.saveNeeded
    
    def insert_node(self, parent, text, abspath):
        if (text.endswith(".mmformula") or os.path.isdir(abspath)):
            try:
                if (os.path.isdir(abspath)):
                    os.listdir(abspath)
                node = self.loadFormulaTree.insert(parent, 'end', text=text, open=False)
                if os.path.isdir(abspath):
                    self.loadFormulaNodes[node] = abspath
                    self.loadFormulaTree.insert(node, 'end')
            except Exception:
                pass    #Don't show directories we don't have permission to access
    def open_node(self, event):
        node = self.loadFormulaTree.focus()
        abspath = self.loadFormulaNodes.pop(node, None)
        if abspath:
            self.loadFormulaTree.delete(self.loadFormulaTree.get_children(node))
            for p in os.listdir(abspath):
                self.insert_node(node, p, os.path.join(abspath, p))
    
    def load_graphLatex(self, e=None):
        try:
            tmptext = self.latexIn
            tmptext2 = self.latexIn
            split_text = tmptext.split(r"\\")
            tmptext = ""
            first = True
            for t in split_text:
                if first:
                    tmptext = "$"+t.rstrip()+"$"
                    first = False
                else:
                    tmptext += "\n" + "$"+t.rstrip()+"$"
            if tmptext2 == "":
               with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.load_latexAx.clear()
                    self.load_latexAx.axis("off")
                    self.load_latexCanvas.draw()
            elif tmptext2 != '\\' and tmptext2[-1] != '\\':
                windowWidth = self.load_latexAx.get_window_extent().transformed(self.load_latexFig.dpi_scale_trans.inverted()).width*self.latexFig.dpi
                windowHeight = self.load_latexAx.get_window_extent().transformed(self.load_latexFig.dpi_scale_trans.inverted()).height*self.latexFig.dpi
                with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.load_latexAx.clear()
                    rLatex = self.load_latexFig.canvas.get_renderer()
                    latexSize = 30
                    latexHeight = 0.5
                    tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                    bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                        self.load_latexAx.clear()
                        if (latexHeight <= 0.05):
                            latexHeight = 0
                            while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                self.load_latexAx.clear()
                                latexSize -=1
                                if (latexSize <= 0):
                                    latexSize = 1
                                    break
                                tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                bb = tLatex.get_window_extent(renderer=rLatex)
                            break
                        latexHeight -= 0.05
                        tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    while (bb.width > windowWidth):
                        self.load_latexAx.clear()
                        latexSize -= 1
                        if (latexSize <= 0):
                            latexSize = 1
                            break
                        tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                        bb = tLatex.get_window_extent(renderer=rLatex)
                    self.load_latexAx.axis("off")
                    self.load_latexCanvas.draw()
        except Exception:
            pass
    
    def clickFormula(self, event):
        try:
            item = self.loadFormulaTree.selection()[0]        
            if (self.loadFormulaTree.item(item,"text").endswith(".mmformula")):
                path_to_child = self.loadFormulaTree.item(item,"text")
                parent_iid = self.loadFormulaTree.parent(item)
                parent_node = self.loadFormulaTree.item(parent_iid)['text']
                while parent_iid != '':
                    path_to_child = os.path.join(parent_node, path_to_child)
                    parent_iid = self.loadFormulaTree.parent(parent_iid)
                    parent_node = self.loadFormulaTree.item(parent_iid)['text']
                try:
                    toLoad = open(path_to_child)
                    firstLine = toLoad.readline()
                    fileVersion = 0
                    if "version: 1.1" in firstLine:
                        fileVersion = 1
                    elif "version: 1.2" in firstLine:
                        fileVersion = 2
                    if (fileVersion == 0):
                        fileToLoad = firstLine.split("filename: ")[1][:-1]
                    else:
                        fileToLoad = toLoad.readline().split("filename: ")[1][:-1]
                    for i in range(12):
                        toLoad.readline()
                    formulaIn = ""
                    while True:
                        nextLineIn = toLoad.readline()
                        if (fileVersion == 0):
                            if "mmparams:" in nextLineIn:
                                break
                            else:
                                formulaIn += nextLineIn
                        elif (fileVersion == 1):
                            if "description:" in nextLineIn:
                                break
                            else:
                                formulaIn += nextLineIn
                        elif (fileVersion == 2):
                            if "imports:" in nextLineIn:
                                break
                            else:
                                formulaIn += nextLineIn
                    descriptionIn = ""
                    imports = ""
                    if (fileVersion == 2):
                        while True:
                            nextLineIn = toLoad.readline()
                            if "description:" in nextLineIn:
                                break
                            else:
                                imports += nextLineIn
                    while True:
                        nextLineIn = toLoad.readline()
                        if "latex:" in nextLineIn:
                            break
                        else:
                            descriptionIn += nextLineIn
                    self.loadFormulaDescription.configure(state="normal")
                    self.loadFormulaDescription.delete("1.0", tk.END)
                    self.loadFormulaDescription.insert("1.0", descriptionIn.rstrip())
                    self.loadFormulaDescription.configure(state="disabled")
                    self.loadFormulaCode.configure(state="normal")
                    self.loadFormulaCode.delete("1.0", tk.END)
                    self.loadFormulaCode.insert("1.0", formulaIn.rstrip())
                    self.loadFormulaCode.configure(state="disabled")
                    self.loadFormula.configure(state="normal")
                    self.latexIn = toLoad.readline().rstrip()
                    self.load_latexAx.clear()
                    self.load_latexAx.axis("off")
                    self.load_latexCanvas.draw()
                    try:
                        tmptext = self.latexIn
                        tmptext2 = self.latexIn
                        split_text = tmptext.split(r"\\")
                        tmptext = ""
                        first = True
                        for t in split_text:
                            if first:
                                tmptext = "$"+t.rstrip()+"$"
                                first = False
                            else:
                                tmptext += "\n" + "$"+t.rstrip()+"$"
                        if tmptext2 == "":
                           with plt.rc_context({'mathtext.fontset': "cm"}):
                                self.load_latexAx.clear()
                                self.load_latexAx.axis("off")
                                self.load_latexCanvas.draw()
                        elif tmptext2 != '\\' and tmptext2[-1] != '\\':
                            windowWidth = self.load_latexAx.get_window_extent().transformed(self.load_latexFig.dpi_scale_trans.inverted()).width*self.latexFig.dpi
                            windowHeight = self.load_latexAx.get_window_extent().transformed(self.load_latexFig.dpi_scale_trans.inverted()).height*self.latexFig.dpi
                            with plt.rc_context({'mathtext.fontset': "cm"}):
                                self.load_latexAx.clear()
                                rLatex = self.load_latexFig.canvas.get_renderer()
                                latexSize = 30
                                latexHeight = 0.5
                                tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                bb = tLatex.get_window_extent(renderer=rLatex)
                                while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                    self.load_latexAx.clear()
                                    if (latexHeight <= 0.05):
                                        latexHeight = 0
                                        while (bb.height > (windowHeight*(1-latexHeight) + 9)):
                                            self.load_latexAx.clear()
                                            latexSize -=1
                                            if (latexSize <= 0):
                                                latexSize = 1
                                                break
                                            tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                            bb = tLatex.get_window_extent(renderer=rLatex)
                                        break
                                    latexHeight -= 0.05
                                    tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)
                                    bb = tLatex.get_window_extent(renderer=rLatex)
                                while (bb.width > windowWidth):
                                    self.load_latexAx.clear()
                                    latexSize -= 1
                                    if (latexSize <= 0):
                                        latexSize = 1
                                        break
                                    tLatex = self.load_latexAx.text(0.01, latexHeight, tmptext, transform=self.load_latexAx.transAxes, color=self.foregroundColor, fontsize = latexSize)  
                                    bb = tLatex.get_window_extent(renderer=rLatex)
                                self.load_latexAx.axis("off")
                                self.load_latexCanvas.draw()
                    except Exception:
                        pass
                    toLoad.close()
                except Exception:
                   messagebox.showerror("File error", "Error 62: \nThere was an error loading or reading the file")
            else:
                self.loadFormulaDescription.configure(state="normal")
                self.loadFormulaDescription.delete("1.0", tk.END)
                self.loadFormulaDescription.configure(state="disabled")
                self.loadFormulaCode.configure(state="normal")
                self.loadFormulaCode.delete("1.0", tk.END)
                self.loadFormulaCode.configure(state="disabled")
                self.loadFormula.configure(state="disabled")
                with plt.rc_context({'mathtext.fontset': "cm"}):
                    self.load_latexAx.clear()
                    self.load_latexAx.axis("off")
                    self.load_latexCanvas.draw()
        except Exception:
            pass    #Ignore clicks that would be out-of-range
    
    def enterFormula(self, path_to_child):
        try:
            toLoad = open(path_to_child)
            firstLine = toLoad.readline()
            fileVersion = 0
            if "version: 1.1" in firstLine:
                fileVersion = 1
            elif "version: 1.2" in firstLine:
                fileVersion = 2
            if (fileVersion == 0):
                fileToLoad = firstLine.split("filename: ")[1][:-1]
            else:
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
                if (fileVersion == 0):
                    if "mmparams:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
                elif (fileVersion == 1):
                    if "description:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
                elif (fileVersion == 2):
                    if "imports:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
            if (self.loadFormulaParamsVariable.get()):
                for i in range(len(self.paramNameEntries)):
                    self.paramNameEntries[i].grid_remove()
                    self.paramNameLabels[i].grid_remove()
                    self.paramValueEntries[i].grid_remove()
                    self.paramValueLabels[i].grid_remove()
                    self.paramComboboxes[i].grid_remove()
                    self.paramDeleteButtons[i].grid_remove()
                self.numParams = 0
                self.paramNameValues.clear()
                self.paramNameEntries.clear()
                self.paramNameLabels.clear()
                self.paramValueLabels.clear()
                self.paramComboboxValues.clear()
                self.paramComboboxes.clear()
                self.paramValueValues.clear()
                self.paramValueEntries.clear()
                self.paramDeleteButtons.clear()
                self.paramListbox.delete(0, tk.END)
                self.multistartCheckboxVariables.clear()
                self.multistartSpacingVariables.clear()
                self.multistartLowerVariables.clear()
                self.multistartUpperVariables.clear()
                self.multistartNumberVariables.clear()
                self.multistartCustomVariables.clear()
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
            imports = ""
            if (fileVersion == 2):
                while True:
                    nextLineIn = toLoad.readline()
                    if "description:" in nextLineIn:
                        break
                    else:
                        imports += nextLineIn
            toImport = imports.rstrip().split('\n')
            for tI in toImport:
                if (len(tI) > 0):
                    self.importPathListbox.insert(tk.END, tI)
            if (fileVersion != 0):
                self.formulaDescriptionEntry.delete("1.0", tk.END)
                descriptionIn = ""
                while True:
                    nextLineIn = toLoad.readline()
                    if "latex:" in nextLineIn:
                        break
                    else:
                        descriptionIn += nextLineIn
                latexIn = toLoad.readline()
                self.formulaDescriptionEntry.insert("1.0", descriptionIn.rstrip())
                self.formulaDescriptionLatexEntry.delete("1.0", tk.END)
                self.formulaDescriptionLatexEntry.insert("1.0", latexIn.rstrip())
                self.latexAx.clear()
                self.latexAx.axis("off")
                self.latexCanvas.draw()
                self.graphLatex()
                toLoad.readline()
            if (self.loadFormulaParamsVariable.get()):
                while True:
                    p = toLoad.readline()
                    if not p:
                        break
                    else:
                        p = p[:-1]
                        try:
                            pname, pval, pcombo, plimup, plimlow, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                            self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                            self.multistartLowerVariables.append(tk.StringVar(self, plower))
                            self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                            self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                            self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                            self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                            self.upperLimits.append(tk.StringVar(self, plimup))
                            self.lowerLimits.append(tk.StringVar(self, plimlow))
                        except Exception:
                            try:
                                pname, pval, pcombo, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                                self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                                self.multistartLowerVariables.append(tk.StringVar(self, plower))
                                self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                                self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                                self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                                self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                                self.upperLimits.append(tk.StringVar(self, "inf"))
                                self.lowerLimits.append(tk.StringVar(self, "-inf"))
                            except Exception:
                                pname, pval, pcombo = p.split("~=~")
                                self.multistartCheckboxVariables.append(tk.IntVar(self, 0))
                                self.multistartLowerVariables.append(tk.StringVar(self, "1E-5"))
                                self.multistartUpperVariables.append(tk.StringVar(self, "1E5"))
                                self.multistartSpacingVariables.append(tk.StringVar(self, "Logarithmic"))
                                self.multistartCustomVariables.append(tk.StringVar(self, "0.1,1,10,100,1000"))
                                self.multistartNumberVariables.append(tk.StringVar(self, "10"))
                                self.upperLimits.append(tk.StringVar(self, "inf"))
                                self.lowerLimits.append(tk.StringVar(self, "-inf"))
                        self.numParams += 1
                        if (self.numParams == 1):
                            self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[0])
                            self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[0])
                            self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[0])
                            self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[0])
                            self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[0])
                            self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[0])
                            self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[0])
                            self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[0])
                            if (self.multistartCheckboxVariables[0].get() == 0):
                                self.multistartSpacing.configure(state="disabled")
                                self.multistartUpperEntry.configure(state="disabled")
                                self.multistartLowerEntry.configure(state="disabled")
                                self.multistartNumberEntry.configure(state="disabled")
                                self.multistartCustomEntry.configure(state="disabled")
                            else:
                                if (self.multistartSpacingVariables[0].get() == "Custom"):
                                    self.multistartSpacing.configure(state="readonly")
                                    self.multistartUpperEntry.configure(state="normal")
                                    self.multistartLowerEntry.configure(state="normal")
                                    self.multistartNumberEntry.configure(state="normal")
                                    self.multistartCustomEntry.configure(state="normal")
                                    self.multistartLowerFrame.pack_forget()
                                    self.multistartUpperFrame.pack_forget()
                                    self.multistartNumberFrame.pack_forget()
                                    self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                else:
                                    self.multistartSpacing.configure(state="readonly")
                                    self.multistartUpperEntry.configure(state="normal")
                                    self.multistartLowerEntry.configure(state="normal")
                                    self.multistartNumberEntry.configure(state="normal")
                                    self.multistartCustomEntry.configure(state="normal")
                        self.paramNameValues.append(tk.StringVar(self, pname))
                        self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10, validate="all", validatecommand=self.valcom))
                        self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                        self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed", "custom"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                        self.paramComboboxes[-1].bind("<<ComboboxSelected>>", partial(self.limitChanged, self.numParams-1))
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
                        self.paramListbox.insert(tk.END, pname)
                        if (self.paramComboboxValues[0].get() == "fixed"):
                            self.advancedLowerLimitEntry.configure(state="disabled")
                            self.advancedUpperLimitEntry.configure(state="disabled")
            toLoad.close()
            self.paramListbox.selection_clear(0, tk.END)
            self.paramListbox.select_set(0)
            self.paramListboxSelection = 0
            self.paramListbox.event_generate("<<ListboxSelect>>")
            
            if (self.loadFormulaOtherVariable.get()):
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
                self.noiseEntryValue.set(alphaVal)
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
            if (self.loadFormulaCodeVariable.get()):
                self.customFormula.delete("1.0", tk.END)
                self.customFormula.insert("1.0", formulaIn.rstrip())    #Insert custom formula into code box and remove trailing whitespace
            self.keyup("")
            self.loadFormulaWindow.withdraw()
        except Exception:
            messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
    
    def loadFormulaCommand(self, e=None):
        item = self.loadFormulaTree.selection()[0]  
        item_text = self.loadFormulaTree.item(item,"text")
        if (item_text.endswith(".mmformula")):
            path_to_child = self.loadFormulaTree.item(item,"text")
            parent_iid = self.loadFormulaTree.parent(item)
            parent_node = self.loadFormulaTree.item(parent_iid)['text']
            while parent_iid != '':
                path_to_child = os.path.join(parent_node, path_to_child)
                parent_iid = self.loadFormulaTree.parent(parent_iid)
                parent_node = self.loadFormulaTree.item(parent_iid)['text']
        try:
            self.loadedFormula = item_text[:-10]
            toLoad = open(path_to_child)
            firstLine = toLoad.readline()
            fileVersion = 0
            if "version: 1.1" in firstLine:
                fileVersion = 1
            elif "version: 1.2" in firstLine:
                fileVersion = 2
            if (fileVersion == 0):
                fileToLoad = firstLine.split("filename: ")[1][:-1]
            else:
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
                if (fileVersion == 0):
                    if "mmparams:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
                elif (fileVersion == 1):
                    if "description:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
                elif (fileVersion == 2):
                    if "imports:" in nextLineIn:
                        break
                    else:
                        formulaIn += nextLineIn
            if (self.loadFormulaParamsVariable.get()):
                for i in range(len(self.paramNameEntries)):
                    self.paramNameEntries[i].grid_remove()
                    self.paramNameLabels[i].grid_remove()
                    self.paramValueEntries[i].grid_remove()
                    self.paramValueLabels[i].grid_remove()
                    self.paramComboboxes[i].grid_remove()
                    self.paramDeleteButtons[i].grid_remove()
                self.numParams = 0
                self.paramNameValues.clear()
                self.paramNameEntries.clear()
                self.paramNameLabels.clear()
                self.paramValueLabels.clear()
                self.paramComboboxValues.clear()
                self.paramComboboxes.clear()
                self.paramValueValues.clear()
                self.paramValueEntries.clear()
                self.paramDeleteButtons.clear()
                self.paramListbox.delete(0, tk.END)
                self.multistartCheckboxVariables.clear()
                self.multistartSpacingVariables.clear()
                self.multistartLowerVariables.clear()
                self.multistartUpperVariables.clear()
                self.multistartNumberVariables.clear()
                self.multistartCustomVariables.clear()
                self.advancedLowerLimitEntry.configure(state="normal")
                self.advancedUpperLimitEntry.configure(state="normal")
            imports = ""
            if (fileVersion == 2):
                while True:
                    nextLineIn = toLoad.readline()
                    if "description:" in nextLineIn:
                        break
                    else:
                        imports += nextLineIn
            toImport = imports.rstrip().split('\n')
            for tI in toImport:
                if (len(tI) > 0):
                    self.importPathListbox.insert(tk.END, tI)
            if (fileVersion != 0):
                self.formulaDescriptionEntry.delete("1.0", tk.END)
                descriptionIn = ""
                while True:
                    nextLineIn = toLoad.readline()
                    if "latex:" in nextLineIn:
                        break
                    else:
                        descriptionIn += nextLineIn
                latexIn = toLoad.readline()
                self.formulaDescriptionEntry.insert("1.0", descriptionIn.rstrip())
                self.formulaDescriptionLatexEntry.delete("1.0", tk.END)
                self.formulaDescriptionLatexEntry.insert("1.0", latexIn.rstrip())
                self.latexAx.clear()
                self.latexAx.axis("off")
                self.latexCanvas.draw()
                self.graphLatex()
                toLoad.readline()
            if (self.loadFormulaParamsVariable.get()):
                while True:
                    p = toLoad.readline()
                    if not p:
                        break
                    else:
                        p = p[:-1]
                        try:
                            pname, pval, pcombo, plimup, plimlow, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                            self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                            self.multistartLowerVariables.append(tk.StringVar(self, plower))
                            self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                            self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                            self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                            self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                            self.upperLimits.append(tk.StringVar(self, plimup))
                            self.lowerLimits.append(tk.StringVar(self, plimlow))
                        except Exception:
                            try:
                                pname, pval, pcombo, pcheck, pspacing, pupper, plower, pnumber, pcustom = p.split("~=~")
                                self.multistartCheckboxVariables.append(tk.IntVar(self, int(pcheck)))
                                self.multistartLowerVariables.append(tk.StringVar(self, plower))
                                self.multistartUpperVariables.append(tk.StringVar(self, pupper))
                                self.multistartSpacingVariables.append(tk.StringVar(self, pspacing))
                                self.multistartCustomVariables.append(tk.StringVar(self, pcustom))
                                self.multistartNumberVariables.append(tk.StringVar(self, pnumber))
                                self.upperLimits.append(tk.StringVar(self, "inf"))
                                self.lowerLimits.append(tk.StringVar(self, "-inf"))
                            except Exception:
                                pname, pval, pcombo = p.split("~=~")
                                self.multistartCheckboxVariables.append(tk.IntVar(self, 0))
                                self.multistartLowerVariables.append(tk.StringVar(self, "1E-5"))
                                self.multistartUpperVariables.append(tk.StringVar(self, "1E5"))
                                self.multistartSpacingVariables.append(tk.StringVar(self, "Logarithmic"))
                                self.multistartCustomVariables.append(tk.StringVar(self, "0.1,1,10,100,1000"))
                                self.multistartNumberVariables.append(tk.StringVar(self, "10"))
                                self.upperLimits.append(tk.StringVar(self, "inf"))
                                self.lowerLimits.append(tk.StringVar(self, "-inf"))
                        self.numParams += 1
                        if (self.numParams == 1):
                            self.multistartCheckbox.configure(variable=self.multistartCheckboxVariables[0])
                            self.multistartLowerEntry.configure(textvariable=self.multistartLowerVariables[0])
                            self.multistartUpperEntry.configure(textvariable=self.multistartUpperVariables[0])
                            self.multistartSpacing.configure(textvariable=self.multistartSpacingVariables[0])
                            self.multistartNumberEntry.configure(textvariable=self.multistartNumberVariables[0])
                            self.multistartCustomEntry.configure(textvariable=self.multistartCustomVariables[0])
                            self.advancedLowerLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.advancedUpperLimitFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.multistartCheckbox.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartSpacingFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartLowerFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.multistartUpperFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5)
                            self.multistartNumberFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                            self.advancedLowerLimitEntry.configure(textvariable=self.lowerLimits[0])
                            self.advancedUpperLimitEntry.configure(textvariable=self.upperLimits[0])
                            if (self.multistartCheckboxVariables[0].get() == 0):
                                self.multistartSpacing.configure(state="disabled")
                                self.multistartUpperEntry.configure(state="disabled")
                                self.multistartLowerEntry.configure(state="disabled")
                                self.multistartNumberEntry.configure(state="disabled")
                                self.multistartCustomEntry.configure(state="disabled")
                            else:
                                if (self.multistartSpacingVariables[0].get() == "Custom"):
                                    self.multistartSpacing.configure(state="readonly")
                                    self.multistartUpperEntry.configure(state="normal")
                                    self.multistartLowerEntry.configure(state="normal")
                                    self.multistartNumberEntry.configure(state="normal")
                                    self.multistartCustomEntry.configure(state="normal")
                                    self.multistartLowerFrame.pack_forget()
                                    self.multistartUpperFrame.pack_forget()
                                    self.multistartNumberFrame.pack_forget()
                                    self.multistartCustomFrame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=5, pady=5)
                                else:
                                    self.multistartSpacing.configure(state="readonly")
                                    self.multistartUpperEntry.configure(state="normal")
                                    self.multistartLowerEntry.configure(state="normal")
                                    self.multistartNumberEntry.configure(state="normal")
                                    self.multistartCustomEntry.configure(state="normal")
                        self.paramNameValues.append(tk.StringVar(self, pname))
                        self.paramNameEntries.append(ttk.Entry(self.pframe, textvariable=self.paramNameValues[self.numParams-1], width=10, validate="all", validatecommand=self.valcom))
                        self.paramNameLabels.append(tk.Label(self.pframe, text="Name: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramValueLabels.append(tk.Label(self.pframe, text="Value: ", background=self.backgroundColor, foreground=self.foregroundColor))
                        self.paramComboboxValues.append(tk.StringVar(self, pcombo))
                        self.paramComboboxes.append(ttk.Combobox(self.pframe, textvariable=self.paramComboboxValues[self.numParams-1], value=("+", "+ or -", "-", "fixed", "custom"), justify=tk.CENTER, state="readonly", exportselection=0, width=6))
                        self.paramComboboxes[-1].bind("<<ComboboxSelected>>", partial(self.limitChanged, self.numParams-1))
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
                        self.paramListbox.insert(tk.END, pname)
                        if (self.paramComboboxValues[0].get() == "fixed"):
                            self.advancedLowerLimitEntry.configure(state="disabled")
                            self.advancedUpperLimitEntry.configure(state="disabled")
            toLoad.close()
            self.paramListbox.selection_clear(0, tk.END)
            self.paramListbox.select_set(0)
            self.paramListboxSelection = 0
            self.paramListbox.event_generate("<<ListboxSelect>>")
            
            if (self.loadFormulaOtherVariable.get()):
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
                self.noiseEntryValue.set(alphaVal)
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
            if (self.loadFormulaCodeVariable.get()):
                self.customFormula.delete("1.0", tk.END)
                self.customFormula.insert("1.0", formulaIn.rstrip())    #Insert custom formula into code box and remove trailing whitespace
            self.keyup("")
            self.loadFormulaWindow.withdraw()
        except Exception:
            messagebox.showerror("File error", "Error 42: \nThere was an error loading or reading the file")
    
    def closeWindows(self):
        self.upperSpinboxVariable.set(str(self.upDelete))
        self.lowerSpinboxVariable.set(str(self.lowDelete))
        self.freqWindow.withdraw()
        self.formulaDescriptionWindow.withdraw()
        self.importPathWindow.withdraw()
        self.loadFormulaWindow.withdraw()
        self.simulationWindow.withdraw()
        self.paramPopup.withdraw()
        self.advancedOptionsWindow.withdraw()
        try:
            self.bootstrapLabel.destroy()
            self.bootstrapEntry.destroy()
            self.bootstrapButton.destroy()
            self.bootstrapCancel.destroy()
            self.bootstrapPopup.destroy()
        except Exception:
            pass
        for resultWindow in self.resultsWindows:
            try:
                resultWindow.destroy()
            except Exception:
                pass
        for resultPlotBig in self.resultPlotBigs:
            try:
                resultPlotBig.destroy()
            except Exception:
                pass
        for resultPlotBigFig in self.resultPlotBigFigs:
            try:
                resultPlotBigFig.clear()
            except Exception:
                pass
        for resultPlot in self.resultPlots:
            try:
                resultPlot.destroy()
            except Exception:
                pass
        for resultPlotFig in self.resultPlotFigs:
            try:
                resultPlotFig.clear()
            except Exception:
                pass
        for nyCanvasButton in self.saveNyCanvasButtons:
            try:
                nyCanvasButton.destroy()
            except Exception:
                pass
        for nyCanvasButton_ttp in self.saveNyCanvasButton_ttps:
            try:
                del nyCanvasButton_ttp
            except Exception:
                pass
        self.alreadyPlotted = False
        for s in self.simPlots:
            try:
                s.destroy()
            except Exception:
                pass
        for s in self.simPlotFigs:
            try:
                s.clear()
                plt.close(s)
            except Exception:
                pass
        for s in self.simPlotBigs:
            try:
                s.destroy()
            except Exception:
                pass
        for s in self.simPlotBigFigs:
            try:
                s.clear()
                plt.close(s)
            except Exception:
                pass
        for s in self.simSaveNyCanvasButtons:
            try:
                s.destroy()
            except Exception:
                pass
        for s in self.simSaveNyCanvasButton_ttps:
            try:
                del s
            except Exception:
                pass
    
    def cancelThreads(self):
        for t in self.currentThreads:
            t.terminate()