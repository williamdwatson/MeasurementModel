# -*- coding: utf-8 -*-
"""
Created on Mon Aug  6 15:03:45 2018

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
from tkinter import messagebox
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilenames, asksaveasfile
import os, sys, threading, queue
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import errorFitting
import pyperclip
#--------------------------------pyperclip-----------------------------------
#     Source: https://pypi.org/project/pyperclip/
#     Email: al@inventwithpython.com
#     By: Al Sweigart
#     License: BSD
#----------------------------------------------------------------------------

class FileExtensionError(Exception):
    """Exception for if the file has an unknown extension"""
    pass

class FileLengthError(Exception):
    """Exception for if the data in the file is not of the same length"""
    pass

class ThreadedTask(threading.Thread):
    def __init__(self, queue, w, c, stdr, stdj, r, j, modz, sig, g, rec, avgre, detrend):
        threading.Thread.__init__(self)
        self.queue = queue
        self.we = w
        self.ch = c
        self.sdr = stdr
        self.sdj = stdj
        self.rdat = r
        self.jdat = j
        self.moddat = modz
        self.sigm = sig
        self.guess = g
        self.choiceRe = rec
        self.averageRe = avgre
        self.dc = detrend
    def run(self):
        alph, bet, gamm, delt, sigalph, sigbet, siggamm, sigdelt, chi, mape, cov = errorFitting.findErrorFit(self.we, self.ch, self.sdr, self.sdj, self.rdat, self.jdat, self.moddat, self.sigm, self.guess, self.choiceRe, self.averageRe, self.dc)
        self.queue.put((alph, bet, gamm, delt, sigalph, sigbet, siggamm, sigdelt, chi, mape, cov))

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

class eF(tk.Frame):
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
        
        self.noneLoaded = True
        self.resPlots = []
        self.wdata = []
        self.avgRe = []
        self.rres = []
        self.jres = []
        self.zr = []
        self.zj = []
        self.modz = []
        self.sigma = []
        self.numFiles = 0
        self.fileNames = []
        self.copyDetrend = "Off"
        self.copyWeighting = "Variance"
        self.copyMA = "5 point"
        self.covar = []
        
        def addFile():
            n = askopenfilenames(initialdir=self.topGUI.getCurrentDirectory(), filetypes =[("Residual Error File (*.mmerrors)", "*.mmerrors")],title = "Choose a file")
            if (len(n) != 0):
                try:
                    alreadyWarned = False
                    directory = os.path.dirname(str(n[0]))
                    self.topGUI.setCurrentDirectory(directory)
                    for name in n:
                        fname, fext = os.path.splitext(name)
                        if (fext != ".mmerrors"):
                            raise FileExtensionError
                        with open(name,'r') as UseFile:
                            filetext = UseFile.read()
                            lines = filetext.splitlines()
                        if ("frequency" in lines[2].lower()):
                            data = np.loadtxt(name, skiprows=3)
                        else:
                            data = np.loadtxt(name)
                        re_in = data[1,0]
                        w_in = data[2:,0]
                        rstddev_in = data[2:,1]
                        jstddev_in = data[2:,2]
                        r_in = data[2:,3]
                        j_in = data[2:,4]
                        z_in = data[2:,5]
                        sigma_in = data[2:,6]
                        if (not (len(w_in) == len(rstddev_in) and len(rstddev_in) == len(jstddev_in) and len(jstddev_in) == len(r_in) and len(r_in) == len(j_in) and len(j_in) == len(z_in) and len(z_in) == len(sigma_in))):
                            raise FileLengthError
                        for i in range(len(w_in)):
                            float(re_in)
                            float(w_in[i])
                            float(rstddev_in[i])
                            float(jstddev_in[i])
                            float(r_in[i])
                            float(j_in[i])
                            float(z_in[i])
                            float(sigma_in[i])
                        if (name in self.fileListbox.get(0, tk.END) and not alreadyWarned):
                            alreadyWarned = True
                            if (len(n) > 1):
                                messagebox.showwarning("File already loaded", "Warning: One or more of the selected files is already loaded and is being loaded again")
                            else:
                                messagebox.showwarning("File already loaded", "Warning: The selected file is already loaded and is being loaded again")
                        self.fileListbox.insert(tk.END, name)
                        if (self.noneLoaded):
                            self.noneLoaded = False
                        self.wdata.append(w_in)
                        self.avgRe.append(re_in)
                        self.rres.append(rstddev_in)
                        self.jres.append(jstddev_in)
                        self.zr.append(r_in)
                        self.zj.append(j_in)
                        self.modz.append(z_in)
                        self.sigma.append(sigma_in)
                        self.numFiles += 1
                        self.fileNames.append(os.path.basename(name))
                        self.regressButton.configure(state="normal")
                except FileExtensionError:
                    messagebox.showerror("File error", "Error 30:\nThe file has an unknown extension")
                except FileLengthError:
                    messagebox.showerror("Length error", "Error 31:\nThe number of data do not match")
                except Exception:
                    messagebox.showerror("File error", "Error 33:\nThere was an error reading the file")
                
        def removeFile():
            if (self.fileListbox.size() == 0):
                self.noneLoaded = True
                self.wdata = []
                self.rres = []
                self.jres = []
                self.zr = []
                self.zj = []
                self.modz = []
                self.sigma = []
                self.avgRe = []
                self.fileNames = []
                self.numFiles = 0
                self.regressButton.configure(state="disabled")
                self.plotButton.grid_remove()
            else:
                items = list(map(int, self.fileListbox.curselection()))
                if (len(items) != 0):
                    del self.wdata[items[0]]
                    del self.rres[items[0]]
                    del self.jres[items[0]]
                    del self.zr[items[0]]
                    del self.zj[items[0]]
                    del self.modz[items[0]]
                    del self.sigma[items[0]]
                    del self.avgRe[items[0]]
                    del self.fileNames[items[0]]
                    self.fileListbox.delete(tk.ANCHOR)
                    self.numFiles -= 1
                    if (self.numFiles < 1):
                        self.regressButton.configure(state="disabled")
                        self.plotButton.grid_remove()
                if (self.fileListbox.size() == 0):
                    self.noneLoaded = True
                    self.wdata = []
                    self.avgRe = []
                    self.rres = []
                    self.jres = []
                    self.zr = []
                    self.zj = []
                    self.modz = []
                    self.sigma = []
                    self.fileNames = []
                    self.numFiles = 0
        
        def popup_inputFiles(event):
            try:
                self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu.grab_release()
        
        def process_queue():
            try:
                self.alpha, self.beta, self.gamma, self.delta, self.sigmaAlpha, self.sigmaBeta, self.sigmaGamma, self.sigmaDelta, chi, mape, self.covar = self.queue.get(0)
                if (self.alpha == "-"):
                    messagebox.showerror("Fitting error", "Error 34:\nThe fitting failed")
                    return
                self.resultsView.delete(*self.resultsView.get_children())
                if (self.chosenParams[0]):
                    if (abs(self.sigmaAlpha*2*100/self.alpha) >= 100):
                        self.resultsView.insert("", tk.END, text="", values=("\u03B1", "%.7g"%self.alpha, "%.4g"%self.sigmaAlpha, "%.3g"%abs(self.sigmaAlpha*2*100/self.alpha)+"%"), tags=("bad",))
                    else:
                        self.resultsView.insert("", tk.END, text="", values=("\u03B1", "%.7g"%self.alpha, "%.4g"%self.sigmaAlpha, "%.3g"%abs(self.sigmaAlpha*2*100/self.alpha)+"%"))
                    self.alphaEntryVariable.set("%.4g"%self.alpha)
                if (self.chosenParams[1]):
                    if (abs(self.sigmaBeta*2*100/self.beta) >= 100):
                        self.resultsView.insert("", tk.END, text="", values=("\u03B2", "%.7g"%self.beta, "%.4g"%self.sigmaBeta, "%.3g"%abs(self.sigmaBeta*2*100/self.beta)+"%"), tags=("bad",))
                    else:
                        self.resultsView.insert("", tk.END, text="", values=("\u03B2", "%.7g"%self.beta, "%.4g"%self.sigmaBeta, "%.3g"%abs(self.sigmaBeta*2*100/self.beta)+"%"))
                    self.betaEntryVariable.set("%.4g"%self.beta)
                if (self.chosenParams[2]):
                    if (abs(self.sigmaGamma*2*100/self.gamma) >= 100):
                        self.resultsView.insert("", tk.END, text="", values=("\u03B3", "%.7g"%self.gamma, "%.4g"%self.sigmaGamma, "%.3g"%abs(self.sigmaGamma*2*100/self.gamma)+"%"), tags=("bad",))
                    else:
                        self.resultsView.insert("", tk.END, text="", values=("\u03B3", "%.7g"%self.gamma, "%.4g"%self.sigmaGamma, "%.3g"%abs(self.sigmaGamma*2*100/self.gamma)+"%"))
                    self.gammaEntryVariable.set("%.4g"%self.gamma)
                if (self.chosenParams[3]):
                    if (abs(self.sigmaDelta*2*100/self.delta) >= 100):
                        self.resultsView.insert("", tk.END, text="", values=("\u03B4", "%.7g"%self.delta, "%.4g"%self.sigmaDelta, "%.3g"%abs(self.sigmaDelta*2*100/self.delta)+"%"), tags=("bad",))
                    else:
                        self.resultsView.insert("", tk.END, text="", values=("\u03B4", "%.7g"%self.delta, "%.4g"%self.sigmaDelta, "%.3g"%abs(self.sigmaDelta*2*100/self.delta)+"%"))
                    self.deltaEntryVariable.set("%.4g"%self.delta)
                self.resultsView.tag_configure("bad", background="yellow")
                self.resultsChi.configure(text="Mean absolute %% error = %.4g%%"%abs(mape))
                self.resultsFrame.grid(column=0, row=5, pady=5, sticky="W")
                self.plotButton.configure(state="normal")
                self.plotButton.grid(column=0, row=6, pady=5, sticky="W")
                self.copyDetrend = self.detrendComboboxVariable.get()
                self.copyWeighting = self.weightingComboboxVariable.get()
                self.copyMA = self.movingAverageComboboxVariable.get()
            except queue.Empty:
                self.after(100, process_queue)
        
        def regressErrors():        
            choices = [False, False, False, False]
            fittingGuesses = [.1, .1, .1, .1]
            reChoice = False
            
            weight = 0
            if (self.weightingComboboxVariable.get() == "Variance"):
                if (self.movingAverageComboboxVariable.get() == "None"):
                    weight = 1
                elif (self.movingAverageComboboxVariable.get() == "3 point"):
                    for i in range(len(self.rres)):
                        if (len(self.rres[i]) < 3):
                            messagebox.showerror("Variance error", "Error 35:\nThere are too few frequencies to use a 3-point moving average")
                            return
                    weight = 2
                elif (self.movingAverageComboboxVariable.get() == "5 point"):
                    for i in range(len(self.rres)):
                        if (len(self.rres[i]) < 5):
                            messagebox.showerror("Variance error", "Error 36:\nThere are too few frequencies to use a 5-point moving average")
                            return
                    weight = 3
            
            if (self.alphaCheckboxVariable.get() == 1):
                choices[0] = True
                if (self.alphaEntryVariable.get() == ""):
                    fittingGuesses[0] = 0
                else:
                    try:
                        fittingGuesses[0] = float(self.alphaEntryVariable.get())
                    except Exception:
                        messagebox.showerror("Value error", "Error 37:\nThe alpha value guess is invalid")
                        return
            if (self.betaCheckboxVariable.get() == 1):
                choices[1] = True
                if (self.betaEntryVariable.get() == ""):
                    fittingGuesses[1] = 0
                else:
                    try:
                        fittingGuesses[1] = float(self.betaEntryVariable.get())
                    except Exception:
                        messagebox.showerror("Value error", "Error 38:\nThe beta value guess is invalid")
                        return
            if (self.gammaCheckboxVariable.get() == 1):
                choices[2] = True
                if (self.gammaEntryVariable.get() == ""):
                    fittingGuesses[2] = 0
                else:
                    try:
                        fittingGuesses[2] = float(self.gammaEntryVariable.get())
                    except Exception:
                        messagebox.showerror("Value error", "Error 39:\nThe gamma value guess is invalid")
                        return
            if (self.deltaCheckboxVariable.get() == 1):
                choices[3] = True
                if (self.deltaEntryVariable.get() == ""):
                    fittingGuesses[3] = 0
                else:
                    try:
                        fittingGuesses[3] = float(self.deltaEntryVariable.get())
                    except Exception:
                        messagebox.showerror("Value error", "Error 40:\nThe delta value guess is invalid")
                        return
            
            if (self.reCheckboxVariable.get() == 1 and self.betaCheckboxVariable.get() == 1):
                reChoice = True
            
            if (not any(choices)):
                messagebox.showerror("Fitting error", "Error 41:\nAt least one parameter must be chosen for fitting")
                return
            self.chosenParams = choices
            
            detrendChoice = 1
            if (self.detrendComboboxVariable.get() == "On"):
                detrendChoice = 3

            self.queue = queue.Queue()
            ThreadedTask(self.queue, weight, choices, self.rres, self.jres, self.zr, self.zj, self.modz, self.sigma, fittingGuesses, reChoice, self.avgRe, detrendChoice).start()
            self.after(100, process_queue)
        
        def plotResiduals():
            for i in range(self.numFiles):
                resPlot = tk.Toplevel()
                self.resPlots.append(resPlot)
                resPlot.title("Standard Deviations Plot: " + self.fileNames[i])
                resPlot.iconbitmap(resource_path("img/elephant3.ico"))
                
                with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                    figError = Figure(figsize=(5,4), dpi=100)
                    a = figError.add_subplot(111)
                    figError.set_facecolor(self.backgroundColor)
                    a.set_facecolor(self.backgroundColor)
                    a.yaxis.set_ticks_position("both")
                    a.yaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")     #Make the ticks point inwards
                    a.xaxis.set_ticks_position("both")
                    a.xaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")     #Make the ticks point inwards
                    imagColor = "tab:orange"
                    realColor = "tab:blue"
                    fitColor = "green"
                    if (self.topGUI.getTheme() == "dark"):
                        imagColor = "gold"
                        realColor = "cyan"
                        fitColor = "springgreen"
                    else:
                        imagColor = "tab:orange"
                        realColor = "tab:blue"
                        fitColor = "green"
                    a.plot(self.wdata[i], self.rres[i], "o", markerfacecolor="None", color=realColor, label="Real")
                    a.plot(self.wdata[i], self.jres[i], "^", markerfacecolor="None", color=imagColor, label="Imaginary")
                    model = np.zeros(len(self.wdata[i]))
                    for j in range(len(self.wdata[i])):
                        if (self.chosenParams[0]):
                            model[j] += self.alpha*abs(self.zj[i][j])
                        if (self.chosenParams[1]):
                            model[j] += self.beta*abs(self.zr[i][j])
                        if (self.chosenParams[2]):
                            model[j] += self.gamma*abs(self.modz[i][j])**2
                        if (self.chosenParams[3]):
                            model[j] += self.delta
                    a.plot(self.wdata[i], model, color=fitColor)
                    
                    a.set_xscale("log")
                    a.set_yscale("log")
                    a.set_title("Standard Deviations: " + self.fileNames[i], color=self.foregroundColor)
                    a.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                    a.set_ylabel("σr, σj / Ω", color=self.foregroundColor)
                    a.legend()
                    figError.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                    errorCanvasInput = FigureCanvasTkAgg(figError, resPlot)
                    errorCanvasInput.draw()
                    errorCanvasInput.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                    toolbar = NavigationToolbar2Tk(errorCanvasInput, resPlot)    #Enables the zoom and move toolbar for the plot
                    toolbar.update()
                errorCanvasInput._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        def checkboxCommand():  #This logic was bothering me so I hard-coded it
            a = self.alphaCheckboxVariable.get() == 1
            b = self.betaCheckboxVariable.get() == 1
            re = self.reCheckboxVariable.get() == 1
            g = self.gammaCheckboxVariable.get() == 1
            d = self.deltaCheckboxVariable.get() == 1
            if (a):
                self.alphaEntry.grid(column=2, row=0, sticky="W")
                self.alphaEqual.grid(column=1, row=0, sticky="W")
            else:
                self.alphaEntry.grid_remove()
                self.alphaEqual.grid_remove()
            if (b):
                self.betaEntry.grid(column=2, row=1, sticky="W")
                self.betaEqual.grid(column=1, row=1, sticky="W")
                self.reCheckbox.grid(column=0, row=2, sticky="W")
            else:
                self.betaEntry.grid_remove()
                self.betaEqual.grid_remove()
                self.reCheckbox.grid_remove()
            if (g):
                self.gammaEntry.grid(column=2, row=3, sticky="W")
                self.gammaEqual.grid(column=1, row=3, sticky="W")
            else:
                self.gammaEntry.grid_remove()
                self.gammaEqual.grid_remove()
            if (d):
                self.deltaEntry.grid(column=2, row=4, sticky="W")
                self.deltaEqual.grid(column=1, row=4, sticky="W")
            else:
                self.deltaEntry.grid_remove()
                self.deltaEqual.grid_remove()
            
            #---All checkboxes---
            if (a and b and re and g and d):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr| - R\u2091> + \u03B3<|Z|>\u00B2 + \u03B4"
            #---All but one---
            elif (a and b and re and g and not d):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr| - R\u2091> + \u03B3<|Z|>\u00B2"
            elif (a and b and re and d and not g):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr| - R\u2091> + \u03B4"
            elif (a and b and g and d and not re):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr|> + \u03B3<|Z|>\u00B2 + \u03B4"
            elif (a and g and d and not b):
                self.regressText = "\u03B1<|Zj|> + \u03B3<|Z|>\u00B2 + \u03B4"
            elif (b and re and g and d and not a):
                self.regressText = "\u03B2<|Zr| - R\u2091> + \u03B3<|Z|>\u00B2 + \u03B4"
            #---Only three---
            elif (a and b and g and not (re or d)):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr|> + \u03B3<|Z|>\u00B2"
            elif (a and b and d and not (re or g)):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr|> + \u03B4"
            elif (b and g and d and not (a or re)):
                self.regressText = "\u03B2<Zr> + \u03B3<|Z|>\u00B2 + \u03B4"
            #---Only two---
            elif (a and b and not (re or g or d)):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr|>"
            elif (a and b and re and not (g or d)):
                self.regressText = "\u03B1<|Zj|> + \u03B2<|Zr| - R\u2091>"
            elif (a and g and not (b or d)):
                self.regressText = "\u03B1<|Zj|> + \u03B3<|Z|>\u00B2"
            elif (a and d and not (b or g)):
                self.regressText = "\u03B1<|Zj|> + \u03B4"
            elif (b and g and not (re or a or d)):
                self.regressText = "\u03B2<|Zr|> + \u03B3<|Z|>\u00B2"
            elif (b and d and not (re or a or g)):
                self.regressText = "\u03B2<|Zr|> + \u03B4"
            elif (b and re and g and not (a or d)):
                self.regressText = "\u03B2<|Zr| - R\u2091> + \u03B3<|Z|>\u00B2"
            elif (b and re and d and not (a or g)):
                self.regressText = "\u03B2<|Zr| - R\u2091> + \u03B4"
            elif (g and d and not (a or b)):
                self.regressText = "\u03B3<|Z|>\u00B2 + \u03B4"
            #---Only one---
            elif (a and not (b or g or d)):
                self.regressText = "\u03B1<|Zj|>"
            elif ((b and re) and not (a or g or d)):
                self.regressText = "\u03B2<|Zr| - R\u2091>"
            elif (b and not (re or a or g or d)):
                self.regressText = "\u03B2<|Zr|>"
            elif (g and not (a or b or d)):
                self.regressText = "\u03B3<|Z|>\u00B2"
            elif (d and not (a or b or g)):
                self.regressText = "\u03B4"
            #---No checkboxes---
            else:
                self.regressText = ""
            
            self.regressFunction.configure(text=self.regressText)
        
        def checkWeight(event):
            if (self.weightingComboboxVariable.get() == "None"):
                self.movingAverageCombobox.grid_remove()
                self.movingAverageLabel.grid_remove()
            else:
                self.movingAverageLabel.grid(column=0, row=1, pady=2, sticky="W")
                self.movingAverageCombobox.grid(column=1, row=1)
        
        def handle_click(event):
            if self.resultsView.identify_region(event.x, event.y) == "separator":
                if (self.resultsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def handle_motion(event):
            if self.resultsView.identify_region(event.x, event.y) == "separator":
                if (self.resultsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def copyResults():
            self.copyResultsButton.configure(text="Copied")
            self.clipboard_clear()
            stringToCopy = ""
            for name in self.fileListbox.get(0, tk.END):
                stringToCopy += name + "\n"
            stringToCopy += "Parameter\tValue\tStd. Dev.\n"
            for child in self.resultsView.get_children():
                resultName, resultValue, resultStdDev, result95 = self.resultsView.item(child, 'values')
                stringToCopy += resultName + "\t" + resultValue + "\t" + resultStdDev + "\n"
            v = np.sqrt(np.diag(self.covar))
            outer_v = np.outer(v, v)
            correl = np.zeros((len(self.covar), len(self.covar)))
            for i in range(len(self.covar)):        #Could do self.covar/outer_v, but possibility exists of divide-by-0 error
                for j in range(len(self.covar)):
                    if (outer_v[i][j] == 0):
                        correl[i][j] = 0
                    else:
                        correl[i][j] = self.covar[i][j] / outer_v[i][j]
            stringToCopy += "\n---Correlation---\n\tAlpha\tBeta\tGamma\tDelta\n"
            stringToCopy += "Alpha\t" + str(correl[0][0]) + "\t" + str(correl[0][1]) + "\t" + str(correl[0][2]) + "\t" + str(correl[0][3])
            stringToCopy += "\nBeta\t" + str(correl[1][0]) + "\t" + str(correl[1][1]) + "\t" + str(correl[1][2]) + "\t" + str(correl[1][3])
            stringToCopy += "\nGamma\t" + str(correl[2][0]) + "\t" + str(correl[2][1]) + "\t" + str(correl[2][2]) + "\t" + str(correl[2][3])
            stringToCopy += "\nDelta\t" + str(correl[3][0]) + "\t" + str(correl[3][1]) + "\t" + str(correl[3][2]) + "\t" + str(correl[3][3]) + "\n"
            stringToCopy += "\nDetrend option:\t" + self.copyDetrend + "\nWeighting:\t" + self.copyWeighting
            if (self.copyWeighting == "Variance"):
                stringToCopy += "\n" + self.copyMA + " moving average"
            pyperclip.copy(stringToCopy)
            self.after(500, lambda : self.copyResultsButton.configure(text="Copy values and std. devs."))
        
        self.fileLabel = tk.Label(self, text="Files to use:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.fileLabel.grid(column=0, row=0, sticky="W")
        self.fileFrame = tk.Frame(self, bg=self.backgroundColor)
        self.fileListboxScrollbar = ttk.Scrollbar(self.fileFrame, orient=tk.VERTICAL)
        self.fileListboxScrollbarHorizontal = ttk.Scrollbar(self.fileFrame, orient=tk.HORIZONTAL)
        self.fileListbox = tk.Listbox(self.fileFrame, height=3, width=65, selectmode=tk.BROWSE, activestyle='none', yscrollcommand=self.fileListboxScrollbar.set, xscrollcommand=self.fileListboxScrollbarHorizontal.set)
        self.fileListboxScrollbar['command'] = self.fileListbox.yview
        self.fileListboxScrollbarHorizontal['command'] = self.fileListbox.xview
        self.fileBrowse = ttk.Button(self.fileFrame, text="Add file(s)", command=addFile)
        self.fileRemove = ttk.Button(self.fileFrame, text="Remove", command=removeFile)#lambda lb=self.fileListbox: lb.delete(tk.ANCHOR))
        self.fileListbox.grid(column=0, row=0, rowspan=2)
        self.fileListboxScrollbar.grid(column=1, row=0, rowspan=2, sticky="NS")
        self.fileListboxScrollbarHorizontal.grid(column=0, row=2, sticky="EW")
        self.fileBrowse.grid(column=2, row=0, padx=2)
        self.fileRemove.grid(column=2, row=1, padx=2)
        self.fileFrame.grid(column=0, row=1, sticky="W")
        addFiles_ttp = CreateToolTip(self.fileBrowse, 'Browse for one or more .mmerrors files')
        removeFile_ttp = CreateToolTip(self.fileRemove, 'Remove selected file from the list')
        
        self.popup_menu = tk.Menu(self.fileListbox, tearoff=0)
        self.popup_menu.add_command(label="Remove file", command=removeFile)
        self.fileListbox.bind("<Button-3>", popup_inputFiles)
        
        self.regressFunctionFrame = tk.Frame(self, bg=self.backgroundColor)
        self.regressFunctionLabel = tk.Label(self.regressFunctionFrame, text="Error structure model: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.regressText = "\u03B3<|Z|>\u00B2 + \u03B4"
        self.regressFunction = tk.Label(self.regressFunctionFrame, text=self.regressText, bg=self.backgroundColor, fg=self.foregroundColor)
        self.regressFunctionLabel.grid(column=0, row=0, sticky="W")
        self.regressFunction.grid(column=1, row=0, padx=2)
        self.regressFunctionFrame.grid(column=0, row=2, pady=5, sticky="W")
        
        self.regressionParamsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.alphaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorAlpha())
        self.alphaCheckbox = ttk.Checkbutton(self.regressionParamsFrame, text="\u03B1", variable=self.alphaCheckboxVariable, command=checkboxCommand)
        self.alphaEqual = tk.Label(self.regressionParamsFrame, text="=  ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.alphaEntryVariable = tk.StringVar(self, "0.1")
        self.alphaEntry = ttk.Entry(self.regressionParamsFrame, textvariable=self.alphaEntryVariable, width=10)
        
        def popup_alpha(event):
            try:
                self.popup_menuA.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuA.grab_release()
        def paste_alpha():
            if (self.focus_get() != self.alphaEntry):
                self.alphaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.alphaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_alpha():
            if (self.focus_get() != self.alphaEntry):
                pyperclip.copy(self.alphaEntry.get())
            else:
                try:
                    stringToCopy = self.alphaEntry.selection_get()
                except Exception:
                    stringToCopy = self.alphaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuA = tk.Menu(self.alphaEntry, tearoff=0)
        self.popup_menuA.add_command(label="Copy", command=copy_alpha)
        self.popup_menuA.add_command(label="Paste", command=paste_alpha)
        self.alphaEntry.bind("<Button-3>", popup_alpha)
        
        self.betaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorBeta())
        self.betaCheckbox = ttk.Checkbutton(self.regressionParamsFrame, text="\u03B2", variable=self.betaCheckboxVariable, command=checkboxCommand)
        self.betaEqual = tk.Label(self.regressionParamsFrame, text="=  ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.betaEntryVariable = tk.StringVar(self, "0.1")
        self.betaEntry = ttk.Entry(self.regressionParamsFrame, textvariable=self.betaEntryVariable, width=10)
        
        def popup_beta(event):
            try:
                self.popup_menuB.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuB.grab_release()
        def paste_beta():
            if (self.focus_get() != self.betaEntry):
                self.betaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.betaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_beta():
            if (self.focus_get() != self.betaEntry):
                pyperclip.copy(self.betaEntry.get())
            else:
                try:
                    stringToCopy = self.betaEntry.selection_get()
                except Exception:
                    stringToCopy = self.betaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuB = tk.Menu(self.betaEntry, tearoff=0)
        self.popup_menuB.add_command(label="Copy", command=copy_beta)
        self.popup_menuB.add_command(label="Paste", command=paste_beta)
        self.betaEntry.bind("<Button-3>", popup_beta)
        
        self.reCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorRe())
        self.reCheckbox = ttk.Checkbutton(self.regressionParamsFrame, text="R\u2091", variable=self.reCheckboxVariable, command=checkboxCommand)
        self.gammaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorGamma())
        self.gammaCheckbox = ttk.Checkbutton(self.regressionParamsFrame, text="\u03B3", variable=self.gammaCheckboxVariable, command=checkboxCommand)
        self.gammaEqual = tk.Label(self.regressionParamsFrame, text="= ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.gammaEntryVariable = tk.StringVar(self, "0.1")
        self.gammaEntry = ttk.Entry(self.regressionParamsFrame, textvariable=self.gammaEntryVariable, width=10)
        
        def popup_gamma(event):
            try:
                self.popup_menuG.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuG.grab_release()
        def paste_gamma():
            if (self.focus_get() != self.gammaEntry):
                self.gammaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.gammaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_gamma():
            if (self.focus_get() != self.gammaEntry):
                pyperclip.copy(self.gammaEntry.get())
            else:
                try:
                    stringToCopy = self.gammaEntry.selection_get()
                except Exception:
                    stringToCopy = self.gammaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuG = tk.Menu(self.gammaEntry, tearoff=0)
        self.popup_menuG.add_command(label="Copy", command=copy_gamma)
        self.popup_menuG.add_command(label="Paste", command=paste_gamma)
        self.gammaEntry.bind("<Button-3>", popup_gamma)
        
        self.deltaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorDelta())
        self.deltaCheckbox = ttk.Checkbutton(self.regressionParamsFrame, text="\u03B4", variable=self.deltaCheckboxVariable, command=checkboxCommand)
        self.deltaEqual = tk.Label(self.regressionParamsFrame, text="= ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.deltaEntryVariable = tk.StringVar(self, "0.1")
        self.deltaEntry = ttk.Entry(self.regressionParamsFrame, textvariable=self.deltaEntryVariable, width=10)
        
        def popup_delta(event):
            try:
                self.popup_menuD.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menuD.grab_release()
        def paste_delta():
            if (self.focus_get() != self.deltaEntry):
                self.deltaEntry.insert(tk.END, pyperclip.paste())
            else:
                self.deltaEntry.insert(tk.INSERT, pyperclip.paste())
        def copy_delta():
            if (self.focus_get() != self.deltaEntry):
                pyperclip.copy(self.deltaEntry.get())
            else:
                try:
                    stringToCopy = self.deltaEntry.selection_get()
                except Exception:
                    stringToCopy = self.deltaEntry.get()
                pyperclip.copy(stringToCopy)
        self.popup_menuD = tk.Menu(self.deltaEntry, tearoff=0)
        self.popup_menuD.add_command(label="Copy", command=copy_delta)
        self.popup_menuD.add_command(label="Paste", command=paste_delta)
        self.deltaEntry.bind("<Button-3>", popup_delta)
        
        self.optionsFrame = tk.Frame(self.regressionParamsFrame, bg=self.backgroundColor)
        self.weightingLabel = tk.Label(self.optionsFrame, text="Weighting: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.weightingComboboxVariable = tk.StringVar(self, self.topGUI.getErrorWeighting())
        self.weightingCombobox = ttk.Combobox(self.optionsFrame, textvariable=self.weightingComboboxVariable, value=("None", "Variance"), state="readonly", exportselection=0, width=10)
        self.weightingCombobox.bind("<<ComboboxSelected>>", checkWeight)
        self.movingAverageLabel = tk.Label(self.optionsFrame, text="Moving Average: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.movingAverageComboboxVariable = tk.StringVar(self, self.topGUI.getMovingAverage())
        self.movingAverageCombobox = ttk.Combobox(self.optionsFrame, textvariable=self.movingAverageComboboxVariable, value=("None", "3 point", "5 point"), state="readonly", exportselection=0, width=10)
        self.detrendLabel = tk.Label(self.optionsFrame, text="Detrend: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.detrendComboboxVariable = tk.StringVar(self, self.topGUI.getDetrend())
        self.detrendCombobox = ttk.Combobox(self.optionsFrame, textvariable=self.detrendComboboxVariable, value=("Off", "On"), state="readonly", exportselection=0, width=10)
        self.weightingLabel.grid(column=0, row=0, sticky="W")
        self.weightingCombobox.grid(column=1, row=0)
        self.movingAverageLabel.grid(column=0, row=1, pady=2, sticky="W")
        self.movingAverageCombobox.grid(column=1, row=1)
        self.detrendLabel.grid(column=0, row=2, pady=(0,2), sticky="W")
        self.detrendCombobox.grid(column=1, row=2)
        self.alphaCheckbox.grid(column=0, row=0, sticky="W")
        self.betaCheckbox.grid(column=0, row=1, pady=2, sticky="W")
        self.gammaCheckbox.grid(column=0, row=3, pady=2, sticky="W")
        self.gammaEqual.grid(column=1, row=3, sticky="W")
        self.gammaEntry.grid(column=2, row=3, sticky="W")
        self.deltaCheckbox.grid(column=0, row=4, sticky="W")
        self.deltaEqual.grid(column=1, row=4, sticky="W")
        self.deltaEntry.grid(column=2, row=4, sticky="W")
        self.optionsFrame.grid(column=0, row=5, columnspan=4, pady=10, sticky="W")
        self.regressionParamsFrame.grid(column=0, row=3, pady=10, sticky="W")
        checkboxCommand()
        weighting_ttp = CreateToolTip(self.weightingCombobox, 'Which weighting is used in the regression')
        movingAverage_ttp = CreateToolTip(self.movingAverageCombobox, 'Which moving average is used when calculating the variance for the weighting')
        detrend_ttp = CreateToolTip(self.detrendCombobox, 'Which detrending is performed on the data prior to regression')
        
        self.regressButton = ttk.Button(self, text="Regress errors", width=25, state="disabled", command=regressErrors)
        self.regressButton.grid(column=0, row=4, pady=10, sticky="W")
        regress_ttp = CreateToolTip(self.regressButton, 'Regress the current error structure to the loaded data using a Levenberg–Marquardt algorithm')
        
        self.resultsFrame = tk.Frame(self, bg=self.backgroundColor)
        def fixed_map(option):
            # Fix for setting text colour for Tkinter 8.6.9
            # From: https://core.tcl.tk/tk/info/509cafafae
            #
            # Returns the style map for 'option' with any styles starting with
            # ('!disabled', '!selected', ...) filtered out.
        
            # style.map() returns an empty list for missing options, so this
            # should be future-safe.
            return [elm for elm in style.map('Treeview', query_opt=option) if
              elm[:2] != ('!disabled', '!selected')]
        
        style = ttk.Style()
        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))
        self.resultsView = ttk.Treeview(self.resultsFrame, columns=("param", "value", "sigma", "percent"), height=5, selectmode="browse")
        self.resultsChi = tk.Label(self.resultsFrame, bg=self.backgroundColor, fg=self.foregroundColor)
        self.resultsView.heading("param", text="Parameter")
        self.resultsView.heading("value", text="Value")
        self.resultsView.heading("sigma", text="Std. Dev.")
        self.resultsView.heading("percent", text="95% CI")
        self.resultsView.column("#0", width=0)
        self.resultsView.column("param", width=120, anchor=tk.CENTER)
        self.resultsView.column("value", width=120, anchor=tk.CENTER)
        self.resultsView.column("sigma", width=120, anchor=tk.CENTER)
        self.resultsView.column("percent", width=120, anchor=tk.CENTER)
        self.copyResultsButton = ttk.Button(self.resultsFrame, text="Copy values and std. devs.", command=copyResults, width=25)
        self.resultsView.grid(column=0, row=0, sticky="W")
        self.resultsView.bind("<Button-1>", handle_click)
        self.resultsView.bind("<Motion>", handle_motion)
        self.resultsChi.grid(column=0, row=1, pady=3, sticky="W")
        self.copyResultsButton.grid(column=0, row=1, pady=3, sticky="E")
        copyResultsButton_ttp = CreateToolTip(self.copyResultsButton, 'Copy the result values and standard deviations as a spreadsheet')
        
        self.plotButton = ttk.Button(self, text="Plot", width=20, command=plotResiduals)
        plot_ttp = CreateToolTip(self.plotButton, 'Plot the data and the fitted error structure')
        
         #---Close all popups---
        self.closeAllPopupsButton = ttk.Button(self, text="Close all popups", command=self.topGUI.closeAllPopups)
        self.closeAllPopupsButton.grid(column=0, row=7, sticky="W", pady=10)
        closeAllPopups_ttp = CreateToolTip(self.closeAllPopupsButton, 'Close all open popup windows')
        
    def setThemeLight(self):
        self.foregroundColor = "#000000"
        self.backgroundColor = "#FFFFFF"
        self.configure(background="#FFFFFF")
        self.fileLabel.configure(background="#FFFFFF", foreground="#000000")
        self.fileFrame.configure(background="#FFFFFF")
        self.regressFunctionFrame.configure(background="#FFFFFF")
        self.regressFunction.configure(background="#FFFFFF", foreground="#000000")
        self.regressFunctionLabel.configure(background="#FFFFFF", foreground="#000000")
        self.regressionParamsFrame.configure(background="#FFFFFF")
        self.alphaEqual.configure(background="#FFFFFF", foreground="#000000")
        self.betaEqual.configure(background="#FFFFFF", foreground="#000000")
        self.gammaEqual.configure(background="#FFFFFF", foreground="#000000")
        self.deltaEqual.configure(background="#FFFFFF", foreground="#000000")
        self.optionsFrame.configure(background="#FFFFFF")
        self.weightingLabel.configure(background="#FFFFFF", foreground="#000000")
        self.movingAverageLabel.configure(background="#FFFFFF", foreground="#000000")
        self.resultsFrame.configure(background="#FFFFFF")
        self.resultsChi.configure(background="#FFFFFF", foreground="#000000")
        self.detrendLabel.configure(background="#FFFFFF", foreground="#000000")
                                 
    def setThemeDark(self):
        self.foregroundColor = "#FFFFFF"
        self.backgroundColor = "#424242"
        self.configure(background="#424242")
        self.fileLabel.configure(background="#424242", foreground="#FFFFFF")
        self.fileFrame.configure(background="#424242")
        self.regressFunctionFrame.configure(background="#424242")
        self.regressFunction.configure(background="#424242", foreground="#FFFFFF")
        self.regressFunctionLabel.configure(background="#424242", foreground="#FFFFFF")
        self.regressionParamsFrame.configure(background="#424242")
        self.alphaEqual.configure(background="#424242", foreground="#FFFFFF")
        self.betaEqual.configure(background="#424242", foreground="#FFFFFF")
        self.gammaEqual.configure(background="#424242", foreground="#FFFFFF")
        self.deltaEqual.configure(background="#424242", foreground="#FFFFFF")
        self.optionsFrame.configure(background="#424242")
        self.weightingLabel.configure(background="#424242", foreground="#FFFFFF")
        self.movingAverageLabel.configure(background="#424242", foreground="#FFFFFF")
        self.resultsFrame.configure(background="#424242")
        self.resultsChi.configure(background="#424242", foreground="#FFFFFF")
        self.detrendLabel.configure(background="#424242", foreground="#FFFFFF")
    
    def closeWindows(self):
        for resPlot in self.resPlots:
            try:
                resPlot.destroy()
            except Exception:
                pass
    
    def errorEnter(self, n):
        try:
            with open(n,'r') as UseFile:
                filetext = UseFile.read()
                lines = filetext.splitlines()
            if ("frequency" in lines[2].lower()):
                data = np.loadtxt(n, skiprows=3)
            else:
                data = np.loadtxt(n)
            re_in = data[1,0]
            w_in = data[2:,0]
            rstddev_in = data[2:,1]
            jstddev_in = data[2:,2]
            r_in = data[2:,3]
            j_in = data[2:,4]
            z_in = data[2:,5]
            sigma_in = data[2:,5]
            if (not (len(w_in) == len(rstddev_in) and len(rstddev_in) == len(jstddev_in) and len(jstddev_in) == len(r_in) and len(r_in) == len(j_in) and len(j_in) == len(z_in) and len(z_in) == len(sigma_in))):
                raise FileLengthError
            for i in range(len(w_in)):
                float(re_in)
                float(w_in[i])
                float(rstddev_in[i])
                float(jstddev_in[i])
                float(r_in[i])
                float(j_in[i])
                float(z_in[i])
                float(sigma_in[i])
            self.fileListbox.insert(tk.END, n)
            if (self.noneLoaded):
                self.wdataLength = len(w_in)
                self.noneLoaded = False
            self.wdata.append(w_in)
            self.avgRe.append(re_in)
            self.rres.append(rstddev_in)
            self.jres.append(jstddev_in)
            self.zr.append(r_in)
            self.zj.append(j_in)
            self.modz.append(z_in)
            self.sigma.append(sigma_in)
            self.numFiles += 1
            self.fileNames.append(os.path.basename(n))
            self.regressButton.configure(state="normal")
        except FileExtensionError:
            messagebox.showerror("File error", "Error 30:\nThe file has an unknown extension")
        except FileLengthError:
            messagebox.showerror("Length error", "Error 31:\nThe number of data do not match")
        except Exception:
            messagebox.showerror("File error", "Error 33:\nThere was an error reading the file")
    
    def errorsEnter(self, n):
        try:
            alreadyWarned = False
            for name in n:
                fname, fext = os.path.splitext(name)
                if (fext != ".mmerrors"):
                    raise FileExtensionError
                with open(name,'r') as UseFile:
                    filetext = UseFile.read()
                    lines = filetext.splitlines()
                if ("frequency" in lines[2].lower()):
                    data = np.loadtxt(name, skiprows=3)
                else:
                    data = np.loadtxt(name)
                re_in = data[1,0]
                w_in = data[2:,0]
                rstddev_in = data[2:,1]
                jstddev_in = data[2:,2]
                r_in = data[2:,3]
                j_in = data[2:,4]
                z_in = data[2:,5]
                sigma_in = data[2:,5]
                if (not (len(w_in) == len(rstddev_in) and len(rstddev_in) == len(jstddev_in) and len(jstddev_in) == len(r_in) and len(r_in) == len(j_in) and len(j_in) == len(z_in) and len(z_in) == len(sigma_in))):
                    raise FileLengthError
                for i in range(len(w_in)):
                    float(re_in)
                    float(w_in[i])
                    float(rstddev_in[i])
                    float(jstddev_in[i])
                    float(r_in[i])
                    float(j_in[i])
                    float(z_in[i])
                    float(sigma_in[i])
                if (name in self.fileListbox.get(0, tk.END) and not alreadyWarned):
                    alreadyWarned = True
                    if (len(n) > 1):
                        messagebox.showwarning("File already loaded", "Warning: One or more of the selected files is already loaded and is being loaded again")
                    else:
                        messagebox.showwarning("File already loaded", "Warning: The selected file is already loaded and is being loaded again")
                self.fileListbox.insert(tk.END, name)
                if (self.noneLoaded):
                    self.noneLoaded = False
                self.wdata.append(w_in)
                self.avgRe.append(re_in)
                self.rres.append(rstddev_in)
                self.jres.append(jstddev_in)
                self.zr.append(r_in)
                self.zj.append(j_in)
                self.modz.append(z_in)
                self.sigma.append(sigma_in)
                self.numFiles += 1
                self.fileNames.append(os.path.basename(name))
                self.regressButton.configure(state="normal")
        except FileExtensionError:
            messagebox.showerror("File error", "Error 30:\nThe file has an unknown extension")
        except FileLengthError:
            messagebox.showerror("Length error", "Error 31:\nThe number of data do not match")
        except Exception:
            messagebox.showerror("File error", "Error 33:\nThere was an error reading the file")