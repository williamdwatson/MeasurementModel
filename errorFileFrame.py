# -*- coding: utf-8 -*-
"""
Created on Fri Oct  5 11:18:25 2018

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
import os, sys
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

class FileExtensionError(Exception):
    """Exception for files with an incorrect extension"""
    pass

class FileLengthError(Exception):
    """Exception for files with different lengths"""
    pass

class FrequencyMismatchError(Exception):
    """Exception for files with different frequencies"""
    pass

class CreateToolTip(object):
    """
    Create a tooltip for a given widget
    Code from: https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter with answer by crxguy52
    """
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #Miliseconds
        self.wraplength = 180   #Pixels
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

class eFF(tk.Frame):
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
        
        self.noneLoaded = True  #Keep track of if any files have been loaded
        self.resPlots = []      #List of open plots
        self.wdata = []         #Frequency data
        self.re = []            #Ohmic resistances
        self.cap = []           #Capacitances
        self.rres = []          #Real residuals
        self.jres = []          #Imaginary residuals
        self.zr = []            #Real data
        self.zj = []            #Imaginary data
        self.wdataLength = 0    #Length of the data currently present
        self.numFiles = 0       #Number of files present
        self.standardDevsI = [] #Imaginary standard deviations
        self.standardDevsR = [] #Real standard deviations
        
        def addFile():
            """
            Add files to be combined into a .mmerrors file
            """
            #---Get file(s) to be added---
            n = askopenfilenames(initialdir=self.topGUI.getCurrentDirectory(), filetypes =[("Residuals File (*.mmresiduals)", "*.mmresiduals")], title = "Choose a file")
            if (len(n) != 0):   #If len(n) == 0, then no file chosen
                try:
                    alreadyWarned = False   #Used to check if warning that a file has already been loaded
                    #---Set the current directory to the directory where the new file is located---
                    directory = os.path.dirname(str(n[0]))
                    self.topGUI.setCurrentDirectory(directory)
                    #---Loop through each file and check if they have the right extension, number of data, and frequencies---
                    for name in n:
                        fname, fext = os.path.splitext(name)
                        #---Check if the file has the right extension---
                        if (fext != ".mmresiduals"):
                            raise FileExtensionError
                        #---Load the file in using numpy---
                        with open(name,'r') as UseFile:
                            filetext = UseFile.read()
                            lines = filetext.splitlines()
                        if ("frequency" in lines[1].lower()):
                            data = np.loadtxt(name, skiprows=2)
                        else:
                            data = np.loadtxt(name)
                        #---Data is located in specific locations---
                        re_in = data[0,0]
                        cap_in = data[0,1]
                        w_in = data[1:,0]
                        rres_in = data[1:,1]
                        jres_in = data[1:,2]
                        r_in = data[1:,3]
                        j_in = data[1:,4]
                        #---Check if the lengths of all the data match---
                        if (not (len(w_in) == len(rres_in) and len(rres_in) == len(jres_in) and len(jres_in) == len(r_in) and len(r_in) == len(j_in))):
                            raise FileLengthError
                        #---Verify that all the data are numeric---
                        for i in range(len(w_in)):
                            float(re_in)
                            float(cap_in)
                            float(w_in[i])
                            float(rres_in[i])
                            float(jres_in[i])
                            float(r_in[i])
                            float(j_in[i])
                        #---Check that the file length and all frequencies match those already present---
                        if (not self.noneLoaded):
                            if (len(w_in) != self.wdataLength):
                                raise FileLengthError
                            else:
                                for i in range(len(w_in)):
                                    if (self.wdata[0][i] != w_in[i]):
                                        raise FrequencyMismatchError
                        #---Warn if a file has already been loaded---
                        if (name in self.fileListbox.get(0, tk.END) and not alreadyWarned):
                            alreadyWarned = True
                            if (len(n) > 1):
                                messagebox.showwarning("File already loaded", "Warning: One or more of the selected files is already loaded and is being loaded again")
                            else:
                                messagebox.showwarning("File already loaded", "Warning: The selected file is already loaded and is being loaded again")
                        #---Add the file to the multiselect and its data to the lists---
                        self.fileListbox.insert(tk.END, name)
                        if (self.noneLoaded):
                            self.wdataLength = len(w_in)
                            self.noneLoaded = False
                        self.re.append(re_in)
                        self.cap.append(cap_in)
                        self.wdata.append(w_in)
                        self.rres.append(rres_in)
                        self.jres.append(jres_in)
                        self.zr.append(r_in)
                        self.zj.append(j_in)
                        self.numFiles += 1
                    #---Calculate the new average values---
                    if (self.numFiles > 1):
                        #---Find the average Ohmic resistance---
                        re_avg = 0
                        for i in range(len(self.zr)):
                            re_avg += self.re[i]
                        re_avg /= len(self.zr)
                        
                        #---Find the average capacitance---
                        cap_avg = 0
                        for i in range(len(self.zr)):
                            cap_avg += self.cap[i]
                        cap_avg /= len(self.zr)
                        
                        #---Find the new mean of the real residuals---
                        meansReal = np.zeros(self.wdataLength)
                        for i in range(self.wdataLength):
                            for j in range(len(self.rres)):
                                meansReal[i] += self.rres[j][i]
                            meansReal[i] /= len(self.rres)
                        
                        #---Find the new mean of the imaginary residuals---
                        meansImag = np.zeros(self.wdataLength)
                        for i in range(self.wdataLength):
                            for j in range(len(self.jres)):
                                meansImag[i] += self.jres[j][i]
                            meansImag[i] /= len(self.jres)
                        
                        #---Find the new standard deviation of the real residuals---
                        standardDevsReal = np.zeros(self.wdataLength)
                        for i in range(self.wdataLength):
                            for j in range(len(self.rres)):
                                standardDevsReal[i] += (self.rres[j][i] - meansReal[i])**2
                            standardDevsReal[i] /= len(self.rres)-1
                            standardDevsReal[i] = np.sqrt(standardDevsReal[i])
                        
                        #---Find the new standard deviation of the imaginary residuals---
                        standardDevsImag = np.zeros(self.wdataLength)
                        for i in range(self.wdataLength):
                            for j in range(len(self.jres)):
                                standardDevsImag[i] += (self.jres[j][i] - meansImag[i])**2
                            standardDevsImag[i] /= len(self.jres)-1
                            standardDevsImag[i] = np.sqrt(standardDevsImag[i])
                        
                        #---Add the new data to the GUI---
                        self.statsView.delete(*self.statsView.get_children())
                        for i in range(self.wdataLength):
                            self.statsView.insert("", tk.END, text="", values=(str(self.wdata[0][i]), "%.6g"%meansReal[i], "%.6g"%standardDevsReal[i], "%.6g"%meansImag[i], "%.6g"%standardDevsImag[i]))
                        self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%re_avg)
                        if (cap_avg != 0):
                            self.capAvgLabel.configure(text="|  Average Capacitance = %.5g"%cap_avg)
                        self.standardDevsR = standardDevsReal
                        self.standardDevsI = standardDevsImag
                    elif (self.numFiles == 1):
                        self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%self.re[0])
                        if (self.cap[0] != 0):
                            self.capAvgLabel.configure(text="|  Average Capacitance = %.5g"%self.cap[0])
                
                #---Catch errors and provide the appropriate feedback---
                except FileExtensionError:
                    messagebox.showerror("File error", "Error 26:\nThe file has an unknown extension")
                except FileLengthError:
                    messagebox.showerror("Length error", "Error 27:\nThe number of data do not match")
                except FrequencyMismatchError:
                    messagebox.showerror("Frequency error", "Error 28:\nSome of the frequencies being loaded do not match the already loaded frequencies")
                except Exception:
                    messagebox.showerror("File error", "Error 29:\nThere was an error reading the file")
                if (self.numFiles > 1):
                    self.saveButton.configure(state="normal")
                    self.plotButton.configure(state="normal")
        
        def removeFile():
            """
            Remove a file and its associated data
            """
            #---If there are now no files, reset everything to its original state, clear variables, and disable the appropriate buttons---
            if (self.fileListbox.size() == 0):
                self.noneLoaded = True
                self.wdata = []
                self.rres = []
                self.jres = []
                self.zr = []
                self.zj = []
                self.re = []
                self.cap = []
                self.numFiles = 0
                self.saveButton.configure(state="disabled")
                self.plotButton.configure(state="disabled")
                self.reAvgLabel.configure(text="")
                self.capAvgLabel.configure(text="")
            else:
                items = list(map(int, self.fileListbox.curselection()))     #Find which items are being deleted
                if (len(items) != 0):
                    #---If it's not the last item, do the appropriate deletions---
                    if (items[0] < len(self.wdata)):
                        #---Remove data from the lists--
                        del self.wdata[items[0]]
                        del self.rres[items[0]]
                        del self.jres[items[0]]
                        del self.zr[items[0]]
                        del self.zj[items[0]]
                        del self.re[items[0]]
                        del self.cap[items[0]]
                        self.fileListbox.delete(tk.ANCHOR)
                        self.numFiles -= 1
                        #---If there's only one file left, set the GUI as appropriate---
                        if (self.numFiles == 1):
                            self.statsView.delete(*self.statsView.get_children())
                            self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%self.re[0])
                            if (self.cap[0] != 0):
                                self.capAvgLabel.configure(text="|  Average Capacitance = %.5g"%self.cap[0])
                            self.saveButton.configure(state="disabled")
                            self.plotButton.configure(state="disabled")
                        #---If there are no files left, set reset everything---
                        elif (self.numFiles == 0):
                            self.statsView.delete(*self.statsView.get_children())
                            self.reAvgLabel.configure(text="")
                            self.capAvgLabel.configure(text="")
                            self.saveButton.configure(state="disabled")
                            self.plotButton.configure(state="disabled")
                        #---If there are 2 or more files left, recalculate values as in addFile()---
                        else:
                            re_avg = 0
                            for i in range(len(self.zr)):
                                re_avg += self.re[i]
                            re_avg /= len(self.zr)
                            
                            cap_avg = 0
                            for i in range(len(self.zr)):
                                cap_avg += self.cap[i]
                            cap_avg /= len(self.zr)
                            
                            meansReal = np.zeros(self.wdataLength)
                            for i in range(self.wdataLength):
                                for j in range(len(self.rres)):
                                    meansReal[i] += self.rres[j][i]
                                meansReal[i] /= len(self.rres)
                
                            meansImag = np.zeros(self.wdataLength)
                            for i in range(self.wdataLength):
                                for j in range(len(self.jres)):
                                    meansImag[i] += self.jres[j][i]
                                meansImag[i] /= len(self.jres)
                            
                            standardDevsReal = np.zeros(self.wdataLength)
                            for i in range(self.wdataLength):
                                for j in range(len(self.rres)):
                                    standardDevsReal[i] += (self.rres[j][i] - meansReal[i])**2
                                standardDevsReal[i] /= len(self.rres)-1
                                standardDevsReal[i] = np.sqrt(standardDevsReal[i])
                            
                            standardDevsImag = np.zeros(self.wdataLength)
                            for i in range(self.wdataLength):
                                for j in range(len(self.jres)):
                                    standardDevsImag[i] += (self.jres[j][i] - meansImag[i])**2
                                standardDevsImag[i] /= len(self.jres)-1
                                standardDevsImag[i] = np.sqrt(standardDevsImag[i])
                            
                            self.statsView.delete(*self.statsView.get_children())
                            for i in range(self.wdataLength):
                                self.statsView.insert("", tk.END, text="", values=(str(self.wdata[0][i]), "%.6g"%meansReal[i], "%.6g"%standardDevsReal[i], "%.6g"%meansImag[i], "%.6g"%standardDevsImag[i]))
                            self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%re_avg)
                            if (cap_avg != 0):
                                self.capAvgLabel.configure(text="|  Average Capacitance = %.5g"%cap_avg)
                            
                            self.standardDevsR = standardDevsReal
                            self.standardDevsI = standardDevsImag
                #---If there are now no files, reset everything to its original state, clear variables, and disable the appropriate buttons---
                if (self.fileListbox.size() == 0):
                    self.statsView.delete(*self.statsView.get_children())
                    self.reAvgLabel.configure(text="")
                    self.capAvgLabel.configure(text="")
                    self.noneLoaded = True
                    self.wdata = []
                    self.rres = []
                    self.jres = []
                    self.zr = []
                    self.zj = []
                    self.re = []
                    self.cap = []
                    self.numFiles = 0
        
        def popup_inputFiles(event):
            """
            Right-click popup menu for deleting entries in file list
            """
            try:
                self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu.grab_release()
        
        def saveErrors(e=None):
            """
            Save the combined residuals files as a .mmerrors file
            """
            try:
                if (len(self.zr) == 0):     #If there's no data, don't save
                    return
                #---Calculate the average real, imaginary, and modulus impedances---
                avgRe = 0
                avgR = np.zeros(self.wdataLength)
                avgJ = np.zeros(self.wdataLength)
                avgZ = np.zeros(self.wdataLength)
                for i in range(len(self.zr)):
                    avgRe += self.re[i]
                    for k in range(self.wdataLength):
                        avgR[k] += abs(self.zr[i][k])
                        avgJ[k] += abs(self.zj[i][k])
                        avgZ[k] += np.sqrt(self.zr[i][k]**2 + self.zj[i][k]**2)
                avgR /= len(self.zr)
                avgJ /= len(self.zj)
                avgZ /= len(self.zr)
                avgRe /= len(self.zr)
                
                #---Calculate the standard deviation between the real and imaginary standard deviations (i.e. the std dev of the std devs)---
                sigmasigma = np.zeros(self.wdataLength)
                for i in range(self.wdataLength):
                    avgSigma = (self.standardDevsI[i] + self.standardDevsR[i])/2
                    sigmasigma[i] = np.sqrt((self.standardDevsI[i]-avgSigma)**2 + (self.standardDevsR[i]-avgSigma)**2)
                #---Prepare the string that will be saved to the file---
                stringToSave = "Number of files\t-\t-\t-\t-\t-\t-\n"
                stringToSave += "Avg. Re\t-\t-\t-\t-\t-\t-\n"
                stringToSave += "Frequency\tReal std devs\tImag std devs\tAvg Real\tAvg Imag\tAvg |Z|\tStd devs of std devs\n"
                stringToSave += str(self.numFiles) + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\n"
                stringToSave += str(avgRe) + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\n" 
                for i in range(self.wdataLength):
                    stringToSave += str(self.wdata[0][i]) + "\t" + str(self.standardDevsR[i]) + "\t" + str(self.standardDevsI[i]) + "\t" + str(avgR[i]) + "\t" + str(avgJ[i]) + "\t" + str(avgZ[i]) + "\t" + str(sigmasigma[i]) + "\n"
                #---Get the desired save name from the user---
                defaultSaveName = os.path.splitext(os.path.basename(self.fileListbox.get(0)))[0]
                saveName = asksaveasfile(mode='w', defaultextension=".mmerrors", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory(), filetypes=[("Measurement model errors", ".mmerrors")])
                #---Set the current directory to the directory where the .mmerrors file was saved---
                directory = os.path.dirname(str(saveName))
                self.topGUI.setCurrentDirectory(directory)
                if saveName is None:     #If save is cancelled
                    return
                #---Write to the saved file and then close---
                saveName.write(stringToSave)
                saveName.close()
                #---Update the GUI to indicate a successful save---
                self.saveButton.configure(text="Saved")
                self.after(1000, lambda : self.saveButton.configure(text="Save errors"))
            except Exception:
                messagebox.showerror("Save error", "There was a problem saving the file")
        
        def plotResiduals():
            """
            Display a plot of the residuals
            """
            #---Create the popup plot window and add it to the list of plots (so it can be closed later if needed)---
            resPlot = tk.Toplevel(background=self.backgroundColor)
            self.resPlots.append(resPlot)
            resPlot.title("Standard Deviations Plot")
            resPlot.iconbitmap(resource_path("img/elephant3.ico"))
            
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):  #Set the plot appearance to match the theme
                figError = Figure(figsize=(5,4), dpi=100)
                a = figError.add_subplot(111)
                figError.set_facecolor(self.backgroundColor)
                a.set_facecolor(self.backgroundColor)
                a.yaxis.set_ticks_position("both")
                a.yaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")     #Make the ticks point inwards
                a.xaxis.set_ticks_position("both")
                a.xaxis.set_tick_params(direction="in", color=self.foregroundColor, which="both")     #Make the ticks point inwards
                #---Prepare the plot appearance based on the theme---
                imagColor = "tab:orange"
                realColor = "tab:blue"
                if (self.topGUI.getTheme() == "dark"):
                    imagColor = "gold"
                    realColor = "cyan"
                else:
                    imagColor = "tab:orange"
                    realColor = "tab:blue"
                #---Plot the real and imaginary standard deviations of the residuals---
                a.plot(self.wdata[0], self.standardDevsR, "o", markerfacecolor="None", color=realColor, label="Real")
                a.plot(self.wdata[0], self.standardDevsI, "^", markerfacecolor="None", color=imagColor, label="Imaginary")
                a.set_xscale("log")
                a.set_yscale("log")
                a.set_title("Standard Deviations", color=self.foregroundColor)
                a.set_xlabel("Frequency / Hz", color=self.foregroundColor)
                a.set_ylabel("σ\u1D63, σ\u2C7C / Ω", color=self.foregroundColor)
                a.legend()
                figError.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                errorCanvasInput = FigureCanvasTkAgg(figError, resPlot)
                errorCanvasInput.draw()
                errorCanvasInput.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                toolbar = NavigationToolbar2Tk(errorCanvasInput, resPlot)    #Enables the zoom and move toolbar for the plot
                toolbar.update()
            errorCanvasInput._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            def on_closing():
                """
                Clear the figure before closing the popup
                """
                figError.clear()
                resPlot.destroy()
            
            resPlot.protocol("WM_DELETE_WINDOW", on_closing)

        def handle_click(event):
            """
            Prevent the first (empty) column from being resized
            """
            if self.statsView.identify_region(event.x, event.y) == "separator":
                if (self.statsView.identify_column(event.x) == "#0"):
                    return "break"
        
        def handle_motion(event):
            """
            Prevent the firsst (empty) column from being dragged/resized
            """
            if self.statsView.identify_region(event.x, event.y) == "separator":
                if (self.statsView.identify_column(event.x) == "#0"):
                    return "break"
        
        #---The listbox holding the file names, and the buttons to add and remove files---
        self.fileLabel = tk.Label(self, text="Files to use:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.fileLabel.grid(column=0, row=0, sticky="W")
        self.fileFrame = tk.Frame(self, bg=self.backgroundColor)
        self.fileListboxScrollbar = ttk.Scrollbar(self.fileFrame, orient=tk.VERTICAL)
        self.fileListboxScrollbarHorizontal = ttk.Scrollbar(self.fileFrame, orient=tk.HORIZONTAL)
        self.fileListbox = tk.Listbox(self.fileFrame, height=3, width=65, selectmode=tk.BROWSE, activestyle='none', yscrollcommand=self.fileListboxScrollbar.set, xscrollcommand=self.fileListboxScrollbarHorizontal.set)
        self.fileListboxScrollbar['command'] = self.fileListbox.yview
        self.fileListboxScrollbarHorizontal['command'] = self.fileListbox.xview
        self.fileBrowse = ttk.Button(self.fileFrame, text="Add file(s)", command=addFile)
        self.fileRemove = ttk.Button(self.fileFrame, text="Remove", command=removeFile)
        self.fileListbox.grid(column=0, row=0, rowspan=2)
        self.fileListboxScrollbar.grid(column=1, row=0, rowspan=2, sticky="NS")
        self.fileListboxScrollbarHorizontal.grid(column=0, row=2, sticky="EW")
        self.fileBrowse.grid(column=2, row=0, padx=2)
        self.fileRemove.grid(column=2, row=1, padx=2)
        self.fileFrame.grid(column=0, row=1, sticky="W")
        addFiles_ttp = CreateToolTip(self.fileBrowse, 'Browse for one or more .mmresiduals files')
        removeFile_ttp = CreateToolTip(self.fileRemove, 'Remove selected file from the list')
        
        #---The right-click option for removing a file---
        self.popup_menu = tk.Menu(self.fileListbox, tearoff=0)
        self.popup_menu.add_command(label="Remove file", command=removeFile)
        self.fileListbox.bind("<Button-3>", popup_inputFiles)
        
        #---The treeview holding the real and imaginary residuals and standard deviations of the residuals at each frequency---
        self.statsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.statsViewScrollbar = ttk.Scrollbar(self.statsFrame, orient=tk.VERTICAL)
        self.statsView = ttk.Treeview(self.statsFrame, columns=("freq", "meanR", "stddevR", "meanI", "stddevI"), height=5, selectmode="browse", yscrollcommand=self.statsViewScrollbar.set)
        self.statsView.heading("freq", text="Freq.")
        self.statsView.heading("meanR", text="Real Mean")
        self.statsView.heading("stddevR", text="Real Std. Dev.")
        self.statsView.heading("meanI", text="Imag. Mean")
        self.statsView.heading("stddevI", text="Imag. Std. Dev.")
        self.statsView.column("#0", width=10)
        self.statsView.column("freq", width=60)
        self.statsView.column("meanR", width=100)
        self.statsView.column("stddevR", width=100)
        self.statsView.column("meanI", width=100)
        self.statsView.column("stddevI", width=120)
        self.statsViewScrollbar['command'] = self.statsView.yview
        
        #---The average Ohmic resistance and capacitance---
        self.averageLabelFrame= tk.Frame(self.statsFrame, bg=self.backgroundColor)
        self.reAvgLabel = tk.Label(self.averageLabelFrame, bg=self.backgroundColor, fg=self.foregroundColor)
        self.capAvgLabel = tk.Label(self.averageLabelFrame, bg=self.backgroundColor, fg=self.foregroundColor)
        self.statsView.grid(column=0, row=0, sticky="W")
        self.statsViewScrollbar.grid(column=1, row=0, sticky="NS")
        self.averageLabelFrame.grid(column=0, row=2, sticky="EW")
        self.reAvgLabel.grid(column=0, row=0, sticky="W")
        self.capAvgLabel.grid(column=1, row=0, sticky="W", padx=10)
        self.statsFrame.grid(column=0, row=4, pady=10, sticky="W")
        self.statsView.bind("<Button-1>", handle_click)
        self.statsView.bind("<Motion>", handle_motion)
        
        #---The plot button (enables once at least two files are added)---
        self.plotButton = ttk.Button(self, text="Plot", width=20, state="disabled", command=plotResiduals)
        self.plotButton.grid(column=0, row=5, pady=5, sticky="W")
        plot_ttp = CreateToolTip(self.plotButton, 'Opens a new window with a plot of the residuals')
        
        #---The save button (enables once at least two files are added)---
        self.saveButton = ttk.Button(self, text="Save errors", width=25, state="disabled", command=saveErrors)
        self.saveButton.grid(column=0, row=6, pady=10, sticky="W")
        save_ttp = CreateToolTip(self.saveButton, 'Saves the combined residuals as a .mmerrors file for use in the Error Analysis tab')
        
         #---Close all popups---
        self.closeAllPopupsButton = ttk.Button(self, text="Close all popups", command=self.topGUI.closeAllPopups)
        self.closeAllPopupsButton.grid(column=0, row=7, sticky="W", pady=10)
        closeAllPopups_ttp = CreateToolTip(self.closeAllPopupsButton, 'Close all open popup windows')
        
    def setThemeLight(self):
        """
        Switch to light theme (white background, black text)
        """
        self.foregroundColor = "#000000"
        self.backgroundColor = "#FFFFFF"
        self.configure(background="#FFFFFF")
        self.fileLabel.configure(background="#FFFFFF", foreground="#000000")
        self.fileFrame.configure(background="#FFFFFF")
        self.statsFrame.configure(background="#FFFFFF")
        self.reAvgLabel.configure(background="#FFFFFF", foreground="#000000")
        self.capAvgLabel.configure(background="#FFFFFF", foreground="#000000")
        self.averageLabelFrame.configure(background="#FFFFFF")
                                 
    def setThemeDark(self):
        """
        Switch to dark theme (dark gray background, white text)
        """
        self.foregroundColor = "#FFFFFF"
        self.backgroundColor = "#424242"
        self.configure(background="#424242")
        self.fileLabel.configure(background="#424242", foreground="#FFFFFF")
        self.fileFrame.configure(background="#424242")
        self.statsFrame.configure(background="#424242")
        self.reAvgLabel.configure(background="#424242", foreground="#FFFFFF")
        self.capAvgLabel.configure(background="#424242", foreground="#FFFFFF")
        self.averageLabelFrame.configure(background="#424242")
    
    def saveErrors(self, e=None):
        if (len(self.zr) == 0):
            return
        avgRe = 0
        avgR = np.zeros(self.wdataLength)
        avgJ = np.zeros(self.wdataLength)
        avgZ = np.zeros(self.wdataLength)
        for i in range(len(self.zr)):
            avgRe += self.re[i]
            for k in range(self.wdataLength):
                avgR[k] += abs(self.zr[i][k])
                avgJ[k] += abs(self.zj[i][k])
                avgZ[k] += np.sqrt(self.zr[i][k]**2 + self.zj[i][k]**2)
        avgR /= len(self.zr)
        avgJ /= len(self.zj)
        avgZ /= len(self.zr)
        avgRe /= len(self.zr)
        
        sigmasigma = np.zeros(self.wdataLength)
        for i in range(self.wdataLength):
            avgSigma = (self.standardDevsI[i] + self.standardDevsR[i])/2
            sigmasigma[i] = np.sqrt((self.standardDevsI[i]-avgSigma)**2 + (self.standardDevsR[i]-avgSigma)**2)
        stringToSave = "Number of files\t-\t-\t-\t-\t-\t-\n"
        stringToSave += "Avg. Re\t-\t-\t-\t-\t-\t-\n"
        stringToSave += "Frequency\tReal std devs\tImag std devs\tAvg Real\tAvg Imag\tAvg |Z|\tStd devs of std devs\n"
        stringToSave = str(self.numFiles) + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\n"
        stringToSave += str(avgRe) + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\n" 
        for i in range(self.wdataLength):
            stringToSave += str(self.wdata[0][i]) + "\t" + str(self.standardDevsR[i]) + "\t" + str(self.standardDevsI[i]) + "\t" + str(avgR[i]) + "\t" + str(avgJ[i]) + "\t" + str(avgZ[i]) + "\t" + str(sigmasigma[i]) + "\n"
        defaultSaveName = os.path.splitext(os.path.basename(self.fileListbox.get(0)))[0]
        saveName = asksaveasfile(mode='w', defaultextension=".mmerrors", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory(), filetypes=[("Measurement model errors", ".mmerrors")])
        directory = os.path.dirname(str(saveName))
        self.topGUI.setCurrentDirectory(directory)
        if saveName is None:     #If save is cancelled
            return
        saveName.write(stringToSave)
        saveName.close()
        self.saveButton.configure(text="Saved")
        self.after(1000, lambda : self.saveButton.configure(text="Save errors"))
    
    def bindIt(self, e=None):
        """
        Use Ctrl+s to save error file
        """
        self.bind_all("<Control-s>", self.saveErrors)
    
    def unbindIt(self, e=None):
        """
        Disable using Ctrl+s to save error file
        """
        self.unbind_all("<Control-s>")
    
    def closeWindows(self):
        """
        Close all residual plots present in the resPlots list
        """
        for resPlot in self.resPlots:
            try:
                resPlot.destroy()
            except Exception:
                pass
                         
    def residualEnter(self, n):
        """
        Load a file that was opened with the program (i.e. a file supplied as a command-line argument)
        """
        try:
            #---Load the file using numpy---
            with open(n,'r') as UseFile:
                filetext = UseFile.read()
                lines = filetext.splitlines()
            if ("frequency" in lines[1].lower()):
                data = np.loadtxt(n, skiprows=2)
            else:
                data = np.loadtxt(n)
            re_in = data[0,0]
            cap_in = data[0, 1]
            w_in = data[1:,0]
            rres_in = data[1:,1]
            jres_in = data[1:,2]
            r_in = data[1:,3]
            j_in = data[1:,4]
            #---Check that all the data are the same length---
            if (not (len(w_in) == len(rres_in) and len(rres_in) == len(jres_in) and len(jres_in) == len(r_in) and len(r_in) == len(j_in))):
                raise FileLengthError
            #---Check that all the data are numeric---
            for i in range(len(w_in)):
                float(re_in)
                float(cap_in)
                float(w_in[i])
                float(rres_in[i])
                float(jres_in[i])
                float(r_in[i])
                float(j_in[i])
            #---Set the appropriate variables---
            self.fileListbox.insert(tk.END, n)
            self.wdataLength = len(w_in)
            self.noneLoaded = False
            self.re.append(re_in)
            self.cap.append(cap_in)
            self.wdata.append(w_in)
            self.rres.append(rres_in)
            self.jres.append(jres_in)
            self.zr.append(r_in)
            self.zj.append(j_in)
            self.numFiles += 1
            #---Update the GUI---
            self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%self.re[0])
            if (self.cap[0] != 0):
                self.capAvgLabel.configure(text="|  Average Capacitance = %.5g "%self.cap[0])
        #---Catch errors and display appropriate messages---
        except FileLengthError:
            messagebox.showerror("Length error", "Error 27:\nThe number of data do not match")
        except FrequencyMismatchError:
            messagebox.showerror("Frequency error", "Error 28:\nSome of the frequencies being loaded do not match the already loaded frequencies")
        except Exception:
            messagebox.showerror("File error", "Error 29:\nThere was an error reading the file")
    
    def residualsEnter(self, n):
        """
        Load files that were opened with the program (i.e. files supplied as a command-line argument)
        """
        #---See addFile() under __init__()---
        try:
            alreadyWarned = False
            for name in n:
                fname, fext = os.path.splitext(name)
                if (fext != ".mmresiduals"):
                    raise FileExtensionError
                with open(name,'r') as UseFile:
                    filetext = UseFile.read()
                    lines = filetext.splitlines()
                if ("frequency" in lines[1].lower()):
                    data = np.loadtxt(name, skiprows=2)
                else:
                    data = np.loadtxt(name)
                re_in = data[0,0]
                cap_in = data[0, 1]
                w_in = data[1:,0]
                rres_in = data[1:,1]
                jres_in = data[1:,2]
                r_in = data[1:,3]
                j_in = data[1:,4]
                if (not (len(w_in) == len(rres_in) and len(rres_in) == len(jres_in) and len(jres_in) == len(r_in) and len(r_in) == len(j_in))):
                    raise FileLengthError
                for i in range(len(w_in)):
                    float(re_in)
                    float(cap_in)
                    float(w_in[i])
                    float(rres_in[i])
                    float(jres_in[i])
                    float(r_in[i])
                    float(j_in[i])
                if (not self.noneLoaded):
                    if (len(w_in) != self.wdataLength):
                        raise FileLengthError
                    else:
                        for i in range(len(w_in)):
                            if (self.wdata[0][i] != w_in[i]):
                                raise FrequencyMismatchError
                if (name in self.fileListbox.get(0, tk.END) and not alreadyWarned):
                    alreadyWarned = True
                    if (len(n) > 1):
                        messagebox.showwarning("File already loaded", "Warning: One or more of the selected files is already loaded and is being loaded again")
                    else:
                        messagebox.showwarning("File already loaded", "Warning: The selected file is already loaded and is being loaded again")
                self.fileListbox.insert(tk.END, name)
                if (self.noneLoaded):
                    self.wdataLength = len(w_in)
                    self.noneLoaded = False
                self.re.append(re_in)
                self.cap.append(cap_in)
                self.wdata.append(w_in)
                self.rres.append(rres_in)
                self.jres.append(jres_in)
                self.zr.append(r_in)
                self.zj.append(j_in)
                self.numFiles += 1
            if (self.numFiles > 1):
                re_avg = 0
                for i in range(len(self.zr)):
                    re_avg += self.re[i]
                re_avg /= len(self.zr)
                
                cap_avg = 0
                for i in range(len(self.zr)):
                    cap_avg += self.cap[i]
                cap_avg /= len(self.zr)
                
                meansReal = np.zeros(self.wdataLength)
                for i in range(self.wdataLength):
                    for j in range(len(self.rres)):
                        meansReal[i] += self.rres[j][i]
                    meansReal[i] /= len(self.rres)
    
                meansImag = np.zeros(self.wdataLength)
                for i in range(self.wdataLength):
                    for j in range(len(self.jres)):
                        meansImag[i] += self.jres[j][i]
                    meansImag[i] /= len(self.jres)
                
                standardDevsReal = np.zeros(self.wdataLength)
                for i in range(self.wdataLength):
                    for j in range(len(self.rres)):
                        standardDevsReal[i] += (self.rres[j][i] - meansReal[i])**2
                    standardDevsReal[i] /= len(self.rres)-1
                    standardDevsReal[i] = np.sqrt(standardDevsReal[i])
                
                standardDevsImag = np.zeros(self.wdataLength)
                for i in range(self.wdataLength):
                    for j in range(len(self.jres)):
                        standardDevsImag[i] += (self.jres[j][i] - meansImag[i])**2
                    standardDevsImag[i] /= len(self.jres)-1
                    standardDevsImag[i] = np.sqrt(standardDevsImag[i])
                
                self.statsView.delete(*self.statsView.get_children())
                for i in range(self.wdataLength):
                    self.statsView.insert("", tk.END, text="", values=(str(self.wdata[0][i]), "%.6g"%meansReal[i], "%.6g"%standardDevsReal[i], "%.6g"%meansImag[i], "%.6g"%standardDevsImag[i]))
                self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%re_avg)
                if (cap_avg != 0):
                    self.capAvgLabel.configure(text="|  Average Capacitance = %.5g"%cap_avg)
                self.standardDevsR = standardDevsReal
                self.standardDevsI = standardDevsImag
            elif (self.numFiles == 1):
                self.reAvgLabel.configure(text="Average R\u2091 = %.5g "%self.re[0])
                if (self.cap[0] != 0):
                    self.capAvgLabel.configure(text="|  Average Capacitance = %.5g"%self.cap[0])
                
        except FileExtensionError:
            messagebox.showerror("File error", "Error 26:\nThe file has an unknown extension")
        except FileLengthError:
            messagebox.showerror("Length error", "Error 27:\nThe number of data do not match")
        except FrequencyMismatchError:
            messagebox.showerror("Frequency error", "Error 28:\nSome of the frequencies being loaded do not match the already loaded frequencies")
        except Exception:
            messagebox.showerror("File error", "Error 29:\nThere was an error reading the file")
        if (self.numFiles > 1):
            self.saveButton.configure(state="normal")
            self.plotButton.configure(state="normal")                         