# -*- coding: utf-8 -*-
"""
Created on Wed Jun 20 15:23:25 2018

@author: willdwat
"""
import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename, asksaveasfile
from tkinter import messagebox
import numpy as np
import pyperclip
#--------------------------------pyperclip-----------------------------------
#     Source: https://pypi.org/project/pyperclip/
#     Email: al@inventwithpython.com
#     By: Al Sweigart
#     License: BSD
#----------------------------------------------------------------------------
import os
import sys
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import detect_delimiter
#------------------------------detect_delimiter------------------------------
#    Source: https://pypi.org/project/detect-delimiter/
#    Email: paperless@timmcnamara.co.nz
#    Copyright 2018   Tim McNamara
#    License: Apache License Version 2.0

#    Licensed under the Apache License, Version 2.0 (the "License");
#    you may not use this file except in compliance with the License.
#    You may obtain a copy of the License at

#        http://www.apache.org/licenses/LICENSE-2.0

#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS,
#    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#    See the License for the specific language governing permissions and
#    limitations under the License.
#----------------------------------------------------------------------------

#---Different exceptions to allow more precise error messages---
class BadMultiplierError(Exception):
    pass

class BadColumnError(Exception):
    pass

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

class CreateToolTipX(object):
    """
    create a tooltip for a given widget
    Code from: https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter with answer by crxguy52
    """
    def __init__(self, widget, color, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.foregroundColor = color
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.widget.configure(foreground='red')
        self.schedule()

    def leave(self, event=None):
        self.widget.configure(foreground=self.foregroundColor)
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

class TextLineNumbers(tk.Canvas):
    """
    Create line numbers in a text field
    from Bryan Oakley: https://stackoverflow.com/questions/16369470/tkinter-adding-line-number-to-text-widget
    """
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(self, *args, **kwargs)
        self.textwidget = None
        self.fillColor = "black"
        
    def attach(self, text_widget):
        self.textwidget = text_widget
        
    def setFillColor(self, color):
        self.fillColor = color
        self.redraw()
    
    def redraw(self, *args):
        '''redraw line numbers'''
        self.delete("all")

        i = self.textwidget.index("@0,0")
        while True :
            dline= self.textwidget.dlineinfo(i)
            if dline is None: break
            y = dline[1]
            linenum = str(i).split(".")[0]
            self.create_text(2,y,anchor="nw", fill=self.fillColor, text=linenum)
            i = self.textwidget.index("%s+1line" % i)

class CustomText(tk.Text):
    """
    Create line numbers in a text field
    from Bryan Oakley: https://stackoverflow.com/questions/16369470/tkinter-adding-line-number-to-text-widget
    """
    def __init__(self, *args, **kwargs):
        tk.Text.__init__(self, *args, **kwargs)

        # create a proxy for the underlying widget
        self._orig = self._w + "_orig"
        self.tk.call("rename", self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def _proxy(self, *args):
        # let the actual widget perform the requested action
        cmd = (self._orig,) + args
        result = self.tk.call(cmd)

        # generate an event if something was added or deleted,
        # or the cursor position changed
        if (args[0] in ("insert", "replace", "delete") or 
            args[0:3] == ("mark", "set", "insert") or
            args[0:2] == ("xview", "moveto") or
            args[0:2] == ("xview", "scroll") or
            args[0:2] == ("yview", "moveto") or
            args[0:2] == ("yview", "scroll")
        ):
            self.event_generate("<<Change>>", when="tail")

        # return what the actual widget returned
        return result

class iF(tk.Frame):
    def __init__(self, parent, topOne):

        def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
        
            return os.path.join(base_path, relative_path)
        
        self.topGUI = topOne
        self.parent = parent
        self.real_data = []
        self.imag_data = []
        self.freq_data = []
        self.fileText = ""
        self.currentFile = ""
        self.savedNeeded = False
        if (self.topGUI.getTheme() == "dark"):
            self.backgroundColor = "#424242"
            self.foregroundColor = "white"
        elif (self.topGUI.getTheme() == "light"):
            self.backgroundColor = "white"
            self.foregroundColor = "black"
        tk.Frame.__init__(self, parent, background=self.backgroundColor)
        
        #---Popup menu for input file---
        def popup_inputFileText(event):
            try:
                self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu.grab_release()
        
        #---Popup menu for deleting data at arbitrary frequencies---
        def popup_otherFrequencies(event):
            try:
                self.popup_menu2.tk_popup(event.x_root, event.y_root, 0)
            finally:
                self.popup_menu2.grab_release()
                
        def copyInputFileText():
            try:
                pyperclip.copy(self.inputFileText.selection_get())
            except:
                pass

        def select_inputFileText():
            self.inputFileText.focus_set()
            self.inputFileText.selection_range(0, tk.END)
            
        def OpenFile():
            name = askopenfilename(initialdir=self.topGUI.currentDirectory, filetypes =(("All files", "*.*"),("DTA files (*.dta)", "*.dta"),("DAT files (*.dat)", "*.dat"),("Text files (*.txt)","*.txt"),("CSV files (*.csv)", "*.csv")),title = "Choose a file")
            if (name == ""):
                return "+", "+"
            else:
                try:
                    with open(name,'r') as UseFile:
                        return name, UseFile.read()
                except:     #If there's a problem with the file
                    messagebox.showerror("File error", "Error 5:\nThere was an error opening the file.")
                    return "+", "+"   
        
        def browse():
            n, filetext = OpenFile()
            if (n != '+'):
                directory = os.path.dirname(n)
                self.topGUI.setCurrentDirectory(directory)
                self.inputFileText.configure(state="normal")
                self.inputFileText.delete(0,tk.END)
                self.inputFileText.insert(0, n)
                self.inputFileText.configure(state="readonly")
                self.browseTextBtn.configure(state="normal")
                self.fileText = filetext
                if (self.topGUI.getDetectComments() == 1):
                    numLinesOfComments = 0
                    whatDelimiter = "no"
                    for line in filetext.splitlines():
                        lineNew = ''.join(line.split())
                        #print(line)
                        if (len(lineNew)<5):
                            numLinesOfComments += 1
                            continue
                        elif (sum(c.isdigit() for c in lineNew) < sum(c.isalpha() for c in lineNew)):
                            numLinesOfComments += 1
                            continue
                        else:
                            if (self.topGUI.getDetectDelimiter() == 1):
                                whatDelimiter = detect_delimiter.detect(line)
                                if (whatDelimiter == ","):
                                    self.delimiterVariable.set(",")
                                elif (whatDelimiter == "\t"):
                                    self.delimiterVariable.set("Tab")
                                elif (whatDelimiter == " "):
                                    self.delimiterVariable.set("Space")
                                elif (whatDelimiter == ";"):
                                    self.delimiterVariable.set(";")
                                elif (whatDelimiter == "|"):
                                    self.delimiterVariable.set("|")
                                elif (whatDelimiter == ":"):
                                    self.delimiterVariable.set(":")
                            break #reached the end of comment lines
                    self.commentVariable.set(numLinesOfComments)
        
        #---Delete data in browse text entry field---
        def clearBrowse(event):
            self.inputFileText.configure(state="normal")
            self.inputFileText.delete(0, tk.END)
            self.inputFileText.configure(state="readonly")
            self.browseTextBtn.configure(state="disabled")
            self.fileText = ""
            
        def displayLineFrequency():
            if (self.lineFrequencyCheckboxVariable.get() == 0):
                self.lineFrequencyFrame.grid_remove()
            else:
                self.lineFrequencyFrame.grid(column=1, row=0, columnspan=2, sticky="W")
                
        def displayOtherFrequency():
            if (self.otherFrequencyCheckboxVariable.get() == 0):
                self.otherFrequencyFrame.grid_remove()
            else:
                self.otherFrequencyFrame.grid(column=0, row=1, sticky="W")
        
        def validateComment(P):
            if P == '':
                return True
            if " " in P:
                return False
            if "\t" in P:
                return False
            if "\n" in P:
                return False
            try:
                int(P)
                if (int(P) < 0):     #No negative numbers of comment lines
                    return False
                else:
                    return True
            except:
                return False
        def validateCol(P):
            if P == '':
                return True
            if " " in P:
                return False
            if "\t" in P:
                return False
            if "\n" in P:
                return False
            try:
                int(P)
                if (int(P) <= 0):   #No negative or 0 column numbers
                    return False
                else:
                    return True
            except:
                return False
        
        def validatePlusMinus(P):
            if P == '':         #Allow no +/- value
                return True
            if P == '.':
                return True
            if " " in P:        #Prevent spaces
                return False
            if "\t" in P:
                return False
            if "\n" in P:
                return False
            try:                #Check if the entry is a number
                float(P)
                if (float(P) < 0):
                    return False
                else:
                    return True
            except:
                return False
            
        def validateUnit(P):
            if P == '':
                return True
            if P == '-':
                return True
            if P == '.':
                return True
            if (' ' in P or '\t' in P or '\n' in P):
                return False
            try:
                float(P)
                return True
            except:
                return False
        
        valUnit = (self.register(validateUnit), '%P')
        valCol = (self.register(validateCol), '%P')
        
        #---Add arbitrary frequency range to delete---
        def addFreq():
            if (self.otherFrequencyNumVariable.get() != ""):
                if (self.otherFrequencyErrorVariable.get() != "" and self.otherFrequencyErrorVariable.get() != "."):
                    self.otherFrequencyListbox.insert(tk.END, self.otherFrequencyNumVariable.get() + " ± " + self.otherFrequencyErrorVariable.get())
                else:
                    self.otherFrequencyListbox.insert(tk.END, self.otherFrequencyNumVariable.get() + " ± 0")
            self.otherFrequencyNumVariable.set("")
            self.otherFrequencyErrorVariable.set("")
        
        def loadData():
            if (self.inputFileText.get() != ""):    #If there is a file in the text field
                try:
                    if (self.realUnitVariable.get() == "." or self.imagUnitVariable.get() == "." or self.freqUnitVariable.get() == "."):
                        raise BadMultiplierError
                    if (self.realColVariable.get() == "" or self.imagColVariable.get() == "" or self.freqColVariable.get() == ""):
                        raise BadColumnError
                    if (self.commentNum.get() == ""):
                        numComments = 0
                    else:
                        numComments = int(self.commentNum.get())
                    if (self.delimiterVariable.get() == "Tab"):
                        data = np.loadtxt(self.inputFileText.get(), skiprows=numComments)
                    elif (self.delimiterVariable.get() == "Space"):
                        data = np.loadtxt(self.inputFileText.get(), skiprows=numComments, delimiter=" ")
                    else:
                        data = np.loadtxt(self.inputFileText.get(), skiprows=numComments, delimiter = self.delimiterVariable.get())
                    freq_read = data[:, int(self.freqCol.get())-1]
                    if (self.freqUnitVariable.get() != "" and self.freqUnitVariable.get() != "-"):
                        freq_read *= float(self.freqUnitVariable.get())
                    elif (self.freqUnitVariable.get() == "-"):
                        freq_read *= -1
                    real_read = data[:, int(self.realCol.get())-1]
                    if (self.realUnitVariable.get() != "" and self.realUnitVariable.get() != "-"):
                        real_read *= float(self.realUnitVariable.get())
                    elif (self.realUnitVariable.get() == "-"):
                        real_read *= -1
                    imag_read = data[:, int(self.imagCol.get())-1]
                    if (self.imagUnitVariable.get() != "" and self.imagUnitVariable.get() != "-"):
                        imag_read *= float(self.imagUnitVariable.get())
                    elif (self.imagUnitVariable.get() == "-"):
                        imag_read *= -1
                    freq_in = freq_read.tolist()
                    real_in = real_read.tolist()
                    imag_in = imag_read.tolist()
                    if (not (len(freq_in) == len(real_in) and len(real_in) == len(imag_in))):
                        raise IndexError
                    if (self.lineFrequencyCheckboxVariable.get() != 0):
                        pM = 0
                        if (self.plusMinusVariable.get() != "" and self.plusMinusVariable.get() != "."):
                            pM = int(self.plusMinusVariable.get())
                        freqsToDelete = []
                        if (self.whichLineFrequenciesVariable.get() == "50"):
                            for i in range(len(freq_read)):
                                if (freq_read[i] >= 50-pM and freq_read[i] <= 50+pM):
                                    freqsToDelete.append(i)
                        elif (self.whichLineFrequenciesVariable.get() == "60"):
                            for i in range(len(freq_read)):
                                if (freq_read[i] >= 60-pM and freq_read[i] <= 60+pM):
                                    freqsToDelete.append(i)
                        elif (self.whichLineFrequenciesVariable.get() == "50&100"):
                            for i in range(len(freq_read)):
                                if ((freq_read[i] >= 50-pM and freq_read[i] <= 50+pM) or (freq_read[i] >= 100-pM and freq_read[i] <= 100+pM)):
                                    freqsToDelete.append(i)
                        elif (self.whichLineFrequenciesVariable.get() == "60&120"):
                            for i in range(len(freq_read)):
                                if ((freq_read[i] >= 60-pM and freq_read[i] <= 60+pM) or (freq_read[i] >= 120-pM and freq_read[i] <= 120+pM)):
                                    freqsToDelete.append(i)
                        corrector = 0
                        for i in range(len(freqsToDelete)):
                            freq_in.pop(freqsToDelete[i]-corrector)
                            real_in.pop(freqsToDelete[i]-corrector)
                            imag_in.pop(freqsToDelete[i]-corrector)
                            corrector += 1
                    if (self.otherFrequencyCheckboxVariable.get() != 0 and self.otherFrequencyListbox.size() != 0):
                        freqsOther = []
                        pMOther = []
                        freqsOtherDelete = []
                        for i in range(self.otherFrequencyListbox.size()):
                            freqsOther.append(float(self.otherFrequencyListbox.get(i).split("±")[0]))
                            pMOther.append(float(self.otherFrequencyListbox.get(i).split("±")[1])) 
                        for i in range(len(freq_in)):
                            for j in range (len(freqsOther)):
                                if (freq_in[i] >= freqsOther[j]-pMOther[j] and freq_in[i] <= freqsOther[j]+pMOther[j]):
                                    freqsOtherDelete.append(i)
                        correctorOther = 0
                        for i in range(len(freqsOtherDelete)):
                            freq_in.pop(freqsOtherDelete[i]-correctorOther)
                            real_in.pop(freqsOtherDelete[i]-correctorOther)
                            imag_in.pop(freqsOtherDelete[i]-correctorOther)
                            correctorOther += 1
                    self.real_data = []     #Clear the existing data
                    self.imag_data = []
                    self.freq_data = []
                    self.dataView.delete(*self.dataView.get_children())
                    for i in range(len(freq_in)):
                        self.dataView.insert("", tk.END, text=str(i+1), values=(str(freq_in[i]), str(real_in[i]), str(imag_in[i])))
                        self.freq_data.append(freq_in[i])
                        self.real_data.append(real_in[i])
                        self.imag_data.append(imag_in[i])
                    self.numData.grid(column=0, row=6, sticky="W")
                    self.dataViewFrame.grid(column=0, row=7)
                    self.saveFrame.grid(column=0, row=8, sticky="W", pady=5)
                    self.numData.configure(text="Number of data: " + str(len(freq_in)))
                    self.loadData.configure(text="Reload Data")
                    self.currentFile = self.inputFileText.get()
                    self.savedNeeded = True
                except IndexError:
                    messagebox.showerror("Different lengths", "Error 6:\nThe number of frequencies, real parts, and imaginary parts are not equal")
                except BadMultiplierError:
                    messagebox.showerror("Missing multiplier", "Error 7:\nOne of the multipliers has an invalid value of \".\"")
                except BadColumnError:
                    messagebox.showerror("Missing column", "Error 8:\nOne of the column numbers is missing")
                except:
                    messagebox.showerror("File error", "Error 5:\nThere was an error loading or reading the file")

        def saveAs():
            continueThoughDiff = True
            if (self.currentFile != self.inputFileText.get()):
                continueThoughDiff = messagebox.askokcancel("Different files", "The current file under \"Browse\" has not been loaded, and so will not be saved. If you continue, only the currently loaded data will be saved.")
            if (continueThoughDiff):
                defaultSaveName, ext = os.path.splitext(os.path.basename(self.inputFileText.get()))
                saveName = asksaveasfile(mode='w', defaultextension=".mmfile", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Measurement model file", ".mmfile")])
                directory = os.path.dirname(str(saveName))
                self.topGUI.setCurrentDirectory(directory)
                if saveName is None:     #If save is cancelled
                    return
                for i in range(len(self.freq_data)):
                    saveName.write(str(self.freq_data[i]) + "\t" + str(self.real_data[i]) + "\t" + str(self.imag_data[i]) + "\n")
                saveName.close()
                self.savedNeeded = False
                self.saveButton.configure(text="Saved")
                self.after(1000, lambda : self.saveButton.configure(text="Save as"))
            
        
        def saveContinue():
            continueThoughDiff = True
            if (self.currentFile != self.inputFileText.get()):
                continueThoughDiff = messagebox.askokcancel("Different files", "The current file under \"Browse\" has not been loaded, and so will not be saved. If you continue, only the currently loaded data will be saved.")
            if (continueThoughDiff):
                defaultSaveName, ext = os.path.splitext(os.path.basename(self.inputFileText.get()))
                saveName = asksaveasfile(mode='w', defaultextension=".mmfile", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Measurement model file", ".mmfile")])
                directory = os.path.dirname(str(saveName))
                self.topGUI.setCurrentDirectory(directory)
                if saveName is None:     #If save is cancelled
                    return
                for i in range(len(self.freq_data)):
                    saveName.write(str(self.freq_data[i]) + "\t" + str(self.real_data[i]) + "\t" + str(self.imag_data[i]) + "\n")
                self.savedNeeded = False
                saveName.close()
                self.topGUI.enterMeasureModel(saveName.name)

        #---Plot the data---
        def plotRawData():
            nyPlot = tk.Toplevel(background=self.backgroundColor)
            nyPlot.title("Current Loaded Data")     #Plot window title
            nyPlot.iconbitmap(resource_path('img/elephant3.ico'))
            x = np.array(self.real_data)    #Doesn't plot without this
            y = np.array(self.imag_data)
            with plt.rc_context({'axes.edgecolor':self.foregroundColor, 'xtick.color':self.foregroundColor, 'ytick.color':self.foregroundColor, 'figure.facecolor':self.backgroundColor}):
                fInput = Figure(figsize=(5,4), dpi=100)
                fInput.set_facecolor(self.backgroundColor)
                a = fInput.add_subplot(111)
                a.set_facecolor(self.backgroundColor)
                a.yaxis.set_ticks_position("both")
                a.yaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                a.xaxis.set_ticks_position("both")
                a.xaxis.set_tick_params(direction="in", color=self.foregroundColor)     #Make the ticks point inwards
                plotColor = "tab:blue"
                if (self.topGUI.getTheme() == "dark"):
                    plotColor = "cyan"
                else:
                    plotColor = "tab:blue"
                a.plot(x, -1*y, "o", color=plotColor)
                a.axis("equal")
                a.set_title("Current Loaded Data", color=self.foregroundColor)
                a.set_xlabel("Zr / Ω", color=self.foregroundColor)
                a.set_ylabel("-Zj / Ω", color=self.foregroundColor)
                fInput.subplots_adjust(left=0.14)   #Allows the y axis label to be more easily seen
                nyCanvasInput = FigureCanvasTkAgg(fInput, nyPlot)
                nyCanvasInput.draw()
                nyCanvasInput.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
                toolbar = NavigationToolbar2Tk(nyCanvasInput, nyPlot)    #Enables the zoom and move toolbar for the plot
                toolbar.update()
            nyCanvasInput._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            def on_closing():   #Clear the figure before closing the popup
                fInput.clear()
                nyPlot.destroy()
            
            nyPlot.protocol("WM_DELETE_WINDOW", on_closing)
       
        def handle_click(event):
            if self.dataView.identify_region(event.x, event.y) == "separator":
                if (self.dataView.identify_column(event.x) == "#0"):
                    return "break"
        
        def handle_motion(event):
            if self.dataView.identify_region(event.x, event.y) == "separator":
                if (self.dataView.identify_column(event.x) == "#0"):
                    return "break"
            
        def browseText():
            def _on_change(event):
                self.linenumbers.redraw()
            self.textPopup = tk.Toplevel()
            self.textPopup.title(self.inputFileText.get())
            self.textPopup.iconbitmap(resource_path("img/elephant3.ico"))
            self.text = CustomText(self.textPopup, bg=self.backgroundColor, fg=self.foregroundColor)
            self.vsb = tk.Scrollbar(self.textPopup, orient="vertical", command=self.text.yview)
            self.text.configure(yscrollcommand=self.vsb.set)
            self.linenumbers = TextLineNumbers(self.textPopup, bg=self.backgroundColor, width=30)
            self.linenumbers.attach(self.text)
            self.vsb.pack(side="right", fill="y")
            self.linenumbers.pack(side="left", fill="y")
            self.text.pack(side="right", fill="both", expand=True)
            self.text.bind("<<Change>>", _on_change)
            self.text.bind("<Configure>", _on_change)
            self.text.insert("end", self.fileText)
            self.linenumbers.setFillColor(self.foregroundColor)
            
        s = ttk.Style()
        if (self.topGUI.getTheme() == "dark"):
            s.configure('TCheckbutton', background='#424242', foreground="white")
        elif (self.topGUI.getTheme() == "light"):
            s.configure('TCheckbutton', background='white', foreground="black")
            s.configure('TLabelFrame', background='#424242', foreground="#FFFFFF")
        s.configure('TCombobox', background='white')
        s.configure('TEntry', background='white')
        s.configure('TNotebook', background='white')
        #s.configure
        
        #---Browse for a file---
        self.browseFrame = tk.Frame(self, bg=self.backgroundColor)
        self.browseBtn = ttk.Button(self.browseFrame, text="Browse...", command=browse)
        self.inputFileText = ttk.Entry(self.browseFrame, state="readonly", width=40)
        self.clearBtn = tk.Label(self.browseFrame, text="⨯", bg=self.backgroundColor, fg=self.foregroundColor, font=("Times New Roman", 14), cursor="hand2")
        self.clearBtn.bind("<Button-1>", clearBrowse)
#        self.clearBtn.bind("<Enter>", lambda e: self.clearBtn.configure(foreground="red"))
#        self.clearBtn.bind("<Leave>", lambda e: self.clearBtn.configure(foreground=self.foregroundColor))
        self.browseTextBtn = ttk.Button(self.browseFrame, text="View file", state="disabled", command=browseText)
        self.browseBtn.grid(column=0, row=0)
        self.inputFileText.grid(column=1, row=0, padx=5)
        #self.clearBtn.grid(column=2, row=0)
        self.browseTextBtn.grid(column=0, row=2, pady=4, sticky="W")
        self.browseFrame.grid(column = 0, row=0, pady=(0,3), sticky="W")
        browse_ttp = CreateToolTip(self.browseBtn, 'Browse for an input file')
        browse_text_ttp = CreateToolTip(self.browseTextBtn, 'View file contents')
        clear_ttp = CreateToolTipX(self.clearBtn, self.foregroundColor, 'Clear file')
        
        #---Popup menu on the browse entry---
#        self.popup_menu = tk.Menu(self.inputFileText, tearoff=0)
#        self.popup_menu.add_command(bitmap="@copy.xbm", compound=tk.LEFT, label="Copy", command=copyInputFileText)
#        self.popup_menu.add_command(label="Select text", command=select_inputFileText)
#        self.inputFileText.bind("<Button-3>", popup_inputFileText)
        
        #---Number of comment lines---
        self.commentFrame = tk.Frame(self, bg=self.backgroundColor)
        self.commentLabel = tk.Label(self.commentFrame, text="Number of comment lines:", bg=self.backgroundColor, fg=self.foregroundColor)  
        self.commentVariable = tk.StringVar(self, self.topGUI.getComments())
        vcmd1 = (self.register(validateComment), '%P')
        self.commentNum = ttk.Entry(self.commentFrame, textvariable=self.commentVariable, width=3, validate="all", validatecommand=vcmd1)
        self.commentNum.bind("<FocusIn>", lambda e: self.commentNum.selection_range(0, tk.END))
        self.delimiterLabel = tk.Label(self.commentFrame, text="     Delimiter:", bg=self.backgroundColor, fg=self.foregroundColor)
        self.delimiterVariable = tk.StringVar(self, self.topGUI.getDelimiter())
        self.delimiterCombobox = ttk.Combobox(self.commentFrame, textvariable=self.delimiterVariable, value=("Tab", "Space", ";", ",", "|", ":"), exportselection=0, state="readonly", width=6)
        self.commentLabel.grid(column=0, row=0, sticky="W")
        self.commentNum.grid(column=1, row=0, padx=5, sticky="W")
        self.delimiterLabel.grid(column=2, row=0, padx=5)
        self.delimiterCombobox.grid(column=3, row=0)
        self.commentFrame.grid(column=0, row=1, pady=5, sticky="W")
        commentNum_ttp = CreateToolTip(self.commentNum, 'Number of lines to ignore')
        delimiter_ttp = CreateToolTip(self.delimiterCombobox, 'Which character specifies breaks in the data')
        
        #---Column numbers for data---
        vcmd = (self.register(validatePlusMinus), '%P')
        self.columnsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.freqColLabel = tk.Label(self.columnsFrame, text="Frequencies:      Column", bg=self.backgroundColor, fg=self.foregroundColor)
        self.freqColVariable = tk.StringVar(self, "1")
        self.freqCol = ttk.Entry(self.columnsFrame, textvariable=self.freqColVariable, width=2, validate="all", validatecommand=valCol)
        self.freqCol.bind("<FocusIn>", lambda e: self.freqCol.selection_range(0, tk.END))
        self.freqUnitVariable = tk.StringVar(self, "1")
        self.freqUnitCombobox = ttk.Combobox(self.columnsFrame, textvariable=self.freqUnitVariable, values=("-1", "0.01", "0.1", "1", "1000", "1000000"), validate="all", validatecommand=valUnit, width=8)
        self.freqUnitLabel = tk.Label(self.columnsFrame, text=" Scaling:  ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.realColLabel = tk.Label(self.columnsFrame, text="Real part:           Column", bg=self.backgroundColor, fg=self.foregroundColor)
        self.realColVariable = tk.StringVar(self, "2")
        self.realCol = ttk.Entry(self.columnsFrame, textvariable=self.realColVariable, width=2, validate="all", validatecommand=valCol)
        self.realCol.bind("<FocusIn>", lambda e: self.realCol.selection_range(0, tk.END))
        self.realUnitVariable = tk.StringVar(self, "1")
        self.realUnitCombobox = ttk.Combobox(self.columnsFrame, textvariable=self.realUnitVariable, values=("-1", "0.01", "0.1", "1", "1000", "1000000"), validate="all", validatecommand=valUnit, width=8)
        self.realUnitLabel = tk.Label(self.columnsFrame, text=" Scaling:  ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.imagColLabel = tk.Label(self.columnsFrame, text="Imaginary part: Column", bg=self.backgroundColor, fg=self.foregroundColor)
        self.imagColVariable = tk.StringVar(self, "3")
        self.imagCol = ttk.Entry(self.columnsFrame, textvariable=self.imagColVariable, width=2, validate="all", validatecommand=valCol)
        self.imagCol.bind("<FocusIn>", lambda e: self.imagCol.selection_range(0, tk.END))
        self.imagUnitVariable = tk.StringVar(self, "1")
        self.imagUnitCombobox = ttk.Combobox(self.columnsFrame, textvariable=self.imagUnitVariable, values=("-1", "0.01", "0.1", "1", "1000", "1000000"), validate="all", validatecommand=valUnit, width=8)
        self.imagUnitLabel = tk.Label(self.columnsFrame, text=" Scaling:  ", bg=self.backgroundColor, fg=self.foregroundColor)
#        self.rmCheckboxVariable = tk.IntVar(self, 0)
#        self.rmCheckbox = ttk.Checkbutton(self.columnsFrame, text="Rm", variable=self.rmCheckboxVariable)
#        self.rmColLabel = tk.Label(self.columnsFrame, text=":     Column", bg=self.backgroundColor, fg=self.foregroundColor)
#        self.rmCol = ttk.Entry(self.columnsFrame, exportselection=0, textvariable=self.rmColVariable, width=2, validate="all", validatecommand=valCol)
        self.freqColLabel.grid(column=0, row=0, sticky="W")
        self.freqCol.grid(column=1, row=0)
        self.freqUnitLabel.grid(column=2, row=0)
        self.freqUnitCombobox.grid(column=3, row=0, padx=2)
        self.realColLabel.grid(column=0, row=1, sticky="W")
        self.realCol.grid(column=1, row=1)
        self.realUnitLabel.grid(column=2, row=1)
        self.realUnitCombobox.grid(column=3, row=1, padx=2)
        self.imagColLabel.grid(column=0, row=2, sticky="W")
        self.imagCol.grid(column=1, row=2)
        self.imagUnitLabel.grid(column=2, row=2)
        self.imagUnitCombobox.grid(column=3, row=2, padx=2)
        #self.rmCheckbox.grid(column=0, row=3, sticky="W")
        self.columnsFrame.grid(column=0, row=2, pady=5, sticky="W")
        freqCol_ttp = CreateToolTip(self.freqCol, 'Which column holds the frequency data')
        freqCombobox_ttp = CreateToolTip(self.freqUnitCombobox, 'Multiplier for the frequency data')
        imagCol_ttp = CreateToolTip(self.imagCol, 'Which column holds the imaginary data')
        imagCombobox_ttp = CreateToolTip(self.imagUnitCombobox, 'Multiplier for the imaginary data')
        realCol_ttp = CreateToolTip(self.realCol, 'Which column holds the real data')
        realCombobox_ttp = CreateToolTip(self.realUnitCombobox, 'Multiplier for the real data')
        
        #---Line frequency deletion---
        self.lineFrequencyMasterFrame = tk.Frame(self, bg=self.backgroundColor)
        self.lineFrequencyCheckboxVariable = tk.IntVar(self, 1)
        self.lineFrequencyCheckbox = ttk.Checkbutton(self.lineFrequencyMasterFrame, text="Delete data at line frequency", variable=self.lineFrequencyCheckboxVariable, command=displayLineFrequency)
        self.lineFrequencyCheckbox.grid(column=0, row=0, sticky="W")
        lineFrequencyCheckbox_ttp = CreateToolTip(self.lineFrequencyCheckbox, 'Delete data at a chosen line frequency/first harmonic')
        
        self.lineFrequencyFrame = tk.Frame(self.lineFrequencyMasterFrame, bg=self.backgroundColor)
        self.whichLineFrequenciesVariable = tk.StringVar(self, "60&120")
        self.whichLineFrequencies = ttk.Combobox(self.lineFrequencyFrame, textvariable=self.whichLineFrequenciesVariable, values=("50", "60", "50&100", "60&120"), exportselection=0, state="readonly", width=12)
        self.plusMinusLabel = tk.Label(self.lineFrequencyFrame, text="±", bg=self.backgroundColor, fg=self.foregroundColor)
        self.plusMinusVariable = tk.StringVar(self, "3")
        
        self.plusMinus = ttk.Entry(self.lineFrequencyFrame, textvariable=self.plusMinusVariable, width=3, validate="all", validatecommand=vcmd)
        self.plusMinus.bind("<FocusIn>", lambda e: self.plusMinus.selection_range(0, tk.END))
        self.plusMinusUnits = tk.Label(self.lineFrequencyFrame, text="Hz", bg=self.backgroundColor, fg=self.foregroundColor)
        self.whichLineFrequencies.grid(column=0, row=0)
        self.plusMinusLabel.grid(column=1, row=0)
        self.plusMinus.grid(column=2, row=0)
        self.plusMinusUnits.grid(column=3, row=0)
        self.lineFrequencyFrame.grid(column=1, row=0, columnspan=2, sticky="W", padx=5)
        self.lineFrequencyMasterFrame.grid(column=0, row=3, sticky="W", pady=7)
        plusMinus_ttp = CreateToolTip(self.plusMinus, 'What range of frequencies above and below should be deleted')
        
        #---Other frequency deletion---
        self.otherFrequencyMasterFrame = tk.Frame(self, bg=self.backgroundColor)
        self.otherFrequencyCheckboxVariable = tk.IntVar(self, 0)
        self.otherFrequencyCheckbox = ttk.Checkbutton(self.otherFrequencyMasterFrame, text="Delete data at arbitrary frequency", variable=self.otherFrequencyCheckboxVariable, command=displayOtherFrequency)
        self.otherFrequencyFrame = tk.Frame(self.otherFrequencyMasterFrame, bg=self.backgroundColor)
        self.otherFrequencyListboxScrollbar = ttk.Scrollbar(self.otherFrequencyFrame, orient=tk.VERTICAL)
        self.otherFrequencyListbox = tk.Listbox(self.otherFrequencyFrame, height=3, selectmode=tk.BROWSE, activestyle='none', yscrollcommand=self.otherFrequencyListboxScrollbar.set)
        self.otherFrequencyListboxScrollbar['command'] = self.otherFrequencyListbox.yview
        self.otherFrequencyAddFrame = tk.Frame(self.otherFrequencyFrame, bg=self.backgroundColor)
        self.otherFrequencyNumVariable = tk.StringVar(self, "")
        self.otherFrequencyNum = ttk.Entry(self.otherFrequencyAddFrame, textvariable=self.otherFrequencyNumVariable, width=8, validate="all", validatecommand=vcmd)
        self.otherFrequencyLabel = tk.Label(self.otherFrequencyAddFrame, text="±", bg=self.backgroundColor, fg=self.foregroundColor)
        self.otherFrequencyErrorVariable = tk.StringVar(self, "")
        self.otherFrequencyError = ttk.Entry(self.otherFrequencyAddFrame, textvariable=self.otherFrequencyErrorVariable, width=4, validate="all", validatecommand=vcmd)
        self.otherFrequencyListboxAdd = ttk.Button(self.otherFrequencyAddFrame, text="Add", command=addFreq)
        self.otherFrequencyListboxDelete = ttk.Button(self.otherFrequencyFrame, text="Remove", command=lambda lb=self.otherFrequencyListbox: lb.delete(tk.ANCHOR))
        self.otherFrequencyListbox.grid(column=0,row=0,rowspan=2)
        self.otherFrequencyListboxScrollbar.grid(column=1,row=0,rowspan=2)
        self.otherFrequencyNum.grid(column=0, row=0)
        self.otherFrequencyLabel.grid(column=1, row=0)
        self.otherFrequencyError.grid(column=2, row=0)
        self.otherFrequencyListboxAdd.grid(column=3, row=0, padx=2)
        self.otherFrequencyAddFrame.grid(column=2, row=0, padx=2)
        self.otherFrequencyListboxDelete.grid(column=2, row=1, padx=2, sticky="W")
        self.otherFrequencyCheckbox.grid(column = 0, row=0, sticky="W")
        self.otherFrequencyMasterFrame.grid(column=0, row=4, sticky="W")
        otherFrequency_ttp = CreateToolTip(self.otherFrequencyNum, 'Which frequency to be deleted')
        otherFrequencyPlusMins_ttp = CreateToolTip(self.otherFrequencyError, 'What range of frequencies above and below should be deleted')
        otherFrequencyAdd_ttp = CreateToolTip(self.otherFrequencyListboxAdd, 'Add frequency to list')
        otherFrequencyDelete_ttp = CreateToolTip(self.otherFrequencyListboxDelete, 'Remove selected frequency from list')
        otherFrequencyCheckbox_ttp = CreateToolTip(self.otherFrequencyCheckbox, 'Delete data at any number of arbitrary frequency ranges')
        
        self.popup_menu2 = tk.Menu(self.otherFrequencyListbox, tearoff=0)
        self.popup_menu2.add_command(label="Remove", command=lambda lb=self.otherFrequencyListbox: lb.delete(tk.ANCHOR))
        self.otherFrequencyListbox.bind("<Button-3>", popup_otherFrequencies)
        
        #---Load data button---
        self.loadData = ttk.Button(self, text="Load Data", width=20, command=loadData)
        self.loadData.grid(column=0, row=5, sticky="W", pady=10)
        loadData_ttp = CreateToolTip(self.loadData, 'Load data from file using given options')
        
        #---Number of data label---
        self.numData = tk.Label(self, text="Number of data: ", bg=self.backgroundColor, fg=self.foregroundColor)
                
        #---View current data---
        self.dataViewFrame = tk.Frame(self, bg=self.backgroundColor)
        self.dataViewScrollbar = ttk.Scrollbar(self.dataViewFrame, orient=tk.VERTICAL)
        self.dataView = ttk.Treeview(self.dataViewFrame, columns=("freq", "real", "imag"), height=5, selectmode="browse", yscrollcommand=self.dataViewScrollbar.set)
        self.dataView.heading("freq", text="Frequency")
        self.dataView.heading("real", text="Real part")
        self.dataView.heading("imag", text="Imaginary part")
        self.dataView.column("#0", width=50)
        self.dataView.column("freq", width=150)
        self.dataView.column("imag", width=150)
        self.dataView.column("real", width=150)
        self.dataViewScrollbar['command'] = self.dataView.yview
        self.dataView.grid(column=0, row=0)
        self.dataViewScrollbar.grid(column=1, row=0, sticky="NS")
        self.dataView.bind("<Button-1>", handle_click)
        self.dataView.bind("<Motion>", handle_motion)
        
        #---Save buttons---
        self.saveFrame = tk.Frame(self, bg=self.backgroundColor)
        self.saveButton = ttk.Button(self.saveFrame, text="Save As", command=saveAs)
        self.saveContinueButton = ttk.Button(self.saveFrame, text="Save and Load to Measurement Model", command=saveContinue)
        self.plotButton = ttk.Button(self.saveFrame, text="Nyquist plot", command=plotRawData)
        self.saveButton.grid(column=0, row=0)
        self.saveContinueButton.grid(column=1, row=0, padx=5)
        self.plotButton.grid(column=0, row=1, pady=5)
        saveButton_ttp = CreateToolTip(self.saveButton, 'Save loaded data as .mmfile')
        saveContinue_ttp = CreateToolTip(self.saveContinueButton, 'Save loaded data as .mmfile and load to Measurement Model tab')
        plot_ttp = CreateToolTip(self.plotButton, 'Display Nyquist plot')
    
    def needToSave(self):
        return self.savedNeeded
    
    def setThemeLight(self):
        self.foregroundColor = "#000000"
        self.backgroundColor = "#FFFFFF"
        self.configure(bg="#FFFFFF")
        self.browseFrame.configure(background="#FFFFFF")
        self.columnsFrame.configure(background="#FFFFFF")
        self.commentFrame.configure(background="#FFFFFF")
        self.saveFrame.configure(background="#FFFFFF")
        self.lineFrequencyMasterFrame.configure(background="#FFFFFF")
        self.lineFrequencyFrame.configure(background="#FFFFFF")
        self.otherFrequencyMasterFrame.configure(background="#FFFFFF")
        self.otherFrequencyFrame.configure(background="#FFFFFF")
        self.otherFrequencyAddFrame.configure(background="#FFFFFF")
        self.otherFrequencyLabel.configure(background="#FFFFFF", foreground="#000000")
        self.commentLabel.configure(background="#FFFFFF", foreground="#000000")
        self.delimiterLabel.configure(background="#FFFFFF", foreground="#000000")
        self.freqColLabel.configure(background="#FFFFFF", foreground="#000000")
        self.freqUnitLabel.configure(background="#FFFFFF", foreground="#000000")
        self.imagColLabel.configure(background="#FFFFFF", foreground="#000000")
        self.realColLabel.configure(background="#FFFFFF", foreground="#000000")
        self.imagUnitLabel.configure(background="#FFFFFF", foreground="#000000")
        self.realUnitLabel.configure(background="#FFFFFF", foreground="#000000")
        self.plusMinusLabel.configure(background="#FFFFFF", foreground="#000000")
        self.plusMinusUnits.configure(background="#FFFFFF", foreground="#000000")
        self.clearBtn.configure(background="#FFFFFF", foreground="#000000")
        self.numData.configure(background="#FFFFFF", foreground="#000000")
        try:
            self.text.configure(background="#FFFFFF", foreground="#000000")
            self.linenumbers.configure(background="#FFFFFF")
            self.linenumbers.setFillColor("#000000")
        except:
            pass
        s = ttk.Style()
        s.configure('TCheckbutton', background='#FFFFFF', foreground="#000000")
    
    def setThemeDark(self):
        self.foregroundColor = "#FFFFFF"
        self.backgroundColor = "#424242"
        self.configure(bg="#424242")
        self.browseFrame.configure(background="#424242")
        self.columnsFrame.configure(background="#424242")
        self.commentFrame.configure(background="#424242")
        self.saveFrame.configure(background="#424242")
        self.lineFrequencyMasterFrame.configure(background="#424242")
        self.lineFrequencyFrame.configure(background="#424242")
        self.otherFrequencyMasterFrame.configure(background="#424242")
        self.otherFrequencyFrame.configure(background="#424242")
        self.otherFrequencyAddFrame.configure(background="#424242")
        self.otherFrequencyLabel.configure(background="#424242", foreground="#FFFFFF")
        self.commentLabel.configure(background="#424242", foreground="#FFFFFF")
        self.delimiterLabel.configure(background="#424242", foreground="#FFFFFF")
        self.freqColLabel.configure(background="#424242", foreground="#FFFFFF")
        self.freqUnitLabel.configure(background="#424242", foreground="#FFFFFF")
        self.imagColLabel.configure(background="#424242", foreground="#FFFFFF")
        self.realColLabel.configure(background="#424242", foreground="#FFFFFF")
        self.imagUnitLabel.configure(background="#424242", foreground="#FFFFFF")
        self.realUnitLabel.configure(background="#424242", foreground="#FFFFFF")
        self.plusMinusLabel.configure(background="#424242", foreground="#FFFFFF")
        self.plusMinusUnits.configure(background="#424242", foreground="#FFFFFF")
        self.clearBtn.configure(background="#424242", foreground="#FFFFFF")
        self.numData.configure(background="#424242", foreground="#FFFFFF")
        try:
            self.text.configure(background="#424242", foreground="#FFFFFF")
            self.linenumbers.configure(background="#424242")
            self.linenumbers.setFillColor("#FFFFFF")
        except:
            pass
        s = ttk.Style()
        s.configure('TCheckbutton', background='#424242', foreground="white")
    
    def OpenFileNoPrompt(self, name):
        try:
            with open(name,'r') as UseFile:
                return name, UseFile.read()
        except:     #If there's a problem with the file
            messagebox.showerror("File error", "Error 5:\nThere was an error opening the file.")
            return "+", "+"
    
    def saveAs(self, e=None):
        if (len(self.real_data) == 0):
            return
        continueThoughDiff = True
        if (self.currentFile != self.inputFileText.get()):
            continueThoughDiff = messagebox.askokcancel("Different files", "The current file under \"Browse\" has not been loaded, and so will not be saved. If you continue, only the currently loaded data will be saved.")
        if (continueThoughDiff):
            defaultSaveName, ext = os.path.splitext(os.path.basename(self.inputFileText.get()))
            saveName = asksaveasfile(mode='w', defaultextension=".mmfile", initialfile=defaultSaveName, initialdir=self.topGUI.getCurrentDirectory, filetypes=[("Measurement model file", ".mmfile")])
            directory = os.path.dirname(str(saveName))
            self.topGUI.setCurrentDirectory(directory)
            if saveName is None:     #If save is cancelled
                return
            for i in range(len(self.freq_data)):
                saveName.write(str(self.freq_data[i]) + "\t" + str(self.real_data[i]) + "\t" + str(self.imag_data[i]) + "\n")
            saveName.close()
            self.savedNeeded = False
            self.saveButton.configure(text="Saved")
            self.after(1000, lambda : self.saveButton.configure(text="Save as"))
    
    def bindIt(self, e=None):
        self.bind_all("<Control-s>", self.saveAs)
    
    def unbindIt(self, e=None):
        self.unbind_all("<Control-s>")
            
    def inputEnter(self, fileName):
        n, filetext = self.OpenFileNoPrompt(fileName)
        fname, fext = os.path.splitext(fileName)
        directory = os.path.dirname(str(fileName))
        self.topGUI.setCurrentDirectory(directory)
        self.inputFileText.configure(state="normal")
        self.inputFileText.delete(0,tk.END)
        self.inputFileText.insert(0, n)
        self.inputFileText.configure(state="readonly")
        self.browseTextBtn.configure(state="normal")
        self.fileText = filetext
        if (self.topGUI.getDetectComments() == 1):
            numLinesOfComments = 0
            whatDelimiter = "no"
            for line in filetext.splitlines():
                lineNew = ''.join(line.split())
                #print(line)
                if (len(lineNew)<5):
                    numLinesOfComments += 1
                    continue
                elif (sum(c.isdigit() for c in lineNew) < sum(c.isalpha() for c in lineNew)):
                    numLinesOfComments += 1
                    continue
                else:
                    if (self.topGUI.getDetectDelimiter() == 1):
                        whatDelimiter = detect_delimiter.detect(line)
                        if (whatDelimiter == ","):
                            self.delimiterVariable.set(",")
                        elif (whatDelimiter == "\t"):
                            self.delimiterVariable.set("Tab")
                        elif (whatDelimiter == " "):
                            self.delimiterVariable.set("Space")
                        elif (whatDelimiter == ";"):
                            self.delimiterVariable.set(";")
                        elif (whatDelimiter == "|"):
                            self.delimiterVariable.set("|")
                        elif (whatDelimiter == ":"):
                            self.delimiterVariable.set(":")
                    break #reached the end of comment lines
            self.commentVariable.set(numLinesOfComments)
    