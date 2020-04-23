# -*- coding: utf-8 -*-
"""
Created on Fri Sep  7 11:37:56 2018

@author: willdwat
"""
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory
import tkinter.ttk as ttk
import tkinter.colorchooser as colorchooser
import configparser
import os

class FileExtensionError(Exception):
    pass

class FileLengthError(Exception):
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

class sF(tk.Frame):
    def __init__(self, parent, topOne):
        self.parent = parent
        self.topGUI = topOne
        if (self.topGUI.getTheme() == "dark"):
            self.backgroundColor = "#424242"
            self.foregroundColor = "#FFFFFF"
            s = ttk.Style()
            s.configure('TCheckbutton', background='#424242', foreground="#FFFFFF")
            s.configure('TNotebook', background='#424242')
        elif (self.topGUI.getTheme() == "light"):
            self.backgroundColor = "#FFFFFF"
            self.foregroundColor = "#000000"
            s = ttk.Style()
            s.configure('TCheckbutton', background='#FFFFFF', foreground="#000000")
            s.configure('TNotebook', background='#FFFFFF')
        self.barColor = self.topGUI.getBarColor()
        self.activeColor = self.topGUI.getAccentColor()
        self.highlightColor = self.topGUI.getHighlightColor()
        self.themeChosen = self.topGUI.getTheme()
        self.ellipseColor = self.topGUI.getEllipseColor()
        
        tk.Frame.__init__(self, parent, background=self.backgroundColor)
        
        def lightTheme(e):
            self.lightLabel.configure(relief="sunken")
            self.darkLabel.configure(relief="raised")
            self.themeChosen = 'light'
        def darkTheme(e):
            self.lightLabel.configure(relief="raised")
            self.darkLabel.configure(relief="sunken")
            self.themeChosen = 'dark'
            
        def pickBarColor(e):
            color = colorchooser.askcolor(self.barColor, title="Choose bar color") 
            if (color[1] is not None):
                self.barColorLabel.configure(bg=color[1])
                self.barColor = color[1]
        def pickHighlightColor(e):
            color = colorchooser.askcolor(self.highlightColor, title="Choose highlight color") 
            if (color[1] is not None):
                self.highlightColorLabel.configure(bg=color[1])
                self.highlightColor = color[1]
        def pickActiveColor(e):
            color = colorchooser.askcolor(self.activeColor, title="Choose active color") 
            if (color[1] is not None):
                self.activeColorLabel.configure(bg=color[1])
                self.activeColor = color[1]
        def pickEllipseColor(e):
            color = colorchooser.askcolor(self.ellipseColor, title="Choose error line/ellipse color") 
            if (color[1] is not None):
                self.defaultEllipseColorLabel.configure(bg=color[1])
                self.ellipseColor = color[1]
        
        def applySettings():
            if (self.themeChosen == "light"):
                self.configure(bg="#FFFFFF")
                #self.themeFrame.configure(bg="#FFFFFF")
                self.themeLabel.configure(bg="#FFFFFF")
                self.themeLabel.configure(fg="#000000")
                self.themeFrame.configure(background="#FFFFFF")
                self.lightLabel.configure(relief="sunken")
                self.darkLabel.configure(relief="raised")
                #self.accentsFrame.configure(bg="#FFFFFF")
                self.accentsLabel.configure(bg="#FFFFFF")
                self.accentsLabel.configure(fg="#000000")
                self.barLabel.configure(bg="#FFFFFF")
                self.barLabel.configure(fg="#000000")
                self.highlightLabel.configure(bg="#FFFFFF")
                self.highlightLabel.configure(fg="#000000")
                self.activeLabel.configure(bg="#FFFFFF")
                self.activeLabel.configure(fg="#000000")
                self.defaultMC.configure(bg="#FFFFFF", fg="#000000")
                #self.inputLabelFrame.configure(bg="#FFFFFF", fg="#000000")
                self.defaultComments.configure(bg="#FFFFFF", fg="#000000")
                self.defaultDelimiter.configure(bg="#FFFFFF", fg="#000000")
                #self.modelLabelFrame.configure(bg="#FFFFFF", fg="#000000")
                self.defaultMC.configure(bg="#FFFFFF", fg="#000000")
                self.defaultFit.configure(bg="#FFFFFF", fg="#000000")
                self.defaultWeighting.configure(bg="#FFFFFF", fg="#000000")
                self.defaultAlpha.configure(bg="#FFFFFF", fg="#000000")
                #self.errorLabelFrame.configure(bg="#FFFFFF", fg="#000000")
                self.defaultParamChoices.configure(bg="#FFFFFF", fg="#000000")
                self.defaultErrorWeighting.configure(bg="#FFFFFF", fg="#000000")
                self.defaultMovingAverage.configure(bg="#FFFFFF", fg="#000000")
                self.defaultEllipse.configure(bg="#FFFFFF", fg="#000000")
                self.tabOverall.configure(background="#FFFFFF")
                self.tabInput.configure(background="#FFFFFF")
                self.tabModel.configure(background="#FFFFFF")
                self.tabError.configure(background="#FFFFFF")
                self.tabCustom.configure(background="#FFFFFF")
                self.defaultDirectoryLabel.configure(bg="#FFFFFF", fg="#000000")
                self.defaultDetrendLabel.configure(bg="#FFFFFF", fg="#000000")
                self.defaultTabLabel.configure(bg="#FFFFFF", fg="#000000")
                s = ttk.Style()
                s.configure('TCheckbutton', background='#FFFFFF', foreground="#000000")
                s.configure('TNotebook', background='#FFFFFF')
                self.topGUI.setThemeLight()
            elif (self.themeChosen == "dark"):
                self.configure(bg="#424242")
                #self.themeFrame.configure(bg="#424242")
                self.themeLabel.configure(bg="#424242")
                self.themeLabel.configure(fg="#FFFFFF")
                self.themeFrame.configure(background="#424242")
                self.lightLabel.configure(relief="raised")
                self.darkLabel.configure(relief="sunken")
                #self.accentsFrame.configure(bg="#424242")
                self.accentsLabel.configure(bg="#424242")
                self.accentsLabel.configure(fg="#FFFFFF")
                self.barLabel.configure(bg="#424242")
                self.barLabel.configure(fg="#FFFFFF")
                self.highlightLabel.configure(bg="#424242")
                self.highlightLabel.configure(fg="#FFFFFF")
                self.activeLabel.configure(bg="#424242")
                self.activeLabel.configure(fg="#FFFFFF")
                self.defaultMC.configure(bg="#424242", fg="#FFFFFF")
                #self.inputLabelFrame.configure(bg="#424242", fg="#FFFFFF")
                self.defaultComments.configure(bg="#424242", fg="#FFFFFF")
                self.defaultDelimiter.configure(bg="#424242", fg="#FFFFFF")
                #self.modelLabelFrame.configure(bg="#424242", fg="#FFFFFF")
                self.defaultMC.configure(bg="#424242", fg="#FFFFFF")
                self.defaultFit.configure(bg="#424242", fg="#FFFFFF")
                self.defaultWeighting.configure(bg="#424242", fg="#FFFFFF")
                self.defaultAlpha.configure(bg="#424242", fg="#FFFFFF")
                #self.errorLabelFrame.configure(bg="#424242", fg="#FFFFFF")
                self.defaultParamChoices.configure(bg="#424242", fg="#FFFFFF")
                self.defaultErrorWeighting.configure(bg="#424242", fg="#FFFFFF")
                self.defaultMovingAverage.configure(bg="#424242", fg="#FFFFFF")
                self.defaultEllipse.configure(bg="#424242", fg="#FFFFFF")
                self.tabOverall.configure(background="#424242")
                self.tabInput.configure(background="#424242")
                self.tabModel.configure(background="#424242")
                self.tabError.configure(background="#424242")
                self.tabCustom.configure(background="#424242")
                self.defaultDirectoryLabel.configure(bg="#424242", fg="#FFFFFF")
                self.defaultDetrendLabel.configure(bg="#424242", fg="#FFFFFF")
                self.defaultTabLabel.configure(bg="#424242", fg="#FFFFFF")
                s = ttk.Style()
                s.configure('TCheckbutton', background='#424242', foreground="#FFFFFF")
                s.configure('TNotebook', background='#424242')
                self.topGUI.setThemeDark()
            
            self.topGUI.setBarColor(self.barColor)
            self.barColorLabel.configure(bg=self.barColor)
            self.topGUI.setHighlightColor(self.highlightColor)
            self.highlightColorLabel.configure(bg=self.highlightColor)
            self.topGUI.setActiveColor(self.activeColor)
            self.activeColorLabel.configure(bg=self.activeColor)
            self.topGUI.setEllipseColor(self.ellipseColor)
            self.defaultEllipseColorLabel.configure(bg=self.ellipseColor)
            self.topGUI.setCurrentDirectory(self.defaultDirectoryEntry.get())
            self.topGUI.setDefaultDirectory(self.defaultDirectoryEntry.get())
            self.topGUI.setInputExitAlert(self.inputExitAlertVariable.get())
            self.topGUI.setCustomFormulaExitAlert(self.customFormulaExitAlertVariable.get())
            self.topGUI.setFreqLoad(self.defaultFreqLoadVariable.get())
            self.topGUI.setFreqUndo(self.defaultFreqUndoVariable.get())
            self.topGUI.setFreqLoadCustom(self.customFreqLoadVariable.get())
            self.topGUI.setScroll(self.defaultScrollVariable.get())
        
        def resetDefaults():
            a = messagebox.askokcancel("Reset all settings?", "Are you sure you want to reset all settings to the defaults?")
            if (a):
                self.themeChosen = "light"
                self.barColor = "#333333"
                self.activeColor = "#737373"
                self.highlightColor = "#999999"
                self.defaultDetectCommentsCheckboxVariable.set(1)
                self.defaultCommentsVariable.set("0")
                self.defaultDetectDelimiterCheckboxVariable.set(1)
                self.defaultDelimiterVariable.set("Tab")
                self.defaultMCVariable.set("1000")
                self.defaultFitVariable.set("Complex")
                self.defaultWeightingVariable.set("Modulus")
                self.defaultAlphaVariable.set("1")
                self.defaultAlphaCheckboxVariable.set(0)
                self.defaultBetaCheckboxVariable.set(0)
                self.defaultReCheckboxVariable.set(0)
                self.defaultGammaCheckboxVariable.set(1)
                self.defaultDeltaCheckboxVariable.set(1)
                self.defaultErrorWeightingVariable.set("Variance")
                self.defaultMovingAverageVariable.set("5 point")
                self.defaultDetrendVariable.set("Off")
                self.ellipseColor = "#FF0000"
                self.defaultDirectoryEntry.configure(state="normal")
                self.defaultDirectoryEntry.delete(0,tk.END)
                self.defaultDirectoryEntry.insert(0, "C:\\")
                self.defaultDirectoryEntry.configure(state="readonly")
                self.inputExitAlertVariable.set(0)
                self.customFormulaExitAlertVariable.set(0)
                self.defaultTabVariable.set("Input file")
                self.defaultFreqUndoVariable.set(0)
                self.defaultFreqLoadVariable.set(0)
                self.customFreqLoadVariable.set(0)
                self.defaultScrollVariable.set(1)
                config = configparser.ConfigParser()
                config['settings'] = {'theme': self.themeChosen, 'bar': self.barColor, 'highlight': self.highlightColor, 'accent': self.activeColor, 'dir': "C:\\", 'tab': self.defaultTabVariable.get(), 'scroll': self.defaultScrollVariable.get()}
                config['input'] = {'detect': self.defaultDetectCommentsCheckboxVariable.get(),'comments': self.defaultCommentsVariable.get(), 'delimiter': self.defaultDelimiterVariable.get(), 'detectDelimiter': self.defaultDetectDelimiterCheckboxVariable.get()}
                config['model'] = {'mc': self.defaultMCVariable.get(), 'fit': self.defaultFitVariable.get(), 'weight': self.defaultWeightingVariable.get(), 'alpha': self.defaultAlphaVariable.get(), 'ellipse': self.ellipseColor, 'freqLoad': self.defaultFreqLoadVariable.get(), 'freqUndo': self.defaultFreqUndoVariable.get()}
                config['error'] = {'detrend': self.defaultDetrendVariable.get(), 'alphaError': self.defaultAlphaCheckboxVariable.get(), 'betaError': self.defaultBetaCheckboxVariable.get(), 'reError': self.defaultReCheckboxVariable.get(), 'gammaError': self.defaultGammaCheckboxVariable.get(), 'deltaError': self.defaultDeltaCheckboxVariable.get(), 'errorWeighting': self.defaultErrorWeightingVariable.get(), 'errorMA': self.defaultMovingAverageVariable.get()}
                config['custom'] = {'askCustomExit': self.customFormulaExitAlertVariable.get(), 'freqLoadCustom': self.customFreqLoadVariable.get()}
                if (os.path.exists(os.getenv('LOCALAPPDATA')+r"\MeasurementModel")):
                    with open(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini", 'w+') as configfile:
                        config.write(configfile)
                        configfile.close()
                else:
                    os.makedirs(os.getenv('LOCALAPPDATA')+r"\MeasurementModel")
                    with open(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini", 'w') as configfile:
                        config.write(configfile)
                        configfile.close()
                applySettings()

        def saveSettings():
            try:
                numMCDefault = int(self.defaultMCVariable.get())
                if (numMCDefault <= 0):
                    raise Exception
            except:
                messagebox.showerror("Value error", "Error 49:\nPlease enter a positive integer for number of simulations")
                return
            try:
                numAlphaDefault = float(self.defaultAlphaVariable.get())
                if (numAlphaDefault <= 0):
                    raise Exception
            except:
                messagebox.showerror("Value error", "Error 50:\nPlease enter a positive number for the assumed noise (α)")
                return
            try:
                numCommentDefault = int(self.defaultCommentsVariable.get())
                if (numCommentDefault < 0):
                    raise Exception
            except:
                messagebox.showerror("Value error", "Error 51:\nPlease enter a positive integer for number of comment lines")
                return
            try:
                config = configparser.ConfigParser()
                config['settings'] = {'theme': self.themeChosen, 'bar': self.barColor, 'highlight': self.highlightColor, 'accent': self.activeColor, 'dir': self.defaultDirectoryEntry.get(), 'tab': self.defaultTabVariable.get(), 'scroll': self.defaultScrollVariable.get()}
                config['input'] = {'detect': self.defaultDetectCommentsCheckboxVariable.get(), 'comments': numCommentDefault, 'delimiter': self.defaultDelimiterVariable.get(), 'detectDelimiter': self.defaultDetectDelimiterCheckboxVariable.get(), 'askInputExit': self.inputExitAlertVariable.get()}
                config['model'] = {'mc': numMCDefault, 'fit': self.defaultFitVariable.get(), 'weight': self.defaultWeightingVariable.get(), 'alpha': numAlphaDefault, 'ellipse': self.ellipseColor, 'freqLoad': self.defaultFreqLoadVariable.get(), 'freqUndo': self.defaultFreqUndoVariable.get()}
                config['error'] = {'detrend': self.defaultDetrendVariable.get(), 'alphaError': self.defaultAlphaCheckboxVariable.get(), 'betaError': self.defaultBetaCheckboxVariable.get(), 'reError': self.defaultReCheckboxVariable.get(), 'gammaError': self.defaultGammaCheckboxVariable.get(), 'deltaError': self.defaultDeltaCheckboxVariable.get(), 'errorWeighting': self.defaultErrorWeightingVariable.get(), 'errorMA': self.defaultMovingAverageVariable.get()}
                config['custom'] = {'askCustomExit': self.customFormulaExitAlertVariable.get(), 'freqLoadCustom': self.customFreqLoadVariable.get()}
                if (os.path.exists(os.getenv('LOCALAPPDATA')+r"\MeasurementModel")):
                    with open(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini", 'w+') as configfile:
                        config.write(configfile)
                        configfile.close()
                else:
                    os.makedirs(os.getenv('LOCALAPPDATA')+r"\MeasurementModel")
                    with open(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini", 'w') as configfile:
                        config.write(configfile)
                        configfile.close()
                applySettings()
                self.saveButton.configure(text="Saved")
                self.after(500, lambda: self.saveButton.configure(text="Save and Apply"))
            except:
                messagebox.showerror("File error", "Error 52:\nError in applying or saving settings")
        
        def defaultCommentsCommand():
            if (self.defaultDetectCommentsCheckboxVariable.get() == 1):
                self.defaultDetectDelimiterCheckbox.grid(column=0, row=2, pady=2)
            else:
                self.defaultDetectDelimiterCheckbox.grid_remove()            
#        styleNotebook = ttk.Style()
#        styleNotebook.theme_create( "yummy", parent="alt", settings={
#        "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0] } },
#        "TNotebook.Tab": {
#            "configure": {"padding": [5, 1], "background": "#d2ffd2" },
#            "map":       {"background": [("selected", "#dd0202")],
#                          "expand": [("selected", [1, 1, 1, 0])] } } } )
        
        def checkRe():
            if (self.defaultReCheckboxVariable.get() == 1):
                self.defaultBetaCheckboxVariable.set(1)
        
        def checkBeta():
            if (self.defaultBetaCheckboxVariable.get() == 0):
                self.defaultReCheckboxVariable.set(0)
        
        def findNewDirectory():
            folder = askdirectory(initialdir=self.topGUI.currentDirectory)
            folder_str = str(folder)
            if len(folder_str) == 0:
                pass
            else:
                self.defaultDirectoryEntry.configure(state="normal")
                self.defaultDirectoryEntry.delete(0,tk.END)
                self.defaultDirectoryEntry.insert(0, folder_str)
                self.defaultDirectoryEntry.configure(state="readonly")
        
        self.notebook = ttk.Notebook(self)
        self.tabOverall = tk.Frame(self.notebook, background=self.backgroundColor)
        self.tabInput = tk.Frame(self.notebook, background=self.backgroundColor)
        self.tabModel = tk.Frame(self.notebook, background=self.backgroundColor)
        self.tabError = tk.Frame(self.notebook, background=self.backgroundColor)
        self.tabCustom = tk.Frame(self.notebook, background=self.backgroundColor)
        self.notebook.add(self.tabOverall, text="Overall")
        self.notebook.add(self.tabInput, text="Input")
        self.notebook.add(self.tabModel, text="Model")
        self.notebook.add(self.tabError, text="Errors")
        self.notebook.add(self.tabCustom, text="Custom")
        self.notebook.grid(column=0, row=0, columnspan=2, sticky="W")
        
        #self.themeFrame = tk.Frame(self, bg=self.backgroundColor)
        self.themeLabel = tk.Label(self.tabOverall, text="Theme: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.themeFrame = tk.Frame(self.tabOverall, background=self.backgroundColor)
        if (self.topGUI.getTheme() == "dark"):
            self.lightLabel = tk.Label(self.themeFrame, text=" Light ", fg="black", bg="white", cursor="hand2", borderwidth=3, relief="raised")
            self.darkLabel = tk.Label(self.themeFrame, text=" Dark ", fg="white", bg="#424242", cursor="hand2", borderwidth=3, relief="sunken")
        elif (self.topGUI.getTheme() == "light"):
            self.lightLabel = tk.Label(self.themeFrame, text=" Light ", fg="black", bg="white", cursor="hand2", borderwidth=3, relief="sunken")
            self.darkLabel = tk.Label(self.themeFrame, text=" Dark ", fg="white", bg="#424242", cursor="hand2", borderwidth=3, relief="raised")

        self.lightLabel.bind("<Button-1>", lightTheme)
        self.darkLabel.bind("<Button-1>", darkTheme)
        self.themeLabel.grid(column=0, row=0, sticky="W", padx=2)
        self.lightLabel.grid(column=0, row=0)
        self.darkLabel.grid(column=1, row=0)
        self.themeFrame.grid(column=1, row=0)
        #self.themeFrame.grid(column=0, row=0, sticky="W")
        light_ttp = CreateToolTip(self.lightLabel, 'Light theme (white background, black text)')
        dark_ttp = CreateToolTip(self.darkLabel, 'Dark theme (dark grey background, white text)')
        
        #self.accentsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.accentsLabel = tk.Label(self.tabOverall, text="Accent theme colors: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.barThemeLabel = tk.Label(self.tabOverall, text="Bar color: ", fg=self.foregroundColor, bg=self.backgroundColor)
#        self.themeVariable = tk.StringVar(self, "Gray")
#        self.themeCombobox = ttk.Combobox(self.accentsFrame, textvariable=self.themeVariable, value=("Gray", "Spring green", "Light blue", "Sienna", "Slate blue", "Peach", "Orchid", "Custom"), exportselection=0, state="readonly", width=15)
#        self.barVariable = tk.StringVar(self, "Dark gray")
#        self.barCombobox = ttk.Combobox(self.accentsFrame, textvariable=self.barVariable, value=("White", "Pearl", "Lightest gray", "Lighter gray", "Light gray", "Gray", "Dark gray", "Darker gray", "Black"), exportselection=0, state="readonly", width=15)
        
        self.barLabel = tk.Label(self.tabOverall, text="Side bar color: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.barColorLabel = tk.Label(self.tabOverall, bg=self.barColor, borderwidth=2, relief="solid", cursor="hand2", width=5)
        self.barColorLabel.bind("<Button-1>", pickBarColor)
        self.highlightLabel = tk.Label(self.tabOverall, text="Highlight: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.highlightColorLabel = tk.Label(self.tabOverall, bg=self.highlightColor, borderwidth=2, relief="solid", cursor="hand2", width=5)
        self.highlightColorLabel.bind("<Button-1>", pickHighlightColor)
        self.activeLabel = tk.Label(self.tabOverall, text="Active: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.activeColorLabel = tk.Label(self.tabOverall, bg=self.activeColor, borderwidth=2, relief="solid", cursor="hand2", width=5)
        self.activeColorLabel.bind("<Button-1>", pickActiveColor)
        self.defaultTabLabel = tk.Label(self.tabOverall, text="Tab: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultTabVariable = tk.StringVar(self, "Input file")
        if (self.topGUI.getDefaultTab() == 1):
            self.defaultTabVariable.set("Measurement model")
        elif (self.topGUI.getDefaultTab() == 2):
            self.defaultTabVariable.set("Error file preparation")
        elif (self.topGUI.getDefaultTab() == 3):
            self.defaultTabVariable.set("Error analysis")
        elif (self.topGUI.getDefaultTab() == 4):
            self.defaultTabVariable.set("Custom fitting")
        elif (self.topGUI.getDefaultTab() == 5):
            self.defaultTabVariable.set("Settings")
        elif (self.topGUI.getDefaultTab() == 6):
            self.defaultTabVariable.set("Help and about")
        self.defaultTab = ttk.Combobox(self.tabOverall, textvariable=self.defaultTabVariable, width=20, value=("Input file", "Measurement model", "Error file preparation", "Error analysis", "Custom fitting", "Settings", "Help and about"), exportselection=0, state="readonly")
        self.defaultDirectoryLabel = tk.Label(self.tabOverall, text="Directory: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultDirectoryEntry = ttk.Entry(self.tabOverall, state="readonly", width=22)
        self.defaultDirectoryButton = ttk.Button(self.tabOverall, text="Browse...", command=findNewDirectory)
        self.defaultScrollVariable = tk.IntVar(self, self.topGUI.getScroll())
        self.defaultScroll = ttk.Checkbutton(self.tabOverall, variable=self.defaultScrollVariable, text="Change tab on scroll")
#        self.accentsLabel.grid(column=0, row=1, sticky="W")
#        self.themeCombobox.grid(column=1, row=0)
#        self.barThemeLabel.grid(column=0, row=1, sticky="W")
#        self.barCombobox.grid(column=1, row=1)
        self.barLabel.grid(column=0, row=2, sticky="W", pady=5)
        self.barColorLabel.grid(column=1, row=2)
        #self.highlightLabel.grid(column=1, row=3, sticky="W", pady=2)
        #self.highlightColorLabel.grid(column=2, row=3)
        #self.activeLabel.grid(column=1, row=4, sticky="W", pady=2)
        #self.activeColorLabel.grid(column=2, row=4)
        self.defaultTabLabel.grid(column=0, row=3, sticky="W")
        self.defaultTab.grid(column=1, row=3)
        self.defaultDirectoryLabel.grid(column=0, row=5, pady=5, sticky="W")
        self.defaultDirectoryEntry.grid(column=1, row=5, columnspan=2, sticky="W")
        self.defaultDirectoryButton.grid(column=3, row=5, padx=2, sticky="W")
        self.defaultScroll.grid(column=0, row=6, pady=2, sticky="W")
        #self.accentsFrame.grid(column=0, row=1, sticky="W", columnspan=2, pady=5)
        barColor_ttp = CreateToolTip(self.barColorLabel, 'Choose color of side and top bars')
        hightlightColor_ttp = CreateToolTip(self.highlightColorLabel, 'Choose color of tabs when hovered over')
        activeColor_ttp = CreateToolTip(self.activeColorLabel, 'Choose color of tabs when active')
        defaultTab_ttp = CreateToolTip(self.defaultTab, 'Which tab is opened first when program starts')
        directory_ttp = CreateToolTip(self.defaultDirectoryEntry, 'Default directory for files')
        directoryBtn_ttp = CreateToolTip(self.defaultDirectoryButton, 'Default directory for files')
        scroll_ttp = CreateToolTip(self.defaultScroll, 'Whether or not to use the scrollwheel to change tabs when mouse is over the navigation pane')
        
        self.defaultDirectoryEntry.configure(state="normal")
        self.defaultDirectoryEntry.delete(0,tk.END)
        self.defaultDirectoryEntry.insert(0, self.topGUI.getDefaultDirectory())
        self.defaultDirectoryEntry.configure(state="readonly")
        
        #self.inputLabelFrame = tk.LabelFrame(self, text="Defaults for File Tools", bg=self.backgroundColor, fg=self.foregroundColor, padx=2, pady=2)
        self.defaultDetectCommentsCheckboxVariable = tk.IntVar(self, self.topGUI.getDetectComments())
        self.defaultDetectCommentsCheckbox = ttk.Checkbutton(self.tabInput, variable=self.defaultDetectCommentsCheckboxVariable, text="Detect number of comment lines", command=defaultCommentsCommand)
        self.defaultComments = tk.Label(self.tabInput, text="Number of comment lines: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.defaultCommentsVariable = tk.StringVar(self, self.topGUI.getComments())
        self.defaultCommentsEntry = ttk.Entry(self.tabInput, textvariable=self.defaultCommentsVariable, width=6, exportselection=0)
        self.defaultDelimiter = tk.Label(self.tabInput, text="Delimiter: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.defaultDelimiterVariable = tk.StringVar(self, self.topGUI.getDelimiter())
        self.defaultDelimiterCombobox = ttk.Combobox(self.tabInput, textvariable=self.defaultDelimiterVariable, value=("Tab", "Space", ";", ",", "|", ":"), exportselection=0, state="readonly", width=6)
        self.defaultDetectDelimiterCheckboxVariable = tk.IntVar(self, self.topGUI.getDetectDelimiter())
        self.defaultDetectDelimiterCheckbox = ttk.Checkbutton(self.tabInput, variable=self.defaultDetectDelimiterCheckboxVariable, text="Detect delimiter")
        self.inputExitAlertVariable = tk.IntVar(self, self.topGUI.getInputExitAlert())
        self.inputExitAlertCheckbox = ttk.Checkbutton(self.tabInput, variable=self.inputExitAlertVariable, text="Alert on close if unsaved")
        #Columns and scaling, line frequency deletion
        self.defaultDetectCommentsCheckbox.grid(column=0, row=0, columnspan=2, sticky="W")
        self.defaultComments.grid(column=0, row=1, pady=2, sticky="W")
        self.defaultCommentsEntry.grid(column=1, row=1, sticky="W")
        self.defaultDetectDelimiterCheckbox.grid(column=0, row=2, pady=2, sticky="W")
        self.defaultDelimiter.grid(column=0, row=3, pady=2, sticky="W")
        self.defaultDelimiterCombobox.grid(column=1, row=3, sticky="W")
        self.inputExitAlertCheckbox.grid(column=0, row=4, sticky="W")
        #self.inputLabelFrame.grid(column=0, row=2, sticky="W", pady=5)
        comments_ttp = CreateToolTip(self.defaultDetectCommentsCheckbox, 'Whether or not the number of comment lines should be automatically detected by default')
        commentsEntry_ttp = CreateToolTip(self.defaultCommentsEntry, 'The default number of comment lines to ignore')
        delimiter_ttp = CreateToolTip(self.defaultDetectDelimiterCheckbox, 'Whether or not the delimiter should be automatically detected')
        delimiterEntry_ttp = CreateToolTip(self.defaultDelimiterCombobox, 'The default delimiter')
        inputExit_ttp = CreateToolTip(self.inputExitAlertCheckbox, 'Whether or not the program should alert before closing if a file has been loaded but not saved')
        
        #self.modelLabelFrame = tk.LabelFrame(self, text="Defaults for Measurement Model", bg=self.backgroundColor, fg=self.foregroundColor, padx=2, pady=2)
        self.defaultMC = tk.Label(self.tabModel, text="Number of simulations: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultMCVariable = tk.StringVar(self, self.topGUI.getMC())
        self.defaultMCEntry = ttk.Entry(self.tabModel, textvariable=self.defaultMCVariable, width=6, exportselection=0)
        self.defaultFit = tk.Label(self.tabModel, text="Fit type: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultFitVariable = tk.StringVar(self, self.topGUI.getFit())
        self.defaultFitCombobox = ttk.Combobox(self.tabModel, textvariable=self.defaultFitVariable, value=("Real", "Imaginary", "Complex"), exportselection=0, state="readonly", width=15)
        self.defaultWeighting = tk.Label(self.tabModel, text="Weighting type: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultWeightingVariable = tk.StringVar(self, self.topGUI.getWeight())
        self.defaultWeightingCombobox = ttk.Combobox(self.tabModel, textvariable=self.defaultWeightingVariable, value=("None", "Modulus", "Proportional", "Error model"), exportselection=0, state="readonly", width=15)
        self.defaultAlpha = tk.Label(self.tabModel, text="Assumed noise level (α): ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.defaultAlphaVariable = tk.StringVar(self, self.topGUI.getAlpha())
        self.defaultAlphaEntry = ttk.Entry(self.tabModel, textvariable=self.defaultAlphaVariable, width=6, exportselection=0)
        self.defaultEllipse = tk.Label(self.tabModel, text="Error lines/ellipse color: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultEllipseColorLabel = tk.Label(self.tabModel, bg=self.ellipseColor, borderwidth=2, relief="solid", cursor="hand2", width=5)
        self.defaultEllipseColorLabel.bind("<Button-1>", pickEllipseColor)
        self.defaultFreqLoadVariable = tk.IntVar(self, self.topGUI.getFreqLoad())
        self.defaultFreqUndoVariable = tk.IntVar(self, self.topGUI.getFreqUndo())
        self.defaultFreqLoad = ttk.Checkbutton(self.tabModel, variable=self.defaultFreqLoadVariable, text="Keep frequency range on loading new file")
        self.defaultFreqUndo = ttk.Checkbutton(self.tabModel, variable=self.defaultFreqUndoVariable, text="Undo frequency range with undo button")
        self.defaultMC.grid(column=0, row=0, sticky="W")
        self.defaultMCEntry.grid(column=1, row=0, sticky="W")
        self.defaultFit.grid(column=0, row=1, pady=2, sticky="W")
        self.defaultFitCombobox.grid(column=1, row=1, sticky="W")
        self.defaultWeighting.grid(column=0, row=2, sticky="W")
        self.defaultWeightingCombobox.grid(column=1, row=2, sticky="W")
        self.defaultAlpha.grid(column=0, row=3, pady=2, sticky="W")
        self.defaultAlphaEntry.grid(column=1, row=3, sticky="W")
        self.defaultEllipse.grid(column=0, row=4, pady=(2,0), sticky="W")
        self.defaultEllipseColorLabel.grid(column=1, row=4, sticky="W")
        self.defaultFreqLoad.grid(column=0, row=5, pady=(2,0), sticky="W")
        self.defaultFreqUndo.grid(column=0, row=6, pady=(2,0), sticky="W")
        #self.modelLabelFrame.grid(column=0, row=3, sticky="W", pady=5)
        mc_ttp = CreateToolTip(self.defaultMCEntry, 'The default number of Monte Carlo simulations')
        fit_ttp = CreateToolTip(self.defaultFitCombobox, 'The default fit type')
        weighting_ttp = CreateToolTip(self.defaultWeightingCombobox, 'The default fitting weighting')
        alpha_ttp = CreateToolTip(self.defaultAlphaEntry, 'The default assumed noise (multiplied by weighting)')
        ellipse_ttp = CreateToolTip(self.defaultEllipseColorLabel, 'The default ellipse color (used to denote confidence interval on Nyquist plots)')
        freqLoad_ttp = CreateToolTip(self.defaultFreqLoad, 'Whether or not the adjusted frequency range should be kept when a new file is loaded')
        freqUndo_ttp = CreateToolTip(self.defaultFreqUndo, 'Whether or not the adjusted frequency range should be changed to what it was previously when the Undo feature is used')
        
        #self.errorLabelFrame = tk.LabelFrame(self, text="Defaults for Error Fitting", bg=self.backgroundColor, fg=self.foregroundColor, padx=2, pady=2)
        self.defaultParamChoices = tk.Label(self.tabError, text="Error Model Parameters: ", fg=self.foregroundColor, bg=self.backgroundColor)
        self.defaultAlphaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorAlpha())
        self.defaultAlphaCheckbox = ttk.Checkbutton(self.tabError, variable=self.defaultAlphaCheckboxVariable, text="\u03B1")
        self.defaultBetaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorBeta())
        self.defaultBetaCheckbox = ttk.Checkbutton(self.tabError, variable=self.defaultBetaCheckboxVariable, command=checkBeta, text="\u03B2")
        self.defaultReCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorRe())
        self.defaultReCheckbox = ttk.Checkbutton(self.tabError, variable=self.defaultReCheckboxVariable, command=checkRe, text="Re")
        self.defaultGammaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorGamma())
        self.defaultGammaCheckbox = ttk.Checkbutton(self.tabError, variable=self.defaultGammaCheckboxVariable, text="\u03B3")
        self.defaultDeltaCheckboxVariable = tk.IntVar(self, self.topGUI.getErrorDelta())
        self.defaultDeltaCheckbox = ttk.Checkbutton(self.tabError, variable=self.defaultDeltaCheckboxVariable, text="\u03B4")
        self.defaultErrorWeighting = tk.Label(self.tabError, text="Variance: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.defaultErrorWeightingVariable = tk.StringVar(self, self.topGUI.getErrorWeighting())
        self.defaultErrorWeightingCombobox = ttk.Combobox(self.tabError, textvariable=self.defaultErrorWeightingVariable, value=("None", "Variance"), exportselection=0, state="readonly", width=15)
        self.defaultMovingAverage = tk.Label(self.tabError, text="Moving Average: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.defaultMovingAverageVariable = tk.StringVar(self, self.topGUI.getMovingAverage())
        self.defaultMovingAverageCombobox = ttk.Combobox(self.tabError, textvariable=self.defaultMovingAverageVariable, value=("None", "3 point", "5 point"), exportselection=0, state="readonly", width=15)
        self.defaultDetrendLabel = tk.Label(self.tabError, text="Detrend: ", bg=self.backgroundColor, fg=self.foregroundColor)
        self.defaultDetrendVariable = tk.StringVar(self, self.topGUI.getDetrend())
        self.defaultDetrendCombobox = ttk.Combobox(self.tabError, textvariable=self.defaultDetrendVariable, value=("Off", "On"), exportselection=0, state="readonly", width=15)
        self.defaultParamChoices.grid(column=0, row=0, sticky="W")
        self.defaultAlphaCheckbox.grid(column=1, row=0, sticky="W")
        self.defaultBetaCheckbox.grid(column=1, row=1, sticky="W")
        self.defaultReCheckbox.grid(column=2, row=1, sticky="W")
        self.defaultGammaCheckbox.grid(column=1, row=2, sticky="W")
        self.defaultDeltaCheckbox.grid(column=1, row=3, sticky="W")
        self.defaultErrorWeighting.grid(column=0, row=4, sticky="W")
        self.defaultErrorWeightingCombobox.grid(column=1, row=4, columnspan=2, sticky="W")
        self.defaultMovingAverage.grid(column=0, row=5, sticky="W", pady=(2,0))
        self.defaultMovingAverageCombobox.grid(column=1, row=5, columnspan=2, sticky="W")
        self.defaultDetrendLabel.grid(column=0, row=6, sticky="W")
        self.defaultDetrendCombobox.grid(column=1, row=6, sticky="W")
        #self.errorLabelFrame.grid(column=0, row=4, sticky="W", pady=5)
        errA_ttp = CreateToolTip(self.defaultAlphaCheckbox, 'Whether \u03B1 should be used in the error structure by default (Zj)')
        errB_ttp = CreateToolTip(self.defaultBetaCheckbox, 'Whether \u03B2 should be used in the error structure by default (Zr)')
        errB_ttp = CreateToolTip(self.defaultReCheckbox, 'Whether the ohmic resistance should be subtracted from Zr by default')
        errG_ttp = CreateToolTip(self.defaultGammaCheckbox, 'Whether \u03B3 should be used in the error structure by default |Z|\u00B2')
        errD_ttp = CreateToolTip(self.defaultDeltaCheckbox, 'Whether \u03B4 should be used in the error structure by default (constant)')
        errorWeighting_ttp = CreateToolTip(self.defaultErrorWeightingCombobox, 'Default weighting for error structure regression')
        ma_ttp = CreateToolTip(self.defaultMovingAverageCombobox, 'Default moving average choice for error structure regression (only if Variance weighting is chosen)')
        detrend_ttp = CreateToolTip(self.defaultDetrendCombobox, 'Whether or not detrending should be done by default')
        
        self.customFormulaExitAlertVariable = tk.IntVar(self, self.topGUI.getCustomFormulaExitAlert())
        self.customFormulaExitAlertCheckbox = ttk.Checkbutton(self.tabCustom, variable=self.customFormulaExitAlertVariable, text="Alert on close if unsaved")
        self.customFreqLoadVariable = tk.IntVar(self, self.topGUI.getFreqLoadCustom())
        self.customFreqLoad = ttk.Checkbutton(self.tabCustom, variable=self.customFreqLoadVariable, text="Keep frequency range on loading new file")
        self.customFormulaExitAlertCheckbox.grid(column=0, row=0, sticky="W")
        self.customFreqLoad.grid(column=0, row=1, pady=2, sticky="W")
        customAlert_ttp = CreateToolTip(self.customFormulaExitAlertCheckbox, 'Whether or not the program should alert prior to closing if an unsaved formula is present')
        customFreqLoad_ttp = CreateToolTip(self.customFreqLoad, 'Whether or not the adjusted frequency range should be kept when a new file is loaded')
        
        self.saveButton = ttk.Button(self, text="Save and Apply", width=20, command=saveSettings)
        self.saveButton.grid(column=0, row=5, sticky="W", pady=2)
        save_ttp = CreateToolTip(self.saveButton, 'Save and apply all settings')
        
        self.resetButton = ttk.Button(self, text="Reset Defaults", width=20, command=resetDefaults)
        self.resetButton.grid(column=1, row=5, sticky="W", pady=5, padx=2)
        reset_ttp = CreateToolTip(self.resetButton, 'Reset all settings to their default values')
    
    def applySettings(self):
        if (self.themeChosen == "light"):
            self.configure(bg="#FFFFFF")
            #self.themeFrame.configure(bg="#FFFFFF")
            self.themeLabel.configure(bg="#FFFFFF")
            self.themeLabel.configure(fg="#000000")
            self.themeFrame.configure(background="#FFFFFF")
            self.lightLabel.configure(relief="sunken")
            self.darkLabel.configure(relief="raised")
            #self.accentsFrame.configure(bg="#FFFFFF")
            self.accentsLabel.configure(bg="#FFFFFF")
            self.accentsLabel.configure(fg="#000000")
            self.barLabel.configure(bg="#FFFFFF")
            self.barLabel.configure(fg="#000000")
            self.highlightLabel.configure(bg="#FFFFFF")
            self.highlightLabel.configure(fg="#000000")
            self.activeLabel.configure(bg="#FFFFFF")
            self.activeLabel.configure(fg="#000000")
            self.defaultMC.configure(bg="#FFFFFF", fg="#000000")
            #self.inputLabelFrame.configure(bg="#FFFFFF", fg="#000000")
            self.defaultComments.configure(bg="#FFFFFF", fg="#000000")
            self.defaultDelimiter.configure(bg="#FFFFFF", fg="#000000")
            #self.modelLabelFrame.configure(bg="#FFFFFF", fg="#000000")
            self.defaultMC.configure(bg="#FFFFFF", fg="#000000")
            self.defaultFit.configure(bg="#FFFFFF", fg="#000000")
            self.defaultWeighting.configure(bg="#FFFFFF", fg="#000000")
            self.defaultAlpha.configure(bg="#FFFFFF", fg="#000000")
            #self.errorLabelFrame.configure(bg="#FFFFFF", fg="#000000")
            self.defaultParamChoices.configure(bg="#FFFFFF", fg="#000000")
            self.defaultErrorWeighting.configure(bg="#FFFFFF", fg="#000000")
            self.defaultMovingAverage.configure(bg="#FFFFFF", fg="#000000")
            self.defaultEllipse.configure(bg="#FFFFFF", fg="#000000")
            self.tabOverall.configure(background="#FFFFFF")
            self.tabInput.configure(background="#FFFFFF")
            self.tabModel.configure(background="#FFFFFF")
            self.tabError.configure(background="#FFFFFF")
            self.tabCustom.configure(background="#FFFFFF")
            self.defaultDirectoryLabel.configure(bg="#FFFFFF", fg="#000000")
            self.defaultDetrendLabel.configure(bg="#FFFFFF", fg="#000000")
            self.defaultTabLabel.configure(bg="#FFFFFF", fg="#000000")
            s = ttk.Style()
            s.configure('TCheckbutton', background='#FFFFFF', foreground="#000000")
            s.configure('TNotebook', background='#FFFFFF')
            self.topGUI.setThemeLight()
        elif (self.themeChosen == "dark"):
            self.configure(bg="#424242")
            #self.themeFrame.configure(bg="#424242")
            self.themeLabel.configure(bg="#424242")
            self.themeLabel.configure(fg="#FFFFFF")
            self.themeFrame.configure(background="#424242")
            self.lightLabel.configure(relief="raised")
            self.darkLabel.configure(relief="sunken")
            #self.accentsFrame.configure(bg="#424242")
            self.accentsLabel.configure(bg="#424242")
            self.accentsLabel.configure(fg="#FFFFFF")
            self.barLabel.configure(bg="#424242")
            self.barLabel.configure(fg="#FFFFFF")
            self.highlightLabel.configure(bg="#424242")
            self.highlightLabel.configure(fg="#FFFFFF")
            self.activeLabel.configure(bg="#424242")
            self.activeLabel.configure(fg="#FFFFFF")
            self.defaultMC.configure(bg="#424242", fg="#FFFFFF")
            #self.inputLabelFrame.configure(bg="#424242", fg="#FFFFFF")
            self.defaultComments.configure(bg="#424242", fg="#FFFFFF")
            self.defaultDelimiter.configure(bg="#424242", fg="#FFFFFF")
            #self.modelLabelFrame.configure(bg="#424242", fg="#FFFFFF")
            self.defaultMC.configure(bg="#424242", fg="#FFFFFF")
            self.defaultFit.configure(bg="#424242", fg="#FFFFFF")
            self.defaultWeighting.configure(bg="#424242", fg="#FFFFFF")
            self.defaultAlpha.configure(bg="#424242", fg="#FFFFFF")
            #self.errorLabelFrame.configure(bg="#424242", fg="#FFFFFF")
            self.defaultParamChoices.configure(bg="#424242", fg="#FFFFFF")
            self.defaultErrorWeighting.configure(bg="#424242", fg="#FFFFFF")
            self.defaultMovingAverage.configure(bg="#424242", fg="#FFFFFF")
            self.defaultEllipse.configure(bg="#424242", fg="#FFFFFF")
            self.tabOverall.configure(background="#424242")
            self.tabInput.configure(background="#424242")
            self.tabModel.configure(background="#424242")
            self.tabError.configure(background="#424242")
            self.tabCustom.configure(background="#424242")
            self.defaultDirectoryLabel.configure(bg="#424242", fg="#FFFFFF")
            self.defaultDetrendLabel.configure(bg="#424242", fg="#FFFFFF")
            self.defaultTabLabel.configure(bg="#424242", fg="#FFFFFF")
            s = ttk.Style()
            s.configure('TCheckbutton', background='#424242', foreground="#FFFFFF")
            s.configure('TNotebook', background='#424242')
            self.topGUI.setThemeDark()
        
        self.topGUI.setBarColor(self.barColor)
        self.barColorLabel.configure(bg=self.barColor)
        self.topGUI.setHighlightColor(self.highlightColor)
        self.highlightColorLabel.configure(bg=self.highlightColor)
        self.topGUI.setActiveColor(self.activeColor)
        self.activeColorLabel.configure(bg=self.activeColor)
        self.topGUI.setEllipseColor(self.ellipseColor)
        self.defaultEllipseColorLabel.configure(bg=self.ellipseColor)
        self.topGUI.setCurrentDirectory(self.defaultDirectoryEntry.get())
        self.topGUI.setDefaultDirectory(self.defaultDirectoryEntry.get())
        self.topGUI.setInputExitAlert(self.inputExitAlertVariable.get())
        self.topGUI.setCustomFormulaExitAlert(self.customFormulaExitAlertVariable.get())
        self.topGUI.setFreqLoad(self.defaultFreqLoadVariable.get())
        self.topGUI.setFreqUndo(self.defaultFreqUndoVariable.get())
        self.topGUI.setFreqLoadCustom(self.customFreqLoadVariable.get())
        self.topGUI.setScroll(self.defaultScrollVariable.get())
    
    def saveSettings(self, e=None):
        try:
            numMCDefault = int(self.defaultMCVariable.get())
            if (numMCDefault <= 0):
                raise Exception
        except:
            messagebox.showerror("Value error", "Error 49:\nPlease enter a positive integer for number of simulations")
            return
        try:
            numAlphaDefault = float(self.defaultAlphaVariable.get())
            if (numAlphaDefault <= 0):
                raise Exception
        except:
            messagebox.showerror("Value error", "Error 50:\nPlease enter a positive number for the assumed noise (α)")
            return
        try:
            numCommentDefault = int(self.defaultCommentsVariable.get())
            if (numCommentDefault < 0):
                raise Exception
        except:
            messagebox.showerror("Value error", "Error 51:\nPlease enter a positive integer for number of comment lines")
            return
        try:
            config = configparser.ConfigParser()
            config['settings'] = {'theme': self.themeChosen, 'bar': self.barColor, 'highlight': self.highlightColor, 'accent': self.activeColor, 'dir': self.defaultDirectoryEntry.get(), 'tab': self.defaultTabVariable.get(), 'scroll': self.defaultScrollVariable.get()}
            config['input'] = {'detect': self.defaultDetectCommentsCheckboxVariable.get(), 'comments': numCommentDefault, 'delimiter': self.defaultDelimiterVariable.get(), 'detectDelimiter': self.defaultDetectDelimiterCheckboxVariable.get(), 'askInputExit': self.inputExitAlertVariable.get()}
            config['model'] = {'mc': numMCDefault, 'fit': self.defaultFitVariable.get(), 'weight': self.defaultWeightingVariable.get(), 'alpha': numAlphaDefault, 'ellipse': self.ellipseColor, 'freqLoad': self.defaultFreqLoadVariable.get(), 'freqUndo': self.defaultFreqUndoVariable.get()}
            config['error'] = {'detrend': self.defaultDetrendVariable.get(), 'alphaError': self.defaultAlphaCheckboxVariable.get(), 'betaError': self.defaultBetaCheckboxVariable.get(), 'reError': self.defaultReCheckboxVariable.get(), 'gammaError': self.defaultGammaCheckboxVariable.get(), 'deltaError': self.defaultDeltaCheckboxVariable.get(), 'errorWeighting': self.defaultErrorWeightingVariable.get(), 'errorMA': self.defaultMovingAverageVariable.get()}
            config['custom'] = {'askCustomExit': self.customFormulaExitAlertVariable.get(), 'freqLoadCustom': self.customFreqLoadVariable.get()}
            if (os.path.exists(os.getenv('LOCALAPPDATA')+r"\MeasurementModel")):
                with open(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini", 'w+') as configfile:
                    config.write(configfile)
                    configfile.close()
            else:
                os.makedirs(os.getenv('LOCALAPPDATA')+r"\MeasurementModel")
                with open(os.getenv('LOCALAPPDATA')+r"\MeasurementModel\settings.ini", 'w') as configfile:
                    config.write(configfile)
                    configfile.close()
            self.applySettings()
            self.saveButton.configure(text="Saved")
            self.after(500, lambda: self.saveButton.configure(text="Save and Apply"))
        except:
            messagebox.showerror("File error", "Error 52:\nError in applying or saving settings")
    
    def bindIt(self, e=None):
        self.bind_all("<Control-s>", self.saveSettings)
    
    def unbindIt(self, e=None):
        self.unbind_all("<Control-s>")