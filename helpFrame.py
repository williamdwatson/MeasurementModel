# -*- coding: utf-8 -*-
"""
Created on Wed Jun 20 15:07:52 2018

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
import webbrowser, sys, os

class CreateToolTip(object):
    """
    create a tooltip for a given widget
    Code from: https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter with answer by crxguy52
    """
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 300   #pixels
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

class hF(tk.Frame):
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
        
        #---Appearance---
        if (self.topGUI.getTheme() == "dark"):
            self.backgroundColor = "#424242"
            self.foregroundColor = "white"
            self.linkColor = "skyblue"
        elif (self.topGUI.getTheme() == "light"):
            self.backgroundColor = "white"
            self.foregroundColor = "black"
            self.linkColor = "blue"
        tk.Frame.__init__(self, parent, background=self.backgroundColor)
        
        #---Links to other libraries---
        self.aboutCodeFrame = tk.Frame(self, background=self.backgroundColor)
        self.aboutCode = tk.Label(self.aboutCodeFrame, text="Modules used outside the", background=self.backgroundColor, fg=self.foregroundColor)
        self.aboutCodeLink = tk.Label(self.aboutCodeFrame, text=" standard library:", background=self.backgroundColor, fg=self.linkColor, cursor="hand2")
        self.aboutCodeLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://docs.python.org/3/library/"))
        aboutCodeLink_ttp = CreateToolTip(self.aboutCodeLink, 'https://docs.python.org/3/library/')
        self.detectLink = tk.Label(self, text="Delimiter detection: detect_delimiter")
        self.detectLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.detectLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://pypi.org/project/detect-delimiter/"))
        detect_ttp = CreateToolTip(self.detectLink, 'https://pypi.org/project/detect-delimiter/')
        self.copyLink = tk.Label(self, text="Copying to clipboard: Pyperclip")
        self.copyLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.copyLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://pypi.org/project/pyperclip/"))
        copy_ttp = CreateToolTip(self.copyLink, 'https://pypi.org/project/pyperclip/')
        self.lmLink = tk.Label(self, text="Fitting: LMFIT")
        self.lmLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.lmLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"http://doi.org/10.5281/zenodo.11813"))
        lm_ttp = CreateToolTip(self.lmLink, 'http://doi.org/10.5281/zenodo.11813')
        self.rsLink = tk.Label(self, text="Frequency adjuster: rangeSlider")
        self.rsLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.rsLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://github.com/halsafar/rangeslider"))
        rs_ttp = CreateToolTip(self.rsLink, 'https://github.com/halsafar/rangeslider')
        self.pilLink = tk.Label(self, text="Images: PIL (Pillow)")
        self.pilLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.pilLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://pillow.readthedocs.io/en/stable/"))
        pil_ttp = CreateToolTip(self.pilLink, 'https://pillow.readthedocs.io/en/stable/')
        self.numpyLink = tk.Label(self, text="Numerical manipulations: numpy")
        self.numpyLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.numpyLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://numpy.org/"))
        numpy_ttp = CreateToolTip(self.numpyLink, 'https://numpy.org/')
        self.comLink = tk.Label(self, text="COM: comtypes")
        self.comLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.comLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://pypi.org/project/comtypes/"))
        com_ttp = CreateToolTip(self.comLink, 'https://pypi.org/project/comtypes/')
        self.sciLink = tk.Label(self, text="Other fitting: scipy")
        self.sciLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.sciLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://www.scipy.org/"))
        scipy_ttp = CreateToolTip(self.sciLink, 'https://www.scipy.org/')
        self.pltLink = tk.Label(self, text="Graphing: matplotlib")
        self.pltLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.pltLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://matplotlib.org/"))
        matplotlib_ttp = CreateToolTip(self.pltLink, 'https://matplotlib.org/')
        self.impedanceLink = tk.Label(self, text="File detection: impedance.py")
        self.impedanceLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.impedanceLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://github.com/ECSHackWeek/impedance.py"))
        impedance_ttp = CreateToolTip(self.impedanceLink, 'https://github.com/ECSHackWeek/impedance.py')
        self.galvaniLink = tk.Label(self, text="MPR decoding: galvani")
        self.galvaniLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.galvaniLink.bind("<Button-1>", lambda e: webbrowser.open_new(r"https://pypi.org/project/galvani/"))
        galvani_ttp = CreateToolTip(self.galvaniLink, 'https://pypi.org/project/galvani/')
        
        #---Help and about links---
        self.helpLink = tk.Label(self, text="User manual")
        self.helpLink.bind('<Button-1>', lambda e: webbrowser.open_new('file://' + os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Measurement_model_guide.pdf')))
        self.helpLink.configure(background=self.backgroundColor, cursor="hand2", fg=self.linkColor)
        self.aboutUsFrame = tk.Frame(self, bg=self.backgroundColor)
        self.orazem = tk.Label(self.aboutUsFrame, text="Dr. Mark Orazem", bg=self.backgroundColor, fg=self.foregroundColor)
        self.orazemPhone = tk.Label(self.aboutUsFrame, text="Phone: 352-392-6207", bg=self.backgroundColor, fg=self.foregroundColor)
        self.orazemEmail = tk.Label(self.aboutUsFrame, text="Email: morazem@che.ufl.edu", cursor="hand2", bg=self.backgroundColor, fg=self.linkColor)
        self.orazemEmail.bind('<Button-1>', lambda e: webbrowser.open("mailto: morazem@che.ufl.edu"))
        self.william = tk.Label(self.aboutUsFrame, text="William Watson", bg=self.backgroundColor, fg=self.foregroundColor)
        self.williamEmail = tk.Label(self.aboutUsFrame, text="Email: williamdwatson1@gmail.com", cursor="hand2", bg=self.backgroundColor, fg=self.linkColor)
        self.williamEmail.bind('<Button-1>', lambda e: webbrowser.open("mailto: williamdwatson1@gmail.com"))
        self.orazem.grid(column=0, row=0, sticky="W")
        self.orazemPhone.grid(column=1, row=0, sticky="W", padx=3)
        self.orazemEmail.grid(column=2, row=0, sticky="W")
        self.william.grid(column=0, row=1, sticky="W", pady=3)
        self.williamEmail.grid(column=1, row=1, sticky="W", columnspan=2, padx=3)
        help_ttp = CreateToolTip(self.helpLink, 'Open a PDF user manual')
        
        #---Copyright information---
        self.copyrightLabel = tk.Label(self, text="©Copyright 2020 University of Florida Research Foundation, Inc. All Rights Reserved.\n\nThis program is free software: you can redistribute it and/or modify \
it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.\n\n\
\
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the \
GNU General Public License for more details.\n\n\
You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.", justify=tk.LEFT, wraplength=600, bg=self.backgroundColor, fg=self.foregroundColor)
        
        self.gnuLicenseLink = tk.Label(self, text="GNU General Public License", cursor="hand2", bg=self.backgroundColor, fg=self.linkColor)
        self.gnuLicenseLink.bind('<Button-1>', lambda e: webbrowser.open_new("gpl-3.0-standalone.html"))
        gnuLicense_ttp = CreateToolTip(self.gnuLicenseLink, 'GNU General Public License')
        self.aboutCodeFrame.grid(column=0, row=1, columnspan=3, sticky="W")
        self.aboutCode.grid(column=0, row=0, sticky="W")
        self.aboutCodeLink.grid(column=1, row=0, sticky="W")
        self.detectLink.grid(column=0, row=2, sticky="W")
        self.copyLink.grid(column=0, row=3, sticky="W")
        self.galvaniLink.grid(column=0, row=4, sticky="W")
        self.impedanceLink.grid(column=0, row=5, sticky="W")
        self.lmLink.grid(column=0, row=6, sticky="W")
        self.rsLink.grid(column=0, row=7, sticky="W")
        self.pilLink.grid(column=0, row=8, sticky="W")
        self.numpyLink.grid(column=0, row=9, sticky="W")
        self.comLink.grid(column=0, row=10, sticky="W")
        self.sciLink.grid(column=0, row=11, sticky="W")
        self.pltLink.grid(column=0, row=12, sticky="W")
        self.helpLink.grid(column=0, row=0, sticky="W", pady=7)
        self.aboutUsFrame.grid(column=0, row=13, sticky="W", pady=7)
        self.copyrightLabel.grid(column=0, row=14, sticky="W")
        self.gnuLicenseLink.grid(column=0, row=15, sticky="W", pady=7)
    
    def setThemeLight(self):
        """
        Change the appearance to light theme

        Returns
        -------
        None

        """
        self.configure(bg="#FFFFFF")
        self.aboutUsFrame.configure(background="#FFFFFF")
        self.orazem.configure(background="#FFFFFF", foreground="#000000")
        self.orazemPhone.configure(background="#FFFFFF", foreground="#000000")
        self.orazemEmail.configure(background="#FFFFFF", foreground="blue")
        self.william.configure(background="#FFFFFF", foreground="#000000")
        self.williamEmail.configure(background="#FFFFFF", foreground="blue")
        self.helpLink.configure(background="#FFFFFF", foreground="blue")
        self.aboutCode.configure(background="#FFFFFF", foreground="#000000")
        self.copyrightLabel.configure(background="#FFFFFF", foreground="#000000")
        self.aboutCodeFrame.configure(background="#FFFFFF")
        self.aboutCodeLink.configure(background="#FFFFFF", foreground="blue")
        self.detectLink.configure(background="#FFFFFF", foreground="blue")
        self.copyLink.configure(background="#FFFFFF", foreground="blue")
        self.galvaniLink.configure(background="#FFFFFF", foreground="blue")
        self.impedanceLink.configure(background="#FFFFFF", foreground="blue")
        self.lmLink.configure(background="#FFFFFF", foreground="blue")
        self.rsLink.configure(background="#FFFFFF", foreground="blue")
        self.pilLink.configure(background="#FFFFFF", foreground="blue")
        self.numpyLink.configure(background="#FFFFFF", foreground="blue")
        self.comLink.configure(background="#FFFFFF", foreground="blue")
        self.sciLink.configure(background="#FFFFFF", foreground="blue")
        self.pltLink.configure(background="#FFFFFF", foreground="blue")
        self.gnuLicenseLink.configure(background="#FFFFFF", foreground="blue")
                                  
    def setThemeDark(self):
        """
        Change the appearance to dark theme
        
        Returns
        -------
        None
        
        """
        self.configure(bg="#424242")
        self.aboutUsFrame.configure(background="#424242")
        self.orazem.configure(background="#424242", foreground="#FFFFFF")
        self.orazemPhone.configure(background="#424242", foreground="#FFFFFF")
        self.orazemEmail.configure(background="#424242", foreground="skyblue")
        self.william.configure(background="#424242", foreground="#FFFFFF")
        self.williamEmail.configure(background="#424242", foreground="skyblue")
        self.helpLink.configure(background="#424242", foreground="skyblue")
        self.aboutCode.configure(background="#424242", foreground="#FFFFFF")
        self.copyrightLabel.configure(background="#424242", foreground="#FFFFFF")
        self.aboutCodeFrame.configure(background="#424242")
        self.aboutCodeLink.configure(background="#424242", foreground="skyblue")
        self.detectLink.configure(background="#424242", foreground="skyblue")
        self.copyLink.configure(background="#424242", foreground="skyblue")
        self.galvaniLink.configure(background="#424242", foreground="skyblue")
        self.impedanceLink.configure(background="#424242", foreground="skyblue")
        self.lmLink.configure(background="#424242", foreground="skyblue")
        self.rsLink.configure(background="#424242", foreground="skyblue")
        self.pilLink.configure(background="#424242", foreground="skyblue")
        self.numpyLink.configure(background="#424242", foreground="skyblue")
        self.comLink.configure(background="#424242", foreground="skyblue")
        self.sciLink.configure(background="#424242", foreground="skyblue")
        self.pltLink.configure(background="#424242", foreground="skyblue")
        self.gnuLicenseLink.configure(background="#424242", foreground="skyblue")