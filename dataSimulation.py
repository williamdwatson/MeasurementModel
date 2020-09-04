# -*- coding: utf-8 -*-
"""
Created on Tue May 26 22:39:33 2020

@author: willd
"""


import numpy as np
from numpy import nan    #This can be a necessary import if errors occur

def run_simulations(w, formula, paramNames, paramValues):
    try:
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(len(paramNames)):
            exec(paramNames[i] + " = " + str(float(paramValues[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return Zreal, Zimag
    except:
        return "-", "-"