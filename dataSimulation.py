# -*- coding: utf-8 -*-
"""
Created on Tue Oct  8 16:19:24 2019

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

import numpy as np
from numpy import nan    #This can be a necessary import if errors occur

def run_simulations(w, formula, paramNames, paramValues):
    """
    Parameters
    ----------
    w : list of floats
        The frequencies at which to produce data
    
    formula : string
        The custom formula to execute as Python code
    
    paramNames : list of strings
        The parameter names
    
    paramValues : list of floats
        The values of the parameters
    
    Returns
    -------
    Zreal, Zimag : tuple
        Tuple of arrays of real and imaginary impedance data, respecively; or ('-', '-') in the case of an error
    
    """
    try:
        Zreal = []
        Zimag = []
        freq = w.copy()
        
        #---Assign each parameter name its value---
        for i in range(len(paramNames)):
            exec(paramNames[i] + " = " + str(float(paramValues[i])))
        ldict = locals()                    #Dictionary of variables used in exec
        exec(formula, globals(), ldict)     #Execute the code
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return Zreal, Zimag
    except Exception:
        return "-", "-"