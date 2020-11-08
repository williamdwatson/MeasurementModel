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
from numpy.random import normal
import lmfit
from numpy import nan    #This can be a necessary import if errors occur
import itertools, threading, sys
import multiprocessing as mp

def mp_complex(guesses, sharedList, index, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams):
    """
    Function for multiprocessing a complex fit
    
    Parameters
    ----------
    guesses : list of floats
        The parameter initial guesses.
    sharedList : list
        List of succesful parameter fits.
    index : int
        Which index of the shared list to use
    parameters : lmfit Parameters
        The lmfit Parameters used when fitting
    numParams : int
        Number of parameters.
    weight : int
        Which weighting to use when fitting; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
    assumedNoise : float
        The assumed noise used when weighting.    
    formula : string
        The custom Python code.
    w : array of floats
        Data frequencies.
    ZrIn : array of floats
        Data real impedance.
    ZjIn : array of floats
        Data imaginary impedance.
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    percentVal : multiprocessing Manager int
        Current percent completion.
    paramNames : list of strings
        Parameter names.    
    fitType : int
        Which fit type to use.
    errorParams : list of floats
        The parameters used for error model weighting.
    
    Returns
    -------
    None (results are appended to a shared list)

    """
    def diffComplex(params):
        """
        The objective function for complex fitting.
        
        Parameters
        ----------
        params : lmfit Parameters
            The parameters for the current fitting iteration
        
        yVals : The data values
        
        Returns
        -------
        residuals : list/array
            Array of the residuals (the known values minus the model values over the weighting)
            
        """
        return (Z_append - model(params))/weightCode(params)
    def model(p):
        """
        The model used in complex fitting

        Parameters
        ----------
        p : lmfit Parameters
            The parameters for the current fitting iteration.

        Returns
        -------
        model : numpy array
            Numpy array of the real impedance, then the imaginary impedance appended to that.

        """
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        #---Assign variable names their values, then exec the Python string---
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return np.append(Zreal, Zimag)
    def weightCode(p):
        """
        Calculate weighting used in regression for complex fits

        Parameters
        ----------
        p : lmfit Parameters
            The parameters to be fit.

        Returns
        -------
        weight : numpy array
            The weighting for regression, as a single array; the first half will be the real weighting, with the imaginary weighting appended at the end.
            For proportional weighting, the weighting is a repeat, with the real/imaginary weighting first, then the real/imaginary weighting appended again.

        """
        V = np.ones(len(w))                 #Default - no weighting
        if (weight == 1):                   #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2):                 #Proportional weighting
            Vj = np.ones(len(w))
            for i in range(len(V)):
                if (fitType != 1):          #If the fit type isn't imaginary use both Zr and Zj
                    V[i] = assumedNoise*ZrIn[i]
                else:                       #If the fit is imaginary use only Zj in weighting
                    V[i] = assumedNoise*ZjIn[i]
                Vj[i] = assumedNoise*ZjIn[i]
        elif (weight == 3):                 #Error structure weighting
            for i in range(len(V)):
                V[i] = errorParams[0]*ZjIn[i] + (errorParams[1]*ZrIn[i] - errorParams[2]) + errorParams[3]*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2) + errorParams[4]
        elif (weight == 4):                 #Custom weighting
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        if (weight == 2):
            return np.append(V, Vj)
        else:
            return np.append(V, V)
    current_best_val = 1E308                                            #Holds the current minimum chi-squared statistic
    current_best_params = guesses[0]                                    #Holds the current minimum parameters
    current_best_result = 0                                             #Holds the lmfit result for the current minimum
    for i in range(len(guesses)):
        with percentVal.get_lock():                                     #Multiprocessing-safe increment of percent completion
            percentVal.value += 1
        for j in range(numParams):                                      #Set the parameters to their appropriate initial guesses
            parameters.get(paramNames[j]).set(value=guesses[i][j])
        minimized = lmfit.minimize(diffComplex, parameters)             #Perform the fitting
        #---After the fitting, see if it's better than previous fittings, and if so update the variables---
        if (minimized.chisqr < current_best_val):
            current_best_val = minimized.chisqr
            current_best_params = guesses
            current_best_result = minimized
    sharedList[index] = [current_best_val, current_best_result, current_best_params]    #Report the best result
        
def mp_real(guesses, sharedList, index, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams):
    """
    Function for multiprocessing a real fit
    
    Parameters
    ----------
    guesses : list of floats
        The parameter initial guesses.
    sharedList : list
        List of succesful parameter fits.
    index : int
        Which index of the shared list to use
    parameters : lmfit Parameters
        The lmfit Parameters used when fitting
    numParams : int
        Number of parameters.
    weight : int
        Which weighting to use when fitting; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
    assumedNoise : float
        The assumed noise used when weighting.    
    formula : string
        The custom Python code.
    w : array of floats
        Data frequencies.
    ZrIn : array of floats
        Data real impedance.
    ZjIn : array of floats
        Data imaginary impedance.
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    percentVal : multiprocessing Manager int
        Current percent completion.
    paramNames : list of strings
        Parameter names.    
    fitType : int
        Which fit type to use.
    errorParams : list of floats
        The parameters used for error model weighting.
    
    Returns
    -------
    None (results are appended to a shared list)

    """
    def diffReal(params):
        """
        The objective function for real fitting.
        
        Parameters
        ----------
        params : lmfit Parameters
            The parameters for the current fitting iteration
        
        yVals : The data values
        
        Returns
        -------
        residuals : list/array
            Array of the residuals (the known values minus the model values over the weighting)
            
        """
        return (ZrIn - modelReal(params))/weightCodeHalf(params)
    def weightCodeHalf(p):
        """
        Calculate weighting used in regression for real fits

        Parameters
        ----------
        p : lmfit Parameters
            The parameters to be fit.

        Returns
        -------
        weight : numpy array
            The weighting for regression, as a single array. This array is not appended (unlike weightCode), and so is half the length.

        """
        V = np.ones(len(w))             #Default (weight == 0) is no weighting (weights are 1)
        if (weight == 1):               #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2):             #Proportional weighting
            for i in range(len(V)):
                V[i] = assumedNoise*ZrIn[i]
        elif (weight == 3):             #Error structure weighting
            pass
        elif (weight == 4):             #Custom weighting
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        return V
    def modelReal(p):
        """
        The model used in real fitting

        Parameters
        ----------
        p : lmfit Parameters
            The parameters for the current fitting iteration.

        Returns
        -------
        model : list/array
            Array of the real impedance

        """
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        #---Assign variable names their values, then exec the Python string---
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return Zreal
    current_best_val = 1E308                                            #Holds the current minimum chi-squared statistic
    current_best_params = guesses[0]                                    #Holds the current minimum parameters
    current_best_result = 0                                             #Holds the lmfit result for the current minimum
    for i in range(len(guesses)):
        with percentVal.get_lock():                                     #Multiprocessing-safe increment of percent completion
            percentVal.value += 1
        for j in range(numParams):                                      #Set the parameters to their appropriate initial guesses
            parameters.get(paramNames[j]).set(value=guesses[i][j])
        minimized = lmfit.minimize(diffReal, parameters)                #Perform the fitting
        #---After the fitting, see if it's better than previous fittings, and if so update the variables---
        if (minimized.chisqr < current_best_val):
            current_best_val = minimized.chisqr
            current_best_params = guesses
            current_best_result = minimized
    sharedList[index] = [current_best_val, current_best_result, current_best_params]    #Report the best result

def mp_imag(guesses, sharedList, index, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams):
    """
    Function for multiprocessing a imaginary fit
    
    Parameters
    ----------
    guesses : list of floats
        The parameter initial guesses.
    sharedList : list
        List of succesful parameter fits.
    index : int
        Which index of the shared list to use
    parameters : lmfit Parameters
        The lmfit Parameters used when fitting
    numParams : int
        Number of parameters.
    weight : int
        Which weighting to use when fitting; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
    assumedNoise : float
        The assumed noise used when weighting.    
    formula : string
        The custom Python code.
    w : array of floats
        Data frequencies.
    ZrIn : array of floats
        Data real impedance.
    ZjIn : array of floats
        Data imaginary impedance.
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    percentVal : multiprocessing Manager int
        Current percent completion.
    paramNames : list of strings
        Parameter names.    
    fitType : int
        Which fit type to use.
    errorParams : list of floats
        The parameters used for error model weighting.
    
    Returns
    -------
    None (results are appended to a shared list)

    """
    
    def diffImag(params):
        """
        The objective function for imaginary fitting.
        
        Parameters
        ----------
        params : lmfit Parameters
            The parameters for the current fitting iteration
        
        yVals : The data values
        
        Returns
        -------
        residuals : list/array
            Array of the residuals (the known values minus the model values over the weighting)
            
        """
        return (ZjIn - modelImag(params))/weightCodeHalf(params)
    def weightCodeHalf(p):
        """
        Calculate weighting used in regression for imaginary fits

        Parameters
        ----------
        p : lmfit Parameters
            The parameters to be fit.

        Returns
        -------
        weight : numpy array
            The weighting for regression, as a single array. This array is not appended (unlike weightCode), and so is half the length.

        """
        V = np.ones(len(w))             #Default (weight == 0) is no weighting (weights are 1)
        if (weight == 1):               #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2):             #Proportional weighting
            for i in range(len(V)):
                V[i] = assumedNoise*ZjIn[i]
        elif (weight == 3):             #Error structure weighting
            pass
        elif (weight == 4):             #Custom weighting
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        return V
    def modelImag(p):
        """
        The model used in imaginary fitting

        Parameters
        ----------
        p : lmfit Parameters
            The parameters for the current fitting iteration.

        Returns
        -------
        model : list/array
            Array of the imaginary impedance

        """
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        #---Assign variable names their values, then exec the Python string---
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return Zimag
    current_best_val = 1E308                                            #Holds the current minimum chi-squared statistic
    current_best_params = guesses[0]                                    #Holds the current minimum parameters
    current_best_result = 0                                             #Holds the lmfit result for the current minimum
    for i in range(len(guesses)):
        with percentVal.get_lock():                                     #Multiprocessing-safe increment of percent completion
            percentVal.value += 1
        for j in range(numParams):                                      #Set the parameters to their appropriate initial guesses
            parameters.get(paramNames[j]).set(value=guesses[i][j])
        minimized = lmfit.minimize(diffImag, parameters)                #Perform the fitting
        #---After the fitting, see if it's better than previous fittings, and if so update the variables---
        if (minimized.chisqr < current_best_val):
            current_best_val = minimized.chisqr
            current_best_params = guesses
            current_best_result = minimized
    sharedList[index] = [current_best_val, current_best_result, current_best_params]    #Report the best result
    
class customFitter:
    def __init__(self):
        self.keepGoing = True
        self.processes = []
        
    def terminateProcesses(self):
        """
        Kill each currently running process
        
        Returns
        -------
        None
        
        """
        self.keepGoing = False
        for process in self.processes:
            process.terminate()
    
    def findFit(self, extra_imports, listPercent, fitType, numMonteCarlo, weight, assumedNoise, formula, w, ZrIn, ZjIn, paramNames, paramGuesses, paramBounds, errorParams):
        """
        Regress the parameters for a custom fit. The custom code is run using the exec command.

        Parameters
        ----------
        extra_imports : list of strings
            Extra import paths.
        listPercent : list of ints
            List used to communicate precent completion for the GUI progress bar.
        fitType : int
            Fit type; 1 is real, 2 is imaginary, 3 is complex.
        numMonteCarlo : int
            Number of Monte Carlo calculations to perform.        
        weight : int
            Weighting type; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
        assumedNoise : float
            Assumed noise for the weighting.
        formula : string
            Custom Python code.
        w : array of floats
            Data frequency.
        ZrIn : array of floats
            Data real impedance.
        ZjIn : array of floats
            Data imaginary impedance.
        paramNames : list of strings
            Parameter names.
        paramGuesses : list of floats
            Parameter initial guesses.
        paramBounds : list of strings
            Parameter bounds; '-' for negative, '+' for positive, 'f' for fixed.
        errorParams : list of floats
            Parameters for error weighting.

        Returns
        -------
        result : array of floats
            Parameter fits.
        sigma : array of floats
            Parameter standard errors
        standardDevsReal : array of floats
            Real standard deviations
        standardDevsImag : array of floats
            Imaginary standard deviations
        minimized.chisqr : float
            Chi-squared statistic of fit
        minimized.aic : float
            Akaike Information Criterion of fit
        ToReturnReal : array of floats
            Real model values
        ToReturnImag : array of floats
            Imaginary model values
        current_best_params : list
            Best parameter guesses (for use if confidence interval bootstrapping is necessary later)
        weightingToReturn : array of floats
            Weighting at each frequency

        """
        
        np.random.seed(1234)                    #Use constant seed to ensure reproducibility of Monte Carlo simulations
        Z_append = np.append(ZrIn, ZjIn)
        numParams = len(paramNames)
        #---Add extra imports to the search path of Python; this allows the custom Python code to import new modules---
        for extra in extra_imports:
            sys.path.append(extra)
        def diffComplex(params):
            """
            The objective function for complex fitting.
            
            Parameters
            ----------
            params : lmfit Parameters
                The parameters for the current fitting iteration
            
            Returns
            -------
            residuals : list/array
                Array of the residuals (the known values minus the model values over the weighting)
            
            """
            if (not self.keepGoing):            #Check if the fitting has been cancelled
                return
            return (Z_append - model(params))/weightCode(params)
        
        def diffImag(params):
            """
            The objective function for imaginary fitting.
            
            Parameters
            ----------
            params : lmfit Parameters
                The parameters for the current fitting iteration
            
            Returns
            -------
            residuals : list/array
                Array of the residuals (the known values minus the model values over the weighting)
            
            """
            if (not self.keepGoing):            #Check if the fitting has been cancelled
                return
            return (ZjIn - modelImag(params))/weightCodeHalf(params)
        
        def diffReal(params):
            """
            The objective function for real fitting.
            
            Parameters
            ----------
            params : lmfit Parameters
                The parameters for the current fitting iteration
            
            Returns
            -------
            residuals : list/array
                Array of the residuals (the known values minus the model values over the weighting)
            
            """
            if (not self.keepGoing):            #Check if the fitting has been cancelled
                return
            return (ZrIn - modelReal(params))/weightCodeHalf(params)
        
        def weightCode(p):
            """
            Calculate weighting used in regression for complex fits

            Parameters
            ----------
            p : lmfit Parameters
                The parameters to be fit.

            Returns
            -------
            weight : numpy array
                The weighting for regression, as a single array; the first half will be the real weighting, with the imaginary weighting appended at the end.
                For proportional weighting, the weighting is a repeat, with the real/imaginary weighting first, then the real/imaginary weighting appended again.

            """
            V = np.ones(len(w))                 #Default - no weighting
            if (weight == 1):                   #Modulus weighting
                for i in range(len(V)):
                    V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
            elif (weight == 2):                 #Proportional weighting
                Vj = np.ones(len(w))
                for i in range(len(V)):
                    if (fitType != 1):          #If the fit type isn't imaginary use both Zr and Zj
                        V[i] = assumedNoise*ZrIn[i]
                    else:                       #If the fit is imaginary use only Zj in weighting
                        V[i] = assumedNoise*ZjIn[i]
                    Vj[i] = assumedNoise*ZjIn[i]
            elif (weight == 3):                 #Error structure weighting
                for i in range(len(V)):
                    V[i] = errorParams[0]*ZjIn[i] + (errorParams[1]*ZrIn[i] - errorParams[2]) + errorParams[3]*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2) + errorParams[4]
            elif (weight == 4):                 #Custom weighting
                p.valuesdict()
                Zr = ZrIn.copy()
                Zj = ZjIn.copy()
                Zreal = []
                Zimag = []
                weighting = []
                freq = w.copy()
                #---Assign variable names their values, then exec the Python string---
                for i in range(numParams):
                    exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
                ldict = locals()
                exec(formula, globals(), ldict)
                V = ldict['weighting']
            if (weight == 2):
                return np.append(V, Vj)
            else:
                return np.append(V, V)
        
        def weightCodeHalf(p):
            """
            Calculate weighting used in regression for real or imaginary fits

            Parameters
            ----------
            p : lmfit Parameters
                The parameters to be fit.

            Returns
            -------
            weight : numpy array
                The weighting for regression, as a single array. This array is not appended (unlike weightCode), and so is half the length.

            """
            V = np.ones(len(w))             #Default (weight == 0) is no weighting (weights are 1)
            if (weight == 1):               #Modulus weighting
                for i in range(len(V)):
                    V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
            elif (weight == 2):             #Proportional weighting
                for i in range(len(V)):
                    if (fitType != 1):
                        V[i] = assumedNoise*ZrIn[i]
                    else:
                        V[i] = assumedNoise*ZjIn[i]
            elif (weight == 3):             #Error structure weighting
                pass
            elif (weight == 4):             #Custom weighting
                p.valuesdict()
                Zr = ZrIn.copy()
                Zj = ZjIn.copy()
                Zreal = []
                Zimag = []
                weighting = []
                freq = w.copy()
                #---Assign variable names their values, then exec the Python string---
                for i in range(numParams):
                    exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
                ldict = locals()
                exec(formula, globals(), ldict)
                V = ldict['weighting']
            return V
        
        def model(p):
            """
            The model used in complex fitting

            Parameters
            ----------
            p : lmfit Parameters
                The parameters for the current fitting iteration.

            Returns
            -------
            model : numpy array
                Numpy array of the real impedance, then the imaginary impedance appended to that.

            """
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            return np.append(Zreal, Zimag)
        
        def modelImag(p):
            """
            The model used in imaginary fitting

            Parameters
            ----------
            p : lmfit Parameters
                The parameters for the current fitting iteration.

            Returns
            -------
            model : list/array
                Array of the imaginary impedance

            """
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            return Zimag
        
        def modelReal(p):
            """
            The model used in real fitting

            Parameters
            ----------
            p : lmfit Parameters
                The parameters for the current fitting iteration.

            Returns
            -------
            model : list/array
                Array of the real impedance

            """
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            return Zreal
        
        if (not self.keepGoing):                                        #Check if the fitting has been cancelled yet
            return
        
        #---Create the lmfit Parameters---
        parameters = lmfit.Parameters()
        parameterCorrector = 0
        for i in range(numParams):
            parameters.add(paramNames[i], value=paramGuesses[i][0])
            if (paramBounds[i] == "-"):                                 #Negative parameters
                parameters[paramNames[i]].max = 0
            elif (paramBounds[i] == "+"):                               #Positive parameters
                parameters[paramNames[i]].min = 0
            elif (paramBounds[i] == "f"):                               #Fixed parameters
                parameters[paramNames[i]].vary = False
            elif (len(paramBounds[i]) > 1):                             #Custom bounds
                upper, lower = paramBounds[i].split(";")
                if (upper != "inf"):                                    #Set non-infinite upper bounds; otherwise, keep the default (no upper bound)
                    parameters[paramNames[i]].max = float(upper)
                if (lower != "-inf"):
                    parameters[paramNames[i]].min = float(lower)        #Set non-infinite lower bounds; otherwise, keep the default (no lower bound)
        
        if (not self.keepGoing):                                        #Check if the fitting has been cancelled yet
            return

        try:
            guesses = list(itertools.product(*paramGuesses))            #Create combinations of all initial guesses
            current_best_val = 1E308
            current_best_params = guesses[0]
            current_best_result = 0
            #---If the number of parameters is low enough, don't use multiprocessing---
            if (numParams * len(guesses) < 1000):
                #---Loop through each combination of initial guesses and see if the regression result from those is better than the previous best---
                for i in range(len(guesses)):
                    for j in range(numParams):
                        parameters.get(paramNames[j]).set(value=guesses[i][j])
                    if (fitType == 3):      #Complex fitting
                        minimized = lmfit.minimize(diffComplex, parameters)
                    elif (fitType == 2):    #Imaginary fitting
                        minimized = lmfit.minimize(diffImag, parameters)
                    elif (fitType == 1):    #Real fitting
                        minimized = lmfit.minimize(diffReal, parameters)
                    #---Compare the current regression to the previous best, and set variables if this one is better---
                    if (minimized.chisqr < current_best_val):
                        current_best_val = minimized.chisqr
                        current_best_params = guesses
                        current_best_result = minimized
                    listPercent.append(1)                               #Communicate percent completion to the GUI
            #---If the number of parameters is high enough, use multiprocessing for increased speed at the cost of high overhead---
            else:
                numCores = mp.cpu_count()                               #Find the number of logical CPUs
                splitArray = np.array_split(guesses, numCores)          #Split the intial guess array so a roughly even number run on each CPU
                #---Prepare the multiprocessing Manager to share values between processes---
                manager = mp.Manager()
                vals = manager.list()
                percentVal = mp.Value("i", 0)                            #The percent completion, for use by the GUI
                for i in range(numCores):
                    vals.append([])
                if (fitType == 3):                                       #Complex fit
                    for i in range(numCores):
                        if (not self.keepGoing):                         #Check if the fitting has been cancelled
                            return
                        #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                        self.processes.append(mp.Process(target=mp_complex, args=(splitArray[i], vals, i, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams)))
                        listPercent.append(2)
                    for p in self.processes:
                        if (not self.keepGoing):                         #Check if the fitting has been cancelled
                            return
                        p.start()
                        listPercent.append(2)
                    def checkVal():
                        """Self-calling thread which checks processes progress"""
                        #---Check percent completion from shared variable, then append to list to communicate with GUI---
                        for i in range(percentVal.value+1 - len(listPercent)):
                            listPercent.append(1)
                        #---Initiate self-calling every 0.1 seconds while not at 100% completion---
                        if (percentVal.value < listPercent[0]):
                            timer = threading.Timer(0.1, checkVal)
                            timer.start()
                    checkVal()                                           #Start thread to check progress
                    #---Weight for all processes to end---
                    for p in self.processes:
                        p.join()
                elif (fitType == 2):
                    for i in range(numCores):
                        if (not self.keepGoing):                         #Check if the fitting has been cancelled
                            return
                        #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                        self.processes.append(mp.Process(target=mp_imag, args=(splitArray[i], vals, i, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams)))
                        listPercent.append(2)
                    for p in self.processes:
                        if (not self.keepGoing):                         #Check if the fitting has been cancelled
                            return
                        p.start()
                        listPercent.append(2)
                    def checkVal():
                        """Self-calling thread which checks processes progress"""
                        #---Check percent completion from shared variable, then append to list to communicate with GUI---
                        for i in range(percentVal.value+1 - len(listPercent)):
                            listPercent.append(1)
                        #---Initiate self-calling every 0.1 seconds while not at 100% completion---
                        if (percentVal.value < listPercent[0]):
                            timer = threading.Timer(0.1, checkVal)
                            timer.start()
                    checkVal()                                           #Start thread to check progress
                    #---Weight for all processes to end---
                    for p in self.processes:
                        p.join()
                elif (fitType == 1):
                    for i in range(numCores):
                        if (not self.keepGoing):                         #Check if the fitting has been cancelled
                            return
                        #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                        self.processes.append(mp.Process(target=mp_real, args=(splitArray[i], vals, i, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams)))
                        listPercent.append(2)
                    for p in self.processes:
                        if (not self.keepGoing):                         #Check if the fitting has been cancelled
                            return
                        p.start()
                        listPercent.append(2)
                    def checkVal():
                        """Self-calling thread which checks processes progress"""
                        #---Check percent completion from shared variable, then append to list to communicate with GUI---
                        for i in range(percentVal.value+1 - len(listPercent)):
                            listPercent.append(1)
                        #---Initiate self-calling every 0.1 seconds while not at 100% completion---
                        if (percentVal.value < listPercent[0]):
                            timer = threading.Timer(0.1, checkVal)
                            timer.start()
                    checkVal()                                           #Start thread to check progress
                    #---Weight for all processes to end---
                    for p in self.processes:
                        p.join()
                lowest = 1E9                                             #Current lowest chi-squared value
                lowestIndex = -1                                         #Current index of the result of the best fit
                #---Check if each fit is better than the previous best, and set the variables if it is---
                for i in range(len(vals)):
                    if (vals[i][0] < lowest):
                        lowest = vals[i][0]
                        lowestIndex = i
                current_best_params = vals[lowestIndex][2]
                current_best_result = vals[lowestIndex][1]
                current_best_val = vals[lowestIndex][0]
        except SystemExit as inst:                                       #If the thread is killed, return indicating that
            return "@", "@", "@", "@", "@", "@", "@", "@", "@", "@"
        except Exception as inst:                                        #If a problem occurss, return indicating that
            return "b", "b", "b", "b", "b", "b", "b", "b", "b", "b"

        if (current_best_result != 0):
            minimized = current_best_result
        if (not minimized.success):     #If the fitting fails, return indicating that
            return "^", "^", "^", "^", "^", "^", "^", "^", "^", "^"
        
        #---Set the parameters to return---
        fitted = minimized.params.valuesdict()
        result = []
        sigma = []
        for i in range(numParams):
            result.append(fitted[paramNames[i]])
            sigma.append(minimized.params[paramNames[i]].stderr)

        #---Calculate newly fitted values---
        ToReturnReal = []
        ToReturnImag = []
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        weightingToReturn = 1
        if (weight == 4):
            weighting = []
        #---Assign variable names their values, then exec the Python string---
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(result[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        if (weight == 4):
            V = ldict['weighting']
        ToReturnReal = Zreal
        ToReturnImag = Zimag
        if (weight == 4):
            weightingToReturn = V
        
        #---If there's a None for at least one parameter standard error---
        if None in sigma:
            if (ToReturnReal[0] == 1E300):              #A value of 1E300 indicates an error
                return "b", "b", "b", "b", "b", "b", "b", "b", "b", "b"
            return result, "-", "-", "-", minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag, current_best_params, weightingToReturn
        
        for s in sigma:
            if (np.isnan(s)):
                if (ToReturnReal[0] == 1E300):          #A value of 1E300 indicates an error
                    return "b", "b", "b", "b", "b", "b", "b", "b", "b", "b"
                return result, "-", "-", "-", minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag, current_best_params, weightingToReturn
        
        if (not self.keepGoing):                        #Check if the fitting has been cancelled
            return

        #---Create the arrays to hold the Monte Carlo results---
        randomParams = np.zeros((numParams, numMonteCarlo))
        randomlyCalculatedReal = np.zeros((len(w), numMonteCarlo))
        randomlyCalculatedImag = np.zeros((len(w), numMonteCarlo))
        
        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        for i in range(numParams):
            randomParams[i] = normal(result[i], abs(sigma[i]), numMonteCarlo)
        
        if (not self.keepGoing):                        #Check if the fitting has been cancelled
            return
        
        #---Calculate impedance values based on the random paramaters---
        for i in range(numMonteCarlo):
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            #---Assign variable names their values, then exec the Python string---
            for k in range(numParams):
                exec(paramNames[k] + " = " + str(float(randomParams[k][i])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            for j in range(len(w)):
                randomlyCalculatedReal[j][i] = Zreal[j]
                randomlyCalculatedImag[j][i] = Zimag[j]
        
        #---Find standard deviations of Monte Carlo results---
        standardDevsReal = np.zeros(len(w))
        standardDevsImag = np.zeros(len(w))
        for i in range(len(w)):
            standardDevsReal[i] = np.std(randomlyCalculatedReal[i])
            standardDevsImag[i] = np.std(randomlyCalculatedImag[i])
        
        return result, sigma, standardDevsReal, standardDevsImag, minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag, current_best_params, weightingToReturn