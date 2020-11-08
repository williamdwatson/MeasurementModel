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
import multiprocessing as mp
import threading, sys

def mp_complex(sharedList, numBootstrap, perVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams):
    """
    Function for multiprocessing the bootstrapping of parameter errors after a complex fit
    
    Parameters
    ----------
    sharedList : list
        List of succesful parameter fits.
    numBootstrap : int
        Number of fits to perform.
    perVal : multiprocessing Manager int
        Current percent completion.
    numParams : int
        Number of parameters.
    paramNames : list of strings
        Parameter names.
    r_in : array of floats
        Previous parameter fits.
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    paramGuesses : list of floats
        Initial guesses for parameters.
    paramBounds : list of strings
        Parameter bounds: '-' for negative, '+' for positive, 'f' for fixed.
    w : array of floats
        Data frequencies.
    weight : int
        Which weighting to use when fitting; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
    assumedNoise : float
        The assumed noise used when weighting.
    ZrIn : array of floats
        Data real impedance.
    ZjIn : array of floats
        Data imaginary impedance.
    fitType : int
        Which fit type to use.
    formula : string
        The custom Python code.
    errorParams : list of floats
        The parameters used for error model weighting.

    Returns
    -------
    None (results are appended to a shared list)

    """
    #---Set up the parameters for fitting---
    parameters = lmfit.Parameters()
    for i in range(numParams):
        parameters.add(paramNames[i], value=paramGuesses[i])
        if (paramBounds[i] == "-"):
            parameters[paramNames[i]].max = 0
        elif (paramBounds[i] == "+"):
            parameters[paramNames[i]].min = 0
        elif (paramBounds[i] == "f"):
            parameters[paramNames[i]].vary = False    
    parametersRes = lmfit.Parameters()
    for i in range(numParams):
        parametersRes.add(paramNames[i], value=r_in[i])
    
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
        V = np.ones(len(w))                #Default - no weighting
        if (weight == 1):                  #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2):                #Proportional weighting
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
    
    def diffComplex(params, yVals):
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
        return (yVals - model(params))/weightCode(params)
    residuals = diffComplex(parametersRes, Z_append)
    residualsStdDevR = np.std(residuals[:len(residuals)//2])
    residualsStdDevJ = np.std(residuals[len(residuals)//2:])
    for i in range(numBootstrap):
        with perVal.get_lock():                                           #Multiprocessing-safe increment of percent completion
            perVal.value += 1
        random_deltas_r = normal(0, residualsStdDevR, len(Z_append)//2)   #Generate Gaussian distribution from residuals standard deviation
        random_deltas_j = normal(0, residualsStdDevJ, len(Z_append)//2)   #Generate Gaussian distribution from residuals standard deviation
        random_deltas = np.append(random_deltas_r, random_deltas_j)
        yTrial = Z_append + random_deltas                                 #Add newly generated random values to the known impedance
        minimized = lmfit.minimize(diffComplex, parameters, args=(yTrial,))    #Fit the newly randomized data
        #---If the fitting is successful, append the new results to the shared list---
        if (minimized.success):         
            fitted = minimized.params.valuesdict()
            toAppendResults = []
            for j in range(numParams):
                toAppendResults.append(fitted[paramNames[j]])
            sharedList.append(toAppendResults)

def mp_imag(sharedList, numBootstrap, perVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams):
    """
    Function for multiprocessing the bootstrapping of parameter errors after an imaginary fit
    
    Parameters
    ----------
    sharedList : list
        List of succesful parameter fits.
    numBootstrap : int
        Number of fits to perform.
    perVal : multiprocessing Manager int
        Current percent completion.
    numParams : int
        Number of parameters.
    paramNames : list of strings
        Parameter names.
    r_in : array of floats
        Previous parameter fits.
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    paramGuesses : list of floats
        Initial guesses for parameters.
    paramBounds : list of strings
        Parameter bounds: '-' for negative, '+' for positive, 'f' for fixed.
    w : array of floats
        Data frequencies.
    weight : int
        Which weighting to use when fitting; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
    assumedNoise : float
        The assumed noise used when weighting.
    ZrIn : array of floats
        Data real impedance.
    ZjIn : array of floats
        Data imaginary impedance.
    fitType : int
        Which fit type to use.
    formula : string
        The custom Python code.
    errorParams : list of floats
        The parameters used for error model weighting.

    Returns
    -------
    None (results are appended to a shared list)

    """
    #---Set up the parameters for fitting---
    parameters = lmfit.Parameters()
    for i in range(numParams):
        parameters.add(paramNames[i], value=paramGuesses[i])
        if (paramBounds[i] == "-"):
            parameters[paramNames[i]].max = 0
        elif (paramBounds[i] == "+"):
            parameters[paramNames[i]].min = 0
        elif (paramBounds[i] == "f"):
            parameters[paramNames[i]].vary = False    
    parametersRes = lmfit.Parameters()
    for i in range(numParams):
        parametersRes.add(paramNames[i], value=r_in[i])

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
    
    def diffImag(params, yVals):
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
        return (yVals - modelImag(params))/weightCodeHalf(params)
    residuals = diffImag(parametersRes, ZjIn)
    residualsStdDev = np.std(residuals)
    for i in range(numBootstrap):
        with perVal.get_lock():                                 #Multiprocessing-safe increment of percent completion
            perVal.value += 1
        random_deltas = normal(0, residualsStdDev, len(ZjIn))   #Generate Gaussian distribution from residuals standard deviation
        yTrial = ZjIn + random_deltas                           #Add newly generated random values to the known imaginary impedance
        minimized = lmfit.minimize(diffImag, parameters, args=(yTrial,))    #Fit the newly randomized data
        #---If the fitting is successful, append the new results to the shared list---
        if (minimized.success):
            fitted = minimized.params.valuesdict()
            toAppendResults = []
            for j in range(numParams):
                toAppendResults.append(fitted[paramNames[j]])
            sharedList.append(toAppendResults)

def mp_real(sharedList, numBootstrap, perVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams):
    """
    Function for multiprocessing the bootstrapping of parameter errors after a complex fit
    
    Parameters
    ----------
    sharedList : list
        List of succesful parameter fits.
    numBootstrap : int
        Number of fits to perform.
    perVal : multiprocessing Manager int
        Current percent completion.
    numParams : int
        Number of parameters.
    paramNames : list of strings
        Parameter names.
    r_in : array of floats
        Previous parameter fits.
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    paramGuesses : list of floats
        Initial guesses for parameters.
    paramBounds : list of strings
        Parameter bounds: '-' for negative, '+' for positive, 'f' for fixed.
    w : array of floats
        Data frequencies.
    weight : int
        Which weighting to use when fitting; 0 for no weighting (weights are unity), 1 for modulus, 2 for proportional, 3 for error structure, 4 for custom.
    assumedNoise : float
        The assumed noise used when weighting.
    ZrIn : array of floats
        Data real impedance.
    ZjIn : array of floats
        Data imaginary impedance.
    fitType : int
        Which fit type to use.
    formula : string
        The custom Python code.
    errorParams : list of floats
        The parameters used for error model weighting.

    Returns
    -------
    None (results are appended to a shared list)

    """
    #---Set up the parameters for fitting---
    parameters = lmfit.Parameters()
    for i in range(numParams):
        parameters.add(paramNames[i], value=paramGuesses[i])
        if (paramBounds[i] == "-"):
            parameters[paramNames[i]].max = 0
        elif (paramBounds[i] == "+"):
            parameters[paramNames[i]].min = 0
        elif (paramBounds[i] == "f"):
            parameters[paramNames[i]].vary = False    
    parametersRes = lmfit.Parameters()
    for i in range(numParams):
        parametersRes.add(paramNames[i], value=r_in[i])
    
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
    
    def diffReal(params, yVals):
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
        return (yVals - modelReal(params))/weightCodeHalf(params)
    residuals = diffReal(parametersRes, ZrIn)
    residualsStdDev = np.std(residuals)
    for i in range(numBootstrap):
        with perVal.get_lock():                                 #Multiprocessing-safe increment of percent completion
            perVal.value += 1
        random_deltas = normal(0, residualsStdDev, len(ZrIn))   #Generate Gaussian distribution from residuals standard deviation
        yTrial = ZrIn + random_deltas                           #Add newly generated random values to the known real impedance
        minimized = lmfit.minimize(diffReal, parameters, args=(yTrial,))    #Fit the newly randomized data
        #---If the fitting is successful, append the new results to the shared list---
        if (minimized.success):
            fitted = minimized.params.valuesdict()
            toAppendResults = []
            for j in range(numParams):
                toAppendResults.append(fitted[paramNames[j]])
            sharedList.append(toAppendResults)

class bootstrapFitter:
    def __init__(self):
        self.processes = []
        self.keepGoing = True
    
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
    
    def findFit(self, extra_imports, listPercent, numBootstrap, r_in, s_in, sdR_in, sdI_in, chi_in, aic_in, realF_in, imagF_in, fitType, numMonteCarlo, weight, assumedNoise, formula, w, ZrIn, ZjIn, paramNames, paramGuesses, paramBounds, errorParams):
        """
        Perform bootstrap calculations of parameter standard errors. First, the known impedance data is changed by adding a random Gaussian delta to each point, centered at the known value with a sigma based on the standard deviation of the previous residual errors.
        Then, new fittings are performed using the previous parameter fits as initial guesses. The standard error is then found from all of the new fits.

        Parameters
        ----------
        extra_imports : list of strings
            Extra import paths.
        listPercent : list of ints
            List used to communicate precent completion for the GUI progress bar.
        numBootstrap : int
            Number of calculations to perform.
        r_in : list of floats
            Previous parameter fits.
        s_in : list of floats
            Previous parameter standard errors.
        sdR_in : array of floats
            Real confidence interval.
        sdI_in : array of floats
            Imaginary confidence interval.
        chi_in : float
            Previous chi-squared statistic.
        aic_in : float
            Previous Akaike Information Criterion.
        realF_in : array of floats
            Previous real model values.
        imagF_in : array of floats
            Previous imaginary model values.
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
        r_in : array of floats
            Parameter fits.
        sigma : array of floats
            Parameter standard errors
        standardDevsReal : array of floats
            Real standard deviations
        standardDevsImag : array of floats
            Imaginary standard deviations
        chi_in : float
            Chi-squared statistic of fit
        aic_in : float
            Akaike Information Criterion of fit
        ToReturnReal : array of floats
            Real model values
        ToReturnImag : array of floats
            Imaginary model values
        weightingToReturn : array of floats
            Weighting at each frequency

        """
        np.random.seed(1234)                #Use constant seed to ensure reproducibility of Monte Carlo simulations
        Z_append = np.append(ZrIn, ZjIn)
        numParams = len(paramNames)         #The number of parameters in the fitting
        
        #---Add extra imports to the search path of Python; this allows the custom Python code to import new modules---
        for extra in extra_imports:
            sys.path.append(extra)
        
        def diffComplex(params, yVals):
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
            if (not self.keepGoing):            #Check if the fitting has been cancelled
                return
            return (yVals - model(params))/weightCode(params)
        
        def diffImag(params, yVals):
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
            if (not self.keepGoing):            #Check if the fitting has been cancelled
                return
            return (yVals - modelImag(params))/weightCodeHalf(params)
        
        def diffReal(params, yVals):
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
            if (not self.keepGoing):            #Check if the fitting has been cancelled
                return
            return (yVals - modelReal(params))/weightCodeHalf(params)
        
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
            V = np.ones(len(w))                #Default - no weighting
            if (weight == 1):                  #Modulus weighting
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
                #---Assign variable names their values, then exec the Python code string---
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
            V = np.ones(len(w))     #Default (weight == 0) is no weighting (weights are 1)
            if (weight == 1):       #Modulus weighting
                for i in range(len(V)):
                    V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
            elif (weight == 2):     #Proportional weighting
                for i in range(len(V)):
                    V[i] = assumedNoise*ZrIn[i]
            elif (weight == 3):     #Error structure weighting
                pass
            elif (weight == 4):     #Custom weighting
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
        
        if (not self.keepGoing):    #Check if the fitting has been cancelled yet
            return
        
        #---Create the lmfit Parameters---
        parameters = lmfit.Parameters()
        for i in range(numParams):
            parameters.add(paramNames[i], value=paramGuesses[i])
            if (paramBounds[i] == "-"):                             #Negative parameters
                parameters[paramNames[i]].max = 0
            elif (paramBounds[i] == "+"):                           #Positive parameters
                parameters[paramNames[i]].min = 0
            elif (paramBounds[i] == "f"):                           #Fixed parameters
                parameters[paramNames[i]].vary = False
            elif (len(paramBounds[i]) > 1):                         #Custom bounds
                upper, lower = paramBounds[i].split(";")
                if (upper != "inf"):                                #Set non-infinite upper bounds; otherwise, keep the default (no upper bound)
                    parameters[paramNames[i]].max = float(upper)
                if (lower != "-inf"):
                    parameters[paramNames[i]].min = float(lower)    #Set non-infinite lower bounds; otherwise, keep the default (no lower bound)
        
        if (not self.keepGoing):    #Check if the fitting has been cancelled yet
            return

        try:
            if (numBootstrap <= 1000):      #Determine whether or not to use multiprocessing
                #---Calculate residuals---
                if (fitType == 3):          #Complex fit
                    #---Create parameters for fitting---
                    parametersRes = lmfit.Parameters()
                    for i in range(numParams):
                        parametersRes.add(paramNames[i], value=r_in[i])
                    residuals = diffComplex(parametersRes, Z_append)
                    residualsStdDevR = np.std(residuals[:len(residuals)//2])
                    residualsStdDevJ = np.std(residuals[len(residuals)//2:])
                    paramResultArray = []
                    #---Perform bootstrap calculations---
                    for i in range(numBootstrap):
                        listPercent.append(1)
                        random_deltas_r = normal(0, residualsStdDevR, len(Z_append)//2)   #Generate Gaussian distribution from residuals standard deviation
                        random_deltas_j = normal(0, residualsStdDevJ, len(Z_append)//2)   #Generate Gaussian distribution from residuals standard deviation
                        random_deltas = np.append(random_deltas_r, random_deltas_j)
                        yTrial = Z_append + random_deltas                                 #Add newly generated random values to the known impedance
                        minimized = lmfit.minimize(diffComplex, parameters, args=(yTrial,))     #Fit the newly randomized data
                        #---If the fitting succeeded, add that new set of fits to the list---
                        if (minimized.success):
                            fitted = minimized.params.valuesdict()
                            toAppendResults = []
                            for j in range(numParams):
                                toAppendResults.append(fitted[paramNames[j]])
                            paramResultArray.append(toAppendResults)
                    sigma = np.std(paramResultArray, axis=0)                                #Calculate parameter errors as the standard devation of the fits
                elif (fitType == 2):            #Imaginary fit
                     #---Create parameters for fitting---
                    parametersRes = lmfit.Parameters()
                    for i in range(numParams):
                        parametersRes.add(paramNames[i], value=r_in[i])
                    residuals = diffImag(parametersRes, ZjIn)
                    residualsStdDev = np.std(residuals)
                    paramResultArray = []
                    for i in range(numBootstrap):
                        listPercent.append(1)
                        random_deltas = normal(0, residualsStdDev, len(ZjIn))   #Generate Gaussian distribution from residuals standard deviation
                        yTrial = ZjIn + random_deltas                           #Add newly generated random values to the known impedance
                        minimized = lmfit.minimize(diffImag, parameters, args=(yTrial,))     #Fit the newly randomized data
                        #---If the fitting succeeded, add that new set of fits to the list---
                        if (minimized.success):
                            fitted = minimized.params.valuesdict()
                            toAppendResults = []
                            for j in range(numParams):
                                toAppendResults.append(fitted[paramNames[j]])
                            paramResultArray.append(toAppendResults)
                    sigma = np.std(paramResultArray, axis=0)                    #Calculate parameter errors as the standard devation of the fits
                elif (fitType == 1):            #Real fit
                 #---Create parameters for fitting---
                    parametersRes = lmfit.Parameters()
                    for i in range(numParams):
                        parametersRes.add(paramNames[i], value=r_in[i])
                    residuals = diffReal(parametersRes, ZrIn)
                    residualsStdDev = np.std(residuals)
                    paramResultArray = []
                    for i in range(numBootstrap):
                        listPercent.append(1)
                        random_deltas = normal(0, residualsStdDev, len(ZrIn))   #Generate Gaussian distribution from residuals standard deviation
                        yTrial = ZrIn + random_deltas                           #Add newly generated random values to the known impedance
                        minimized = lmfit.minimize(diffReal, parameters, args=(yTrial,))     #Fit the newly randomized data
                         #---If the fitting succeeded, add that new set of fits to the list---
                        if (minimized.success):
                            fitted = minimized.params.valuesdict()
                            toAppendResults = []
                            for j in range(numParams):
                                toAppendResults.append(fitted[paramNames[j]])
                            paramResultArray.append(toAppendResults)
                    sigma = np.std(paramResultArray, axis=0)                    #Calculate parameter errors as the standard devation of the fits
            else:   #Use multiprocessing if more than a certain number of trials requested
                numCores = mp.cpu_count()                           #Get number of CPUs available
                manager = mp.Manager()                              #Create multiprocessing Manager for shared variables
                vals = manager.list()
                percentVal = mp.Value("i", 0)                       #Shared percent completion variable for progress bar
                numRunsPer = int(np.ceil(numBootstrap/numCores))    #Number of runs needed per CPU
                if (fitType == 3):          #Complex fit
                    for i in range(numCores):
                        if (not self.keepGoing):    #Check if fitting has been cancelled
                            return
                        #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                        self.processes.append(mp.Process(target=mp_complex, args=(vals, numRunsPer, percentVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams)))
                        listPercent.append(2)
                    for p in self.processes:
                        if (not self.keepGoing):
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
                    checkVal()  #Start thread to check progress
                    #---Weight for all processes to end---
                    for p in self.processes:
                        p.join()
                    sigma = np.std(vals, axis=0)                #Calculate parameter errors as the standard devation of the fits 
                elif (fitType == 2):       #Imaginary fit
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                         #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                        self.processes.append(mp.Process(target=mp_imag, args=(vals, numRunsPer, percentVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams)))
                        listPercent.append(2)
                    for p in self.processes:
                        if (not self.keepGoing):
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
                    checkVal()  #Start thread to check progress
                    #---Weight for all processes to end---
                    for p in self.processes:
                        p.join()
                    sigma = np.std(vals, axis=0)               #Calculate parameter errors as the standard devation of the fits
                elif (fitType == 1):       #Real fit
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                        self.processes.append(mp.Process(target=mp_real, args=(vals, numRunsPer, percentVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams)))
                        listPercent.append(2)
                    for p in self.processes:
                        if (not self.keepGoing):
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
                    checkVal()
                    #---Weight for all processes to end---
                    for p in self.processes:
                        p.join()
                    sigma = np.std(vals, axis=0)               #Calculate parameter errors as the standard devation of the fits
        except SystemExit:       #If cancelled
            return "@", "@", "@", "@", "@", "@", "@", "@"
        except Exception:        #If there's a problem
            return "b", "b", "b", "b", "b", "b", "b", "b"
        
         #---Calculate newly fitted model and weighting values---
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
            exec(paramNames[i] + " = " + str(float(r_in[i])))
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
        
        #---Create the arrays to hold the Monte Carlo results---
        randomParams = np.zeros((numParams, numMonteCarlo))
        randomlyCalculatedReal = np.zeros((len(w), numMonteCarlo))
        randomlyCalculatedImag = np.zeros((len(w), numMonteCarlo))
        
        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        for i in range(numParams):
            randomParams[i] = normal(r_in[i], abs(sigma[i]), numMonteCarlo)
        
        if (not self.keepGoing):        #Check whether fitting has been cancelled
            return
        
        #---Calculate values based on the random parameters---
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
    
        standardDevsReal = np.zeros(len(w))
        standardDevsImag = np.zeros(len(w))
        #---Find standard deviations of Monte Carlo results---
        for i in range(len(w)):
            standardDevsReal[i] = np.std(randomlyCalculatedReal[i])
            standardDevsImag[i] = np.std(randomlyCalculatedImag[i])
        
        return r_in, sigma, standardDevsReal, standardDevsImag, chi_in, aic_in, ToReturnReal, ToReturnImag, weightingToReturn