# -*- coding: utf-8 -*-
"""
Created on Thu Jan 16 17:49:10 2020

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
#--------------------------------------lm_fit----------------------------------------
#     Source: https://lmfit.github.io/lmfit-py/index.html
#     OR: https://zenodo.org/record/11813#.Xl7RCqhKhPa
#     By: Newville, Matthew; Stensitzki, Till; Allen, Daniel B.; Ingargiola, Antonino
#     License: MIT
#------------------------------------------------------------------------------------
import multiprocessing as mp
import threading

def mp_complex(guesses, sharedList, index, parameters, numVoigtElements, Z_append, V_append, w, perVal):
    mp_current_min = 1E9
    mp_current_best = 0
    def diffComplex(params):
        return (Z_append - model(params))/V_append
    def model(p):
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
        for i in range(2, numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
        return np.append(actual.real, actual.imag)
    def recursive_fit_mp(voigt_number, recursiveGuesses):
        if (voigt_number == 1):
            for i in range(len(recursiveGuesses[0])):
                for j in range(len(recursiveGuesses[1])):
                    for k in range(len(recursiveGuesses[2])):
                        parameters.get("Re").set(value=recursiveGuesses[0][i])
                        parameters.get("R1").set(value=recursiveGuesses[1][j])
                        parameters.get("T1").set(value=recursiveGuesses[2][k])
                        fittedMin = lmfit.minimize(diffComplex, parameters)
                        nonlocal mp_current_min
                        nonlocal mp_current_best
                        with perVal.get_lock():
                            perVal.value += 1
                        if (fittedMin.chisqr < mp_current_min and fittedMin.success):
                            mp_current_min = fittedMin.chisqr
                            mp_current_best = fittedMin
        else:
            for z in range(len(recursiveGuesses[2*voigt_number-1])):
                for y in range(len(recursiveGuesses[2*voigt_number])):
                    parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                    parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                    recursive_fit_mp(voigt_number-1, recursiveGuesses)
    recursive_fit_mp(numVoigtElements, guesses)
    sharedList[index] = [mp_current_min, mp_current_best]

def mp_imag(guesses, sharedList, index, parameters, numVoigtElements, Zj, V, w, perVal):
    mp_current_min = 1E9
    mp_current_best = 0
    def diffImag(params):
        return (Zj - modelImag(params))/V
    def modelImag(p):
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
        for i in range(2, numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
        return actual.imag
    def recursive_fit_imag_mp(voigt_number, recursiveGuesses):
        if (voigt_number == 1):
            for i in range(len(recursiveGuesses[0])):
                for j in range(len(recursiveGuesses[1])):
                    for k in range(len(recursiveGuesses[2])):
                        parameters.get("Re").set(value=recursiveGuesses[0][i])
                        parameters.get("R1").set(value=recursiveGuesses[1][j])
                        parameters.get("T1").set(value=recursiveGuesses[2][k])
                        fittedMin = lmfit.minimize(diffImag, parameters)
                        nonlocal mp_current_min
                        nonlocal mp_current_best
                        with perVal.get_lock():
                            perVal.value += 1
                        if (fittedMin.chisqr < mp_current_min and fittedMin.success):
                            mp_current_min = fittedMin.chisqr
                            mp_current_best = fittedMin
        else:
            for z in range(len(recursiveGuesses[2*voigt_number-1])):
                for y in range(len(recursiveGuesses[2*voigt_number])):
                    parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                    parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                    recursive_fit_imag_mp(voigt_number-1, recursiveGuesses)
    recursive_fit_imag_mp(numVoigtElements, guesses)
    sharedList[index] = [mp_current_min, mp_current_best]

def mp_real(guesses, sharedList, index, parameters, numVoigtElements, Zr, V, w, perVal):
    mp_current_min = 1E9
    mp_current_best = 0
    def diffReal(params):
        return (Zr - modelReal(params))/V
    def modelReal(p):
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
        for i in range(2, numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
        return actual.real 
    def recursive_fit_real_mp(voigt_number, recursiveGuesses):
        if (voigt_number == 1):
            for i in range(len(recursiveGuesses[0])):
                for j in range(len(recursiveGuesses[1])):
                    for k in range(len(recursiveGuesses[2])):
                        parameters.get("Re").set(value=recursiveGuesses[0][i])
                        parameters.get("R1").set(value=recursiveGuesses[1][j])
                        parameters.get("T1").set(value=recursiveGuesses[2][k])
                        fittedMin = lmfit.minimize(diffReal, parameters)
                        nonlocal mp_current_min
                        nonlocal mp_current_best
                        with perVal.get_lock():
                            perVal.value += 1
                        if (fittedMin.chisqr < mp_current_min and fittedMin.success):
                            mp_current_min = fittedMin.chisqr
                            mp_current_best = fittedMin
        else:
            for z in range(len(recursiveGuesses[2*voigt_number-1])):
                for y in range(len(recursiveGuesses[2*voigt_number])):
                    parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                    parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                    recursive_fit_real_mp(voigt_number-1, recursiveGuesses)
    recursive_fit_real_mp(numVoigtElements, guesses)
    sharedList[index] = [mp_current_min, mp_current_best]

class fitter:
    
    def __init__(self):
        self.processes = []
        self.keepGoing = True
    
    def terminateProcesses(self):
        self.keepGoing = False
        for process in self.processes:
            process.terminate()
    
    def findFit(self, w, Zr, Zj, numVoigtElements, numMonteCarlo, choice, assumed_noise, Rm, fitType, lowerBounds, initialGuesses, constants, upperBounds, listPercent, *error_params):
        """
        Finds the fit for the measurement model

        Parameters
        ----------
        w : numpy array
            1D numpy array of frequencies
        Zr : numpy array
            1D numpy array of real impedance data
        Zj : numpy array
            1D numpy array of imaginary impedance data
        numVoigtElements : int
            Number of Voigt elements to fit
        numMonteCarlo : int
            Number of Monte Carlo simulations to perform
        choice : int
            What weighting type to use; 0 for no weighting, 1 for modulus, 2 for proportional, 3 for error model
        assumed_noise : float
            Assumed noise for modulus weighting
        Rm : float
            Modifier for error model weighting gamma term
        fitType : int
            Type of fit to perform; 0 for complex, 1 for imaginary, 2 for real
        lowerBounds : numpy array
            1D numpy array of lower bounds for each parameter
        initialGuesses : list
            List of float initial guesses
        constants : list
            List of constant parameters
        upperBounds : numpy array
            1D numpy array of upper bounds for each parameter
        listPercent : list
            List to modify for communication with main thread
        error_params : floats
            Float error parameters (alpha, beta, betaRe, gamma, and delta)
        
        Returns
        -------
        result : list
            List of parameter results
        sigma : list
            List of parameter standard errors
        standardDevsReal : numpy array
            1D numpy array of real Monte Carlo standard deviations
        standardDevsImag : numpy array
            1D numpy array of imaginary Monte Carlo standard deviations
        Zzero : complex
            Predicted impedance at 0 frequency
        ZzeroSigma : complex
            Standard error on `Zzero`
        Zpolar : complex
            Polarization impedance
        ZpolarSigma : complex
            Standard error on `Zpolar`
        float :
            Chi-squared of minimization
        cor : numpy array
            2D numpy array of parameter correlations
        float :
            Akaike Information Criterion
        """
        np.random.seed(1234)        #Use constant seed to ensure reproducibility of Monte Carlo simulations
        self.constants = np.sort(constants)
        self.Z_append = np.append(Zr, Zj)
        self.numParams = numVoigtElements*2 + 1
        self.w = w
        self.numVoigtElements= numVoigtElements
        self.Zr = Zr
        self.Zj = Zj

        self.V = np.ones(len(w))                 #Default no weighting (all weights are 1) if choice==0
        if (choice == 1):                        #Modulus weighting
            self.V = assumed_noise*np.sqrt(Zr**2 + Zj**2)
        elif (choice == 2):                      #Proportional weighting
            self.Vj = assumed_noise*Zj
            if (fitType != 1):                   #If the fit type isn't imaginary use both Zr and Zj
                self.V = assumed_noise*Zr
            else:                                #If the fit is imaginary use only Zj in weighting
                self.V = assumed_noise*Zj
        elif (choice == 3):                      #Error model weighting
            alpha = error_params[0]
            beta = error_params[1]
            betaRe = error_params[2]
            gamma = error_params[3]
            delta = error_params[4]
            self.V = alpha*np.abs(Zj) + (beta*np.abs(Zr) - betaRe) + gamma*(np.sqrt(Zr**2 + Zj**2)**2)/Rm + delta
        if (choice != 2):
            self.V_append = np.append(self.V, self.V)
        else:
            self.V_append = np.append(self.V, self.Vj)
        
        parameters = lmfit.Parameters()
        parameterCorrector = 0
        parameters.add("Re", value=initialGuesses[0][0], min=lowerBounds[0])
        for i in range(1, numVoigtElements+1):
            if (lowerBounds[i+parameterCorrector] == np.NINF):
                parameters.add("R"+str(i), value=initialGuesses[i+parameterCorrector][0])
            else:
                parameters.add("R"+str(i), value=initialGuesses[i+parameterCorrector][0], min=lowerBounds[i+parameterCorrector])
            if (upperBounds[i+parameterCorrector] != np.inf):
                parameters['R'+str(i)].max = 0
            parameters.add("T"+str(i), value=initialGuesses[i+1+parameterCorrector][0], min=0.0)
            parameterCorrector += 1
        if (constants[0] != -1):
            for constant in constants:
                if (constant == 0):
                    parameters['Re'].vary = False
                elif (constant%2 == 0):
                    parameters['T'+str(int(constant/2))].vary = False
                else:
                    parameters['R'+str(int(np.ceil(constant/2)))].vary = False
        
        current_min = 1E9
        current_best = 0
        def recursive_fit(voigt_number, recursiveGuesses):
            if (not self.keepGoing):
                return
            if (voigt_number == 1):
                for i in range(len(recursiveGuesses[0])):
                    for j in range(len(recursiveGuesses[1])):
                        for k in range(len(recursiveGuesses[2])):
                            parameters.get("Re").set(value=recursiveGuesses[0][i])
                            parameters.get("R1").set(value=recursiveGuesses[1][j])
                            parameters.get("T1").set(value=recursiveGuesses[2][k])
                            fittedMin = lmfit.minimize(self.diffComplex, parameters)
                            nonlocal current_min
                            nonlocal current_best
                            listPercent.append(1)
                            if (fittedMin.chisqr < current_min and fittedMin.success):
                                current_min = fittedMin.chisqr
                                current_best = fittedMin
            else:
                for z in range(len(recursiveGuesses[2*voigt_number-1])):
                    for y in range(len(recursiveGuesses[2*voigt_number])):
                        parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                        parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                        recursive_fit(voigt_number-1, recursiveGuesses)
        
        def recursive_fit_imag(voigt_number, recursiveGuesses):
            if (not self.keepGoing):
                return
            if (voigt_number == 1):
                for i in range(len(recursiveGuesses[0])):
                    for j in range(len(recursiveGuesses[1])):
                        for k in range(len(recursiveGuesses[2])):
                            parameters.get("Re").set(value=recursiveGuesses[0][i])
                            parameters.get("R1").set(value=recursiveGuesses[1][j])
                            parameters.get("T1").set(value=recursiveGuesses[2][k])
                            fittedMin = lmfit.minimize(self.diffImag, parameters)
                            nonlocal current_min
                            nonlocal current_best
                            listPercent.append(1)
                            if (fittedMin.chisqr < current_min and fittedMin.success):
                                current_min = fittedMin.chisqr
                                current_best = fittedMin
            else:
                for z in range(len(recursiveGuesses[2*voigt_number-1])):
                    for y in range(len(recursiveGuesses[2*voigt_number])):
                        parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                        parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                        recursive_fit_imag(voigt_number-1, recursiveGuesses)
        
        def recursive_fit_real(voigt_number, recursiveGuesses):
            if (not self.keepGoing):
                return
            if (voigt_number == 1):
                for i in range(len(recursiveGuesses[0])):
                    for j in range(len(recursiveGuesses[1])):
                        for k in range(len(recursiveGuesses[2])):
                            parameters.get("Re").set(value=recursiveGuesses[0][i])
                            parameters.get("R1").set(value=recursiveGuesses[1][j])
                            parameters.get("T1").set(value=recursiveGuesses[2][k])
                            fittedMin = lmfit.minimize(self.diffReal, parameters)
                            nonlocal current_min
                            nonlocal current_best
                            listPercent.append(1)
                            if (fittedMin.chisqr < current_min and fittedMin.success):
                                current_min = fittedMin.chisqr
                                current_best = fittedMin
            else:
                for z in range(len(recursiveGuesses[2*voigt_number-1])):
                    for y in range(len(recursiveGuesses[2*voigt_number])):
                        parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                        parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                        recursive_fit_real(voigt_number-1, recursiveGuesses)
     
        numCombos = 1
        numCores = mp.cpu_count()
        for g in initialGuesses:
            numCombos *= len(g)
        if (numVoigtElements * numCombos > 1000):
            longestIndex = initialGuesses.index(max(initialGuesses, key=len))       #Find largest set of initial guesses
            splitArray = np.array_split(initialGuesses[longestIndex], numCores)     #Split the array at the index of the longest set of initial Guesses
            toRun = []                                                              #Initialize the array that will be used for multi-core processing
            for i in range(numCores):
                toAppend = []
                for j in range(len(initialGuesses)):
                    if (j == longestIndex):
                        toAppend.append(splitArray[i].tolist())
                    else:
                        toAppend.append(initialGuesses[j])
                toRun.append(toAppend)
            manager = mp.Manager()
            vals = manager.list()
            percentVal = mp.Value("i", 0)
            for i in range(numCores):
                vals.append([])
            if (fitType == 0):
                for i in range(numCores):
                    if (not self.keepGoing):
                        return
                    self.processes.append(mp.Process(target=mp_complex, args=(toRun[i], vals, i, parameters, numVoigtElements, self.Z_append, self.V_append, w, percentVal)))
                    listPercent.append(2)
                for p in self.processes:
                    if (not self.keepGoing):
                        return
                    p.start()
                    listPercent.append(2)
                def checkVal():
                    for i in range(percentVal.value+1 - len(listPercent)):
                        listPercent.append(1)
                    if (percentVal.value < listPercent[0]):
                        timer = threading.Timer(0.1, checkVal)
                        timer.start()
                checkVal()
                for p in self.processes:
                    p.join()
            if (fitType == 1):
                parameters['Re'].vary = False
                for i in range(numCores):
                    if (not self.keepGoing):
                        return
                    self.processes.append(mp.Process(target=mp_imag, args=(toRun[i], vals, i, parameters, numVoigtElements, Zj, self.V, w, percentVal)))
                    listPercent.append(2)
                for p in self.processes:
                    if (not self.keepGoing):
                        return
                    p.start()
                    listPercent.append(2)
                def checkVal():
                    for i in range(percentVal.value+1 - len(listPercent)):
                        listPercent.append(1)
                    if (percentVal.value < listPercent[0]):
                        timer = threading.Timer(0.1, checkVal)
                        timer.start()
                checkVal()
                for p in self.processes:
                    p.join()
            if (fitType == 2):
                for i in range(numCores):
                    if (not self.keepGoing):
                        return
                    self.processes.append(mp.Process(target=mp_real, args=(toRun[i], vals, i, parameters, numVoigtElements, Zr, self.V, w, percentVal)))
                    listPercent.append(2)
                for p in self.processes:
                    if (not self.keepGoing):
                        return
                    p.start()
                    listPercent.append(2)
                def checkVal():
                    for i in range(percentVal.value+1 - len(listPercent)):
                        listPercent.append(1)
                    if (percentVal.value < listPercent[0]):
                        timer = threading.Timer(0.1, checkVal)
                        timer.start()
                checkVal()
                for p in self.processes:
                    p.join()
            lowest = 1E9
            lowestIndex = -1
            for i in range(len(vals)):
                if (vals[i][0] < lowest):
                    lowest = vals[i][0]
                    lowestIndex = i
            minimized = vals[lowestIndex][1]
            current_min = vals[lowestIndex][0]
        else:
            if (fitType == 0):      #Complex fitting
                recursive_fit(numVoigtElements, initialGuesses)
                minimized = current_best
            elif (fitType == 1):    #Imaginary fitting
                parameters['Re'].vary = False
                recursive_fit_imag(numVoigtElements, initialGuesses)
                minimized = current_best
            elif (fitType == 2):    #Real fitting
                recursive_fit_real(numVoigtElements, initialGuesses)
                minimized = current_best
            
        if (current_min == 1E9):     #If the fitting fails
            return "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^"
        
        fitted = minimized.params.valuesdict()
        result = [fitted['Re']]
        sigma = [minimized.params['Re'].stderr]
        for i in range(1, numVoigtElements+1):
            result.append(fitted['R'+str(i)])
            result.append(fitted['T'+str(i)])
            sigma.append(minimized.params['R'+str(i)].stderr)
            sigma.append(minimized.params['T'+str(i)].stderr)
        if None in sigma:
            Zzero = result[0] + result[1]
            for i in range(3, len(result), 2):
                Zzero += result[i]
            Zpolar = 0
            for i in range(1, len(result), 2):
                Zpolar += result[i]
            cvar = np.zeros((self.numParams, self.numParams))
            return result, "-", "-", "-", Zzero, "-", Zpolar, "-", minimized.chisqr, cvar, minimized.aic
        
        cor = np.zeros((self.numParams, self.numParams))
        if (minimized.params['Re'].correl == None):
            cor[:,0] = 0
            cor[0,:] = 0
            cor[0][0] = 1
        else:
            cor[0][0] = 1
            parameterCorrector = 0
            for i in range(1, numVoigtElements+1):
                try:
                    cor[0][i+parameterCorrector] = minimized.params['Re'].correl['R'+str(i)]
                except KeyError:
                    cor[0][i+parameterCorrector] = 0
                try:
                    cor[0][i+1+parameterCorrector] = minimized.params['Re'].correl['T'+str(i)]
                except KeyError:
                    cor[0][i+1+parameterCorrector] = 0
                parameterCorrector += 1
        parameterCorrectorA = 0
        for i in range(1, numVoigtElements+1):
            #---R---
            if (minimized.params['R'+str(i)].correl == None):
                cor[i+parameterCorrectorA,:] = 0
                cor[i+parameterCorrectorA][i+parameterCorrectorA] = 1
            else:
                try:
                    cor[i+parameterCorrectorA][0] = minimized.params['R'+str(i)].correl['Re']
                except KeyError:
                    cor[i+parameterCorrectorA][0] = 0
                cor[i+parameterCorrectorA][i+parameterCorrectorA] = 1
                parameterCorrectorB = 0
                for j in range(1, numVoigtElements+1):
                    
                    if (i != j):
                        try:
                            cor[i+parameterCorrectorA][j+parameterCorrectorB] = minimized.params['R'+str(i)].correl['R'+str(j)]
                        except KeyError:
                            cor[i+parameterCorrectorA][j+parameterCorrectorB] = 0
                    try:
                        cor[i+parameterCorrectorA][j+1+parameterCorrectorB] = minimized.params['R'+str(i)].correl['T'+str(j)]
                    except KeyError:
                        cor[i+parameterCorrectorA][j+1+parameterCorrectorB] = 0
                    parameterCorrectorB += 1
                    
            #---Tau---
            if (minimized.params['T'+str(i)].correl == None):
                cor[i+1+parameterCorrectorA,:] = 0
                cor[i+1+parameterCorrectorA][i+1+parameterCorrectorA] = 1
            else:
                try:
                    cor[i+1+parameterCorrectorA][0] = minimized.params['T'+str(i)].correl['Re']
                except KeyError:
                    cor[i+1+parameterCorrectorA][0] = 0
                cor[i+1+parameterCorrectorA][i+1+parameterCorrectorA] = 1
                
                parameterCorrectorB = 0
                for j in range(1, numVoigtElements+1):         
                    try:
                        cor[i+1+parameterCorrectorA][j+parameterCorrectorB] = minimized.params['T'+str(i)].correl['R'+str(j)]
                    except KeyError:
                        cor[i+1+parameterCorrectorA][j+parameterCorrectorB] = 0
                    if (i != j):
                        try:
                            cor[i+1+parameterCorrectorA][j+1+parameterCorrectorB] = minimized.params['T'+str(i)].correl['T'+str(j)]
                        except KeyError:
                            cor[i+1+parameterCorrectorA][j+1+parameterCorrectorB] = 0
                    parameterCorrectorB += 1    
                
            parameterCorrectorA += 1
    
        
        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        randomParams = np.zeros((self.numParams, numMonteCarlo))
        randomlyCalculated = np.zeros((len(w), numMonteCarlo), dtype=np.complex128)
        
        for i in range(self.numParams):
            randomParams[i] = normal(result[i], abs(sigma[i]), numMonteCarlo)
        
        #---Calculate impedance values based on the random paramaters---
        with np.errstate(divide='ignore', invalid='ignore'):    #Ignore results with divide-by-zero errors and whatnot
            for i in range(numMonteCarlo):
                for j in range(len(w)):
                    randomlyCalculated[j][i] = randomParams[0][i] + (randomParams[1][i]/(1+(1j*w[j]*2*np.pi*randomParams[2][i])))
                    for k in range(3, self.numParams, 2):
                        randomlyCalculated[j][i] += (randomParams[k][i]/(1+(1j*w[j]*2*np.pi*randomParams[k+1][i])))
    
        standardDevsReal = np.zeros(len(w))
        standardDevsImag = np.zeros(len(w))
        for i in range(len(w)):
            standardDevsReal[i] = np.std(randomlyCalculated[i].real)
            standardDevsImag[i] = np.std(randomlyCalculated[i].imag)
        
        #---Calculate the predicted impedance at 0 frequency---
        Zzero = result[0] + result[1]
        ZzeroSigma = sigma[0]**2 + sigma[1]**2
        for i in range(3, len(result), 2):
            Zzero += result[i]
            ZzeroSigma += sigma[i]**2
        ZzeroSigma = np.sqrt(ZzeroSigma)
        
        #---Calculate the polarization impedance---
        Zpolar = 0
        ZpolarSigma = 0
        for i in range(1, len(result), 2):
            Zpolar += result[i]
            ZpolarSigma += sigma[i]**2
        ZpolarSigma = np.sqrt(ZpolarSigma)
                
        #---Return everything---
        return result, sigma, standardDevsReal, standardDevsImag, Zzero, ZzeroSigma, Zpolar, ZpolarSigma, minimized.chisqr, cor, minimized.aic
    
    def diffComplex(self, params):
        return (self.Z_append - self.model(params))/self.V_append
    
    def diffImag(self, params):
        return (self.Zj - self.modelImag(params))/self.V
    
    def diffReal(self, params):
        return (self.Zr - self.modelReal(params))/self.V
    
    #---Model used in complex fitting; returns appendation of real and imaginary parts---
    def model(self, p):
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*self.w*2*np.pi*t1)))
        for i in range(2, self.numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*self.w*2*np.pi*p['T'+str(i)])))
        return np.append(actual.real, actual.imag)
        
    #---Model used in imaginary fitting; returns only imaginary part---
    def modelImag(self, p):
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*self.w*2*np.pi*t1)))
        for i in range(2, self.numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*self.w*2*np.pi*p['T'+str(i)])))
        return actual.imag
    
    #---Model used in real fitting; returns only real part---
    def modelReal(self, p):
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*self.w*2*np.pi*t1)))
        for i in range(2, self.numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*self.w*2*np.pi*p['T'+str(i)])))
        return actual.real
