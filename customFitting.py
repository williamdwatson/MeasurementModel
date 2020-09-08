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
import itertools
import multiprocessing as mp
import threading
import sys

def mp_complex(guesses, sharedList, index, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams):
    def diffComplex(params):
        return (Z_append - model(params))/weightCode(params)
    def model(p):
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return np.append(Zreal, Zimag)
    def weightCode(p):
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
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        if (weight == 2):
            return np.append(V, Vj)
        else:
            return np.append(V, V)
    current_best_val = 1E308
    current_best_params = guesses[0]
    current_best_result = 0
    for i in range(len(guesses)):
        with percentVal.get_lock():
            percentVal.value += 1
        for j in range(numParams):
            parameters.get(paramNames[j]).set(value=guesses[i][j])
        minimized = lmfit.minimize(diffComplex, parameters)
        if (minimized.chisqr < current_best_val):
            current_best_val = minimized.chisqr
            current_best_params = guesses
            current_best_result = minimized
    sharedList[index] = [current_best_val, current_best_result, current_best_params]
        
def mp_real(guesses, sharedList, index, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams):
    def diffReal(params):
        return (ZrIn - modelReal(params))/weightCodeHalf(params)
    def weightCodeHalf(p):
        V = np.ones(len(w)) #Default (weight == 0) is no weighting (weights are 1)
        if (weight == 1):   #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2): #Proportional weighting
            for i in range(len(V)):
                V[i] = assumedNoise*ZrIn[i]
        elif (weight == 3): #Error structure weighting
            pass
        elif (weight == 4): #Custom weighting
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        return V
    def modelReal(p):
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return Zreal
    current_best_val = 1E308
    current_best_params = guesses[0]
    current_best_result = 0
    for i in range(len(guesses)):
        with percentVal.get_lock():
            percentVal.value += 1
        for j in range(numParams):
            parameters.get(paramNames[j]).set(value=guesses[i][j])
        minimized = lmfit.minimize(diffReal, parameters)
        if (minimized.chisqr < current_best_val):
            current_best_val = minimized.chisqr
            current_best_params = guesses
            current_best_result = minimized
    sharedList[index] = [current_best_val, current_best_result, current_best_params]

def mp_imag(guesses, sharedList, index, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams):
    def diffImag(params):
        return (ZjIn - modelImag(params))/weightCodeHalf(params)
    def weightCodeHalf(p):
        V = np.ones(len(w)) #Default (weight == 0) is no weighting (weights are 1)
        if (weight == 1):   #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2): #Proportional weighting
            for i in range(len(V)):
                V[i] = assumedNoise*ZjIn[i]
        elif (weight == 3): #Error structure weighting
            pass
        elif (weight == 4): #Custom weighting
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        return V
    def modelImag(p):
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return Zimag
    current_best_val = 1E308
    current_best_params = guesses[0]
    current_best_result = 0
    for i in range(len(guesses)):
        with percentVal.get_lock():
            percentVal.value += 1
        for j in range(numParams):
            parameters.get(paramNames[j]).set(value=guesses[i][j])
        minimized = lmfit.minimize(diffImag, parameters)
        if (minimized.chisqr < current_best_val):
            current_best_val = minimized.chisqr
            current_best_params = guesses
            current_best_result = minimized
    sharedList[index] = [current_best_val, current_best_result, current_best_params]
    
class customFitter:
    def __init__(self):
        self.keepGoing = True
        self.processes = []
        
    def terminateProcesses(self):
        self.keepGoing = False
        for process in self.processes:
            process.terminate()
    
    def findFit(self, extra_imports, listPercent, fitType, numMonteCarlo, weight, assumedNoise, formula, w, ZrIn, ZjIn, paramNames, paramGuesses, paramBounds, errorParams):
        np.random.seed(1234)                    #Use constant seed to ensure reproducibility of Monte Carlo simulations
        Z_append = np.append(ZrIn, ZjIn)
        numParams = len(paramNames)
        for extra in extra_imports:
            sys.path.append(extra)
        def diffComplex(params):
            if (not self.keepGoing):
                return
            return (Z_append - model(params))/weightCode(params)
        
        def diffImag(params):
            if (not self.keepGoing):
                return
            return (ZjIn - modelImag(params))/weightCodeHalf(params)
        
        def diffReal(params):
            if (not self.keepGoing):
                return
            return (ZrIn - modelReal(params))/weightCodeHalf(params)
        
        #---Used to calculate weights---
        def weightCode(p):
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
                for i in range(numParams):
                    exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
                ldict = locals()
                exec(formula, globals(), ldict)
                V = ldict['weighting']
            return V
        
        #---Model used in complex fitting; returns appendation of real and imaginary parts---
        def model(p):
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            return np.append(Zreal, Zimag)
         
        #---Model used in imaginary fitting; returns only imaginary part---
        def modelImag(p):
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            return Zimag
    
        #---Model used in real fitting; returns only real part---
        def modelReal(p):
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            return Zreal
        
        if (not self.keepGoing):
            return
        
        parameters = lmfit.Parameters()
        parameterCorrector = 0
        for i in range(numParams):
            parameters.add(paramNames[i], value=paramGuesses[i][0])
            if (paramBounds[i] == "-"):
                parameters[paramNames[i]].max = 0
            elif (paramBounds[i] == "+"):
                parameters[paramNames[i]].min = 0
            elif (paramBounds[i] == "f"):
                parameters[paramNames[i]].vary = False
            elif (len(paramBounds[i]) > 1):
                upper, lower = paramBounds[i].split(";")
                if (upper != "inf"):
                    parameters[paramNames[i]].max = float(upper)
                if (lower != "-inf"):
                    parameters[paramNames[i]].min = float(lower)
        
        if (not self.keepGoing):
            return

        try:
            guesses = list(itertools.product(*paramGuesses))
            current_best_val = 1E308
            current_best_params = guesses[0]
            current_best_result = 0
            if (numParams * len(guesses) < 1000):
                for i in range(len(guesses)):
                    for j in range(numParams):
                        parameters.get(paramNames[j]).set(value=guesses[i][j])
                    if (fitType == 3):      #Complex fitting
                        minimized = lmfit.minimize(diffComplex, parameters)
                    elif (fitType == 2):    #Imaginary fitting
                        minimized = lmfit.minimize(diffImag, parameters)
                    elif (fitType == 1):    #Real fitting
                        minimized = lmfit.minimize(diffReal, parameters)
                    if (minimized.chisqr < current_best_val):
                        current_best_val = minimized.chisqr
                        current_best_params = guesses
                        current_best_result = minimized
                    listPercent.append(1)
            else:
                numCores = mp.cpu_count()
                splitArray = np.array_split(guesses, numCores)
                manager = mp.Manager()
                vals = manager.list()
                percentVal = mp.Value("i", 0)
                for i in range(numCores):
                    vals.append([])
                if (fitType == 3):
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        self.processes.append(mp.Process(target=mp_complex, args=(splitArray[i], vals, i, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams)))
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
                elif (fitType == 2):
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        self.processes.append(mp.Process(target=mp_imag, args=(splitArray[i], vals, i, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams)))
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
                elif (fitType == 1):
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        self.processes.append(mp.Process(target=mp_real, args=(splitArray[i], vals, i, parameters, numParams, weight, assumedNoise, formula, w, ZrIn, ZjIn, Z_append, percentVal, paramNames, fitType, errorParams)))
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
                current_best_params = vals[lowestIndex][2]
                current_best_result = vals[lowestIndex][1]
                current_best_val = vals[lowestIndex][0]
        except SystemExit as inst:
            return "@", "@", "@", "@", "@", "@", "@", "@", "@", "@"
        except Exception as inst:
            return "b", "b", "b", "b", "b", "b", "b", "b", "b", "b"

        if (current_best_result != 0):
            minimized = current_best_result
        if (not minimized.success):     #If the fitting fails
            return "^", "^", "^", "^", "^", "^", "^", "^", "^", "^"
    
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

        if None in sigma:
            if (ToReturnReal[0] == 1E300):
                return "b", "b", "b", "b", "b", "b", "b", "b", "b", "b"
            return result, "-", "-", "-", minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag, current_best_params, weightingToReturn
        
        for s in sigma:
            if (np.isnan(s)):
                if (ToReturnReal[0] == 1E300):
                    return "b", "b", "b", "b", "b", "b", "b", "b", "b", "b"
                return result, "-", "-", "-", minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag, current_best_params, weightingToReturn
        
        if (not self.keepGoing):
            return

        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        randomParams = np.zeros((numParams, numMonteCarlo))
        randomlyCalculatedReal = np.zeros((len(w), numMonteCarlo))
        randomlyCalculatedImag = np.zeros((len(w), numMonteCarlo))
        
        for i in range(numParams):
            randomParams[i] = normal(result[i], abs(sigma[i]), numMonteCarlo)
        
        if (not self.keepGoing):
            return
        
        #---Calculate impedance values based on the random paramaters---
        for i in range(numMonteCarlo):
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            freq = w.copy()
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
        for i in range(len(w)):
            standardDevsReal[i] = np.std(randomlyCalculatedReal[i])
            standardDevsImag[i] = np.std(randomlyCalculatedImag[i])
        
        return result, sigma, standardDevsReal, standardDevsImag, minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag, current_best_params, weightingToReturn