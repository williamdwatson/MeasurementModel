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
#from scipy.optimize import least_squares
from numpy.random import normal
import lmfit
#import matplotlib.pyplot as plt
#import time
import multiprocessing as mp
import threading

def mp_complex(sharedList, numBootstrap, perVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams):
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
    #---Used to calculate weights---
    def weightCode(p):
        V = np.ones(len(w))     #Default - no weighting
        if (weight == 1):       #Modulus weighting
            for i in range(len(V)):
                V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
        elif (weight == 2):     #Proportional weighting
            Vj = np.ones(len(w))
            for i in range(len(V)):
                if (fitType != 1):          #If the fit type isn't imaginary use both Zr and Zj
                    V[i] = assumedNoise*ZrIn[i]
                else:                       #If the fit is imaginary use only Zj in weighting
                    V[i] = assumedNoise*ZjIn[i]
                Vj[i] = assumedNoise*ZjIn[i]
        elif (weight == 3):     #Error structure weighting
            for i in range(len(V)):
                V[i] = errorParams[0]*ZjIn[i] + (errorParams[1]*ZrIn[i] - errorParams[2]) + errorParams[3]*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2) + errorParams[4]
        elif (weight == 4):     #Custom weighting
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
    #---Model used in complex fitting; returns appendation of real and imaginary parts---
    def model(p):
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(numParams):
            #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return np.append(Zreal, Zimag)
    def diffComplex(params, yVals):
        return (yVals - model(params))/weightCode(params)
    residuals = diffComplex(parametersRes, Z_append)
    residualsStdDevR = np.std(residuals[:len(residuals)//2])
    residualsStdDevJ = np.std(residuals[len(residuals)//2:])
    for i in range(numBootstrap):
        with perVal.get_lock():
            perVal.value += 1
        random_deltas_r = normal(0, residualsStdDevR, len(Z_append)//2)
        random_deltas_j = normal(0, residualsStdDevJ, len(Z_append)//2)
        random_deltas = np.append(random_deltas_r, random_deltas_j)
        yTrial = Z_append + random_deltas
        minimized = lmfit.minimize(diffComplex, parameters, args=(yTrial,))
        if (minimized.success):
            
            fitted = minimized.params.valuesdict()
            toAppendResults = []
            for j in range(numParams):
                toAppendResults.append(fitted[paramNames[j]])
            sharedList.append(toAppendResults)

def mp_imag(sharedList, numBootstrap, perVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams):
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
    #---Used to calculate weights---
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
    def modelImag(p):
        p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(numParams):
            #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        #print(Zreal)
        return Zimag
    def diffImag(params, yVals):
        return (yVals - modelImag(params))/weightCodeHalf(params)
    residuals = diffImag(parametersRes, ZjIn)
    residualsStdDev = np.std(residuals)
    for i in range(numBootstrap):
        with perVal.get_lock():
            perVal.value += 1
        random_deltas = normal(0, residualsStdDev, len(ZjIn))
        yTrial = ZjIn + random_deltas
        minimized = lmfit.minimize(diffImag, parameters, args=(yTrial,))
        if (minimized.success):
            fitted = minimized.params.valuesdict()
            toAppendResults = []
            for j in range(numParams):
                toAppendResults.append(fitted[paramNames[j]])
            sharedList.append(toAppendResults)

def mp_real(sharedList, numBootstrap, perVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams):
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
    #---Used to calculate weights---
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
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(numParams):
            #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        #print(Zreal)
        return Zreal
    def diffReal(params, yVals):
        return (yVals - modelReal(params))/weightCodeHalf(params)
    residuals = diffReal(parametersRes, ZrIn)
    residualsStdDev = np.std(residuals)
    for i in range(numBootstrap):
        with perVal.get_lock():
            perVal.value += 1
        random_deltas = normal(0, residualsStdDev, len(ZrIn))
        yTrial = ZrIn + random_deltas
        minimized = lmfit.minimize(diffReal, parameters, args=(yTrial,))
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
        self.keepGoing = False
        for process in self.processes:
            process.terminate()
    
    def findFit(self, listPercent, numBootstrap, r_in, s_in, sdR_in, sdI_in, chi_in, aic_in, realF_in, imagF_in, fitType, numMonteCarlo, weight, assumedNoise, formula, w, ZrIn, ZjIn, paramNames, paramGuesses, paramBounds, errorParams):
        np.random.seed(1234)        #Use constant seed to ensure reproducibility of Monte Carlo simulations
        Z_append = np.append(ZrIn, ZjIn)
        numParams = len(paramNames)
        
        def diffComplex(params, yVals):
            if (not self.keepGoing):
                return
            return (yVals - model(params))/weightCode(params)
        
        def diffImag(params, yVals):
            if (not self.keepGoing):
                return
            return (yVals - modelImag(params))/weightCodeHalf(params)
        
        def diffReal(params, yVals):
            if (not self.keepGoing):
                return
            return (yVals - modelReal(params))/weightCodeHalf(params)
        
        #---Used to calculate weights---
        def weightCode(p):
            V = np.ones(len(w))     #Default - no weighting
            if (weight == 1):       #Modulus weighting
                for i in range(len(V)):
                    V[i] = assumedNoise*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2)
            elif (weight == 2):     #Proportional weighting
                Vj = np.ones(len(w))
                for i in range(len(V)):
                    if (fitType != 1):          #If the fit type isn't imaginary use both Zr and Zj
                        V[i] = assumedNoise*ZrIn[i]
                    else:                       #If the fit is imaginary use only Zj in weighting
                        V[i] = assumedNoise*ZjIn[i]
                    Vj[i] = assumedNoise*ZjIn[i]
            elif (weight == 3):     #Error structure weighting
                for i in range(len(V)):
                    V[i] = errorParams[0]*ZjIn[i] + (errorParams[1]*ZrIn[i] - errorParams[2]) + errorParams[3]*np.sqrt(ZrIn[i]**2 + ZjIn[i]**2) + errorParams[4]
            elif (weight == 4):     #Custom weighting
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
        
        #---Model used in complex fitting; returns appendation of real and imaginary parts---
        def model(p):
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = [] #np.zeros(len(w))
            Zimag = [] #np.zeros(len(w))
            freq = w.copy()
            for i in range(numParams):
                #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
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
            Zreal = [] #np.zeros(len(w))
            Zimag = [] #np.zeros(len(w))
            freq = w.copy()
            for i in range(numParams):
                #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            #print(Zreal)
            return Zimag
    
        #---Model used in real fitting; returns only real part---
        def modelReal(p):
            p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = [] #np.zeros(len(w))
            Zimag = [] #np.zeros(len(w))
            freq = w.copy()
            for i in range(numParams):
                #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
                exec(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            ldict = locals()
            exec(formula, globals(), ldict)
            Zreal = ldict['Zreal']
            Zimag = ldict['Zimag']
            #print(Zreal)
            return Zreal
        
        if (not self.keepGoing):
            return
        
        parameters = lmfit.Parameters()
        parameterCorrector = 0
        for i in range(numParams):
            parameters.add(paramNames[i], value=paramGuesses[i])
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
            if (numBootstrap <= 1000):
                #---Calculate residuals---
                if (fitType == 3):
                    parametersRes = lmfit.Parameters()
                    for i in range(numParams):
                        parametersRes.add(paramNames[i], value=r_in[i])
                    residuals = diffComplex(parametersRes, Z_append)
                    residualsStdDevR = np.std(residuals[:len(residuals)//2])
                    residualsStdDevJ = np.std(residuals[len(residuals)//2:])
                    paramResultArray = []
                    for i in range(numBootstrap):
                        listPercent.append(1)
                        #print(len(listPercent))
                        random_deltas_r = normal(0, residualsStdDevR, len(Z_append)//2)
                        random_deltas_j = normal(0, residualsStdDevJ, len(Z_append)//2)
                        random_deltas = np.append(random_deltas_r, random_deltas_j)
                        yTrial = Z_append + random_deltas
                        minimized = lmfit.minimize(diffComplex, parameters, args=(yTrial,))
                        if (minimized.success):
                            #print(minimized.params)
                            fitted = minimized.params.valuesdict()
                            toAppendResults = []
                            for j in range(numParams):
                                toAppendResults.append(fitted[paramNames[j]])
                            paramResultArray.append(toAppendResults)
                    sigma = np.std(paramResultArray, axis=0)
                elif (fitType == 2):
                    parametersRes = lmfit.Parameters()
                    for i in range(numParams):
                        parametersRes.add(paramNames[i], value=r_in[i])
                    residuals = diffImag(parametersRes, ZjIn)
                    residualsStdDev = np.std(residuals)
                    paramResultArray = []
                    for i in range(numBootstrap):
                        listPercent.append(1)
                        #print(len(listPercent))
                        random_deltas = normal(0, residualsStdDev, len(ZjIn))
                        yTrial = ZjIn + random_deltas
                        minimized = lmfit.minimize(diffImag, parameters, args=(yTrial,))
                        if (minimized.success):
                            #print(minimized.params)
                            fitted = minimized.params.valuesdict()
                            toAppendResults = []
                            for j in range(numParams):
                                toAppendResults.append(fitted[paramNames[j]])
                            paramResultArray.append(toAppendResults)
                    sigma = np.std(paramResultArray, axis=0)
                elif (fitType == 1):
                    parametersRes = lmfit.Parameters()
                    for i in range(numParams):
                        parametersRes.add(paramNames[i], value=r_in[i])
                    residuals = diffReal(parametersRes, ZrIn)
                    residualsStdDev = np.std(residuals)
                    paramResultArray = []
                    for i in range(numBootstrap):
                        listPercent.append(1)
                        #print(len(listPercent))
                        random_deltas = normal(0, residualsStdDev, len(ZrIn))
                        yTrial = ZrIn + random_deltas
                        minimized = lmfit.minimize(diffReal, parameters, args=(yTrial,))
                        if (minimized.success):
                            #print(minimized.params)
                            fitted = minimized.params.valuesdict()
                            toAppendResults = []
                            for j in range(numParams):
                                toAppendResults.append(fitted[paramNames[j]])
                            paramResultArray.append(toAppendResults)
                    sigma = np.std(paramResultArray, axis=0)
            else:   #Use multiprocessing if more than a certain number of trials requested
                numCores = mp.cpu_count()
                manager = mp.Manager()
                vals = manager.list()
                percentVal = mp.Value("i", 0)
                numRunsPer = int(np.ceil(numBootstrap/numCores))
                #for i in range(numCores):
                #    vals.append([])
                if (fitType == 3):
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        self.processes.append(mp.Process(target=mp_complex, args=(vals, numRunsPer, percentVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams)))
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
                    sigma = np.std(vals, axis=0)
                elif (fitType == 2):
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        self.processes.append(mp.Process(target=mp_imag, args=(vals, numRunsPer, percentVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams)))
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
                    sigma = np.std(vals, axis=0)
                elif (fitType == 1):
                    for i in range(numCores):
                        if (not self.keepGoing):
                            return
                        self.processes.append(mp.Process(target=mp_real, args=(vals, numRunsPer, percentVal, numParams, paramNames, r_in, Z_append, paramGuesses, paramBounds, w, weight, assumedNoise, ZrIn, ZjIn, fitType, formula, errorParams)))
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
                    sigma = np.std(vals, axis=0)
        except SystemExit:
            return "@", "@", "@", "@", "@", "@", "@", "@"
        except Exception as inst:
            #print(inst)
            return "b", "b", "b", "b", "b", "b", "b", "b"
        
         #---Calculate newly fitted values---
        ToReturnReal = []
        ToReturnImag = []
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(r_in[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        ToReturnReal = Zreal
        ToReturnImag = Zimag
        
        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        randomParams = np.zeros((numParams, numMonteCarlo))
        randomlyCalculatedReal = np.zeros((len(w), numMonteCarlo))
        randomlyCalculatedImag = np.zeros((len(w), numMonteCarlo))
        
        for i in range(numParams):
            randomParams[i] = normal(r_in[i], abs(sigma[i]), numMonteCarlo)
        
        if (not self.keepGoing):
            return
        
        #---Calculate impedance values based on the random paramaters---
        for i in range(numMonteCarlo):
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = [] #np.zeros(len(w))
            Zimag = [] #np.zeros(len(w))
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
        
        return r_in, sigma, standardDevsReal, standardDevsImag, chi_in, aic_in, ToReturnReal, ToReturnImag
        """
        if (not minimized.success):     #If the fitting fails
            return "^", "^", "^", "^", "^", "^", "^", "^"
    
        fitted = minimized.params.valuesdict()
        result = []
        sigma = []
        for i in range(numParams):
            result.append(fitted[paramNames[i]])
            sigma.append(minimized.params[paramNames[i]].stderr)
        
        #print(result)
        #---Calculate newly fitted values---
        ToReturnReal = []
        ToReturnImag = []
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = []
        Zimag = []
        freq = w.copy()
        for i in range(numParams):
            exec(paramNames[i] + " = " + str(float(result[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        ToReturnReal = Zreal
        ToReturnImag = Zimag
        
        if None in sigma:
            if (ToReturnReal[0] == 1E300):
                return "b", "b", "b", "b", "b", "b", "b", "b"
            return result, "-", "-", "-", minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag
        
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
            Zreal = [] #np.zeros(len(w))
            Zimag = [] #np.zeros(len(w))
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
        
        return result, sigma, standardDevsReal, standardDevsImag, minimized.chisqr, minimized.aic, ToReturnReal, ToReturnImag
    """