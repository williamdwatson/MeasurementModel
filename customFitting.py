# -*- coding: utf-8 -*-
"""
Created on Tue Oct  8 16:19:24 2019

@author: willdwat
"""

import numpy as np
#from scipy.optimize import least_squares
from numpy.random import normal
import lmfit
#import matplotlib.pyplot as plt
#import time
import cmath

class customFitter:
    def __init__(self):
        self.keepGoing = True
    
    def terminateProcesses(self):
        self.keepGoing = False
    
    def findFit(self, fitType, numMonteCarlo, weight, assumedNoise, formula, w, ZrIn, ZjIn, paramNames, paramGuesses, paramBounds, errorParams):
        np.random.seed(1234)        #Use constant seed to ensure reproducibility of Monte Carlo simulations
        Z_append = np.append(ZrIn, ZjIn)
        numParams = len(paramNames)
        
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
            return np.append(Zimag, Zreal)
         
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
        
        if (not self.keepGoing):
            return

        try:
            if (fitType == 3):      #Complex fitting
                minimized = lmfit.minimize(diffComplex, parameters)
            elif (fitType == 2):    #Imaginary fitting
                minimized = lmfit.minimize(diffImag, parameters)
            elif (fitType == 1):    #Real fitting
                minimized = lmfit.minimize(diffReal, parameters)
        except SystemExit:
            return "@", "@", "@", "@", "@", "@", "@", "@"
        except Exception as inst:
            #print(inst)
            return "b", "b", "b", "b", "b", "b", "b", "b"
      
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
        cor = np.zeros((numParams, numParams))
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
        
    #    #---Replace "fitted" values for fixed parameters with their initial guesses---
    #    if (constants[0] != -1):
    #        for constant in constants:
    #            result[constant] = initialGuesses[constant]
                
        #---Return everything---
        return result, sigma, standardDevsReal, standardDevsImag, minimized.chisqr, cor, minimized.aic
        """