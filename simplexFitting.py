# -*- coding: utf-8 -*-
"""
Created on Tue May  5 18:08:43 2020

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
#from numpy.random import normal
#import lmfit
#import cmath
#import itertools
#import multiprocessing as mp
#import threading
import scipy.optimize as opt

def findFit(fitType, weight, assumedNoise, formula, w, ZrIn, ZjIn, paramNames, paramGuesses, errorParams):
    Z_append = np.append(ZrIn, ZjIn)
    numParams = len(paramNames)
    
    def diffComplex(params):
        return np.sum(((Z_append - model(params))**2)/(weightCode(params)**2))
    
    def diffImag(params):
        return np.sum(((ZjIn - modelImag(params))**2)/(weightCodeHalf(params)**2))
    
    def diffReal(params):
        return np.sum(((ZrIn - modelReal(params))**2)/(weightCodeHalf(params)**2))
    
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
            #p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[i])))
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
                if (fitType != 1):
                    V[i] = assumedNoise*ZrIn[i]
                else:
                    V[i] = assumedNoise*ZjIn[i]
        elif (weight == 3): #Error structure weighting
            pass
        elif (weight == 4): #Custom weighting
            #p.valuesdict()
            Zr = ZrIn.copy()
            Zj = ZjIn.copy()
            Zreal = []
            Zimag = []
            weighting = []
            freq = w.copy()
            for i in range(numParams):
                exec(paramNames[i] + " = " + str(float(p[i])))
            ldict = locals()
            exec(formula, globals(), ldict)
            V = ldict['weighting']
        return V
    
    #---Model used in complex fitting; returns appendation of real and imaginary parts---
    def model(p):
        #p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(numParams):
            #print(paramNames[i] + " = " + str(float(paramGuesses[i])))
            exec(paramNames[i] + " = " + str(float(p[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        return np.append(Zreal, Zimag)
     
    #---Model used in imaginary fitting; returns only imaginary part---
    def modelImag(p):
        #p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(numParams):
            #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            exec(paramNames[i] + " = " + str(float(p[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        #print(Zreal)
        return Zimag

    #---Model used in real fitting; returns only real part---
    def modelReal(p):
        #p.valuesdict()
        Zr = ZrIn.copy()
        Zj = ZjIn.copy()
        Zreal = [] #np.zeros(len(w))
        Zimag = [] #np.zeros(len(w))
        freq = w.copy()
        for i in range(numParams):
            #print(paramNames[i] + " = " + str(float(p[paramNames[i]])))
            exec(paramNames[i] + " = " + str(float(p[i])))
        ldict = locals()
        exec(formula, globals(), ldict)
        Zreal = ldict['Zreal']
        Zimag = ldict['Zimag']
        #print(Zreal)
        return Zreal
    
    if (fitType == 3):
        return opt.minimize(diffComplex, paramGuesses, method='Nelder-Mead', options={'maxiter': numParams})