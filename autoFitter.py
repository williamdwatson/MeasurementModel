# -*- coding: utf-8 -*-
"""
Created on Sun Mar 22 21:54:03 2020

@author: William
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

def mp_complex(guesses, sharedList, index, parameters, numVoigtElements, Z_append, V_append, w):
    mp_current_min = 1E9
    mp_current_best = 0
    def diffComplex(params):
        return (Z_append - model(params))/V_append
    def model(p):
        p.valuesdict()
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

def mp_imag(guesses, sharedList, index, parameters, numVoigtElements, Zj, V, w):
    mp_current_min = 1E9
    mp_current_best = 0
    def diffImag(params):
        return (Zj - modelImag(params))/V
    def modelImag(p):
        p.valuesdict()
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
                        if (fittedMin.chisqr < current_min and fittedMin.success):
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

def mp_real(guesses, sharedList, index, parameters, numVoigtElements, Zr, V, w):
    mp_current_min = 1E9
    mp_current_best = 0
    def diffReal(params):
        return (Zr - modelReal(params))/V
    def modelReal(p):
        p.valuesdict()
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
                        if (fittedMin.chisqr < current_min and fittedMin.success):
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

class autoFitter:
    def __init__(self):#, w, Zr, Zj, numVoigtElements, numMonteCarlo, choice, assumed_noise, Rm, fitType, lowerBounds, initialGuesses, constants, upperBounds, *error_params):
        self.processes = []
        self.keepGoing = True
    
    def terminateProcesses(self):
        self.keepGoing = False
        for process in self.processes:
            process.terminate()
    
    def autoFit(self, maxNVE, numMonteCarlo, w, Zr, Zj, listPercent, choice, autoType, *error_params):
        bestResults = []
        bestResultValues = []
        bestNVE = 1
        currentFitType = 0
        currentFitWeighting = choice
        tauBegin = 1/((max(w)-min(w))/(np.log10(abs(max(w)/min(w)))))
        rBegin = (max(Zr)+min(Zr))/2
        cBegin = 1/(Zj[0] * -2 * np.pi * min(w))
        didntWork = False
        for nve in range(1, maxNVE+1):
            listPercent.append(str(nve))
            if not self.keepGoing:
                return
            bL = np.zeros(nve*2 + 1)
            bU = np.full(nve*2 + 1, np.inf)
            bL[0] = np.NINF
            for i in range(1, len(bL)-1, 2):
                bL[i] = np.NINF
            rGuess = rBegin
            tauGuess = tauBegin
            if (nve > 1):
                rGuess = 0
                tauGuess = 0
                for i in range(1, len(bestResultValues), 2):
                    rGuess += bestResultValues[i]
                    if (bestResultValues[i+1] == 0):
                        tauGuess += np.log10(tauBegin)
                    else:
                        tauGuess += np.log10(bestResultValues[i+1])
                rGuess /= (len(bestResultValues)-1)/2
                tauGuess /= (len(bestResultValues)-1)/2
                tauGuess = 10**tauGuess
            ig = []
            if (nve > 1):
                for i in range(len(bestResultValues)):
                    ig.append([bestResultValues[i]])
            else:
                ig.append([rGuess])
            ig.append([rGuess])
            ig.append([tauGuess])
            #---Try default (complex fit, modulus weighting)---
            r = self.findFit(w, Zr, Zj, nve, choice, 1, 1, autoType, listPercent, bL, ig, [-1], bU, error_params[0], error_params[1], error_params[2], error_params[3], error_params[4])
            if (r == "^" or r == "-" or r == "$"):
                #---Try multistart (complex fit, modulus weighting)---
                listPercent.append("b" + str(nve))
                ig[len(ig)-1] = np.logspace(-5, 5, 10)
                r2 = self.findFit(w, Zr, Zj, nve, choice, 1, 1, autoType, listPercent, bL, ig, [-1], bU, error_params[0], error_params[1], error_params[2], error_params[3], error_params[4])
                if (r2 == "^" or r2 == "-" or r2 == "$"):
                    if didntWork:
                        return "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^" 
                    elif (nve > 5 or nve == maxNVE):
                        break
                    elif (nve == 1):
                        bestResultValues = [rBegin, rBegin, tauBegin]
                        didntWork = True
                        continue
                    else:
                        bestResultValues.append(rBegin)
                        bestResultValues.append(tauBegin)
                        didntWork = True
                        continue
                else:
                    didntWork = False
                    bestResults = r2
                    bestNVE = nve
                    currentFitType = 0
                    currentFitWeighting = 1
                    fitted = r2.params.valuesdict()
                    bestResultValues = [fitted['Re']]
                    for i in range(1, nve+1):
                        bestResultValues.append(fitted['R'+str(i)])
                        bestResultValues.append(fitted['T'+str(i)])
            else:
                didntWork = False
                bestResults = r
                bestNVE = nve
                currentFitType = 0
                currentFitWeighting = 1
                fitted = r.params.valuesdict()
                bestResultValues = [fitted['Re']]
                for i in range(1, nve+1):
                    bestResultValues.append(fitted['R'+str(i)])
                    bestResultValues.append(fitted['T'+str(i)])
        
        fitted = bestResults.params.valuesdict()
        result = [fitted['Re']]
        sigma = [bestResults.params['Re'].stderr]
        numParams = 2*bestNVE + 1
        for i in range(1, bestNVE+1):
            result.append(fitted['R'+str(i)])
            result.append(fitted['T'+str(i)])
            sigma.append(bestResults.params['R'+str(i)].stderr)
            sigma.append(bestResults.params['T'+str(i)].stderr)
        
        cor = np.zeros((numParams, numParams))
        if (bestResults.params['Re'].correl == None):
            cor[:,0] = 0
            cor[0,:] = 0
            cor[0][0] = 1
        else:
            cor[0][0] = 1
            parameterCorrector = 0
            for i in range(1, bestNVE+1):
                try:
                    cor[0][i+parameterCorrector] = bestResults.params['Re'].correl['R'+str(i)]
                except KeyError:
                    cor[0][i+parameterCorrector] = 0
                try:
                    cor[0][i+1+parameterCorrector] = bestResults.params['Re'].correl['T'+str(i)]
                except KeyError:
                    cor[0][i+1+parameterCorrector] = 0
                parameterCorrector += 1
        parameterCorrectorA = 0
        for i in range(1, bestNVE+1):
            #---R---
            if (bestResults.params['R'+str(i)].correl == None):
                cor[i+parameterCorrectorA,:] = 0
                cor[i+parameterCorrectorA][i+parameterCorrectorA] = 1
            else:
                try:
                    cor[i+parameterCorrectorA][0] = bestResults.params['R'+str(i)].correl['Re']
                except KeyError:
                    cor[i+parameterCorrectorA][0] = 0
                cor[i+parameterCorrectorA][i+parameterCorrectorA] = 1
                parameterCorrectorB = 0
                for j in range(1, bestNVE+1):
                    
                    if (i != j):
                        try:
                            cor[i+parameterCorrectorA][j+parameterCorrectorB] = bestResults.params['R'+str(i)].correl['R'+str(j)]
                        except KeyError:
                            cor[i+parameterCorrectorA][j+parameterCorrectorB] = 0
                    try:
                        cor[i+parameterCorrectorA][j+1+parameterCorrectorB] = bestResults.params['R'+str(i)].correl['T'+str(j)]
                    except KeyError:
                        cor[i+parameterCorrectorA][j+1+parameterCorrectorB] = 0
                    parameterCorrectorB += 1
                    
            #---Tau---
            if (bestResults.params['T'+str(i)].correl == None):
                cor[i+1+parameterCorrectorA,:] = 0
                cor[i+1+parameterCorrectorA][i+1+parameterCorrectorA] = 1
            else:
                try:
                    cor[i+1+parameterCorrectorA][0] = bestResults.params['T'+str(i)].correl['Re']
                except KeyError:
                    cor[i+1+parameterCorrectorA][0] = 0
                cor[i+1+parameterCorrectorA][i+1+parameterCorrectorA] = 1
                
                parameterCorrectorB = 0
                for j in range(1, bestNVE+1):         
                    try:
                        cor[i+1+parameterCorrectorA][j+parameterCorrectorB] = bestResults.params['T'+str(i)].correl['R'+str(j)]
                    except KeyError:
                        cor[i+1+parameterCorrectorA][j+parameterCorrectorB] = 0
                    if (i != j):
                        try:
                            cor[i+1+parameterCorrectorA][j+1+parameterCorrectorB] = bestResults.params['T'+str(i)].correl['T'+str(j)]
                        except KeyError:
                            cor[i+1+parameterCorrectorA][j+1+parameterCorrectorB] = 0
                    parameterCorrectorB += 1    
                
            parameterCorrectorA += 1
        listPercent.append("c")
        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        randomParams = np.zeros((numParams, numMonteCarlo))
        randomlyCalculated = np.zeros((len(w), numMonteCarlo), dtype=np.complex128)
        
        for i in range(numParams):
            randomParams[i] = normal(result[i], abs(sigma[i]), numMonteCarlo)
        
        #---Calculate impedance values based on the random paramaters---
        with np.errstate(divide='ignore', invalid='ignore'):
            for i in range(numMonteCarlo):
                for j in range(len(w)):
                    randomlyCalculated[j][i] = randomParams[0][i] + (randomParams[1][i]/(1+(1j*w[j]*2*np.pi*randomParams[2][i])))
                    for k in range(3, numParams, 2):
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
        
    #    #---Replace "fitted" values for fixed parameters with their initial guesses---
    #    if (constants[0] != -1):
    #        for constant in constants:
    #            result[constant] = initialGuesses[constant]
                
        #---Return everything---
        return result, sigma, standardDevsReal, standardDevsImag, Zzero, ZzeroSigma, Zpolar, ZpolarSigma, bestResults.chisqr, cor, bestResults.aic, currentFitType, currentFitWeighting
            
    def findFit(self, w, Zr, Zj, numVoigtElements, choice, assumed_noise, Rm, fitType, listPercent, lowerBounds, initialGuesses, constants, upperBounds, *error_params):
        np.random.seed(1234)        #Use constant seed to ensure reproducibility of Monte Carlo simulations
        constants = np.sort(constants)
        Z_append = np.append(Zr, Zj)
        #numParams = numVoigtElements*2 + 1
        
        V = np.ones(len(w))     #Default no weighting (all weights are 1) if choice==0
        if (choice == 1):       #Modulus weighting
            for i in range(len(V)):
                V[i] = assumed_noise*np.sqrt(Zr[i]**2 + Zj[i]**2)
        elif (choice == 2):     #Proportional weighting
            Vj = np.ones(len(w))
            for i in range(len(V)):
                if (fitType != 1):          #If the fit type isn't imaginary use both Zr and Zj
                    V[i] = assumed_noise*Zr[i]
                else:                       #If the fit is imaginary use only Zj in weighting
                    V[i] = assumed_noise*Zj[i]
                Vj[i] = assumed_noise*Zj[i]
        elif (choice == 3):     #Error model weighting
            alpha = error_params[0]
            beta = error_params[1]
            betaRe = error_params[2]
            gamma = error_params[3]
            delta = error_params[4]
            for i in range(len(V)):
                V[i] = alpha*abs(Zj[i]) + (beta*abs(Zr[i]) - betaRe) + gamma*(np.sqrt(Zr[i]**2 + Zj[i]**2)**2)/Rm + delta
        if (choice != 2):
            V_append = np.append(V, V)
        else:
            V_append = np.append(V, Vj)
        
        def diffComplex(params):
            return (Z_append - model(params))/V_append
        
        def diffImag(params):
            return (Zj - modelImag(params))/V
        
        def diffReal(params):
            return (Zr - modelReal(params))/V
        
        #---Model used in complex fitting; returns appendation of real and imaginary parts---
        def model(p):
            p.valuesdict()
            re = p['Re']
            r1 = p['R1']
            t1 = p['T1']
            actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
            for i in range(2, numVoigtElements+1):
                actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
            return np.append(actual.real, actual.imag)
         
        #---Model used in imaginary fitting; returns only imaginary part---
        def modelImag(p):
            p.valuesdict()
            re = p['Re']
            r1 = p['R1']
            t1 = p['T1']
            actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
            for i in range(2, numVoigtElements+1):
                actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
            return actual.imag
        
        #---Model used in real fitting; returns only real part---
        def modelReal(p):
            p.valuesdict()
            re = p['Re']
            r1 = p['R1']
            t1 = p['T1']
            actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
            for i in range(2, numVoigtElements+1):
                actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
            return actual.real    
        
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
                            fittedMin = lmfit.minimize(diffComplex, parameters)
                            nonlocal current_min
                            nonlocal current_best
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
                            fittedMin = lmfit.minimize(diffImag, parameters)
                            nonlocal current_min
                            nonlocal current_best
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
                            fittedMin = lmfit.minimize(diffReal, parameters)
                            nonlocal current_min
                            nonlocal current_best
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
        if ((numVoigtElements < 8 and numCombos > 200) or (numVoigtElements >= 8 and numCombos > 100)):
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
            for i in range(numCores):
                vals.append([])
            if (fitType == 0):
                for i in range(numCores):
                    if (not self.keepGoing):
                        return
                    self.processes.append(mp.Process(target=mp_complex, args=(toRun[i], vals, i, parameters, numVoigtElements, Z_append, V_append, w)))
                for p in self.processes:
                    if (not self.keepGoing):
                        return
                    p.start()
                for p in self.processes:
                    p.join()
            if (fitType == 1):
                parameters['Re'].vary = False
                for i in range(numCores):
                    if (not self.keepGoing):
                        return
                    self.processes.append(mp.Process(target=mp_imag, args=(toRun[i], vals, i, parameters, numVoigtElements, Zj, V, w)))
                for p in self.processes:
                    if (not self.keepGoing):
                        return
                    p.start()
                for p in self.processes:
                    p.join()
            if (fitType == 2):
                for i in range(numCores):
                    if (not self.keepGoing):
                        return
                    self.processes.append(mp.Process(target=mp_real, args=(toRun[i], vals, i, parameters, numVoigtElements, Zr, V, w)))
                for p in self.processes:
                    if (not self.keepGoing):
                        return
                    p.start()
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
            return "^"
        
        fitted = minimized.params.valuesdict()
        result = [fitted['Re']]
        sigma = [minimized.params['Re'].stderr]
        for i in range(1, numVoigtElements+1):
            result.append(fitted['R'+str(i)])
            result.append(fitted['T'+str(i)])
            sigma.append(minimized.params['R'+str(i)].stderr)
            sigma.append(minimized.params['T'+str(i)].stderr)
        if None in sigma:
            return "-"
        for i in range(len(result)):
            if (result[i] != 0):
                if (abs(2*sigma[i]/result[i]) >= 1.0 or np.isnan(sigma[i])):
                    return "$"
                
        #---Return everything---
        return minimized