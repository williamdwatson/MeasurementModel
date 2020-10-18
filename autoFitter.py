# -*- coding: utf-8 -*-
"""
Created on Sun Mar 22 21:54:03 2020

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

def mp_complex(guesses, sharedList, index, parameters, numVoigtElements, Z_append, V_append, w):
    """
    Function for multiprocessing the complex fit of Voigt elements
    
    Parameters
    ----------
    guesses : list of floats
        Initial guesses for parameters.
    sharedList : list
        List of succesful parameter fits.
    index : int
        The index of the shared list to use
    parameters : lmfit Parameters
        The parameters to be used in fitting
    numVoigtElements : int
        The number of Voigt elements for the current fit
    Z_append : numpy array
        Array of real impedance, then imaginary impedance.
    V_append : numpy array
        Array of weights to use when fitting
    w : array of floats
        Data frequencies.

    Returns
    -------
    None (results are appended to a shared list)

    """
    mp_current_min = 1E9    #The current lowest objective function found
    mp_current_best = 0     #The parameters used to fit mp_current_min
    def diffComplex(params):
        """
        The objective function for complex fitting.
        
        Parameters
        ----------
        params : lmfit Parameters
            The parameters for the current fitting iteration
        
        Returns
        -------
        residuals : array
            Array of the residuals (the known values minus the model values over the weighting)
            
        """
        return (Z_append - model(params))/V_append
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
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
        for i in range(2, numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
        return np.append(actual.real, actual.imag)
    def recursive_fit_mp(voigt_number, recursiveGuesses):
        """
        Recursively adds parameters to recursiveGuesses based on the passed-in initial guesses. When the base case with one Voigt element is reached, lmfit
        is used to attempt to minimize the diffComplex objective function, whereafter the chi-squared value is checked to see if that fit is better than the current best.

        Parameters
        ----------
        voigt_number : int
            The current Voigt number of whatever round of recursion.
        recursiveGuesses : list of floats
            The current parameter guesses.

        Returns
        -------
        None (finds minimum for outer function to return)

        """
        #---The base recursive case with one Voigt element, which attempts to minimize diffComplex and see if that result is better than the previous minimum---
        if (voigt_number == 1):
            for i in range(len(recursiveGuesses[0])):
                for j in range(len(recursiveGuesses[1])):
                    for k in range(len(recursiveGuesses[2])):
                        parameters.get("Re").set(value=recursiveGuesses[0][i])
                        parameters.get("R1").set(value=recursiveGuesses[1][j])
                        parameters.get("T1").set(value=recursiveGuesses[2][k])
                        fittedMin = lmfit.minimize(diffComplex, parameters)     #Perform the actual regression
                        #---Check if the minimization is better than any previous---
                        nonlocal mp_current_min
                        nonlocal mp_current_best
                        if (fittedMin.chisqr < mp_current_min and fittedMin.success):
                            mp_current_min = fittedMin.chisqr
                            mp_current_best = fittedMin
        #---The non-base recursive case, which merely adds more parameters to recursiveGuesses based on the passed-in initial guesses---
        else:
            for z in range(len(recursiveGuesses[2*voigt_number-1])):
                for y in range(len(recursiveGuesses[2*voigt_number])):
                    parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                    parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                    recursive_fit_mp(voigt_number-1, recursiveGuesses)     #Recurse with one fewer Voigt element
    recursive_fit_mp(numVoigtElements, guesses)
    sharedList[index] = [mp_current_min, mp_current_best]

def mp_imag(guesses, sharedList, index, parameters, numVoigtElements, Zj, V, w):
    """
    Function for multiprocessing the imaginary fit of Voigt elements
    
    Parameters
    ----------
    guesses : list of floats
        Initial guesses for parameters.
    sharedList : list
        List of succesful parameter fits.
    index : int
        The index of the shared list to use
    parameters : lmfit Parameters
        The parameters to be used in fitting
    numVoigtElements : int
        The number of Voigt elements for the current fit
    Zj : numpy array
        Array of the imaginary impedance.
    V : numpy array
        Array of weights to use when fitting
    w : array of floats
        Data frequencies.

    Returns
    -------
    None (results are appended to a shared list)
    
    """
    mp_current_min = 1E9    #The current lowest objective function found
    mp_current_best = 0     #The parameters used to fit mp_current_min
    def diffImag(params):
        """
        The objective function for imaginary fitting.
        
        Parameters
        ----------
        params : lmfit Parameters
            The parameters for the current fitting iteration
        
        Returns
        -------
        residuals : array
            Array of the residuals (the known values minus the model values over the weighting)
            
        """
        return (Zj - modelImag(params))/V
    def modelImag(p):
        """
        The model used in imaginary fitting

        Parameters
        ----------
        p : lmfit Parameters
            The parameters for the current fitting iteration.

        Returns
        -------
        model : numpy array
            Numpy array the imaginary impedance.

        """
        p.valuesdict()
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
        for i in range(2, numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
        return actual.imag
    def recursive_fit_imag_mp(voigt_number, recursiveGuesses):
        """
        Recursively adds parameters to recursiveGuesses based on the passed-in initial guesses. When the base case with one Voigt element is reached, lmfit
        is used to attempt to minimize the diffImag objective function, whereafter the chi-squared value is checked to see if that fit is better than the current best.

        Parameters
        ----------
        voigt_number : int
            The current Voigt number of whatever round of recursion.
        recursiveGuesses : list of floats
            The current parameter guesses.

        Returns
        -------
        None (finds minimum for outer function to return)

        """
        #---The base recursive case with one Voigt element, which attempts to minimize diffImag and see if that result is better than the previous minimum---
        if (voigt_number == 1):
            for i in range(len(recursiveGuesses[0])):
                for j in range(len(recursiveGuesses[1])):
                    for k in range(len(recursiveGuesses[2])):
                        parameters.get("Re").set(value=recursiveGuesses[0][i])
                        parameters.get("R1").set(value=recursiveGuesses[1][j])
                        parameters.get("T1").set(value=recursiveGuesses[2][k])
                        fittedMin = lmfit.minimize(diffImag, parameters)     #Perform the actual regression
                        #---Check if the minimization is better than any previous---
                        nonlocal mp_current_min
                        nonlocal mp_current_best
                        if (fittedMin.chisqr < current_min and fittedMin.success):
                            mp_current_min = fittedMin.chisqr
                            mp_current_best = fittedMin
        #---The non-base recursive case, which merely adds more parameters to recursiveGuesses based on the passed-in initial guesses---
        else:
            for z in range(len(recursiveGuesses[2*voigt_number-1])):
                for y in range(len(recursiveGuesses[2*voigt_number])):
                    parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                    parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                    recursive_fit_imag_mp(voigt_number-1, recursiveGuesses)     #Recurse with one fewer Voigt element
    recursive_fit_imag_mp(numVoigtElements, guesses)
    sharedList[index] = [mp_current_min, mp_current_best]

def mp_real(guesses, sharedList, index, parameters, numVoigtElements, Zr, V, w):
    """
    Function for multiprocessing the real fit of Voigt elements
    
    Parameters
    ----------
    guesses : list of floats
        Initial guesses for parameters.
    sharedList : list
        List of succesful parameter fits.
    index : int
        The index of the shared list to use
    parameters : lmfit Parameters
        The parameters to be used in fitting
    numVoigtElements : int
        The number of Voigt elements for the current fit
    Zr : numpy array
        Array of the real impedance.
    V : numpy array
        Array of weights to use when fitting
    w : array of floats
        Data frequencies.

    Returns
    -------
    None (results are appended to a shared list)
    
    """
    mp_current_min = 1E9    #The current lowest objective function found
    mp_current_best = 0     #The parameters used to fit mp_current_min
    def diffReal(params):
        """
        The objective function for real fitting.
        
        Parameters
        ----------
        params : lmfit Parameters
            The parameters for the current fitting iteration
        
        Returns
        -------
        residuals : array
            Array of the residuals (the known values minus the model values over the weighting)
            
        """
        return (Zr - modelReal(params))/V
    def modelReal(p):
        """
        The model used in real fitting

        Parameters
        ----------
        p : lmfit Parameters
            The parameters for the current fitting iteration.

        Returns
        -------
        model : numpy array
            Numpy array the real impedance.

        """
        p.valuesdict()
        re = p['Re']
        r1 = p['R1']
        t1 = p['T1']
        actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
        for i in range(2, numVoigtElements+1):
            actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
        return actual.real 
    def recursive_fit_real_mp(voigt_number, recursiveGuesses):
        """
        Recursively adds parameters to recursiveGuesses based on the passed-in initial guesses. When the base case with one Voigt element is reached, lmfit
        is used to attempt to minimize the diffReal objective function, whereafter the chi-squared value is checked to see if that fit is better than the current best.

        Parameters
        ----------
        voigt_number : int
            The current Voigt number of whatever round of recursion.
        recursiveGuesses : list of floats
            The current parameter guesses.

        Returns
        -------
        None (finds minimum for outer function to return)

        """
        #---The base recursive case with one Voigt element, which attempts to minimize diffReal and see if that result is better than the previous minimum---
        if (voigt_number == 1):
            for i in range(len(recursiveGuesses[0])):
                for j in range(len(recursiveGuesses[1])):
                    for k in range(len(recursiveGuesses[2])):
                        parameters.get("Re").set(value=recursiveGuesses[0][i])
                        parameters.get("R1").set(value=recursiveGuesses[1][j])
                        parameters.get("T1").set(value=recursiveGuesses[2][k])
                        fittedMin = lmfit.minimize(diffReal, parameters)        #Perform the actual regression
                        #---Check if the minimization is better than any previous---
                        nonlocal mp_current_min
                        nonlocal mp_current_best
                        if (fittedMin.chisqr < current_min and fittedMin.success):
                            mp_current_min = fittedMin.chisqr
                            mp_current_best = fittedMin
        #---The non-base recursive case, which merely adds more parameters to recursiveGuesses based on the passed-in initial guesses---
        else:
            for z in range(len(recursiveGuesses[2*voigt_number-1])):
                for y in range(len(recursiveGuesses[2*voigt_number])):
                    parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                    parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                    recursive_fit_real_mp(voigt_number-1, recursiveGuesses)     #Recurse with one fewer Voigt element
    recursive_fit_real_mp(numVoigtElements, guesses)
    sharedList[index] = [mp_current_min, mp_current_best]

class autoFitter:
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
    
    def autoFit(self, maxNVE, fixRe, fixReVal, numMonteCarlo, w, Zr, Zj, listPercent, choice, autoType, *error_params):
        """
        Automatically fit Voigt elements based on the user's choices

        Parameters
        ----------
        maxNVE : int
            The maximum number of Voigt elements to fit.
        fixRe : boolean
            Whether the Ohmic resistance should be fixed.
        fixReVal : float
            The value for the Ohmic resistance if it is fixed.
        numMonteCarlo : int
            The number of Monte Carlo simulations to perform.
        w : list of floats
            The frequency data.
        Zr : list of floats
            The real impedance data.
        Zj : list of floats
            The imaginary impedance data.
        listPercent : list
            List used to communicate with the GUI.
        choice : int
            The weighting type: 0 for unity (no weighting), 1 for modulus, 2 for proportional, 3 for error structure.
        autoType : int
            The fitting type: 0 for complex, 1 for imaginary, 2 for real.
        *error_params : list of floats
            The error parameters if using error structure weighting.
        
        Returns
        -------
        result : list of floats
            Fitted parameters in order: [Re, R1, T1, R2, T2, ...].
        sigma : list of floats
            Standard errors for each parameter.
        standardDevsReal : list of floats
            Real Monte Carlo results.
        standardDevsImag : list of floats
            Imaginary Monte Carlo results.
        Zzero : float
            Impedance at 0 frequency.
        ZzeroSigma : float
            Error of impedance at 0 frequency.
        Zpolar : float
            Polarization impedance.
        ZpolarSigma : float
            Error of polarization impedance.
        bestResults.chisqr : float
            Chi-squared statistic for the fit.
        cor : 2D list
            Correlation matrix.
        bestResults.aic : float
            Akaike Information Criterion for the fit.
        currentFitType : int
            Which fit type was used.
        currentFitWeighting : int
            Which weighting was used.
        
        """
        bestResults = []
        bestResultValues = []
        bestNVE = 1
        currentFitType = 0
        currentFitWeighting = choice
        #---The default initial guesses---
        tauBegin = 1/((max(w)-min(w))/(np.log10(abs(max(w)/min(w)))))
        rBegin = (max(Zr)+min(Zr))/2
        cBegin = 1/(Zj[0] * -2 * np.pi * min(w))
        didntWork = False
        for nve in range(1, maxNVE+1):                                 #Loop up to the max number of Voigt elements
            listPercent.append(str(nve))                               #To communicate to the GUI the current Voigt element being tried
            if not self.keepGoing:                                     #Check if the fitting has been cancelled each loop
                return
            #---Set the parameter bounds---
            bL = np.zeros(nve*2 + 1)
            bU = np.full(nve*2 + 1, np.inf)
            bL[0] = np.NINF
            for i in range(1, len(bL)-1, 2):
                bL[i] = np.NINF
            rGuess = rBegin
            tauGuess = tauBegin
            #---Update the initial guesses on each loop based on the newly fit parameters---
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
            r = self.findFit(fixRe, fixReVal, w, Zr, Zj, nve, choice, 1, 1, autoType, listPercent, bL, ig, [-1], bU, error_params[0], error_params[1], error_params[2], error_params[3], error_params[4])
            if (r == "^" or r == "-" or r == "$"):
                #---Try multistart (complex fit, modulus weighting)---
                listPercent.append("b" + str(nve))          #Indicate trying multistart to the GUI
                ig[len(ig)-1] = np.logspace(-5, 5, 10)
                r2 = self.findFit(fixRe, fixReVal, w, Zr, Zj, nve, choice, 1, 1, autoType, listPercent, bL, ig, [-1], bU, error_params[0], error_params[1], error_params[2], error_params[3], error_params[4])
                if (r2 == "^" or r2 == "-" or r2 == "$"):
                    #---If the fit fails again, return failure---
                    if didntWork:
                        return "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^", "^"
                    #---If the fit fails again and we've reached the max number of elements or we have more than 5 currently done, stop trying to fit
                    elif (nve > 5 or nve == maxNVE):
                        break
                    #---If this is the first fit, try again from 2 elements---
                    elif (nve == 1):
                        bestResultValues = [rBegin, rBegin, tauBegin]
                        didntWork = True
                        continue
                    #---If fit 2, 3, or 4 fails, try again from 1 extra element---
                    else:
                        bestResultValues.append(rBegin)
                        bestResultValues.append(tauBegin)
                        didntWork = True
                        continue
                #---If the multistart succeeds, store its results and keep going---
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
            #---If the standard fit succeeds, store its results and keep going---
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
        #---After all fits are done, prepare the results---
        fitted = bestResults.params.valuesdict()
        result = [fitted['Re']]
        sigma = [bestResults.params['Re'].stderr]
        numParams = 2*bestNVE + 1
        for i in range(1, bestNVE+1):
            result.append(fitted['R'+str(i)])
            result.append(fitted['T'+str(i)])
            sigma.append(bestResults.params['R'+str(i)].stderr)
            sigma.append(bestResults.params['T'+str(i)].stderr)
        #---Create the correlation matrix---
        cor = np.zeros((numParams, numParams))
        #---Ohmic resistance correlations---
        if (bestResults.params['Re'].correl == None):   #If Re has no correlations, set it to be 1 with itself and zero elsewhere
            cor[:,0] = 0
            cor[0,:] = 0
            cor[0][0] = 1
        else:
            cor[0][0] = 1   #Correlation of Re with itself is 1
            parameterCorrector = 0  #Use to keep indices in order (as each fit takes two spots: one for R and one for T)
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
        #---Other parameter correlations---
        parameterCorrectorA = 0     #Use to keep indices in order (as each fit takes two spots: one for R and one for T)
        for i in range(1, bestNVE+1):
            #---R---
            if (bestResults.params['R'+str(i)].correl == None):         #If there's not correlation coefficient, set it to 0
                cor[i+parameterCorrectorA,:] = 0
                cor[i+parameterCorrectorA][i+parameterCorrectorA] = 1   #Correlation of a parameter with itself is 1
            else:
                try:
                    cor[i+parameterCorrectorA][0] = bestResults.params['R'+str(i)].correl['Re']
                except KeyError:                                        #If there's not correlation coefficient, set it to 0
                    cor[i+parameterCorrectorA][0] = 0
                cor[i+parameterCorrectorA][i+parameterCorrectorA] = 1   #Correlation of a parameter with itself is 1
                parameterCorrectorB = 0  #Use to keep indices in order (as each fit takes two spots: one for R and one for T)
                for j in range(1, bestNVE+1):
                    if (i != j):
                        try:
                            cor[i+parameterCorrectorA][j+parameterCorrectorB] = bestResults.params['R'+str(i)].correl['R'+str(j)]
                        except KeyError:
                            cor[i+parameterCorrectorA][j+parameterCorrectorB] = 0     #If there's not correlation coefficient, set it to 0
                    try:
                        cor[i+parameterCorrectorA][j+1+parameterCorrectorB] = bestResults.params['R'+str(i)].correl['T'+str(j)]
                    except KeyError:
                        cor[i+parameterCorrectorA][j+1+parameterCorrectorB] = 0       #If there's not correlation coefficient, set it to 0
                    parameterCorrectorB += 1
                    
            #---Tau---
            if (bestResults.params['T'+str(i)].correl == None):
                cor[i+1+parameterCorrectorA,:] = 0                              #If there's not correlation coefficient, set it to 0
                cor[i+1+parameterCorrectorA][i+1+parameterCorrectorA] = 1       #Correlation of a parameter with itself is 1
            else:
                try:
                    cor[i+1+parameterCorrectorA][0] = bestResults.params['T'+str(i)].correl['Re']
                except KeyError:
                    cor[i+1+parameterCorrectorA][0] = 0                         #If there's not correlation coefficient, set it to 0
                cor[i+1+parameterCorrectorA][i+1+parameterCorrectorA] = 1       #Correlation of a parameter with itself is 1
                
                parameterCorrectorB = 0  #Use to keep indices in order (as each fit takes two spots: one for R and one for T)
                for j in range(1, bestNVE+1):         
                    try:
                        cor[i+1+parameterCorrectorA][j+parameterCorrectorB] = bestResults.params['T'+str(i)].correl['R'+str(j)]
                    except KeyError:
                        cor[i+1+parameterCorrectorA][j+parameterCorrectorB] = 0             #If there's not correlation coefficient, set it to 0
                    if (i != j):
                        try:
                            cor[i+1+parameterCorrectorA][j+1+parameterCorrectorB] = bestResults.params['T'+str(i)].correl['T'+str(j)]
                        except KeyError:
                            cor[i+1+parameterCorrectorA][j+1+parameterCorrectorB] = 0       #If there's not correlation coefficient, set it to 0
                    parameterCorrectorB += 1    
                
            parameterCorrectorA += 1
        listPercent.append("c")         #Tell the GUI we're performing Monte Carlo simulations
        #---Calculate numMonteCarlo number of parameters, using a Gaussian distribution---
        randomParams = np.zeros((numParams, numMonteCarlo))
        randomlyCalculated = np.zeros((len(w), numMonteCarlo), dtype=np.complex128)
        
        for i in range(numParams):
            randomParams[i] = normal(result[i], abs(sigma[i]), numMonteCarlo)
        
        #---Calculate impedance values based on the random parameters---
        with np.errstate(divide='ignore', invalid='ignore'):    #Ignore invalid or nan issues
            #---For every frequency at every Monte Carlo simulation, calculate a new impedance based on the random parameters---
            for i in range(numMonteCarlo):
                for j in range(len(w)):
                    randomlyCalculated[j][i] = randomParams[0][i] + (randomParams[1][i]/(1+(1j*w[j]*2*np.pi*randomParams[2][i])))
                    for k in range(3, numParams, 2):
                        randomlyCalculated[j][i] += (randomParams[k][i]/(1+(1j*w[j]*2*np.pi*randomParams[k+1][i])))
        #---Calculate the standard deviations from the Monte Carlo simulations---
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
        return result, sigma, standardDevsReal, standardDevsImag, Zzero, ZzeroSigma, Zpolar, ZpolarSigma, bestResults.chisqr, cor, bestResults.aic, currentFitType, currentFitWeighting
            
    def findFit(self, fixed, fixedVal, w, Zr, Zj, numVoigtElements, choice, assumed_noise, Rm, fitType, listPercent, lowerBounds, initialGuesses, constants, upperBounds, *error_params):
        """
        Finds the parameters that minimize the desired objective function for a round of auto-fitting

        Parameters
        ----------
        fixed : boolean
            Whether the Ohmic resistance should be fixed.
        fixedVal : float
            The value for the Ohmic resistance if it is fixed.
        w : array of floats
            Array of frequency data.
        Zr : array of floats
            Array of real impedance data.
        Zj : array of floats
            Array of imaginary impedance data.
        numVoigtElements : int
            Number of Voigt elements to fit.
        choice : int
            The weighting type: 0 for unity (no weighting), 1 for modulus, 2 for proportional, 3 for error structure.
        assumed_noise : float
            The assumed noise (used to multiply the weighting).
        Rm : float
            The current-measuring resistor, fixed to 1.
        fitType : int
            The fitting type: 0 for complex, 1 for imaginary, 2 for real.
        listPercent : list
            List used to communicate with the GUI.
        lowerBounds : list of floats
            The lower bounds for each parameter.
        initialGuesses : list of floats
            Initial guesses for the parameters.
        constants : list of ints
            Parameters to hold constant.
        upperBounds : list of floats
            The upper bounds for each parameter.
        *error_params : list of floats
            The error parameters if using error structure weighting.

        Returns
        -------
        minimized
            The lmfit minimization, or a specific string signaling an error.

        """
        np.random.seed(1234)                #Use constant seed to ensure reproducibility of Monte Carlo simulations
        constants = np.sort(constants)
        Z_append = np.append(Zr, Zj)
        
        V = np.ones(len(w))                 #Default no weighting (all weights are 1) if choice==0
        if (choice == 1):                   #Modulus weighting
            for i in range(len(V)):
                V[i] = assumed_noise*np.sqrt(Zr[i]**2 + Zj[i]**2)
        elif (choice == 2):                 #Proportional weighting
            Vj = np.ones(len(w))
            for i in range(len(V)):
                if (fitType != 1):          #If the fit type isn't imaginary use both Zr and Zj
                    V[i] = assumed_noise*Zr[i]
                else:                       #If the fit is imaginary use only Zj in weighting
                    V[i] = assumed_noise*Zj[i]
                Vj[i] = assumed_noise*Zj[i]
        elif (choice == 3):                 #Error model weighting
            alpha = error_params[0]
            beta = error_params[1]
            betaRe = error_params[2]
            gamma = error_params[3]
            delta = error_params[4]
            for i in range(len(V)):
                V[i] = alpha*abs(Zj[i]) + (beta*abs(Zr[i]) - betaRe) + gamma*(np.sqrt(Zr[i]**2 + Zj[i]**2)**2)/Rm + delta
        if (choice != 2):               #If weighting isn't proportional, then weighting is the same for real and imaginary
            V_append = np.append(V, V)
        else:                           #If weighting is proportional, it will differ between real and imaginary
            V_append = np.append(V, Vj)
        
        def diffComplex(params):
            """
            The objective function for complex fitting.
            
            Parameters
            ----------
            params : lmfit Parameters
                The parameters for the current fitting iteration
            
            Returns
            -------
            residuals : array
                Array of the residuals (the known values minus the model values over the weighting)
                
            """
            return (Z_append - model(params))/V_append
        
        def diffImag(params):
            """
            The objective function for imaginary fitting.
            
            Parameters
            ----------
            params : lmfit Parameters
                The parameters for the current fitting iteration
            
            Returns
            -------
            residuals : array
                Array of the residuals (the known values minus the model values over the weighting)
                
            """
            return (Zj - modelImag(params))/V
        
        def diffReal(params):
            """
            The objective function for real fitting.
            
            Parameters
            ----------
            params : lmfit Parameters
                The parameters for the current fitting iteration
            
            Returns
            -------
            residuals : array
                Array of the residuals (the known values minus the model values over the weighting)
                
            """
            return (Zr - modelReal(params))/V
        
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
            re = p['Re']
            r1 = p['R1']
            t1 = p['T1']
            actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
            for i in range(2, numVoigtElements+1):
                actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
            return np.append(actual.real, actual.imag)
         
        def modelImag(p):
            """
            The model used in imaginary fitting
    
            Parameters
            ----------
            p : lmfit Parameters
                The parameters for the current fitting iteration.
    
            Returns
            -------
            model : numpy array
                Numpy array the imaginary impedance.
    
            """
            p.valuesdict()
            re = p['Re']
            r1 = p['R1']
            t1 = p['T1']
            actual = re + (r1/(1+(1j*w*2*np.pi*t1)))
            for i in range(2, numVoigtElements+1):
                actual += (p['R'+str(i)]/(1+(1j*w*2*np.pi*p['T'+str(i)])))
            return actual.imag
        
        def modelReal(p):
            """
            The model used in real fitting
    
            Parameters
            ----------
            p : lmfit Parameters
                The parameters for the current fitting iteration.
    
            Returns
            -------
            model : numpy array
                Numpy array the real impedance.
    
            """
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
        #---Set Ohmic resistance parameter---
        parameters.add("Re", value=initialGuesses[0][0], min=lowerBounds[0])
        if (fixed):
            parameters["Re"].value = fixedVal
            parameters["Re"].vary = False
        #---Set parameter bounds---
        for i in range(1, numVoigtElements+1):
            if (lowerBounds[i+parameterCorrector] == np.NINF):
                parameters.add("R"+str(i), value=initialGuesses[i+parameterCorrector][0])
            else:
                parameters.add("R"+str(i), value=initialGuesses[i+parameterCorrector][0], min=lowerBounds[i+parameterCorrector])
            if (upperBounds[i+parameterCorrector] != np.inf):
                parameters['R'+str(i)].max = 0
            parameters.add("T"+str(i), value=initialGuesses[i+1+parameterCorrector][0], min=0.0)
            parameterCorrector += 1
        
        current_min = 1E9    #The current lowest objective function found
        current_best = 0     #The parameters used to fit current_min
        def recursive_fit(voigt_number, recursiveGuesses):
            """
            Recursively adds parameters to recursiveGuesses based on the passed-in initial guesses. When the base case with one Voigt element is reached, lmfit
            is used to attempt to minimize the diffComplex objective function, whereafter the chi-squared value is checked to see if that fit is better than the current best.
    
            Parameters
            ----------
            voigt_number : int
                The current Voigt number of whatever round of recursion.
            recursiveGuesses : list of floats
                The current parameter guesses.
    
            Returns
            -------
            None (finds minimum for outer function to return)
    
            """
            if (not self.keepGoing):    #Check if the fitting has been cancelled
                return
            #---The base recursive case with one Voigt element, which attempts to minimize diffComplex and see if that result is better than the previous minimum---
            if (voigt_number == 1):
                for i in range(len(recursiveGuesses[0])):
                    for j in range(len(recursiveGuesses[1])):
                        for k in range(len(recursiveGuesses[2])):
                            if (not fixed):
                                parameters.get("Re").set(value=recursiveGuesses[0][i])
                            parameters.get("R1").set(value=recursiveGuesses[1][j])
                            parameters.get("T1").set(value=recursiveGuesses[2][k])
                            fittedMin = lmfit.minimize(diffComplex, parameters)     #Perform the actual regression
                            #---Check if the minimization is better than any previous---
                            nonlocal current_min
                            nonlocal current_best
                            if (fittedMin.chisqr < current_min and fittedMin.success):
                                current_min = fittedMin.chisqr
                                current_best = fittedMin
            #---The non-base recursive case, which merely adds more parameters to recursiveGuesses based on the passed-in initial guesses---
            else:
                for z in range(len(recursiveGuesses[2*voigt_number-1])):
                    for y in range(len(recursiveGuesses[2*voigt_number])):
                        parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                        parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                        recursive_fit(voigt_number-1, recursiveGuesses)
        
        def recursive_fit_imag(voigt_number, recursiveGuesses):
            """
            Recursively adds parameters to recursiveGuesses based on the passed-in initial guesses. When the base case with one Voigt element is reached, lmfit
            is used to attempt to minimize the diffImag objective function, whereafter the chi-squared value is checked to see if that fit is better than the current best.
    
            Parameters
            ----------
            voigt_number : int
                The current Voigt number of whatever round of recursion.
            recursiveGuesses : list of floats
                The current parameter guesses.
    
            Returns
            -------
            None (finds minimum for outer function to return)
    
            """
            if (not self.keepGoing):        #Check if the fitting has been cancelled
                return
            #---The base recursive case with one Voigt element, which attempts to minimize diffImag and see if that result is better than the previous minimum---
            if (voigt_number == 1):
                for i in range(len(recursiveGuesses[0])):
                    for j in range(len(recursiveGuesses[1])):
                        for k in range(len(recursiveGuesses[2])):
                            if (not fixed):
                                parameters.get("Re").set(value=recursiveGuesses[0][i])
                            parameters.get("R1").set(value=recursiveGuesses[1][j])
                            parameters.get("T1").set(value=recursiveGuesses[2][k])
                            fittedMin = lmfit.minimize(diffImag, parameters)     #Perform the actual regression
                            #---Check if the minimization is better than any previous---
                            nonlocal current_min
                            nonlocal current_best
                            if (fittedMin.chisqr < current_min and fittedMin.success):
                                current_min = fittedMin.chisqr
                                current_best = fittedMin
            #---The non-base recursive case, which merely adds more parameters to recursiveGuesses based on the passed-in initial guesses---
            else:
                for z in range(len(recursiveGuesses[2*voigt_number-1])):
                    for y in range(len(recursiveGuesses[2*voigt_number])):
                        parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                        parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                        recursive_fit_imag(voigt_number-1, recursiveGuesses)
        
        def recursive_fit_real(voigt_number, recursiveGuesses):
            """
            Recursively adds parameters to recursiveGuesses based on the passed-in initial guesses. When the base case with one Voigt element is reached, lmfit
            is used to attempt to minimize the diffReal objective function, whereafter the chi-squared value is checked to see if that fit is better than the current best.
    
            Parameters
            ----------
            voigt_number : int
                The current Voigt number of whatever round of recursion.
            recursiveGuesses : list of floats
                The current parameter guesses.
    
            Returns
            -------
            None (finds minimum for outer function to return)
    
            """
            if (not self.keepGoing):    #Check if the fitting has been cancelled
                return
            #---The base recursive case with one Voigt element, which attempts to minimize diffReal and see if that result is better than the previous minimum---
            if (voigt_number == 1):
                for i in range(len(recursiveGuesses[0])):
                    for j in range(len(recursiveGuesses[1])):
                        for k in range(len(recursiveGuesses[2])):
                            if (not fixed):
                                parameters.get("Re").set(value=recursiveGuesses[0][i])
                            parameters.get("R1").set(value=recursiveGuesses[1][j])
                            parameters.get("T1").set(value=recursiveGuesses[2][k])
                            fittedMin = lmfit.minimize(diffReal, parameters)        #Perform the actual regression
                            #---Check if the minimization is better than any previous---
                            nonlocal current_min
                            nonlocal current_best
                            if (fittedMin.chisqr < current_min and fittedMin.success):
                                current_min = fittedMin.chisqr
                                current_best = fittedMin
            #---The non-base recursive case, which merely adds more parameters to recursiveGuesses based on the passed-in initial guesses---
            else:
                for z in range(len(recursiveGuesses[2*voigt_number-1])):
                    for y in range(len(recursiveGuesses[2*voigt_number])):
                        parameters.get("R"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number-1][z])
                        parameters.get("T"+str(voigt_number)).set(value=recursiveGuesses[2*voigt_number][y])
                        recursive_fit_real(voigt_number-1, recursiveGuesses)
     
        numCombos = 1
        numCores = mp.cpu_count()   #The number of logical CPUs present
        for g in initialGuesses:    #Calculate the number of parameter combinations
            numCombos *= len(g)
        #---If the number of parameter combinations is too great, use multiprocessing. This increases fitting speed at the cost of overhead---
        if ((numVoigtElements < 8 and numCombos > 200) or (numVoigtElements >= 8 and numCombos > 100)):
            longestIndex = initialGuesses.index(max(initialGuesses, key=len))       #Find largest set of initial guesses
            splitArray = np.array_split(initialGuesses[longestIndex], numCores)     #Split the array at the index of the longest set of initial Guesses
            toRun = []                                                              #Initialize the array that will be used for multi-core processing
            #---For each CPU, create a set of initial guesses to run on it---
            for i in range(numCores):
                toAppend = []
                for j in range(len(initialGuesses)):
                    if (j == longestIndex):
                        toAppend.append(splitArray[i].tolist())
                    else:
                        toAppend.append(initialGuesses[j])
                toRun.append(toAppend)
            #---Prepare shared values for use between CPUs---
            manager = mp.Manager()
            vals = manager.list()
            for i in range(numCores):
                vals.append([])
            #---For complex fitting---
            if (fitType == 0):
                for i in range(numCores):
                    if (not self.keepGoing):    #Check if the fitting has been cancelled
                        return
                    #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                    self.processes.append(mp.Process(target=mp_complex, args=(toRun[i], vals, i, parameters, numVoigtElements, Z_append, V_append, w)))
                for p in self.processes:
                    if (not self.keepGoing):    #Check if the fitting has been cancelled
                        return
                    p.start()
                #---Weight for all processes to end---
                for p in self.processes:
                    p.join()
            #---For imaginary fitting---
            if (fitType == 1):
                parameters['Re'].vary = False
                for i in range(numCores):
                    if (not self.keepGoing):    #Check if the fitting has been cancelled
                        return
                    #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                    self.processes.append(mp.Process(target=mp_imag, args=(toRun[i], vals, i, parameters, numVoigtElements, Zj, V, w)))
                for p in self.processes:
                    if (not self.keepGoing):    #Check if the fitting has been cancelled
                        return
                    p.start()
                #---Weight for all processes to end---
                for p in self.processes:
                    p.join()
            #---For real fitting---
            if (fitType == 2):
                for i in range(numCores):
                    if (not self.keepGoing):    #Check if the fitting has been cancelled
                        return
                    #---Add new processes to list; these will be executed in parallel on different CPUS, which increases speed at the cost of high overhead---
                    self.processes.append(mp.Process(target=mp_real, args=(toRun[i], vals, i, parameters, numVoigtElements, Zr, V, w)))
                for p in self.processes:
                    if (not self.keepGoing):    #Check if the fitting has been cancelled
                        return
                    p.start()
                #---Weight for all processes to end---
                for p in self.processes:
                    p.join()
            lowest = 1E9        #The lowest chi-squared value
            lowestIndex = -1    #The index of the minimization for the lowest chi-squared value
            #---Look through each chi-squared value and see if it's the lowest; if it is, then set lowest and lowestIndex---
            for i in range(len(vals)):
                if (vals[i][0] < lowest):
                    lowest = vals[i][0]
                    lowestIndex = i
            minimized = vals[lowestIndex][1]    #The lmfit minimization for the best result
            current_min = vals[lowestIndex][0]  #The chi-squared value at that minimization
        #---If we're not using multiprocessing---
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
        
        #---Gather the results---
        fitted = minimized.params.valuesdict()
        result = [fitted['Re']]
        sigma = [minimized.params['Re'].stderr]
        for i in range(1, numVoigtElements+1):
            result.append(fitted['R'+str(i)])
            result.append(fitted['T'+str(i)])
            sigma.append(minimized.params['R'+str(i)].stderr)
            sigma.append(minimized.params['T'+str(i)].stderr)
        
        if None in sigma:   #If at least one of the fitted parameters doesn't have a known error, return failure
            return "-"
        for i in range(len(result)):
            if (result[i] != 0):
                if (abs(2*sigma[i]/result[i]) >= 1.0 or np.isnan(sigma[i])):    #If at least one parameter has a 95% CI over 100% or a sigma that's nan, return failure
                    return "$"    
        
        #---Return everything---
        return minimized