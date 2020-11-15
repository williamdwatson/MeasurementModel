# -*- coding: utf-8 -*-
"""
Created on Tue Oct  2 15:33:11 2018

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
from scipy.optimize import least_squares

def findErrorFit(weighting, choices, stdr, stdj, r, j, modz, sigmaIn, guesses, reChoice, re, detrendChoice):
    """
    Find the fit for the error structure of form alpha*|Zj| + beta*|Zr (- Re)| + gamma*|Z| + delta
    
    Parameters
    ----------
    weighting : int
        The weighting choice for the regression; 0 for no weighting (all weights are 1), 1 for variance weighting, otherwise moving average variance
    choices : list of booleans
        Which parameters to use in the model; 0 is alpha, 1 is beta, 2 is gamma, 3 is delta
    stdr : list of lists
        List of lists of floats, with each sublist containing the real standard deviations for a particular file
    stdj : list of lists
        List of lists of floats, with each sublist containing the imaginary standard deviations for a particular file
    r : list of lists
        List of lists of floats, with each sublist containing the real data for a particular file
    j : list of lists
        List of lists of floats, with each sublist containing the imaginary data for a particular file
    modz : list of lists
        List of lists of floats, with each sublist containing the modulus of the impedance data for a particular file
    sigmaIn : list of lists
        List of lists of floats of sigma values from fitting
    guesses: list of floats
        Parameter initial guesses
    reChoice : boolean
        Whether to subtract the ohmic resistance from the real impedance
    re : list of floats
        The ohmic resistance for each file
    detrendChoice : int
        Whether to detrend or not; 1 for no detrending, 3 for constant detrending (2 was for linear, but was removed)
    
    Returns
    -------
    a : float
        Fitted alpha value
    b : float
        Fitted beta value
    g : float
        Fitted gamma value
    d : float
        Fitted delta value
    sa : float
        Standard error in the fitted alpha value
    sb : float
        Standard error in the fitted beta value
    sg : float
        Standard error in the fitted gamma value
    sd : float
        Standard error in the fitted delta value
    chi : float
        Chi-squared value from the fitting
    mpe : float
        Mean percentage error from the fitting
    cov : numpy array
        Array of variable covariances
    """
    stddev = []
    impedReal = []
    impedRealRe = []
    impedImag = []
    impedMod = []
    sigma = []
    res = []
    stddev2 = []
    impedReal2 = []
    impedRealRe2 = []
    impedImag2 = []
    impedMod2 = []
    sigmaDetrend = []
    res2 = []
    sigma2 = []
    if (len(stdr) > 1 and detrendChoice == 3):      #If detrending with multiple files
        sigmaDetrend = []
        for i in range(len(stdr)):
            #---Find fitting weights; append twice (once for real, once for imaginary)---
            if (weighting == 0):                    #No weighting means all weights are one
                sigma.extend(np.ones(2*len(stdr[i])))
                sigmaDetrend.extend(np.ones(len(stdr[i])))
            elif (weighting == 1):                  #Variance weighting means detrend by sigma
                sigma.extend(sigmaIn[i])
                sigma.extend(sigmaIn[i])
                sigmaDetrend.extend(sigmaIn[i])
            else:                                   #Variance weighting with moving average
                sigmaTemp = []
                for k in range(len(stdr[i])):
                    if (weighting == 2):            #Variance weighting with 3-point moving averages
                        if (k == 0):                #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/3)
                        elif (k == len(stdr[i])-1):  #Final point takes average of preceding point
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k-1]+sigmaIn[i][k-2])/3)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/3)
                    elif (weighting == 3):          #Variance weighting with 5-point moving averages
                        if (k == 0):                #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3]+sigmaIn[i][k+4])/5)
                        elif (k == 1):              #Second point takes average of initial and following points
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3])/5)
                        elif (k == len(stdr[i])-2): #Second from last point takes average of final point and preceding points
                            sigmaTemp.append((sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/5)
                        elif (k == len(stdr[i])-1): #Final point takes average of preceding points
                            sigmaTemp.append((sigmaIn[i][k-4]+sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k])/5)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/5)
                sigma.extend(sigmaTemp)
                sigma.extend(sigmaTemp)
                sigmaDetrend.extend(sigmaTemp)
            stddev.extend(np.array(stdr[i]))
            stddev.extend(np.array(stdj[i]))
            impedReal.extend(abs(r[i]))
            impedImag.extend(abs(j[i]))
            impedMod.extend((abs(modz[i]**2)))
            res.extend(np.full(len(r[i]),re[i]))
            impedRealRe.extend((abs(r[i])-res[i]))
            stddev2.extend(abs(stdr[i]))
            stddev2.extend(abs(stdj[i]))
            impedReal2.extend(abs(r[i]))
            impedImag2.extend(abs(j[i]))
            impedMod2.extend(abs(modz[i]))
            res2.extend(np.full(len(r[i]),re[i]))
            impedRealRe2.extend(abs(r[i]) - res[i])
        #---Detrend by subtracting off average first---
        stddev -= np.mean(stddev)
        impedReal -= np.mean(impedReal)
        impedImag -= np.mean(impedImag)
        impedMod -= np.mean(impedMod)
        impedRealRe -= np.mean(impedRealRe)
        
        #---Then detrend by dividing by weights---
        start = 0
        for i in range(len(r)):
            stop = start+len(r[i])
            for j in range(start, stop):
                impedReal[j] /= sigmaDetrend[j]
                impedImag[j] /= sigmaDetrend[j]
                impedMod[j] /= sigmaDetrend[j]
                impedRealRe[j] /= sigmaDetrend[j]
            start = stop
        start = 0
        for i in range(len(stdr)):
            stop = start+2*len(stdr[i])
            for j in range(start, stop):
                stddev[j] /= sigma[j]
            start = stop
    else:
        for i in range(len(stdr)):
            #---Find fitting weights; append twice (once for real, once for imaginary)---
            if (weighting == 0):                    #No weighting means all weights are one
                sigma.extend(np.ones(2*len(stdr[i])))
                sigmaDetrend.extend(np.ones(len(stdr[i])))
            elif (weighting == 1):                  #Variance weighting means detrend by sigma
                sigma.extend(sigmaIn[i])
                sigma.extend(sigmaIn[i])
                sigmaDetrend.extend(sigmaIn[i])
            else:                                   #Variance weighting with moving average
                sigmaTemp = []
                for k in range(len(stdr[i])):
                    if (weighting == 2):            #Variance weighting with 3-point moving averages
                        if (k == 0):                #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/3)
                        elif (k == len(stdr[i])-1): #Final point takes average of preceding point
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k-1]+sigmaIn[i][k-2])/3)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/3)
                    elif (weighting == 3):          #Variance weighting with 5-point moving averages
                        if (k == 0):                #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3]+sigmaIn[i][k+4])/5)
                        elif (k == 1):              #Second point takes average of initial and following points
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3])/5)
                        elif (k == len(stdr[i])-2): #Second from last point takes average of final point and preceding points
                            sigmaTemp.append((sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/5)
                        elif (k == len(stdr[i])-1): #Final point takes average of preceding points
                            sigmaTemp.append((sigmaIn[i][k-4]+sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k])/5)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/5)
                sigma.extend(sigmaTemp)
                sigma.extend(sigmaTemp)
                sigmaDetrend.extend(sigmaTemp)
    
            if (detrendChoice == 1):    #No detrending
                stddev.extend(stdr[i])
                stddev.extend(stdj[i])
                impedReal.extend(abs(r[i]))
                impedImag.extend(abs(j[i]))
                impedMod.extend(abs(modz[i])**2)
                res.extend(np.full(len(r[i]),re[i]))
                impedRealRe.extend(abs(r[i]) - res[i])
            #detrendChoice == 2 was for linear detrending, which was removed
            elif (detrendChoice == 3):  #Average detrending
                for i in range(len(stdr)):
                    stddev.extend(np.array(stdr[i]))
                    stddev.extend(np.array(stdj[i]))
                    impedReal.extend(abs(r[i]))
                    impedImag.extend(abs(j[i]))
                    impedMod.extend((abs(modz[i]**2)))
                    res.extend(np.full(len(r[i]),re[i]))
                    impedRealRe.extend((abs(r[i])-res[i]))
                    stddev2.extend(abs(stdr[i]))
                    stddev2.extend(abs(stdj[i]))
                    impedReal2.extend(abs(r[i]))
                    impedImag2.extend(abs(j[i]))
                    impedMod2.extend(abs(modz[i]))
                    res2.extend(np.full(len(r[i]),re[i]))
                    impedRealRe2.extend(abs(r[i]) - res[i])
                #---Detrend by subtracting off average first---
                stddev -= np.mean(stddev)
                impedReal -= np.mean(impedReal)
                impedImag -= np.mean(impedImag)
                impedMod -= np.mean(impedMod)
                impedRealRe -= np.mean(impedRealRe)
                
                #---Then detrend by dividing by weights---
                start = 0
                for i in range(len(r)):
                    stop = start+len(r[i])
                    for j in range(start, stop):
                        impedReal[j] /= sigmaDetrend[j]
                        impedImag[j] /= sigmaDetrend[j]
                        impedMod[j] /= sigmaDetrend[j]
                        impedRealRe[j] /= sigmaDetrend[j]
                    start = stop
                start = 0
                for i in range(len(stdr)):
                    stop = start+2*len(stdr[i])
                    for j in range(start, stop):
                        stddev[j] /= sigma[j]
                    start = stop
    
    #---If detrending, prep sigma2 for delta fit and set sigma to all ones, as weighting as already been done---
    if (detrendChoice == 3):
        sigma2 = sigma.copy()
        sigma = np.ones(len(stddev))

    def diff(params):
        """
        Residuals function for least squares; residuals are standard deviations minus model stddevs divided by sigmas
        
        Parameters
        ----------
        params : array of floats
            Current function parameters
        
        Returns
        -------
        residuals : array of floats
            Array of residual values
        """
        return (np.array(stddev) - np.array(model(params)))/np.array(sigma)

    def model(p):
        """
        Error structure model, of the form alpha*|Zj| + beta*|Zr (- Re)| + gamma*|Z| + delta
        
        Parameters
        ----------
        p : array of floats
            Current function parameters
        
        Returns
        -------
        error_structure : array of floats
            The error structure given by the current parameters
        """
        toReturn = []
        start = 0
        #---Loop through each file---
        for i in range(len(r)):
            toExtend = []
            #---Loop through each frequency---
            stop = start+len(r[i])
            for k in range(start, stop):
                #---Only add parameters that have been chosen---
                toAdd = 0
                if (choices[0]):
                    toAdd += p[0]*impedImag[k]
                if (choices[1]):
                    if (not reChoice):
                        toAdd += p[1]*impedReal[k]
                    else:
                        toAdd += p[1]*impedRealRe[k]
                if (choices[2]):
                    toAdd += p[2]*impedMod[k]
                if (choices[3] and detrendChoice != 3):
                    toAdd += p[3]
                toExtend.append(toAdd)
            start = stop
            #---Append twice (once for real, once for imaginary)---
            toReturn.extend(toExtend)
            toReturn.extend(toExtend)
        return toReturn

    try:
        minimized = least_squares(diff, guesses, method='lm')           #Find best fit for error structure
        J = minimized.jac                                               #Get fitting Jacobian for future covariance calculations
        #---Remove unused parameters from the Jacobian---
        deletionCorrector = 0
        for i in range(4):
            if (not choices[i]):
                J = np.delete(J, i-deletionCorrector, axis=1)
                deletionCorrector += 1
        #---Calculate covariances---
        cov = np.linalg.pinv(J.T.dot(J)) * (minimized.fun**2).mean()     #Hessian = J . J^T ; covariance = Hessian^-1 
        #---Add covariances of 0 for parameters not fit---
        if (not all(choices)):          
            toMerge = np.zeros((4, 4))
            rowCorrector = 0
            for i in range(4):
                colCorrector = 0
                if (not choices[i]):
                    rowCorrector += 1
                    continue
                for k in range(4):
                    if (not choices[k]):
                        colCorrector += 1
                        continue
                    toMerge[i][k] = cov[i-rowCorrector][k-colCorrector]
            cov = toMerge
        for i in range(4):
            if (cov[i][i] < 0):                    #If there are negatives on the covariance diagonal, something's wrong
                raise ValueError
        sigma_result = np.sqrt(np.diag(cov))       #Result standard deviations are the square roots of the covariance matrix diagonal
    except Exception:
        return "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"     #Indicate an error to the GUI
    
    a, b, g, d = minimized.x
    sa, sb, sg, sd = sigma_result
    
    def diff_delta(de):
        """
        Residuals function for fitting delta if detrending with delta chosen
        
        Parameters
        ----------
        de : array of one float
            Delta value
        
        Returns
        -------
        residuals : array of floats
            Array of residual values
        """
        return (np.array(stddev2) - np.array(delta_model(de[0])))/np.array(sigma2)
    
    def delta_model(de):
        """
        Model for fitting delta if detrending with delta chosen
        
        Parameters
        ----------
        de : array of one float
            Delta value
        
        Returns
        -------
        residuals : array of floats
            Array of residual values
        """
        toReturn = []
        start = 0
        #---Loop through each file---
        for i in range(len(r)):
            toExtend = []
            #---Loop through each frequency---
            stop = start+len(r[i])
            for k in range(start, stop):
                toAdd = 0
                if (choices[0]):
                    toAdd += a*impedImag2[k]
                if (choices[1]):
                    if (not reChoice):
                        toAdd += b*impedReal2[k]
                    else:
                        toAdd += b*impedRealRe2[k]
                if (choices[2]):
                    toAdd += g*impedMod2[k]
                if (choices[3]):                #This should always be true
                    toAdd += de
                toExtend.append(toAdd)
            start = stop
            #---Append twice (once for real, once for imaginary)---
            toReturn.extend(toExtend)
            toReturn.extend(toExtend)
        return toReturn
    
    #---If detrending and also fitting delta, need to replace because detrending will remove constants entirely---
    if (detrendChoice == 3 and choices[3]):
        #---Find delta by least squares---
        minimized_delta = least_squares(diff_delta, [guesses[3]], method='lm')
        d = minimized_delta.x
        #---Recalculate delta parameter standard errors from fit---
        J_delta = minimized_delta.jac
        cov_delta = np.linalg.pinv(J_delta.T.dot(J_delta)) * (minimized_delta.fun**2).mean()
        sd = np.sqrt(np.diag(cov_delta))
    
    #---Create array with real and imaginary values next to each other for mean percent error calculation---
    combined = []
    numPoints = 0
    for i in range(len(stdr)):
        for j in range(len(stdr[i])):
            combined.append(stdr[i][j])
            combined.append(stdj[i][j])
            numPoints += 2
    #---Calculate mean percent error---
    mpe = np.sum(100*abs(np.array(combined) - np.array(model([a, b, g, d])))/np.array(model([a, b, g, d])))
    return a, b, g, d, sa, sb, sg, sd, np.sum(minimized.fun**2), mpe/numPoints, cov