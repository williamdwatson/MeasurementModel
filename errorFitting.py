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
    stddev = []    #np.zeros(2*len(stdr)*wlength)
    impedReal = []     #np.zeros(2*len(stdr)*wlength)
    impedRealRe = []
    impedImag = []
    impedMod = []
    sigma = []     #np.zeros(len(stdr)*wlength)
    res = []
    stddev2 = []    #np.zeros(2*len(stdr)*wlength)
    impedReal2 = []     #np.zeros(2*len(stdr)*wlength)
    impedRealRe2 = []
    impedImag2 = []
    impedMod2 = []
    sigmaDetrend = []     #np.zeros(len(stdr)*wlength)
    res2 = []
    sigma2 = []
    if (len(stdr) > 1 and detrendChoice == 3):      #If detrending with multiple files
        sigmaDetrend = []
        for i in range(len(stdr)):
            if (weighting == 0):        #No weighting
                sigma.extend(np.ones(2*len(stdr[i])))
                sigmaDetrend.extend(np.ones(len(stdr[i])))
            elif (weighting == 1):      #Variance weighting
                sigma.extend(sigmaIn[i])
                sigma.extend(sigmaIn[i])
                sigmaDetrend.extend(sigmaIn[i])
            else:
                sigmaTemp = []
                for k in range(len(stdr[i])):
                    if (weighting == 2):  #Variance weighting with 3-point moving averages
                        if (k == 0):        #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/3)
                        elif (k == len(stdr[i])-1):  #Final point takes average of preceding point
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k-1]+sigmaIn[i][k-2])/3)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/3)
                    elif (weighting == 3):  #Variance weighting with 5-point moving averages
                        if (k == 0):        #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3]+sigmaIn[i][k+4])/5)
                        elif (k == 1):      #Second point takes average of initial and following points
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3])/5)
                        elif (k == len(stdr[i])-2):  #Second from last point takes average of final point and preceding points
                            sigmaTemp.append((sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/5)
                        elif (k == len(stdr[i])-1):  #Final point takes average of preceding points
                            sigmaTemp.append((sigmaIn[i][k-4]+sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k])/5)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/5)
                sigma.extend(sigmaTemp)
                sigma.extend(sigmaTemp)
                sigmaDetrend.extend(sigmaTemp)
            stddev.extend(np.array(stdr[i]))#/np.array(sigmaDetrend))
            stddev.extend(np.array(stdj[i]))#/np.array(sigmaDetrend))
            impedReal.extend(abs(r[i]))#/np.array(sigmaDetrend)))
            impedImag.extend(abs(j[i]))#/np.array(sigmaDetrend)))
            impedMod.extend((abs(modz[i]**2)))#/np.array(sigmaDetrend)))
            res.extend(np.full(len(r[i]),re[i]))
            impedRealRe.extend((abs(r[i])-res[i]))#/np.array(sigmaDetrend))
            stddev2.extend(abs(stdr[i]))
            stddev2.extend(abs(stdj[i]))
            impedReal2.extend(abs(r[i]))
            impedImag2.extend(abs(j[i]))
            impedMod2.extend(abs(modz[i]))
            res2.extend(np.full(len(r[i]),re[i]))
            impedRealRe2.extend(abs(r[i]) - res[i])
        stddev -= np.mean(stddev)
        impedReal -= np.mean(impedReal)
        impedImag -= np.mean(impedImag)
        impedMod -= np.mean(impedMod)
        impedRealRe -= np.mean(impedRealRe)
        
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
            if (weighting == 0):        #No weighting
                sigma.extend(np.ones(2*len(stdr[i])))
                sigmaDetrend.extend(np.ones(len(stdr[i])))
            elif (weighting == 1):      #Variance weighting
                sigma.extend(sigmaIn[i])
                sigma.extend(sigmaIn[i])
                sigmaDetrend.extend(sigmaIn[i])
            else:
                sigmaTemp = []
                for k in range(len(stdr[i])):
                    if (weighting == 2):  #Variance weighting with 3-point moving averages
                        if (k == 0):        #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2])/3)
                        elif (k == len(stdr[i])-1):  #Final point takes average of preceding point
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k-1]+sigmaIn[i][k-2])/3)
                        else:
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/3)
                    elif (weighting == 3):  #Variance weighting with 5-point moving averages
                        if (k == 0):        #Initial point takes average of following points
                            sigmaTemp.append((sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3]+sigmaIn[i][k+4])/5)
                        elif (k == 1):      #Second point takes average of initial and following points
                            sigmaTemp.append((sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1]+sigmaIn[i][k+2]+sigmaIn[i][k+3])/5)
                        elif (k == len(stdr[i])-2):  #Second from last point takes average of final point and preceding points
                            sigmaTemp.append((sigmaIn[i][k-3]+sigmaIn[i][k-2]+sigmaIn[i][k-1]+sigmaIn[i][k]+sigmaIn[i][k+1])/5)
                        elif (k == len(stdr[i])-1):  #Final point takes average of preceding points
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
                    stddev.extend(np.array(stdr[i]))#/np.array(sigmaDetrend))
                    stddev.extend(np.array(stdj[i]))#/np.array(sigmaDetrend))
                    impedReal.extend(abs(r[i]))#/np.array(sigmaDetrend)))
                    impedImag.extend(abs(j[i]))#/np.array(sigmaDetrend)))
                    impedMod.extend((abs(modz[i]**2)))#/np.array(sigmaDetrend)))
                    res.extend(np.full(len(r[i]),re[i]))
                    impedRealRe.extend((abs(r[i])-res[i]))#/np.array(sigmaDetrend))
                    stddev2.extend(abs(stdr[i]))
                    stddev2.extend(abs(stdj[i]))
                    impedReal2.extend(abs(r[i]))
                    impedImag2.extend(abs(j[i]))
                    impedMod2.extend(abs(modz[i]))
                    res2.extend(np.full(len(r[i]),re[i]))
                    impedRealRe2.extend(abs(r[i]) - res[i])
                stddev -= np.mean(stddev)
                impedReal -= np.mean(impedReal)
                impedImag -= np.mean(impedImag)
                impedMod -= np.mean(impedMod)
                impedRealRe -= np.mean(impedRealRe)
                
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
    if (detrendChoice == 3):
        sigma2 = sigma
        sigma = np.ones(len(stddev))

    def diff(params):
        return (np.array(stddev) - np.array(model(params)))/np.array(sigma)

    def model(p):
        toReturn = []
        start = 0
        for i in range(len(r)):
            toExtend = []
            stop = start+len(r[i])
            for k in range(start, stop):
                toAdd = 0
                if (choices[0]):
                    toAdd += p[0]*impedImag[k]
                if (choices[1]):
                    if (not reChoice):
                        toAdd += p[1]*impedReal[k]
                    else:
                        toAdd += p[1]*impedRealRe[k]   #p[1]*abs(impedReal[k]-res[k])
                if (choices[2]):
                    toAdd += p[2]*impedMod[k]
                if (choices[3] and detrendChoice != 3):
                    toAdd += p[3]
                toExtend.append(toAdd)
            start = stop
            toReturn.extend(toExtend)
            toReturn.extend(toExtend)
        return toReturn

    try:
        minimized = least_squares(diff, guesses, method='lm')
        J = minimized.jac
        deletionCorrector = 0
        for i in range(4):
            if (not choices[i]):
                J = np.delete(J, i-deletionCorrector, axis=1)
                deletionCorrector += 1
        cov = np.linalg.pinv(J.T.dot(J)) * (minimized.fun**2).mean()     #Hessian = J . J^T ; covariance = Hessian^-1 
        if (not all(choices)):        #Add covariances of 0 for parameters not fit
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
            if (cov[i][i] < 0):     #If there are negatives on the covariance diagonal, something's wrong
                raise ValueError
        sigma_result = np.sqrt(np.diag(cov))       #The standard deviations
    except:
        return "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"
    
    a, b, g, d = minimized.x
    sa, sb, sg, sd = sigma_result
    
    def diff_delta(de):
        return (np.array(stddev2) - np.array(delta_model(de[0])))/np.array(sigma2)
    
    def delta_model(de):
        toReturn = []
        start = 0
        for i in range(len(r)):
            toExtend = []
            stop = start+len(r[i])
            for k in range(start, stop):
                toAdd = 0
                if (choices[0]):
                    toAdd += a*impedImag2[k]
                if (choices[1]):
                    if (not reChoice):
                        toAdd += b*impedReal2[k]
                    else:
                        toAdd += b*impedRealRe2[k]   #p[1]*abs(impedReal[k]-res[k])
                if (choices[2]):
                    toAdd += g*impedMod2[k]
                if (choices[3]):
                    toAdd += de
                toExtend.append(toAdd)
            start = stop
            toReturn.extend(toExtend)
            toReturn.extend(toExtend)
        return toReturn
    if (detrendChoice == 3 and choices[3]):
        minimized_delta = least_squares(diff_delta, [guesses[3]], method='lm')
        d = minimized_delta.x
        J_delta = minimized_delta.jac
        cov_delta = np.linalg.pinv(J_delta.T.dot(J_delta)) * (minimized_delta.fun**2).mean()
        sd = np.sqrt(np.diag(cov_delta))
        
    combined = []
    numPoints = 0
    for i in range(len(stdr)):
        for j in range(len(stdr[i])):
            combined.append(stdr[i][j])
            combined.append(stdj[i][j])
            numPoints += 2
    mpe = np.sum(100*abs(np.array(combined) - np.array(model([a, b, g, d])))/np.array(model([a, b, g, d])))
    return a, b, g, d, sa, sb, sg, sd, np.sum(minimized.fun**2), mpe/numPoints, cov