# -*- coding: utf-8 -*-
"""
Created on Thu Jan 24 13:06:54 2019
Major revision on June 30th 2020
V 6.04 (2019-04-04)
@author: ludovicraymond
"""
def convert(Val1,Unit1,Unit2):
    if Val1 == None:
        Val2 = None
    elif Unit1 == 'in' and Unit2 == 'mm':
        Val2 = Val1 * 25.4
    elif Unit1 == 'mm' and Unit2 == 'in':
        Val2 = Val1 / 25.4
    elif Unit1 == 'ft' and Unit2 == 'm':
        Val2 = Val1 * 0.3048
    elif Unit1 == 'm' and Unit2 == 'ft':
        Val2 = Val1 / 0.3048
    elif Unit1 == 'lb' and Unit2 == 'N':
        Val2 = Val1 * 4.4482216152605
    elif Unit1 == 'N' and Unit2 == 'lb':
        Val2 = Val1 / 4.4482216152605
    elif Unit1 == 'lbft' and Unit2 == 'Nm':
        Val2 = Val1 * 1.3558179483314
    elif Unit1 == 'Nm' and Unit2 == 'lbft':
        Val2 = Val1 / 1.3558179483314
    elif Unit1 == 'Nm' and Unit2 == 'Nmm':
        Val2 = Val1 * 1000
    elif Unit1 == 'Nmm' and Unit2 == 'Nm':
        Val2 = Val1 / 1000
    elif Unit1 == 'MPa' and Unit2 == 'psf':
        Val2 = Val1 * 20885.46
    elif Unit1 == 'psf' and Unit2 == 'MPa':
        Val2 = Val1 / 20885.46
    elif Unit1 == 'MPa' and Unit2 == 'psi':
        Val2 = Val1 * 145.038
    elif Unit1 == 'psi' and Unit2 == 'MPa':
        Val2 = Val1 / 145.038
    elif Unit1 == None and Unit2 == None:
        Val2 = Val1
    else:
        Val2 = None
        print('Units unsupported by this function')
    return Val2
