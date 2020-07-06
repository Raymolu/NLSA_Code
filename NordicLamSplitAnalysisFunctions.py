# -*- coding: utf-8 -*-
"""
Created on Tue Jan 15 14:39:07 2019
V 6.04 (2019-04-04)
@author: ludovicraymond
"""
#Functions to be use in the Nordic Lam split analysis.

import XLProcessingTableClass as XLTB
from tkinter import messagebox


def try_catch(func):
    def wrapper(*args,**kwargs): 
        try:
            fu = func(*args,**kwargs)
        except:
            print('Function: '+str(func.__name__)+' failed to load\nargs: ' + str(args))
            return
        return fu
    return wrapper

#Generate product spec object
SpecFileName = "NordicLamSpecs.xlsx"
tbNL = XLTB.XLTable(1,SpecFileName)
#Maxh=406.4
#Maxb=88.9
Maxh=2401
Maxb=366

#Evaluate the beam residual bending resistance
@try_catch
def B_Eval(Mf,Mr,hd,h):
    """
        A_bend: Allowed bending (data pulled from lookup in spec table based on user input)
        hd: Hole diameter
        h: from 9.5" to 16" (Height)
        
        Note: there is a concern for bending if there are less than four lamellas left.
        Nordic Lam can have lamellas up to 2" depth. We would try and maintain at least one top and one bottom lamella intact (4" or 101.6mm).
    """
    if hd <= (h-101.6):
        try:
            CMr = (h-hd)/h*Mr
            Eval = Mf/CMr
        except:
            CMr = None
            Eval = None
    else:
        messagebox.showinfo("Bending Calculation Error Message","Please correct the input value: [hd] (hd <= (h-4in))")
        print('Less than minimum top and bottom lamella left (4" (101.6mm))')
        print('repair for tension reinforcement parallel and perpendicular to grain')
        CMr = None
        Eval = None
    #Corrected Allowed bending     
    return Eval, CMr


#Evaluate the Beam residual shear resistance
@try_catch
def S_Eval(NLType,h,b,hd,Vf,KD,KH,Ply,Wf,Wr,ShearA):    
    """
        Nordic-Lam Type based on CSA 086-14
	
        h = member depth mm
        b = member width mm (Input single ply dimension)
        Vf = Maximum Longitudinal shear at hole location (b) N
        hd: Hole depth mm
        KD, KH: Set all to 1.00 and have it indicated in the output notes.
        KSv: Always in dry conditions: 1.00
        KT: Always untreated: 1.00   
        ShearA: Is the shear (a) equation used.
        for single element or built-up beam in Clauses 7.5.7.2(b), 7.5.7.3, and 7.5.7.4.1,
        the effect of all loads acting within a distance from a support equal to the depth of the member need not be taken into account.
        As an alternative for beams less than 2.0 m^3 in volume, the factored shear resistance may be calculated using the equation in Item (b).
        As long as the dimensions are restricted to 3.5" x 16" maximum cross section, the <2.0m^3 will be met.    
    """

    if h <= Maxh and b <= Maxb:
        # Use equation (a) or (b) methodes
        if ShearA == 1 and (Wf != None or Wr != None) and Wf*Wr > 0:
            print('Shear equation values from: 7.5.7.2 (a)')
            CShearRes = Wr * (h-hd)/h
            Eval = Wf/CShearRes
            Fv = 0

        elif ShearA == 0:
        #Equation 7.5.7.2 (b) Simple shear equation
            Phi = 0.9
            fv = float(tbNL.GenDict()[NLType]['Fvx'])
            Fv = fv*(KD*KH*1*1)
            
            #Ag = b × (h - hd) = gross cross-sectional area of member, 
            Ag = b*(h-hd)
            CShearRes=Phi*Fv*2/3*Ag*Ply
            Eval=Vf/CShearRes
        else:
            CShearRes=None 
            Eval=None
            Fv=None            
    else:
        messagebox.showinfo("Shear Calculation Error Message",
                            "Select valid h ("+str(h)[0:5]+") and b ("+str(b)[0:5]+"). Max h = "+str(Maxh)+" Max b = "+str(Maxb))
        CShearRes=None 
        Eval=None
        Fv=None
    return Eval, CShearRes, Fv


#Perpendicular to grain tension calculation
@try_catch
def P_Tension(NLType,h,b,hd,Vf,Mf,Ply):
    """
        h: from 9.5" to 16" (Height), mm
        If h >=17.72in (450mm): The Kt,90 will need to be calculated
        hd: Hole diameter, mm
        Mf: bending force in Nmm (typicaly we input and display Mf in Nm)
    """
    #   Note: For round holes only!!!
    if h <= Maxh and b <= Maxb:
        if hd/h < 0.3:
            try:            
                Bending_tp = 0.008*Mf/((h-hd)/2+0.15*hd)
            except:
                Bending_tp = None
        else:
            Bending_tp = (3*Mf*hd**3*(hd+h))/(4*h**3*(h*hd+h**2+hd**2)) 
            print('Alternate Bending_tp equation used\nMore conservative than DIN 1052 equation')
        # Tension perpendicular to wood fibre resistance
        # Check for Kt90 for resistance reduction based on beam depth.
        Kt90 = min(1,(450/h)**0.5) # mm
#        print('Kt,90: ',round(Kt90,2))
        FtpRes = float(tbNL.GenDict()[NLType]['Ftp']) * Kt90
#        print(FtpRes,' VS ',float(tbNL.GenDict()[NLType]['Ftp']))

        Shear_tp = Vf*hd*(3*h**2-hd**2)/(4*h**3)
        # FtpO is the strain in the wood section with the opening. The divider is multiplied by the number of ply instead of deviding the forces. 
        FtpO = (Shear_tp + Bending_tp)/(0.5*(0.35*hd+0.5*h)*b*Ply)
        FtpNO = Shear_tp + Bending_tp #Perpendicular tension force in N
        if h <= 400:
            Ftp = FtpO
            FtpN = FtpNO
        else:
            #Apply the tension amplifier. ANALYSIS AND DESIGN OF LAMINATED VENEER LUMBER BEAMS WITH HOLES, by Manoochehr Ardalany, University of Canterbury, September 2012 section 8.5
            Ftp = ((h/400)**0.5)*FtpO
            FtpN = ((h/400)**0.5)*FtpNO
            print('Tension strain "MPa" amplified for a beam deeper than 400mm')

        Eval = Ftp/FtpRes
        print('Shear_tp = ',Shear_tp, ' N')
        print('Bending_tp = ',Bending_tp, ' N')
    else:
        messagebox.showinfo("Perpendicular Tension Calculation Error Message","Select valid h ("+str(h)[0:5]+") and b ("+str(b)[0:5]+"). Max h = "+str(Maxh)+" Max b = "+str(Maxb))
        FtpN = None
        FtpRes = None
        Ftp = None
        Eval = None
    return Eval, Ftp, FtpRes, FtpN

#Pass spec data
def NLSpecs():
    NLSpecs=tbNL.XLToTable()
    return NLSpecs
    
#APA 700 equations
    '''
        Limits:
            hd <= 2/3*h
            
            Validate that the hole placement requirements are met.
    '''   
@try_catch
def B700_Eval(Mf,Mr,hd,h):
    if hd <= 2/3*h:
        Eval, CMr = B_Eval(Mf,Mr,hd,h)
    else:
        messagebox.showinfo("APA 700 Bending Calculation Error Message","Please correct the input value: [hd] (hd <= 2/3*h)")
        CMr = None
        Eval = None        
    return Eval, CMr

@try_catch
def S700_Eval(NLType,h,b,hd,Vf,KD,KH,Ply,Wf,Wr,ShearA):   
    if (hd <= 2/3*h and h <= Maxh and b <= Maxb):
        # Use equation (a) or (b) methodes
        print('Shear eq (a) :', ShearA == 1)
        Chv = ((h-hd)/h)**2
        if ShearA == 1 and Wf*Wr > 0:
            print('Shear equation values from: 7.5.7.2 (a)')
            CShearRes = Wr * Chv
            Eval = Wf/CShearRes
            Fv = 0
        else:
        #Equation 7.5.7.2 (b) Simple shear equation with Hole effect factor
            Phi = 0.9
            fv = float(tbNL.GenDict()[NLType]['Fvx'])
            Fv = fv*(KD*KH*1*1)
            
            #Ag = b × (h - hd) = gross cross-sectional area of member, 
            Ag = b*h
            CShearRes=Phi*Fv*2/3*Ag*Chv*Ply
            Eval=Vf/CShearRes        
    else:
        messagebox.showinfo("APA 700 Shear Calculation Error Message","Validate parameters: h ("+str(h)[0:5]+"), b ("
                            +str(b)[0:5]+') and [hd]' +' hd/h = '+ str(round(hd/h,2)) + "." + "\n" +
                            " Max h = "+str(Maxh)+"; Max b = "+str(Maxb) + "; (hd/h <= 2/3 [0.67])")
        CShearRes=None 
        Eval=None
        Fv=None      
    return Eval, CShearRes, Fv

#print(B700_Eval(1000,2000,250,300))
#print(S700_Eval("ES1M1",300,45,250,1000,1,1))