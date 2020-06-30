# -*- coding: utf-8 -*-
"""
Created on Tue Mar  5 16:49:12 2019
V 6.04 (2019-04-04)
@author: ludovicraymond
"""

#Functions to be use in the Nordic Lam split analysis software to design reinforcement options.

import XLProcessingTableClass as XLTB
from tkinter import messagebox
import math as math

#import for tests only
#import tkinter

SpecFileName = "NLRepairSpecs.xlsx"
tblNLRS = XLTB.XLTable(1,SpecFileName)
Spectbl = tblNLRS.GenDict()
testvar = Spectbl['PL400']['GRes']



#Hole reinforcement with screws
def ScrewRepair(Tp,ScrewQty,ScrewRes,ScrewResAbs,ScrewTip,Ply,b,h,hd,Method,Offset):
    '''
        Tp: Tension perpendicular to wood fibres N (For all plies)
        ScrewQty: number of screws per ply per side of hole. Note that all screws must be at the same distance of the opening.
        ScrewResAbs: The tensil resistance of the screw. (Steel resistance in tension)
        
        Equation variables are in these metric units: N, N/mm or mm.
    '''

    # Tension perpendicular to grain offset effect factor.
    if abs(Offset) <= 0.10*h and Offset != 0: # mm
        print('Original tension force: ',round(Tp,2),'N')
        Tp = Tp * (1 + hd/h)
        print('Tension force increased for Offset hole effect')
        print('New tension force: ',round(Tp,2),'N')
    elif Offset == 0:
        print('No hole offset considered')
    else:
        messagebox.showinfo('Design Failure',str(round(Offset,2))+'mm offset should be corrected')
        Tp = 99999
        return 0, 0, 0

    # Thread length calculations based on tension perpendicular to fibres    
    ThreadL = Tp/(ScrewRes * ScrewQty * Ply)
    print(round(ThreadL,2),'mm minimum thread length')

    # Screw Tip Left    
    ScrewTL = max(ThreadL + ScrewTip,hd + 40 - 0.15*hd) # mm
    if ScrewTL <= hd - 0.15*hd + (h-hd)/2 - Offset:
        print(round(ScrewTL,2),'mm left screw tip side section')
    else:
        print('ScrewTL Fail')
        ScrewTL = 99999

    # Screw Head Left
    ScrewHL = max((h-hd)/2 + Offset + 0.15*hd,ThreadL)
    if ScrewHL <= (h-hd)/2 + Offset + 0.15*hd:
        print(round(ScrewHL,2),'mm left screw head side section')
    else:
        print('ScrewHL Fail')
        ScrewHL = 99999
    
### Same Side method (to test)
    if Method == 'Same Side':

        # Screw Tip Right
        ScrewTR = max(ThreadL + ScrewTip, 40 + 0.15*hd) # mm
        if ScrewTR <= (h-hd)/2 - Offset + 0.15*hd:
            print(round(ScrewTR,2),'mm right screw tip side section')
        else:
            print('ScrewTR Fail')
            ScrewTR = 99999
    
        # Screw Head Right
        ScrewHR = h - (h-hd)/2 + Offset - 0.15*hd
        if ScrewHR >= ThreadL:
            print(round(ScrewHR,2),'mm right screw head side section')
        else:
            print('ScrewHR Fail')
            ScrewHR = 99999

### Opposed Side (to test)
    elif Method == 'Opposed Side':

        # Screw Tip Right
        ScrewTR = max(ThreadL + ScrewTip,hd - 0.15*hd + 40) # mm
        if ScrewTR <= (h-hd)/2 + Offset + hd - 0.15*hd:
            print(round(ScrewTR,2),'mm right screw tip side section')
        else:
            print('ScrewTR Fail')
            ScrewTR = 99999

        # Screw Head Right 
        ScrewHR = max((h-hd)/2 - Offset + 0.15*hd,ThreadL)
        if ScrewHR <= (h-hd)/2 - Offset + 0.15*hd:
            print(round(ScrewHR,2),'mm right screw head side section')
        else:
            print('ScrewHR Fail')
            ScrewHR = 99999
    else:
        messagebox.showinfo('Error message','Trouble with the screw orientation selection')
        ScrewTL = 99999
        ScrewHL = 99999
        ScrewTR = 99999
        ScrewHR = 99999
        
    #Screw length calculation on left and right side of hole
    ScrewL = ScrewTL + ScrewHL
    ScrewR = ScrewTR + ScrewHR
    print(round(ScrewL,2),' mm left screw length')
    print(round(ScrewR,2),' mm right screw length')

    MaxScrew = max(ScrewL,ScrewR)
    print(round(MaxScrew,2),' mm design screw length')
    
    if MaxScrew <= h - 50: # mm
        print(round(MaxScrew,4),' mm screw is within beam depth tolerance')
    else:
        print(round(MaxScrew,4),' mm required screw is too long for the beam depth tolerance. Change screw type or quantity')
        messagebox.showinfo('Design Failure',str(round(MaxScrew,2))+' mm required screw is too long for the beam depth tolerance.'
                            +'\nChange screw type or quantity')
        MaxScrew = 99999
    
    #Validate Screw tensil strength is sufficient.
    if ScrewResAbs * ScrewQty >= Tp:
        print('Screw tensil strenght can resist the tension perpendicular to wood fibres')
    else:
        print('Selected screw fails tensil resistance')
        messagebox.showinfo('Design Failure','Selected screw fails tensil resistance')
        return 99999, 99999, Tp
    
    #Evaluate minimum spacing to avoid splitting
    ScrewEdSpa = 3 * ScrewTip # mm
    ScrewSpa = 2.5 * ScrewTip # mm
    if 2 * ScrewEdSpa + ( ScrewQty - 1 ) * ScrewSpa > b:
        print('Impossible to fit all screws','Screw edge distance: ',ScrewEdSpa,' mm and screw spacing: ',ScrewSpa,' mm')
        messagebox.showinfo('Design Failure','Screw placement is invalid\nChange screw quantity')
        MaxScrew = 0

    return MaxScrew, ThreadL, Tp

#Hole reinforcement with panels
def PanelRepair(Tp,Ply,h,hd):
    '''
        Tp: Tension perpendicular to wood fibres N (For all plies)
        Equation variables in N, N/mm2, N/mm, mm or N/nail.
        
        Plywood surface required above tearing area for glue
        Tp is multiplied by 2 because of plywood effect. (Most of the tension force
        is concentrated at tearing point. The plywood surface is far from that point)
        Tp / 2 because of number of plywood.
	'''
    #Parameter from table
    GlueRes = tblNLRS.GenDict()['PL400']['GRes']
    NailRes = tblNLRS.GenDict()['Nail_0.131x3.25']['NRes']
    print('Glue shear resistance (N/mm\N{SUPERSCRIPT TWO}): ',GlueRes,'; Shear resistance per nail (N): ',NailRes)
    
    #Hard coded parameters. (might get turned into user input in the next phase)
    PanelQty = 2
    PanelTpIncrease = 2
    GlueEff = tblNLRS.GenDict()['PL400']['GEff'] # (Glue efficiency factor based on handling)
    Panel_h = h #mm
    InLineSpacing = 76.2 #mm Spaces between nails in one row
    RowSpacing = 38.1 #mm Spacing between rows (Rows are parallèle to the beam length)
		# This Corn_h dimension should always be maxed to a full beam depth plywood.
		#(One Day if its worth it, in repair phase 2, there could be a feature to reduce it's size)
    Corn_h = (h-hd)/2 + 0.15 * hd

    # Calculate minimum Glue area		
    GArea = Tp * PanelTpIncrease / (PanelQty * GlueRes * GlueEff)
    print ('Glue required area to match tension force: ',round(GArea,2),' mm\N{SUPERSCRIPT TWO}')
	#Calculate plywood dimensions and Calcultate minimum number of nail required
		#If one ply, the Nail quantity will be 1 nail / 5800mm2
    	#Minimum of width of panel = hd + 101.6mm
		#Panel_w is a bit overestimated because we consider the outmost edge of the circle as reference
       #instead of the location where split will initiate. This simplifies the equations. (Could be refined in repair phase 2)
    	#Calculate plywood surface required based on nail quantity (Only for 2 or 3 ply)
    GCorn_w = GArea / Corn_h
    
    if Ply == 1:
        Panel_w = max(hd + 101.6, hd + GCorn_w * 2)
        print('1 ply Panel_w based on adhesif shear resistance: ',Panel_w,' mm, Default minimum dimension used:',101.6>GCorn_w * 2)
        Panel_w = Panel_wCheck(Tp,PanelQty,h,hd,Panel_w)
        print('The Panel width after checkup = ',round(Panel_w/1000,2),' m')
        NailQty = math.ceil(Panel_h * Panel_w / 5800)
    elif Ply == 2 or Ply == 3:
        NailQty = math.ceil(Tp * PanelTpIncrease / (PanelQty * NailRes))
        NRows = math.floor(Corn_h/RowSpacing) -1
        NailQty_w = math.ceil(NailQty / NRows)
        NCorn_w = InLineSpacing * (NailQty_w + 1)
        NArea = Corn_h * NCorn_w
        Panel_w1 = max(hd + 101.6, hd + NCorn_w * 2, hd + GCorn_w * 2)
#        print('From nail resistance Panel width: ',Panel_w1)
#        print('NailQty ',NailQty,'//NRows ',NRows,'//NailQty_w ',NailQty_w,'//NCorn_w ',NCorn_w,'//NArea ',NArea)
#        print('NailQty = math.ceil(Tp * PanelTpIncrease / (PanelQty * NailRes))',NailQty,Tp,PanelTpIncrease,PanelQty,NailRes)        
        print('Minimum nail quantity above and bellow the split area: ',NailQty)
        Panel_w = Panel_wCheck(Tp,PanelQty,h,hd,Panel_w1)
        print('The Panel Checkup width: ',round(Panel_w/1000,2),' m')
        print ('Nail panel required area: ',round(NArea,2),' mm\N{SUPERSCRIPT TWO}')
        print('Nail spacing in mm: ',InLineSpacing,' Row spacing: ',RowSpacing)
        NailQty = (math.floor(Panel_h/RowSpacing)-1) * (math.floor(Panel_w/InLineSpacing)-1) - math.floor(hd/RowSpacing) * math.floor(hd/InLineSpacing)
    else:
        messagebox.showinfo('Invalid input',"invalide number of Ply. Enter 1, 2 or 3 plies.")
        Panel_h = 99999
        Panel_w = 99999
        NailQty = 99999
    print('returned values: ','Panel_h: ', round(Panel_h/1000,4),'m Panel_w: ', round(Panel_w/1000,4),'m NailQty: ', NailQty)
    return Panel_h, Panel_w, NailQty

def Panel_wCheck(Tp,PanelQty,h,hd,Panel_w):
#    print('this is the Panel repair function Tp,PanelQty,h,hd,Panel_w',Tp,PanelQty,h,hd,Panel_w)

    #Plywood tension check CSA O86-14 p 109 (9.5.7)
    Phi = 0.60 #instead of 0.95 because they could install the panel at the wrong angle
    #KD,KS and KT set to 1
    PanelTStrength = tblNLRS.GenDict()['CSP18']['PanelTStr'] #N/mm 90 deg worst case, most comon 18.5mm CSP plywood 6 ply
    PanelTRes = Phi * PanelTStrength * (Panel_w - hd ) / 2
    
    print('Panel tensil resistance for a ',round((Panel_w - hd ) / 2,2),'mm section: ',PanelTRes,'N')

    if Tp > PanelTRes * PanelQty:
        messagebox.showinfo('Info Message',"Panel was widened to resist tension")
        Panel_w = (Tp / PanelQty) * 2 /(Phi * PanelTStrength) + hd
#        print('First modification of Panel_w: ',Panel_w)
    else:
        Panel_w = Panel_w
#        print('Unchainged Panel_w: ',Panel_w)

    #Plywood rolling shear check. Ignored because the nails will help avoid it.
    #Panel restriction to disipate tension force	
    if Panel_w > 2 * h or Panel_w > 1220:
        print('Panel width designed: ',round(Panel_w,2),' mm, 2 x beam depth: ',2*h,' mm, Panel too long:',Panel_w > 2 * h)
        print('Panel width too long because greater than 1220mm',Panel_w > 1220)
        messagebox.showinfo('Invalid result',str(round(Panel_w/1000,2)) + ' m panel is too long')
        Panel_w = 99999
    else:
        Panel_w = Panel_w
    return Panel_w