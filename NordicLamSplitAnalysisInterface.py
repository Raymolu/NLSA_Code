# -*- coding: utf-8 -*-
"""
Created on Fri Jan 18 11:17:39 2019
V 6.04 (2019-04-04)
@author: ludovicraymond
"""
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import os
import NordicLamSplitAnalysisFunctions as fu
import NordicLamSplitRepairFunctions as Rfu
import UnitConversion as UC
from UnitConversion import convert
import Report as Re

class MainInt:
        def __init__(self, master): #Calls the user interface
            
            PowerUser = 'ludovicraymond'
            self.master = master
            self.In=str()
            self.Unit = StringVar()
            self.Method = StringVar()
            self.Reinforce = StringVar()
            self.ShearA = StringVar() 
            self.ScrewOri = StringVar()
            self.UnitDict = {}
            self.Units = ['Imperial','Metric']
            self.Methods = ['CSA & Euro','APA 700']
            self.Reinforces = ['ASSY Screw','Glued Panel']
            self.ScrewOris = ['Same Side','Opposed Side']
            ValUnit = ['b','h','hd','Mf','Mr','Vf','Max beam length','Wf','Wr']
            Units = [['in','mm'],['in','mm'],['in','mm'],['lbf-ft','N-m'],['lbf-ft','N-m'],['lbf','N'],['ft','m'],['lbf','N'],['lbf','N']]
            count=0
#-_-# Work on the unit dictionary. See what type of units the functions accepts
            for i in ValUnit:
                self.UnitDict[i]=Units[count]
                count+=1
            self.bList = []
            self.hList = []
            self.Unit.set(self.Units[1])
            self.Method.set(self.Methods[0])
            self.Reinforce.set(self.Reinforces[0])
            self.ScrewOri.set(self.ScrewOris[0])
            self.ShearA.set(0) 
            if os.getlogin() == PowerUser:
                self.Unit.set(self.Units[0])
            self.Un = 0 if self.Unit.get() == self.Units[0] else 1
            self.bImp = [1.75, 3.5]
            self.bMet = [44.45,88.9]
            self.blist = self.bImp if self.Un == 0 else self.bMet               
            self.hImp = [9.5, 11.875, 14, 16]
            self.hMet = [241.3, 301.6, 355.6, 406.4]
            self.hlist = self.hImp if self.Un == 0 else self.hMet

            self.SMLength = "Click analyze"

            self.RepList = []               

            master.title("Nordic Lam Split Analyser")
            master.geometry('1000x470')

        # Sets the repair windows to upper stack level.
            master.bind("<FocusIn>", self.GuiStack)


        #Column 0 Input values
            Col = 0
            Ro = 0
            Width = 12
            Height = 5
            Anchor = 'e'
            Ro+=1
            self.LblProject = Label(master, text="Project Name:", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblNotes = Label(master, text="Notes:", anchor = 'ne', width = Width, height = Height).grid(row=Ro, column=Col)  
            Ro+=4
#            self.CBB1=Checkbutton(master, text=" Bypass", variable=self.ByPass, anchor = Anchor, width = Width-4).grid(row=Ro, column=Col)
            
        #Column 1 Input values
            Col += 1
            Ro = 0
            Width = 150
            Height = 5
            Anchor = 'e'
            Ro+=1
            self.InputProject = Entry(master, font = "Arial 8 bold", width = 60)
            self.InputProject.grid(row=Ro, column=Col, columnspan=2, sticky='w')
            Ro+=1
            self.InputNotes = Text(master, font = "Arial 8", width = Width, height = Height)
            self.InputNotes.grid(row=Ro, column=Col, columnspan=6, sticky=Anchor)
            Width = 15
            Ro+=1
            self.LblHead2 = Label(master, text="User Input").grid(row=Ro, column=Col)
            Ro+=1
            self.Input1 = ttk.Combobox(master, values=(fu.tbNL.SeriesName()),width=Width)
            self.Input1.grid(row=Ro, column=Col)
            self.Input1.insert(0,fu.tbNL.SeriesName()[1])
            Ro+=1            
            self.Input1_1 = ttk.Combobox(master, values=(1,2,3,4),width=Width)
            self.Input1_1.grid(row=Ro, column=Col)
            self.Input1_1.insert(0,1)
            Ro+=1            
            self.Input2 = ttk.Combobox(master, values=(self.blist),width=Width)
            self.Input2.grid(row=Ro, column=Col)
            self.Input2.insert(0,self.blist[0])
            Ro+=1            
            self.Input3 = ttk.Combobox(master, values=(self.hlist),width=Width)
            self.Input3.grid(row=Ro, column=Col)
            self.Input3.insert(0,self.hlist[1])
            Ro+=1            
            self.Input4 = Entry(master, width = Width+3)
            self.Input4.grid(row=Ro, column=Col)
            if self.Unit.get() == self.Units[0]:
                self.Input4.insert(0,4)
            else:
                self.Input4.insert(0,25.4)
            Ro+=1            
            self.Input5 = Entry(master, width = Width+3)
            self.Input5.grid(row=Ro, column=Col)
            Ro+=1            
            self.Input6 = Entry(master, width = Width+3)
            self.Input6.grid(row=Ro, column=Col)
            Ro+=1            
            self.Input7 = Entry(master, width = Width+3)
            self.Input7.grid(row=Ro, column=Col)
            Ro+=1            
            self.Input8 = ttk.Combobox(master, values=(0.65, 1.00, 1.15),width=Width)
            self.Input8.grid(row=Ro, column=Col)
            self.Input8.insert(0,1.00)
            Ro+=1            
            self.Input9 = ttk.Combobox(master, values=(1.00, 1.10),width=Width)
            self.Input9.grid(row=Ro, column=Col)
            self.Input9.insert(0,1.00)
            Ro+=1
            self.ShearMaxLength = StringVar() 
            self.OutShearMaxLength = Label(master, textvariable=self.ShearMaxLength, width = Width).grid(row=Ro, column=Col)
            Ro+=1            
            self.CBB2=Checkbutton(master, text=" Use shear equation 7.5.7.2 (a)", variable=self.ShearA).grid(row=Ro, column=Col, columnspan = 2)
            Ro+=1            
            self.Input10 = Entry(master, width = Width+3)
            self.Input10.grid(row=Ro, column=Col)
            Ro+=1            
            self.Input11 = Entry(master, width = Width+3)
            self.Input11.grid(row=Ro, column=Col)

        #Column 2 Input Labels
            Col += 1
            Ro = 0    
            Anchor = 'w'
            Ro+=1
            Ro+=1
            Width = 32
            Ro+=1
            self.LblHead1 = Label(master, text="Nordic Lam information", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput1 = Label(master, text="Nordic Lam Type (NPG = Architectural)", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput1_1 = Label(master, text="Number of plies", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput2Txt = StringVar()
            self.LblInput2 = Label(master, textvariable=self.LblInput2Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID2 = 'b'
            self.Lbl2 = self.ID2+", Nordic Lam  single ply width, "
            self.LblInput2Txt.set(self.Lbl2+self.UnitDict[self.ID2][self.Un])
            Ro+=1
            self.LblInput3Txt = StringVar()
            self.LblInput3 = Label(master, textvariable=self.LblInput3Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID3 = 'h'
            self.Lbl3 = self.ID3+", Nordic Lam  depth, "
            self.LblInput3Txt.set(self.Lbl3+self.UnitDict[self.ID3][self.Un])
            Ro+=1 
            self.LblInput4Txt = StringVar()
            self.LblInput4 = Label(master, textvariable=self.LblInput4Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID4 = 'hd'
            self.Lbl4 = self.ID4+", Hole diameter, "
            self.LblInput4Txt.set(self.Lbl4+self.UnitDict[self.ID4][self.Un])
            Ro+=1 
            self.LblInput5Txt = StringVar()
            self.LblInput5 = Label(master, textvariable=self.LblInput5Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID5 = 'Mf'
            self.Lbl5 = self.ID5+", Moment force at opening, "
            self.LblInput5Txt.set(self.Lbl5+self.UnitDict[self.ID5][self.Un])
            Ro+=1
            self.LblInput6Txt = StringVar()
            self.LblInput6 = Label(master, textvariable=self.LblInput6Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID6 = 'Mr'
            self.Lbl6 = self.ID6+", Full section moment resistance, "
            self.LblInput6Txt.set(self.Lbl6+self.UnitDict[self.ID6][self.Un])
            Ro+=1
            self.LblInput7Txt = StringVar()
            self.LblInput7 = Label(master, textvariable=self.LblInput7Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID7 = 'Vf'
            self.Lbl7 = self.ID7+", Longitudinal shear 7.5.7.2 (b), "
            self.LblInput7Txt.set(self.Lbl7+self.UnitDict[self.ID7][self.Un])
            Ro+=1
            self.LblInput8 = Label(master, text="KD, duration factor", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput9 = Label(master, text="KH, humidity factor", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
#            self.LblOutShearMaxLength = Label(master, text=" Max beam length to use 7.5.7.2 (b), m", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.LblOutShearMaxLengthTxt = StringVar()
            self.LblOutShearMaxLength = Label(master, textvariable=self.LblOutShearMaxLengthTxt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.IDOutShearMaxLength = 'Max beam length'
            self.LblShearMaxLength = self.IDOutShearMaxLength+" to use 7.5.7.2 (b), "
            self.LblOutShearMaxLengthTxt.set(self.LblShearMaxLength+self.UnitDict[self.IDOutShearMaxLength][self.Un])
            Ro+=2
#            self.LblInput10 = Label(master, text="Wf, shear force, equation 7.5.7.2 (a), N", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.LblInput10Txt = StringVar()
            self.LblInput10 = Label(master, textvariable=self.LblInput10Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID10 = 'Wf'
            self.Lbl10 = self.ID10+", Longitudinal shear 7.5.7.2 (a), "
            self.LblInput10Txt.set(self.Lbl10+self.UnitDict[self.ID10][self.Un])
            Ro+=1
#            self.LblInput11 = Label(master, text="Wr, shear resistance, equation 7.5.7.2 (a), N", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.LblInput11Txt = StringVar()
            self.LblInput11 = Label(master, textvariable=self.LblInput11Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID11 = 'Wr'
            self.Lbl11 = self.ID11+", Shear resistance 7.5.7.2 (a), "
            self.LblInput11Txt.set(self.Lbl11+self.UnitDict[self.ID11][self.Un])

        #Column 3 Output labels
            Col += 1
            Ro = 3
            Width = 36
            Anchor = 'e'
            self.LblHead3 = Label(master, text="Analysis", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput301 = Label(master, text="Bending", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput302 = Label(master, text="Shear", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput303 = Label(master, text="Tension perpendicular to fibers", anchor = Anchor, width = Width).grid(row=Ro, column=Col)

        #Column 4 Output Force
            Col += 1
            Ro = 0
            Width = 10
            Ro+=3
            self.LblHead4 = Label(master, text="Force").grid(row=Ro, column=Col)
            Ro+=1
            self.Out401Mf = StringVar()
            self.Out401 = Label(master, textvariable=self.Out401Mf, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out402Vf = StringVar()
            self.Out402 = Label(master, textvariable=self.Out402Vf, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out403Ftp = StringVar()
            self.Out403 = Label(master, textvariable=self.Out403Ftp, width = Width).grid(row=Ro, column=Col)
            
            Ro+=2
            Anchor = 'w'
            self.RB1=Radiobutton(master, text=self.Units[0], variable=self.Unit, value=self.Units[0], command=self.Radio, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.RB2=Radiobutton(master, text=self.Units[1], variable=self.Unit, value=self.Units[1], command=self.Radio, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=2
            self.RB4=Radiobutton(master, text=self.Methods[0], variable=self.Method, value=self.Methods[0], command = lambda : print(self.Method.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.RB5=Radiobutton(master, text=self.Methods[1], variable=self.Method, value=self.Methods[1], command = lambda : print(self.Method.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=2
            self.RB7=Radiobutton(master, text=self.Reinforces[0], variable=self.Reinforce, value=self.Reinforces[0], command = lambda : print(self.Reinforce.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.RB8=Radiobutton(master, text=self.Reinforces[1], variable=self.Reinforce, value=self.Reinforces[1], command = lambda : print(self.Reinforce.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)

        #Column 5 Output Resistance
            Col += 1
            Ro = 0
            Width = 10
            Ro+=3
            self.LblHead5 = Label(master, text="Resistance").grid(row=Ro, column=Col)
            Ro+=1
            self.Out501CMr = StringVar()
            self.Out501 = Label(master, textvariable=self.Out501CMr, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out502CVr = StringVar()
            self.Out502 = Label(master, textvariable=self.Out502CVr, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out503FtpRes = StringVar()
            self.Out503 = Label(master, textvariable=self.Out503FtpRes, width = Width).grid(row=Ro, column=Col)

        #Reinforcement Output            
            Ro+=10
            Width = 24
            ColumnSpan = 2
            Anchor = 'e'
            self.Out6R1T = StringVar()
            self.Out6R1 = Label(master, textvariable=self.Out6R1T, width = Width, anchor = Anchor).grid(row=Ro, column=Col, columnspan=ColumnSpan)  
            
        #Column 6 Evaluations
            Col += 1
            Ro = 3
            Width = 14
            self.LblHead5 = Label(master, text="Force/Resist").grid(row=Ro, column=Col)
            Ro+=1            
            self.Out601B = StringVar()
            self.Out601 = Label(master, textvariable=self.Out601B, width = Width).grid(row=Ro, column=Col)

            Ro+=1            
            self.Out602V = StringVar()
            self.Out602 = Label(master, textvariable=self.Out602V, width = Width).grid(row=Ro, column=Col)

            Ro+=1            
            self.Out603TP = StringVar()
            self.Out603 = Label(master, textvariable=self.Out603TP, width = Width).grid(row=Ro, column=Col)

        #Buttons
            Ro+=2            
            self.Button1 = Button(master, text='Analyze', command=self.InputProcess, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Button2 = Button(master, text='Analyze & Report', command=self.ReportInt, width = Width).grid(row=Ro, column=Col)
            Ro+=2
            self.Button3 = Button(master, text='Reference & Info', command=self.OpenFile, width = Width).grid(row=Ro, column=Col)
            Ro+=3
            self.Button4 = Button(master, text='Reinforce', command=self.Reinforcement, width = Width).grid(row=Ro, column=Col)
            
            if os.getlogin() == PowerUser:
                self.Input1.delete(0,END)
                self.Input1.insert(0,fu.tbNL.SeriesName()[0])
                self.Reinforce.set(self.Reinforces[1])

            
            
            
        def OpenFile(self):
            #Opens the notes information and reference pdf
            Re.os.startfile('NLSA_Reference.pdf')     

        def InputProcess(self):
            print(str(Re.datetime.datetime.fromtimestamp(Re.time.time()).strftime('%Y%m%d_%H:%M and %Ssec')),'---------------------------------------------------------------------------------------')

#            ByPass= int(self.ByPass.get())
            ShearA= int(self.ShearA.get())            
            NLType = self.Input1.get()
            Project = self.InputProject.get()
            Notes = self.InputNotes.get(0.0,"end-1c") 
            Unit = self.Unit.get()
            self.Un = 0 if self.Unit.get() == 'Imperial' else 1
            List = [self.Input1_1.get(),self.Input2.get(),self.Input3.get(),
                     self.Input4.get(),self.Input5.get(),self.Input6.get(),
                     self.Input7.get(),self.Input8.get(),self.Input9.get()]
            
            #Add the Shear values if shear equation (a) is used
            if ShearA == 1:
                List = List + [self.Input10.get(),self.Input11.get()]
            else:
                List = List + [0, 0]
            
            Index = 0
            for i in List:
                try:
                    List[Index] = float(i)
                except:
                    List[Index] = 0
                    messagebox.showinfo("Error Message","Please correct the input value: ["+i+"] to a number. (float or integer)")
                    print("[",i,"] is not a valid input")
                Index+=1
            Ply,b,h,hd,Mf,Mr,Vf,KD,KH,Wf,Wr = List
            
            # Sets the ply to the lowest int number
            Ply = Rfu.math.floor(Ply)

            #Convert input from imperial units to international units
            if Unit == "Imperial":
                print("The units will be converted from imperial to metric")
            #Convert input data from imperial to metric to do the function calculations in metric.
            # Is this if necessary. Old remenant?
                if self.Unit.get() == 'Imperial':
                    b,h,hd,Mf,Mr,Vf,Wf,Wr = (convert(b,"in","mm"),convert(h,"in","mm"),
                                       convert(hd,"in","mm"),convert(Mf,"lbft","Nm"),
                                       convert(Mr,"lbft","Nm"),convert(Vf,"lb","N"),
                                       convert(Wf,"lb","N"),convert(Wr,"lb","N"))  
            #Generate all the results from the analysis functions based on the method selected.
            MethodUsed = self.Method.get()
            if MethodUsed == self.Methods[0]:
                EvalM, CMr = fu.B_Eval(Mf,Mr,hd,h)
#                EvalV, CShearRes, Fv = fu.S_Eval(NLType,h,b,hd,Vf,KD,KH,Ply,ByPass,Wf,Wr,ShearA)
                EvalV, CShearRes, Fv = fu.S_Eval(NLType,h,b,hd,Vf,KD,KH,Ply,Wf,Wr,ShearA)
            elif MethodUsed == self.Methods[1]:
                EvalM, CMr = fu.B700_Eval(Mf,Mr,hd,h)
#                EvalV, CShearRes, Fv = fu.S700_Eval(NLType,h,b,hd,Vf,KD,KH,Ply,ByPass,Wf,Wr,ShearA)
                EvalV, CShearRes, Fv = fu.S700_Eval(NLType,h,b,hd,Vf,KD,KH,Ply,Wf,Wr,ShearA)
#            Evaltp, Ftp, FtpRes, FtpN = fu.P_Tension(NLType,h,b,hd,Vf,convert(Mf,"Nm","Nmm"),Ply,ByPass)            
            Evaltp, Ftp, FtpRes, FtpN = fu.P_Tension(NLType,h,b,hd,Vf,convert(Mf,"Nm","Nmm"),Ply)            

            #Find the maximum length in m to be able to use shear equation 7.5.7.2 (b)
            self.SMLength = round((2000000000/(h*b))/1000,2)

            #Convert all the outputs in imperial if the interface is in imperial mode.
            PresImp, PresMet = 'psf', 'MPa'
            if self.Unit.get() == 'Imperial':
                CMr = (convert(CMr,"Nm","lbft"))
                Mf = (convert(Mf,"Nm","lbft")) 
                CShearRes = (convert(CShearRes,"N","lb")) 
                Vf = (convert(Vf,"N","lb")) 
                Wf = (convert(Wf,"N","lb")) 
                Wr = (convert(Wr,"N","lb")) 
                Ftp = convert(Ftp,PresMet,PresImp)
                FtpRes = convert(FtpRes,PresMet,PresImp)
                self.SMLength = (convert(self.SMLength,"m","ft"))
                self.FtUnit = PresImp
            else:
                self.FtUnit = PresMet                

            #Set the maximum length to be able to use shear equation 7.5.7.2 (b)
            self.ShearMaxLength.set(str(round(self.SMLength,0)))

            print('Residual Mr = ',CMr,'; Residual shear resistance = ',CShearRes,'; tp Force = ', Ftp, ' MPa ; tp Resistance = ', FtpRes,' MPa')
#            print('Wf for testing = ',Wf)
            #Bending output display on interface
            self.Out401Mf.set (str(round(Mf,))+" "+self.UnitDict[self.ID5][self.Un])
            self.Out501CMr.set (str(round(CMr,))+" "+self.UnitDict[self.ID5][self.Un])
            self.Out601B.set (round(EvalM,2))

            #Shear output display on interface
            if ShearA == 1:
                self.Out402Vf.set (str(round(Wf,))+" "+self.UnitDict[self.ID7][self.Un])
                self.Out502CVr.set (str(round(CShearRes,))+" "+self.UnitDict[self.ID7][self.Un])
                self.Out602V.set (round(EvalV,2))                
            elif ShearA == 0: 
                self.Out402Vf.set (str(round(Vf,))+" "+self.UnitDict[self.ID7][self.Un])
                self.Out502CVr.set (str(round(CShearRes,))+" "+self.UnitDict[self.ID7][self.Un])
                self.Out602V.set (round(EvalV,2))

            #Tension perpendicular to grain output display on interface
            if self.Unit.get() == 'Imperial':
                self.Out403Ftp.set (str(round(Ftp,))+' '+self.FtUnit)
                self.Out503FtpRes.set (str(round(FtpRes,))+' '+self.FtUnit)
            else:
                self.Out403Ftp.set (str(round(Ftp,3))+' '+self.FtUnit)
                self.Out503FtpRes.set (str(round(FtpRes,3))+' '+self.FtUnit)
            self.Out603TP.set (round(Evaltp,2))
            
            #Convert all imperial input units to metric for the report
            if self.Unit.get() == 'Imperial':
                CMr = (convert(CMr,"lbft","Nm"))
                Mf = (convert(Mf,"lbft","Nm")) 
                CShearRes = (convert(CShearRes,"lb","N")) 
                Vf = (convert(Vf,"lb","N"))         
                Wf = (convert(Wf,"lb","N"))
                Wr = (convert(Wr,"lb","N"))
                Ftp = convert(Ftp,PresImp,PresMet)
                FtpRes = convert(FtpRes,PresImp,PresMet)

            if ShearA == 1:
                ShearEq = '7.5.7.2 (a)'
            elif ShearA == 0:
                ShearEq = '7.5.7.2 (b)'
            
#-_-#       # create a dictionary with all the data as dictionary. {'Mf': {'imperial_unit':2222, 'lbft','metric_unit':3333, 'Nm'}}
            # Get the utility function to create dictionaries.
        
            return Project,Notes,NLType,Ply,b,h,hd,MethodUsed,Mf,Mr,CMr,EvalM,ShearEq,Vf,Wf,Wr,CShearRes,EvalV,KD,KH,Ftp,FtpRes,Evaltp,FtpN,Fv

### This is where the Repair GUI code begins:
###
###
        def Reinforcement(self):
            if self.Reinforce.get() == self.Reinforces[0]:
                self.ASSY_ScrewGUI()
            elif  self.Reinforce.get() == self.Reinforces[1]:
                try:
                    self.Glued_PanelWindow.destroy()            
                except:
                    print('Panel reinforcement')                
                Project,Notes,self.NLType,self.Ply,b,self.h,self.hd,MethodUsed,Mf,Mr,CMr,EvalM,ShearEq,Vf,Wf,Wr,CShearRes,EvalV,KD,KH,Ftp,FtpRes,Evaltp,self.FtpN,Fv = self.InputProcess()
                self.Glued_PanelGUI()


### test
        def GuiStack(self,EO):
                
            try:
                self.ASSY_ScrewWindow.lift(self.master)
                self.Glued_PanelWindow.lift(self.master)
            except:
                EO = 0
                
                
### Screw GUI and functions        
#        def ASSY_ScrewGUI(self,FtpN,NLType,Ply,h,hd): #Remove parameters?
        def ASSY_ScrewGUI(self):
            self.ASSY_ScrewWindow = Toplevel(self.master)
            self.ASSY_ScrewWindow .wm_title("Repair parameters")
            self.ASSY_ScrewWindow .wm_geometry('720x250')
#            self.ASSY_ScrewWindow .attributes('-topmost', True)
            
        # Border buffer
            Col = 0
            Ro = 0
            self.Lblbuf = Label(self.ASSY_ScrewWindow , text="", width = 2).grid(row=Ro, column=Col)       


        #Column 0 Input values & buttons
            Col += 1
            Ro = 0
            Width = 18
            Anchor = 'w'
            ScrewTypes = 5 #Quantity of different screw types in the table. Screw types should be the first entries in the tables.
            self.LblCol0 = Label(self.ASSY_ScrewWindow , text="User Input", anchor = 'c').grid(row=Ro, column=Col)       
            Ro+=1
            self.ScrewInput1 = Entry(self.ASSY_ScrewWindow, font = "Arial 8 bold", width = Width)
            self.ScrewInput1.grid(row=Ro, column=Col, sticky=Anchor)
            self.ScrewInput1.insert(0,1)
            Ro+=1
            self.ScrewInput2 = ttk.Combobox(self.ASSY_ScrewWindow, values=(Rfu.tblNLRS.SeriesName()[0:ScrewTypes]), width=Width-3)
            self.ScrewInput2.grid(row=Ro, column=Col, sticky=Anchor)
            self.ScrewInput2.current(0)
            self.ScrewInput2.bind("<<ComboboxSelected>>",self.ScrewUserOut)
            Ro+=1
            self.ScrewInput3 = Entry(self.ASSY_ScrewWindow, font = "Arial 8 bold", width = Width)
            self.ScrewInput3.grid(row=Ro, column=Col, sticky=Anchor)
            self.ScrewInput3.insert(0,0)

        #Column 0 Radio Buttons
            Ro+=2
            self.RBSO1=Radiobutton(self.ASSY_ScrewWindow, text=self.ScrewOris[0], variable=self.ScrewOri, value=self.ScrewOris[0], command = lambda : print(self.Reinforce.get()), anchor = Anchor, width = Width-5).grid(row=Ro, column=Col)
            Ro+=1
            self.RBSO2=Radiobutton(self.ASSY_ScrewWindow, text=self.ScrewOris[1], variable=self.ScrewOri, value=self.ScrewOris[1], command = lambda : print(self.Reinforce.get()), anchor = Anchor, width = Width-5).grid(row=Ro, column=Col)


        #Column 0 Buttons
            Columnspan=2
            Ro+=4            
            self.ScrewButton1 = Button(self.ASSY_ScrewWindow, text='Analyze', command=self.ASSY_ScrewCal, width = Width).grid(row=Ro, column=Col,columnspan=Columnspan)
            Ro+=1
            self.ScrewButton2 = Button(self.ASSY_ScrewWindow, text='Analyze & Report', command=self.ReportScrew, width = Width).grid(row=Ro, column=Col,columnspan=Columnspan)
            
            
        #Column 1 Input Labels
            Col += 1
            Ro = 1
            Width = 20
            Anchor = 'w'
            self.LblIn1 = Label(self.ASSY_ScrewWindow , text="Screw Per side", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1
            self.LblIn2 = Label(self.ASSY_ScrewWindow , text="Screw Type", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1
            self.LblIn3 = Label(self.ASSY_ScrewWindow , text="mm, Hole offset (max 10%)", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1
        
        
        
        #Column 3 Output values
            Col += 1
            Ro = 0
            Width = 18
            Anchor = 'e'
            self.LblCol3 = Label(self.ASSY_ScrewWindow , text="Output", anchor = Anchor, width = Width).grid(row=Ro, column=Col)  
            Ro+=1
            self.ScrewOut1Var = StringVar()
            self.ScrewOut1 = Label(self.ASSY_ScrewWindow , textvariable=self.ScrewOut1Var, anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1
            self.ScrewOut2Var = StringVar()
            self.ScrewOut2 = Label(self.ASSY_ScrewWindow , textvariable=self.ScrewOut2Var, anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1        
            self.ScrewOut3Var = StringVar()
            self.ScrewOut3 = Label(self.ASSY_ScrewWindow , textvariable=self.ScrewOut3Var, anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1        
            self.ScrewOut4Var = StringVar()
            self.ScrewOut4 = Label(self.ASSY_ScrewWindow , textvariable=self.ScrewOut4Var, anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1        
            self.ScrewOut5Var = StringVar()
            self.ScrewOut5 = Label(self.ASSY_ScrewWindow , textvariable=self.ScrewOut5Var, anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1         

        #Column 4 Output labels
            Col += 1
            Ro = 1
            Width = 60
            Anchor = 'w'
            self.ScrewLblOut1 = Label(self.ASSY_ScrewWindow , text="Screw Factored Withdrawal Resistance, N/mm", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1
            self.ScrewLblOut2 = Label(self.ASSY_ScrewWindow , text="Screw Factored Tensile Strength, N", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1        
            self.ScrewLblOut3 = Label(self.ASSY_ScrewWindow , text="Screw Tip length, mm", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1        
            self.ScrewLblOut4 = Label(self.ASSY_ScrewWindow , text="Screw Length, mm (Note: Fully Threaded)", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=1        
            self.ScrewLblOut5 = Label(self.ASSY_ScrewWindow , text="Tension perpendicular to wood fibre, N", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
            Ro+=2         
            self.ScrewLblOut6 = Label(self.ASSY_ScrewWindow , text="Only use fully threaded screws", anchor = Anchor, width = Width,font='arial 10 bold').grid(row=Ro, column=Col)       
            Ro+=1 
        
        #Function to update the User Input dependant outputs (Screw infos)
        def ScrewUserOut(self,EO):
            
            self.NLType = self.Input1.get()
            G = fu.tbNL.GenDict()[self.NLType]['G']
            ScrewTable = Rfu.tblNLRS.GenDict()
            self.ScrewType = self.ScrewInput2.get()
            try:
                self.ScrewOffset = float(self.ScrewInput3.get())
                self.ScrewRes = float(ScrewTable[self.ScrewType][G])
                self.ScrewResAbs = int(ScrewTable[self.ScrewType]['TStr'])
                self.ScrewTip = float(ScrewTable[self.ScrewType]['dia'])
            except:
                self.ScrewOffset = 99999
                self.ScrewRes = 0
                self.ScrewResAbs = 0
                self.ScrewTip = 99999
            self.ScrewOut1Var.set(round(self.ScrewRes,2))
            self.ScrewOut2Var.set(round(self.ScrewResAbs,0))
            self.ScrewOut3Var.set(round(self.ScrewTip,0))
            
        def ASSY_ScrewCal(self):
            self.Input = self.InputProcess()
            Project,Notes,self.NLType,self.Ply,b,self.h,self.hd,MethodUsed,Mf,Mr,CMr,EvalM,ShearEq,Vf,Wf,Wr,CShearRes,EvalV,KD,KH,Ftp,FtpRes,Evaltp,self.FtpN,Fv = self.Input
            self.ScrewUserOut("<<ComboboxSelected>>")  
            
            if abs(self.ScrewOffset) > self.h * 0.1:
                messagebox.showinfo("Error Message","Please correct the offset value: ["+str(self.ScrewOffset)+"] to a value within "+str(round(self.h * 0.1,2)))
                return 
            else:
                try:
                    self.ScrewQty = int(self.ScrewInput1.get())
                    if self.ScrewQty * self.ScrewRes * self.ScrewResAbs * self.ScrewTip > 0:
                        MaxScrew, ThreadL, TpScrew = Rfu.ScrewRepair(self.FtpN,self.ScrewQty,self.ScrewRes,self.ScrewResAbs,self.ScrewTip,self.Ply,b,self.h,self.hd,self.ScrewOri.get(),self.ScrewOffset)
                        print('length: ' + str(round(MaxScrew,4)) + 'mm  Thread: ' + str(round(ThreadL,4)))
                        self.ScrewOut4Var.set(round(MaxScrew,2))
#                        self.ScrewOut5Var.set(round(self.FtpN,2))
                        self.ScrewOut5Var.set(round(TpScrew,2))
                    else:
                        MaxScrew = 99999
                        ThreadL = 99999
                        self.ScrewOut4Var.set(round(MaxScrew,2))
                        print('Input screw quantity and type')
                except:
                    MaxScrew = 99999
                    ThreadL = 99999
                    self.ScrewOut4Var.set(round(MaxScrew,2))
                    print('Error, Input screw quantity and type')
            return self.Reinforce.get(), self.ScrewOri.get(), self.ScrewType, self.ScrewQty, MaxScrew, ThreadL, self.ScrewRes, self.ScrewOffset, TpScrew
        
        def ReportScrew(self):
            try:
                ScrewData = self.ASSY_ScrewCal()
                MaxScrew = ScrewData[4]
                print(MaxScrew)
                Input = self.Input
                Input += ScrewData
            except:
                print('invalid data')
            Re.Report(Input)
            if MaxScrew !=0: self.ASSY_ScrewWindow.destroy()

### Panel GUI and functions    
        def Glued_PanelGUI(self):  
            self.Glued_PanelWindow = Toplevel(self.master)
            self.Glued_PanelWindow .wm_title("Repair parameters")
            self.Glued_PanelWindow .wm_geometry('300x220')    
            self.Glued_PanelWindow .attributes('-topmost', True)    
            
            
            self.Panel_h, self.Panel_w, self.NailQty = Rfu.PanelRepair(self.FtpN,self.Ply,self.h,self.hd)
            
        #Column 0 Input values & buttons
            Col = 0
            Ro = 0
            Width = 300
            Anchor = 'w'
            self.LblPan0 = Label(self.Glued_PanelWindow , text="Results", anchor = 'c').grid(row=Ro, column=Col)       
            Ro+=1
            self.PanelOut1Var = StringVar()
            self.PanelOut1 = Label(self.Glued_PanelWindow, textvariable=self.PanelOut1Var, anchor = Anchor).grid(row=Ro, column=Col)       
            Ro+=1
            self.PanelOut2Var = StringVar()
            self.PanelOut2 = Label(self.Glued_PanelWindow, textvariable=self.PanelOut2Var, anchor = Anchor).grid(row=Ro, column=Col)       
            Ro+=1
            self.PanelButton1 = Button(self.Glued_PanelWindow, text='Report', command=self.ReportPanel).grid(row=Ro, column=Col)
            Ro+=2
            self.PanelButton2 = Button(self.Glued_PanelWindow, text='Reference & Info', command=self.OpenFile).grid(row=Ro, column=Col)
            
            TextOut1Var = ('Repair detail based on full beam height 3/4''"'' CSP'+
                           '\nplywood installed grain perpendicular to beam span'+
                           '\n\nPanel fastened to beam with PL Premium or better structural'+
                           '\nadhesive and 0.131''"'' x 3.25''"'' round head driven nails.')
            TextOut2Var = ('Panel dimensions: ' + str(round(self.Panel_h,1)) + ' mm x ' + str(round(self.Panel_w,1)) + ' mm' +
                           '\nNail minimum quantity: ' + str(self.NailQty) +
                           '\nTension perpendicular to grain: ' + str(round(self.FtpN,0)) + 'N')
            self.PanelOut1Var.set(TextOut1Var)
            self.PanelOut2Var.set(TextOut2Var)
            
        def ReportPanel(self):
            try:
                ReinforceMethod = self.Reinforce.get()
                Input = self.InputProcess()
                PanelData = ReinforceMethod, self.Panel_h, self.Panel_w, self.NailQty
                Input += PanelData
            except:
                print('invalid data')
            
            Re.Report(Input)
            self.Glued_PanelWindow.destroy()            

#Generate an analysis and a report based on the inputed data.        
        def ReportInt(self):
            Input = self.InputProcess()
            Re.Report(Input)

#Convert input between imperial and international systems.
        def Radio(self):
            if self.Units[self.Un] != self.Unit.get():
                self.Un = 0 if self.Unit.get() == 'Imperial' else 1
                if self.Unit.get() == 'Imperial':
                    FUni="mm"
                    SUni="in"
                    FUni2="m"
                    SUni2="ft"
                    FMom='Nm'
                    SMom='lbft'
                    FFor='N'
                    SFor='lb'
                else:
                    FUni="in"
                    SUni="mm"
                    FUni2="ft"
                    SUni2="m"
                    FMom='lbft'
                    SMom='Nm'
                    FFor='lb'
                    SFor='N'
                self.LblInput2Txt.set(self.Lbl2+self.UnitDict[self.ID2][self.Un])            
                self.LblInput3Txt.set(self.Lbl3+self.UnitDict[self.ID3][self.Un])
                self.LblInput4Txt.set(self.Lbl4+self.UnitDict[self.ID4][self.Un])
                self.LblInput5Txt.set(self.Lbl5+self.UnitDict[self.ID5][self.Un])
                self.LblInput6Txt.set(self.Lbl6+self.UnitDict[self.ID6][self.Un])
                self.LblInput7Txt.set(self.Lbl7+self.UnitDict[self.ID7][self.Un])
                self.LblOutShearMaxLengthTxt.set(self.LblShearMaxLength+self.UnitDict[self.IDOutShearMaxLength][self.Un])
                self.LblInput10Txt.set(self.Lbl10+self.UnitDict[self.ID10][self.Un])
                self.LblInput11Txt.set(self.Lbl11+self.UnitDict[self.ID11][self.Un])
                
                if self.Unit.get() == "Imperial":
                    self.blist = self.bImp
                    self.hlist = self.hImp
                else:
                    self.blist = self.bMet
                    self.hlist = self.hMet

                self.Input2['values']=self.blist
                self.Input3['values']=self.hlist

#-_- #Create a dico with inputs and units.        
                InputList = (self.Input2,self.Input3,self.Input4)
                
                for i in InputList:
                    try:
                        self.ConvVal = convert(float(i.get()),FUni,SUni)
                        i.delete(0,END)
                        i.insert(0,round(self.ConvVal,2))
                    except:
                        i.delete(0,END)
                
                InputList = (self.Input5,self.Input6)
                for i in InputList:
                    try:
                        self.ConvVal = convert(float(i.get()),FMom,SMom)
                        i.delete(0,END)
                        i.insert(0,round(self.ConvVal,3))
                    except:
                        i.delete(0,END)

                InputList = (self.Input7,self.Input10,self.Input11)
                for i in InputList:
                    try:
#                        print(i.get())
                        self.ConvVal = convert(float(i.get()),FFor,SFor)
                        i.delete(0,END)
                        i.insert(0,round(self.ConvVal,3))
                    except:
                        i.delete(0,END)

                # Special code for variable label
                InputList = [self.SMLength]
                for i in InputList:
                    try:
#                        print('This is SMLength: ',self.SMLength)
                        self.SMLength = convert(float(self.SMLength),FUni2,SUni2)
                        self.ShearMaxLength.set(str(round(self.SMLength,0)))
                    except:
                        self.ShearMaxLength.set("")
                        print('Invalid data to convert the max length to use equation 7.5.7.2 (b)')
            return

root = Tk()
GUI = MainInt(root)
root.mainloop()            