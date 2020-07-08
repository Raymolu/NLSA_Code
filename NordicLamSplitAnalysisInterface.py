# -*- coding: utf-8 -*-
"""
Created on Fri Jan 18 11:17:39 2019
V 6.04 (2019-04-04)
@author: ludovicraymond
"""
#from tkinter import *
from tkinter import (Label, Entry, Text, ttk, messagebox,
                     Checkbutton, StringVar, Radiobutton,
                     Button, END, Toplevel, Tk)
#from tkinter import ttk
#from tkinter import messagebox
import os
import NordicLamSplitAnalysisFunctions as fu
from NordicLamSplitRepairFunctions import try_round_NA
import NordicLamSplitRepairFunctions as Rfu
from UnitConversion import convert
import Report as Re
import math as math

icon = 'nordic_N.ico' #For exe version


class MainInt:
        def __init__(self, master): #Calls the user interface
            
            PowerUser = 'ludovicraymond'
            self.master = master
            self.In=str()
            self.selected_unit = StringVar()
            self.selected_calc_method = StringVar()
            self.reinforcement_type = StringVar()
            self.shear_method_a_used = StringVar() 
            self.screw_configuration = StringVar()
            self.imperial_data_dico = {}
            self.metric_data_dico = {}
            self.UnitDict = {}
            self.Units = ['Imperial','Metric']
            self.calc_methods = ['CSA & Euro','APA 700']
            self.reinforcement_types = ['ASSY Screw','Glued Panel']
            self.screw_configurations = ['Same Side','Opposed Side']
            ValUnit = ['Nordic_Lam_thickness','Nordic_Lam_depth','hole_diameter',
                       'bending_force_Mf','bending_resistance_Mr','shear_force_Vf',
                       'max_beam_length','shear_force_Wf','shear_resistance_Wr']
            Units = [['in','mm'],['in','mm'],['in','mm'],
                     ['lbft','Nm'],['lbft','Nm'],['lb','N'],
                     ['ft','m'],['lb','N'],['lb','N']]
            count=0

#-_-# Update all the label text to show the proper wording instead of the variable name.
### Make sure max_beam_length displays on the interface when calc done
            
            for i in ValUnit:
                self.UnitDict[i]=Units[count]
                count+=1
            self.Nordic_Lam_thickness_List = []
            self.Nordic_Lam_depth_List = []
            self.selected_unit.set(self.Units[1])
            self.selected_calc_method.set(self.calc_methods[0])
            self.reinforcement_type.set(self.reinforcement_types[0])
            self.screw_configuration.set(self.screw_configurations[0])
            self.shear_method_a_used.set(0) 
            if os.getlogin() == PowerUser:
                self.selected_unit.set(self.Units[0])
            self.selected_unit_index = self.Units.index(self.selected_unit.get())
            self.Nordic_Lam_imperial_thickness_List = [1.75, 3.5]
            self.Nordic_Lam_metric_thickness_List = [44.45,88.9]
            self.Nordic_Lam_thickness_List = self.Nordic_Lam_imperial_thickness_List if self.selected_unit_index== 0 else self.Nordic_Lam_metric_thickness_List               
            self.Nordic_Lam_imperial_depth_List = [9.5, 11.875, 14, 16]
            self.Nordic_Lam_metric_depth_List = [241.3, 301.6, 355.6, 406.4]
            self.Nordic_Lam_depth_List = self.Nordic_Lam_imperial_depth_List if self.selected_unit_index== 0 else self.Nordic_Lam_metric_depth_List

            self.max_length_to_use_shear_resistance_b = "Click analyze"

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
            self.label_project_name = Label(master, text="Project Name:", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.label_project_notes = Label(master, text="Notes:", anchor = 'ne', width = Width, height = Height).grid(row=Ro, column=Col)  
            Ro+=4
#            self.CBB1=Checkbutton(master, text=" Bypass", variable=self.ByPass, anchor = Anchor, width = Width-4).grid(row=Ro, column=Col)
            
        #Column 1 Input values
            Col += 1
            Ro = 0
            Width = 150
            Height = 5
            Anchor = 'e'
            Ro+=1
            self.input_project_name = Entry(master, font = "Arial 8 bold", width = 60)
            self.input_project_name.grid(row=Ro, column=Col, columnspan=2, sticky='w')
            Ro+=1
            self.Input_project_notes = Text(master, font = "Arial 8", width = Width, height = Height)
            self.Input_project_notes.grid(row=Ro, column=Col, columnspan=6, sticky=Anchor)
            Width = 15
            Ro+=1
            self.LblHead2 = Label(master, text="User Input").grid(row=Ro, column=Col)
            Ro+=1
            self.input_Nordic_Lam_type_1 = ttk.Combobox(master, values=(fu.tbNL.SeriesName()),width=Width)
            self.input_Nordic_Lam_type_1.grid(row=Ro, column=Col)
            self.input_Nordic_Lam_type_1.insert(0,fu.tbNL.SeriesName()[1])
            Ro+=1            
            self.input_Nordic_Lam_ply_quantity_1_1 = ttk.Combobox(master, values=(1,2,3,4),width=Width)
            self.input_Nordic_Lam_ply_quantity_1_1.grid(row=Ro, column=Col)
            self.input_Nordic_Lam_ply_quantity_1_1.insert(0,1)
            Ro+=1            
            self.input_Nordic_Lam_thickness_2 = ttk.Combobox(master, values=(self.Nordic_Lam_thickness_List),width=Width)
            self.input_Nordic_Lam_thickness_2.grid(row=Ro, column=Col)
            self.input_Nordic_Lam_thickness_2.insert(0,self.Nordic_Lam_thickness_List[0])
            Ro+=1            
            self.input_Nordic_Lam_depth_3 = ttk.Combobox(master, values=(self.Nordic_Lam_depth_List),width=Width)
            self.input_Nordic_Lam_depth_3.grid(row=Ro, column=Col)
            self.input_Nordic_Lam_depth_3.insert(0,self.Nordic_Lam_depth_List[1])
            Ro+=1            
            self.input_hole_diameter_4 = Entry(master, width = Width+3)
            self.input_hole_diameter_4.grid(row=Ro, column=Col)
            if self.selected_unit.get() == self.Units[0]:
                self.input_hole_diameter_4.insert(0,4)
            else:
                self.input_hole_diameter_4.insert(0,25.4)
            Ro+=1            
            self.input_bending_force_Mf_5 = Entry(master, width = Width+3)
            self.input_bending_force_Mf_5.grid(row=Ro, column=Col)
            Ro+=1            
            self.input_bending_resistance_Mr_6 = Entry(master, width = Width+3)
            self.input_bending_resistance_Mr_6.grid(row=Ro, column=Col)
            Ro+=1            
            self.input_shear_force_Vf_7 = Entry(master, width = Width+3)
            self.input_shear_force_Vf_7.grid(row=Ro, column=Col)
            Ro+=1            
            self.input_K_D_8 = ttk.Combobox(master, values=(0.65, 1.00, 1.15),width=Width)
            self.input_K_D_8.grid(row=Ro, column=Col)
            self.input_K_D_8.insert(0,1.00)
            Ro+=1            
            self.input_K_H_9 = ttk.Combobox(master, values=(1.00, 1.10),width=Width)
            self.input_K_H_9.grid(row=Ro, column=Col)
            self.input_K_H_9.insert(0,1.00)
            Ro+=1
            self.ShearMaxLength = StringVar() 
            self.OutShearMaxLength = Label(master, textvariable=self.ShearMaxLength, width = Width).grid(row=Ro, column=Col)
            Ro+=1            
            self.CBB2=Checkbutton(master, text=" Use shear equation 7.5.7.2 (a)", variable=self.shear_method_a_used).grid(row=Ro, column=Col, columnspan = 2)
            Ro+=1            
            self.input_shear_force_Wf_10 = Entry(master, width = Width+3)
            self.input_shear_force_Wf_10.grid(row=Ro, column=Col)
            Ro+=1            
            self.input_shear_resistance_Wr_11 = Entry(master, width = Width+3)
            self.input_shear_resistance_Wr_11.grid(row=Ro, column=Col)

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
            self.Lbl_input_Nordic_Lam_type_1 = Label(master, text="Nordic Lam Type (NPG = Architectural)", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Lbl_input_Nordic_Lam_ply_quantity_1_1 = Label(master, text="Number of plies", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblInput2Txt = StringVar()
            self.LblInput2 = Label(master, textvariable=self.LblInput2Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID2 = 'Nordic_Lam_thickness'
            self.Lbl2 = self.ID2+", Nordic Lam  single ply width, "
            self.LblInput2Txt.set(self.Lbl2+self.UnitDict[self.ID2][self.selected_unit_index])
            Ro+=1
            self.Lbl_input_Nordic_Lam_depth_3_Txt = StringVar()
            self.Lbl_input_Nordic_Lam_depth_3 = Label(master, textvariable=self.Lbl_input_Nordic_Lam_depth_3_Txt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID3 = 'Nordic_Lam_depth'
            self.Lbl3 = self.ID3+", Nordic Lam  depth, "
            self.Lbl_input_Nordic_Lam_depth_3_Txt.set(self.Lbl3+self.UnitDict[self.ID3][self.selected_unit_index])
            Ro+=1 
            self.label_input_hole_diameter_4_text = StringVar()
            self.label_input_hole_diameter_4 = Label(master, textvariable=self.label_input_hole_diameter_4_text, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID4 = 'hole_diameter'
            self.Lbl4 = self.ID4+", Hole diameter, "
            self.label_input_hole_diameter_4_text.set(self.Lbl4+self.UnitDict[self.ID4][self.selected_unit_index])
            Ro+=1 
            self.label_input_bending_force_Mf_5_text = StringVar()
            self.label_input_bending_force_Mf_5 = Label(master, textvariable=self.label_input_bending_force_Mf_5_text, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID5 = 'bending_force_Mf'
            self.Lbl5 = self.ID5+", Moment force at opening, "
            self.label_input_bending_force_Mf_5_text.set(self.Lbl5+self.UnitDict[self.ID5][self.selected_unit_index])
            Ro+=1
            self.label_input_bending_resistance_Mr_6_text = StringVar()
            self.label_input_bending_resistance_Mr_6 = Label(master, textvariable=self.label_input_bending_resistance_Mr_6_text, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID6 = 'bending_resistance_Mr'
            self.Lbl6 = self.ID6+", Full section moment resistance, "
            self.label_input_bending_resistance_Mr_6_text.set(self.Lbl6+self.UnitDict[self.ID6][self.selected_unit_index])
            Ro+=1
            self.label_input_shear_force_Vf_7_text = StringVar()
            self.label_input_shear_force_Vf_7 = Label(master, textvariable=self.label_input_shear_force_Vf_7_text, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID7 = 'shear_force_Vf'
            self.Lbl7 = self.ID7+", Longitudinal shear 7.5.7.2 (b), "
            self.label_input_shear_force_Vf_7_text.set(self.Lbl7+self.UnitDict[self.ID7][self.selected_unit_index])
            Ro+=1
            self.label_input_K_D_8 = Label(master, text="KD, duration factor", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.label_input_K_H_9 = Label(master, text="KH, humidity factor", anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.LblOutShearMaxLengthTxt = StringVar()
            self.LblOutShearMaxLength = Label(master, textvariable=self.LblOutShearMaxLengthTxt, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.IDOutShearMaxLength = 'max_beam_length'
            self.LblShearMaxLength = self.IDOutShearMaxLength+" to use 7.5.7.2 (b), "
            self.LblOutShearMaxLengthTxt.set(self.LblShearMaxLength+self.UnitDict[self.IDOutShearMaxLength][self.selected_unit_index])
            Ro+=2
            self.label_input_shear_force_Wf_10_text = StringVar()
            self.label_input_shear_force_Wf_10 = Label(master, textvariable=self.label_input_shear_force_Wf_10_text, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID10 = 'shear_force_Wf'
            self.Lbl10 = self.ID10+", Longitudinal shear 7.5.7.2 (a), "
            self.label_input_shear_force_Wf_10_text.set(self.Lbl10+self.UnitDict[self.ID10][self.selected_unit_index])
            Ro+=1
            self.label_input_shear_resistance_Wr_11_text = StringVar()
            self.label_input_shear_resistance_Wr_11 = Label(master, textvariable=self.label_input_shear_resistance_Wr_11_text, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            self.ID11 = 'shear_resistance_Wr'
            self.Lbl11 = self.ID11+", Shear resistance 7.5.7.2 (a), "
            self.label_input_shear_resistance_Wr_11_text.set(self.Lbl11+self.UnitDict[self.ID11][self.selected_unit_index])

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
            self.Out401bending_force_Mf = StringVar()
            self.Out401 = Label(master, textvariable=self.Out401bending_force_Mf, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out402Vf = StringVar()
            self.Out402 = Label(master, textvariable=self.Out402Vf, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out403_tension_strain_perpendicular = StringVar()
            self.Out403 = Label(master, textvariable=self.Out403_tension_strain_perpendicular, width = Width).grid(row=Ro, column=Col)
            
            Ro+=2
            Anchor = 'w'
            self.RB1=Radiobutton(master, text=self.Units[0], variable=self.selected_unit, value=self.Units[0], command=self.Radio, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.RB2=Radiobutton(master, text=self.Units[1], variable=self.selected_unit, value=self.Units[1], command=self.Radio, anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=2
            self.RB4=Radiobutton(master, text=self.calc_methods[0], variable=self.selected_calc_method, value=self.calc_methods[0], command = lambda : print(self.selected_calc_method.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.RB5=Radiobutton(master, text=self.calc_methods[1], variable=self.selected_calc_method, value=self.calc_methods[1], command = lambda : print(self.selected_calc_method.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=2
            self.RB7=Radiobutton(master, text=self.reinforcement_types[0], variable=self.reinforcement_type, value=self.reinforcement_types[0], command = lambda : print(self.reinforcement_type.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.RB8=Radiobutton(master, text=self.reinforcement_types[1], variable=self.reinforcement_type, value=self.reinforcement_types[1], command = lambda : print(self.reinforcement_type.get()), anchor = Anchor, width = Width).grid(row=Ro, column=Col)

        #Column 5 Output Resistance
            Col += 1
            Ro = 0
            Width = 10
            Ro+=3
            self.LblHead5 = Label(master, text="Resistance").grid(row=Ro, column=Col)
            Ro+=1
            self.Out501_reduced_bending_resistance = StringVar()
            self.Out501 = Label(master, textvariable=self.Out501_reduced_bending_resistance, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out502CVr = StringVar()
            self.Out502 = Label(master, textvariable=self.Out502CVr, width = Width).grid(row=Ro, column=Col)
            Ro+=1
            self.Out503_tension_strain_perpendicular_resistance = StringVar()
            self.Out503 = Label(master, textvariable=self.Out503_tension_strain_perpendicular_resistance, width = Width).grid(row=Ro, column=Col)

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
            self.Button2 = Button(master, text='Analyze & Report', command=self.general_report, width = Width).grid(row=Ro, column=Col)
            Ro+=2
            self.Button3 = Button(master, text='Reference & Info', command=self.OpenFile, width = Width).grid(row=Ro, column=Col)
            Ro+=3
            self.Button4 = Button(master, text='Reinforce', command=self.Reinforcement, width = Width).grid(row=Ro, column=Col)
            
            if os.getlogin() == PowerUser:
                self.input_Nordic_Lam_type_1.delete(0,END)
                self.input_Nordic_Lam_type_1.insert(0,fu.tbNL.SeriesName()[0])
                self.reinforcement_type.set(self.reinforcement_types[1])

        def pass_data_from_metric_to_imperial_dico(self):
            '''
                Use this function when the imperial dictionary needs to be updated.
                This should be used everytime the metric data is changed to keep the 
                imperial dictionary up to date.
            '''
            for selected_key in self.metric_data_dico:
                try:
                    self.imperial_data_dico[selected_key]['value'] = (
                            convert(
                                    self.metric_data_dico[selected_key]['value'],
                                    self.metric_data_dico[selected_key]['unit'],
                                    self.imperial_data_dico[selected_key]['unit']))  
                except:
                    print("pass_data_from_metric_to_imperial_dico:\n"\
                          f"[ {selected_key} ] could not be converted")
            pass

        def dico_in_use (self):
            '''
                returns the dico to use based on the user unit type selected
                Use the following variables in the code for eas of use.
                selected_data_dico = dico_in_use()
            '''
            if self.selected_unit.get().lower() == 'imperial':
                return self.imperial_data_dico
            elif self.selected_unit.get().lower() == 'metric':
                return self.metric_data_dico
            pass             

        def round_factor_in_use (self,imperial_rounding, metric_rounding):
            '''
                returns the roun_at variable value.
                Use the following variables in the code for eas of use.
                round_at = round_factor_in_use()
            '''
            if self.selected_unit.get().lower() == 'imperial':
                return self.imperial_data_dico, imperial_rounding
            elif self.selected_unit.get().lower() == 'metric':
                return self.metric_data_dico, metric_rounding
            pass  
        
        def add_receptecals_to_data_dico (self, data_receptacle_list):
            '''
                Uses a list of lists to add items to the data dictionaries.
                [['var_name','imperial units', 'metric units'],]
            '''
            for selected_receptacle in data_receptacle_list:
                self.imperial_data_dico[selected_receptacle[0]] = {'unit':selected_receptacle[1],'value':None}
                self.metric_data_dico[selected_receptacle[0]] =  {'unit':selected_receptacle[2],'value':None}              
            pass
            
        def OpenFile(self):
            #Opens the notes information and reference pdf
            Re.os.startfile('NLSA_Reference.pdf')     

        def InputProcess(self):
            print(str(Re.datetime.datetime.fromtimestamp(
                    Re.time.time()).strftime('%Y%m%d_%H:%M and %Ssec')),
                    '---------------------------------------------------------------------------------------')

            self.selected_unit_index = self.Units.index(self.selected_unit.get())
            selected_unit = self.selected_unit.get().lower()
            selected_calc_method = self.selected_calc_method.get()
            
            numerical_input_dico = {
                            'Nordic_Lam_ply_quantity': self.input_Nordic_Lam_ply_quantity_1_1.get(),
                            'Nordic_Lam_thickness': self.input_Nordic_Lam_thickness_2.get(),
                            'Nordic_Lam_depth': self.input_Nordic_Lam_depth_3.get(),
                            'hole_diameter': self.input_hole_diameter_4.get(),
                            'bending_force_Mf': self.input_bending_force_Mf_5.get(),
                            'bending_resistance_Mr': self.input_bending_resistance_Mr_6.get(),
                            'shear_force_Vf': self.input_shear_force_Vf_7.get(),
                            'shear_force_Wf': self.input_shear_force_Wf_10.get(),
                            'shear_resistance_Wr': self.input_shear_resistance_Wr_11.get(),
                            'K_D': self.input_K_D_8.get(),
                            'K_H': self.input_K_H_9.get()                            
                            }



            first_data_key_list = [
                            ['project_name',None,None],
                            ['Notes',None,None],
                            ['Nordic_Lam_type',None,None],
                            ['Nordic_Lam_ply_quantity',None,None],
                            ['selected_calc_method',None,None],
                            ['shear_method_a_used',None,None],
                            ['reduced_bending_resistance','lbft','Nm'],
                            ['bending_analysis_ratio',None,None],
                            ['shear_equation_reference',None,None],
                            ['shear_reduced_resistance','lb','N'],
                            ['shear_analysis_ratio',None,None],
                            ['K_D',None,None],
                            ['K_H',None,None],
                            ['tension_strain_perpendicular','psi','MPa'],
                            ['tension_strain_perpendicular_resistance','psi','MPa'],
                            ['perpendicular_tension_strain_analysis_ratio',None,None],
                            ['tension_force_perpendicular','lb','N'],
                            ['Fv','psi','MPa'],
                            ['shear_force','lb','N'],
                            ['shear_resistance','lb','N'],
                            ['shear_resistance_Vf','lb','N']
                            ]

            self.add_receptecals_to_data_dico(first_data_key_list)
            
            for selected_dictionary in [self.imperial_data_dico, self.metric_data_dico]:
                selected_dictionary['project_name']['value'] = self.input_project_name.get()
                selected_dictionary['Notes']['value'] = self.Input_project_notes.get(0.0,"end-1c")
                selected_dictionary['Nordic_Lam_type']['value'] = self.input_Nordic_Lam_type_1.get()
                selected_dictionary['selected_calc_method']['value'] = self.selected_calc_method.get()
                selected_dictionary['shear_method_a_used']['value'] = int(self.shear_method_a_used.get()) 

            for selected_key in self.UnitDict:
                self.imperial_data_dico[selected_key] = {'unit':self.UnitDict[selected_key][0],'value':None}
                self.metric_data_dico[selected_key] =  {'unit':self.UnitDict[selected_key][1],'value':None}

            # Start validation of data #
            for i in numerical_input_dico: # Validate if items in the list can be converted to float.
                try:
                    numerical_input_dico[i] = float(numerical_input_dico[i])
                except:
                    if (i == 'shear_force_Vf' and 
                        self.metric_data_dico['shear_method_a_used']['value'] == 1):
                        continue
                    if ((i == 'shear_force_Wf' or i == 'shear_resistance_Wr') and
                        self.metric_data_dico['shear_method_a_used']['value'] == 0):
                        continue

                    error_value = numerical_input_dico[i]
                    numerical_input_dico[i] = None
                    messagebox.showinfo(f'Error Message','Please correct the'\
                    f' input value for the {i} [ {error_value} ] to a number. (float or integer)')
                    print("[",i,"] is not a valid input")
            # End validation of data #
            
            # Sets the ply to the lowest int number
            if numerical_input_dico['Nordic_Lam_ply_quantity'] != None:
                numerical_input_dico['Nordic_Lam_ply_quantity'] = (
                    math.floor(numerical_input_dico['Nordic_Lam_ply_quantity']))
                
            # Set all the numerical inputs in the dico after being validated
            for selected_key in numerical_input_dico:
                if selected_unit == 'imperial':
                    try:
                        self.imperial_data_dico[selected_key]['value'] = numerical_input_dico[selected_key]
                        converted_key_value = convert(numerical_input_dico[selected_key],
                                                      self.imperial_data_dico[selected_key]['unit'],
                                                      self.metric_data_dico[selected_key]['unit'])
                        self.metric_data_dico[selected_key]['value'] = converted_key_value
                    except:
                        pass
                if selected_unit == 'metric':
                    try:
                        self.metric_data_dico[selected_key]['value'] = numerical_input_dico[selected_key]
                        converted_key_value = convert(numerical_input_dico[selected_key],
                                                      self.metric_data_dico[selected_key]['unit'],
                                                      self.imperial_data_dico[selected_key]['unit'])
                        self.imperial_data_dico[selected_key]['value'] = converted_key_value
                    except:
                        pass   

            print(f'Final ply quantity: {self.metric_data_dico["Nordic_Lam_ply_quantity"]["value"]}')

            #Generate all the results from the analysis functions based on the method selected.
            if selected_calc_method == self.calc_methods[0]:
                bending_evaluation_function = fu.B_Eval
                shear_evaluation_function = fu.S_Eval
            
            elif selected_calc_method == self.calc_methods[1]:
                bending_evaluation_function = fu.B700_Eval
                shear_evaluation_function = fu.S700_Eval
            
            try:
                self.metric_data_dico['shear_resistance_Vf']['value'] = (
                         fu.shear_resistance_equation_b(
                            self.metric_data_dico['Nordic_Lam_type']['value'],
                            self.metric_data_dico['Nordic_Lam_depth']['value'],
                            self.metric_data_dico['Nordic_Lam_thickness']['value'],
                            self.metric_data_dico['K_D']['value'],
                            self.metric_data_dico['K_H']['value'],
                            self.metric_data_dico['Nordic_Lam_ply_quantity']['value'],
                            )) # No hole and methode b to calculate Vr without reduction.
            except:
                self.metric_data_dico['shear_resistance_Vf']['value'] = None
                
            try:    
                (self.metric_data_dico['bending_analysis_ratio']['value'],
                 self.metric_data_dico['reduced_bending_resistance']['value']) = (
                         bending_evaluation_function(
                            self.metric_data_dico['bending_force_Mf']['value'],
                            self.metric_data_dico['bending_resistance_Mr']['value'],
                            self.metric_data_dico['hole_diameter']['value'],
                            self.metric_data_dico['Nordic_Lam_depth']['value']))
            except:
                (self.metric_data_dico['bending_analysis_ratio']['value'],
                 self.metric_data_dico['reduced_bending_resistance']['value']) = (None,)*2
                print('bending calculations failed to load due to invalid inputs')

            try:
                (self.metric_data_dico['shear_analysis_ratio']['value'],
                 self.metric_data_dico['shear_reduced_resistance']['value'],
                 self.metric_data_dico['Fv']['value']) = (
                         shear_evaluation_function(
                            self.metric_data_dico['Nordic_Lam_type']['value'],
                            self.metric_data_dico['Nordic_Lam_depth']['value'],
                            self.metric_data_dico['Nordic_Lam_thickness']['value'],
                            self.metric_data_dico['hole_diameter']['value'],
                            self.metric_data_dico['shear_force_Vf']['value'],
                            self.metric_data_dico['K_D']['value'],
                            self.metric_data_dico['K_H']['value'],
                            self.metric_data_dico['Nordic_Lam_ply_quantity']['value'],
                            self.metric_data_dico['shear_force_Wf']['value'],
                            self.metric_data_dico['shear_resistance_Wr']['value'],
                            self.metric_data_dico['shear_method_a_used']['value']))
            except:
                (self.metric_data_dico['shear_analysis_ratio']['value'],
                 self.metric_data_dico['shear_reduced_resistance']['value'],
                 self.metric_data_dico['Fv']['value']) = (None,)*3
                print('shear calculations failed to load due to invalid inputs')

            try:
                (self.metric_data_dico['perpendicular_tension_strain_analysis_ratio']['value'], 
                 self.metric_data_dico['tension_strain_perpendicular']['value'], 
                 self.metric_data_dico['tension_strain_perpendicular_resistance']['value'], 
                 self.metric_data_dico['tension_force_perpendicular']['value'] ) = (
                        fu.P_Tension(
                            self.metric_data_dico['Nordic_Lam_type']['value'],
                            self.metric_data_dico['Nordic_Lam_depth']['value'],
                            self.metric_data_dico['Nordic_Lam_thickness']['value'],
                            self.metric_data_dico['hole_diameter']['value'],
                            self.metric_data_dico['shear_force_Vf']['value'],
                            convert(self.metric_data_dico['bending_force_Mf']['value'],"Nm","Nmm"),
                            self.metric_data_dico['Nordic_Lam_ply_quantity']['value']))
            except:
                (self.metric_data_dico['perpendicular_tension_strain_analysis_ratio']['value'], 
                 self.metric_data_dico['tension_strain_perpendicular']['value'], 
                 self.metric_data_dico['tension_strain_perpendicular_resistance']['value'], 
                 self.metric_data_dico['tension_force_perpendicular']['value'] ) = (None,)*4
                print('perpendicular tension calculations failed to load due to invalid inputs')

            #Find the maximum length in m to be able to use shear equation 7.5.7.2 (b)
            self.metric_data_dico['max_beam_length']['value'] = (
                    round((2000000000 / (
                    self.metric_data_dico['Nordic_Lam_depth']['value']
                    * self.metric_data_dico['Nordic_Lam_thickness']['value']))
                    / 1000, 2))
            self.max_length_to_use_shear_resistance_b = (
                    self.metric_data_dico['max_beam_length']['value']) # updates the gui variable

            # Update the imperial data dictionary
            self.pass_data_from_metric_to_imperial_dico()

            data_dico = self.dico_in_use()
            tension_strain_perpendicular_round_factor = self.round_factor_in_use(0,3)

            #Set the maximum length to be able to use shear equation 7.5.7.2 (b)
            self.ShearMaxLength.set(str(round(data_dico['Nordic_Lam_depth']['value'],0)))

            print('Residual Mr = ',
                  data_dico['reduced_bending_resistance']['value'],'; Residual shear resistance = ',
                  data_dico['shear_reduced_resistance']['value'],'; tp Force = ',
                  data_dico['tension_strain_perpendicular']['value'], ' MPa ; tp Resistance = ',
                  data_dico['tension_strain_perpendicular_resistance']['value'],' MPa')

            #Bending output display on interface
            self.Out401bending_force_Mf.set (
                    f'{try_round_NA(data_dico["bending_force_Mf"]["value"],round_at=0)}'\
                    f' {data_dico["bending_force_Mf"]["unit"]}')
            self.Out501_reduced_bending_resistance.set (
                    f'{try_round_NA(data_dico["reduced_bending_resistance"]["value"],round_at=0)}'\
                    f' {data_dico["reduced_bending_resistance"]["unit"]}')
            self.Out601B.set (
                    f'{try_round_NA(data_dico["bending_analysis_ratio"]["value"],round_at=2)}')

            #Shear output display on interface
            if data_dico['shear_method_a_used']['value'] == 1:
                self.Out402Vf.set (
                        f'{try_round_NA(data_dico["shear_force_Wf"]["value"],round_at=0)}'\
                        f' {data_dico["shear_force_Wf"]["unit"]}')
                data_dico['shear_equation_reference']['value'] = '7.5.7.2 (a)'
            elif data_dico['shear_method_a_used']['value'] == 0:
                self.Out402Vf.set (
                        f'{try_round_NA(data_dico["shear_force_Vf"]["value"],round_at=0)}'\
                        f' {data_dico["shear_force_Vf"]["unit"]}')
                data_dico['shear_equation_reference']['value'] = '7.5.7.2 (b)'

            self.Out502CVr.set (
                        f'{try_round_NA(data_dico["shear_reduced_resistance"]["value"],round_at=0)}'\
                        f' {data_dico["shear_reduced_resistance"]["unit"]}')
            self.Out602V.set (
                        f'{try_round_NA(data_dico["shear_analysis_ratio"]["value"],round_at=2)}')

            #Tension perpendicular to grain output display on interface
            self.Out403_tension_strain_perpendicular.set (
                        f'{try_round_NA(data_dico["tension_strain_perpendicular"]["value"],round_at=tension_strain_perpendicular_round_factor)}'\
                        f' {data_dico["tension_strain_perpendicular"]["unit"]}')
            self.Out503_tension_strain_perpendicular_resistance.set (
                        f'{try_round_NA(data_dico["tension_strain_perpendicular_resistance"]["value"],round_at=tension_strain_perpendicular_round_factor)}'\
                        f' {data_dico["tension_strain_perpendicular_resistance"]["unit"]}')
            self.Out603TP.set (
                        f'{try_round_NA(data_dico["perpendicular_tension_strain_analysis_ratio"]["value"],round_at=2)}')
            return data_dico

###################
# Repair GUI code #
###################

        def Reinforcement(self):
            self.InputProcess()
            
            reinforcement_data_key_list = [
                                        ['reinforcement_type',None,None],
                                        ]

            self.add_receptecals_to_data_dico(reinforcement_data_key_list) 
            self.metric_data_dico['screw_oreinforcement_type']['value'] = (
                    self.reinforcement_type.get())
            self.pass_data_from_metric_to_imperial_dico()
            
            
            
            if self.metric_data_dico['screw_oreinforcement_type']['value'] == (
                    self.reinforcement_types[0]):
                self.ASSY_ScrewGUI()
            elif  self.metric_data_dico['screw_oreinforcement_type']['value'] == (
                    self.reinforcement_types[1]):
                try:
                    self.Glued_PanelWindow.destroy()            
                except:
                    print('Panel reinforcement')                
                self.Glued_PanelGUI()
            pass

        def GuiStack(self,EO):
            try:
                self.ASSY_ScrewWindow.lift(self.master)
                self.Glued_PanelWindow.lift(self.master)
            except:
                EO = 0
            return EO

# Screw GUI and functions
#########################
            
        def ASSY_ScrewGUI(self):
            self.ASSY_ScrewWindow = Toplevel(self.master)
            self.ASSY_ScrewWindow .wm_title("Repair parameters")
            self.ASSY_ScrewWindow .wm_geometry('720x250')
            
        # Border buffer
            Col = 0
            Ro = 0
            self.Lblbuf = Label(self.ASSY_ScrewWindow , text="", width = 2).grid(row=Ro, column=Col)       

        #Column 0 Input values & buttons
            Col += 1
            Ro = 0
            Width = 18
            Anchor = 'w'
            screw_types = 5 #Quantity of different screw types in the table. Screw types should be the first entries in the tables.
            self.LblCol0 = Label(self.ASSY_ScrewWindow , text="User Input", anchor = 'c').grid(row=Ro, column=Col)       
            Ro+=1
            self.ScrewInput1 = Entry(self.ASSY_ScrewWindow, font = "Arial 8 bold", width = Width)
            self.ScrewInput1.grid(row=Ro, column=Col, sticky=Anchor)
            self.ScrewInput1.insert(0,1)
            Ro+=1
            self.ScrewInput2 = ttk.Combobox(self.ASSY_ScrewWindow, values=(Rfu.tblNLRS.SeriesName()[0:screw_types]), width=Width-3)
            self.ScrewInput2.grid(row=Ro, column=Col, sticky=Anchor)
            self.ScrewInput2.current(None)
            self.ScrewInput2.bind("<<ComboboxSelected>>",self.screw_user_selection)
            Ro+=1
            self.screw_offset_input_3 = Entry(self.ASSY_ScrewWindow, font = "Arial 8 bold", width = Width)
            self.screw_offset_input_3.grid(row=Ro, column=Col, sticky=Anchor)
            self.screw_offset_input_3.insert(0,0)

        #Column 0 Radio Buttons
            Ro+=2
            self.RBSO1=Radiobutton(self.ASSY_ScrewWindow, text=self.screw_configurations[0], variable=self.screw_configuration, value=self.screw_configurations[0], command = lambda : print(self.reinforcement_type.get()), anchor = Anchor, width = Width-5).grid(row=Ro, column=Col)
            Ro+=1
            self.RBSO2=Radiobutton(self.ASSY_ScrewWindow, text=self.screw_configurations[1], variable=self.screw_configuration, value=self.screw_configurations[1], command = lambda : print(self.reinforcement_type.get()), anchor = Anchor, width = Width-5).grid(row=Ro, column=Col)


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
            self.LblIn1 = Label(self.ASSY_ScrewWindow , text="Screw/ply on one side of hole", anchor = Anchor, width = Width).grid(row=Ro, column=Col)       
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
        
        def screw_user_selection(self,EO):
            '''
                Function to update the User Input dependant outputs (Screw infos)
            '''
            screw_data_key_list = [
                                ['screw_type',None,None],
                                ['screw_offset','in','mm'],
                                ['screw_pull_out_resistance','lb/in','N/mm'],
                                ['scres_absolute_tension_resistance_TStr','lb','N'],
                                ['screw_tip','in','mm']
                                ]
            self.add_receptecals_to_data_dico(screw_data_key_list)

            G = fu.tbNL.GenDict()[self.input_Nordic_Lam_type_1.get()]['G']
            ScrewTable = Rfu.tblNLRS.GenDict()
            
            def try_table_lookup (item_to_lookup, ScrewTable=ScrewTable):
                try:
                    data = ScrewTable[self.ScrewInput2.get()][item_to_lookup]
                except:
                    data = None
                    print( f"failed to lookup [ {item_to_lookup} ] for {self.ScrewInput2.get()} screw." )
                return data
            
            get_list = [self.ScrewInput2.get(), self.screw_offset_input_3.get(),
                        try_table_lookup (G), try_table_lookup ('TStr'),
                        try_table_lookup ('dia')]
            
            counter = 0
            for selected_key in screw_data_key_list:
                try:
                    if selected_key[0] == 'screw_offset' :
                        self.metric_data_dico[selected_key[0]]['value'] = abs(float(get_list[counter]))
                    else:
                        self.metric_data_dico[selected_key[0]]['value'] = float(get_list[counter])
                except:
                    self.metric_data_dico[selected_key[0]]['value'] = None
                counter+=1

            self.ScrewOut1Var.set(try_round_NA(
                    self.metric_data_dico['screw_pull_out_resistance']['value']
                    ,round_at=2))
            self.ScrewOut2Var.set(try_round_NA(
                    self.metric_data_dico['scres_absolute_tension_resistance_TStr']['value']
                    ))
            self.ScrewOut3Var.set(try_round_NA(
                    self.metric_data_dico['screw_tip']['value']
                    ))
            self.pass_data_from_metric_to_imperial_dico()
###
### In progress
            ### Start updating this function. Handle None inputs.
        def ASSY_ScrewCal(self):
            '''
                Evaluates the screw properties before calling the screw repair
                All calculations in metric units
            '''
            self.InputProcess()
            self.screw_user_selection("<<ComboboxSelected>>")  

### -_- Add screw repair info to the self.metric_data_dico and update the imperial_data_dico for the reports    
            
            screw_repair_data_key_list = [
                                        ['screw_quantity',None,None],
                                        ['screw_configuration',None,None]
                                        ['highest_screw_length','in','mm'],
                                        ['calculated_screw_thread','in','mm'],
                                        ['tension_force_perpendicular_offset_increase','lb','N'],
                                        ]

            self.add_receptecals_to_data_dico(screw_repair_data_key_list)

            if self.metric_data_dico['screw_offset']['value'] > (
                    self.metric_data_dico['Nordic_Lam_depth']['value'] * 0.1):
                messagebox.showinfo("Error Message",
                    f"Please correct the offset value: {self.metric_data_dico['screw_offset']['value']}mm"\
                    f" to a value within {str(round(self.metric_data_dico['Nordic_Lam_depth']['value'] * 0.1,2))}mm")
                return 
            else:
                try:
                    self.metric_data_dico['screw_quantity']['value'] = int(self.ScrewInput1.get())
                    self.metric_data_dico['screw_configuration']['value'] = self.screw_configuration.get()
                    

                    if (self.metric_data_dico['screw_quantity']['value'] *
                        self.metric_data_dico['screw_pull_out_resistance']['value'] *
                        self.metric_data_dico['scres_absolute_tension_resistance_TStr']['value'] *
                        self.metric_data_dico['screw_tip']['value']) > 0:
                        
                        (self.metric_data_dico['highest_screw_length']['value'],
                         self.metric_data_dico['calculated_screw_thread']['value'],
                         self.metric_data_dico['tension_force_perpendicular_offset_increase']['value']) = (
                            Rfu.ScrewRepair(
                            self.metric_data_dico['tension_force_perpendicular']['value'],
                            self.metric_data_dico['screw_quantity']['value'],
                            self.metric_data_dico['screw_pull_out_resistance']['value'],
                            self.metric_data_dico['scres_absolute_tension_resistance_TStr']['value'],
                            self.metric_data_dico['screw_tip']['value'],
                            self.metric_data_dico['Nordic_Lam_ply_quantity']['value'],
                            self.metric_data_dico['Nordic_Lam_thickness']['value'],
                            self.metric_data_dico['Nordic_Lam_depth']['value'],
                            self.metric_data_dico['hole_diameter']['value'],
                            self.metric_data_dico['screw_configuration']['value'],
                            self.metric_data_dico['screw_offset']['value']
                            ))

                        print(f"length: "\
                              f"{round(self.metric_data_dico['highest_screw_length']['value'],4)}"\
                              f"{self.metric_data_dico['highest_screw_length']['unit']}"\
                              f"Thread: "\
                              f"{round(self.metric_data_dico['calculated_screw_thread']['value'],4)}"\
                              f"{self.metric_data_dico['calculated_screw_thread']['unit']}")
                                
                        self.ScrewOut4Var.set(round(
                                self.metric_data_dico['highest_screw_length']['value'],2))
                        self.ScrewOut5Var.set(round(
                                self.metric_data_dico['tension_force_perpendicular_offset_increase'
                                                      ]['value'],2))
                    else:
                        self.metric_data_dico['highest_screw_length']['value'] = None
                        self.metric_data_dico['calculated_screw_thread']['value'] = None
                        self.ScrewOut4Var.set(round(
                                self.metric_data_dico['highest_screw_length']['value'],2))
                        print('Input screw quantity and type')
                except:
                    self.metric_data_dico['highest_screw_length']['value'] = None
                    self.metric_data_dico['calculated_screw_thread']['value'] = None
                    self.ScrewOut4Var.set('NA')
                    print('Error, Input screw quantity and type')

                self.pass_data_from_metric_to_imperial_dico()
            return
        
        def ReportScrew(self):
            try:
                ScrewData = self.ASSY_ScrewCal()
                highest_screw_length = ScrewData[4]
                print(highest_screw_length)
                Input = self.Input
                Input += ScrewData
            except:
                print('invalid data')
            Re.Report(Input)
            if highest_screw_length !=0: self.ASSY_ScrewWindow.destroy()

        ### Panel GUI and functions    
        def Glued_PanelGUI(self):  
            '''
                This is the panel GUI to display the results for a panel reinforcer.                
            '''
            self.Glued_PanelWindow = Toplevel(self.master)
            self.Glued_PanelWindow .wm_title("Repair parameters")
            self.Glued_PanelWindow .attributes('-topmost', True)    

            panel_data_key_list = [
                                ['panel_height','in','mm'],
                                ['panel_width','in','mm'],
                                ['panel_nail_quantity',None,None],
                                ['in_line_nail_spacing','in','mm'],
                                ['nail_row_spacing','in','mm']
                                ]
            self.add_receptecals_to_data_dico(panel_data_key_list)
#            for selected_key in panel_data_key_list:
#                self.imperial_data_dico[selected_key[0]] = {'unit':selected_key[1],'value':None}
#                self.metric_data_dico[selected_key[0]] =  {'unit':selected_key[2],'value':None}

            try:
                (self.metric_data_dico['panel_height']['value'],
                 self.metric_data_dico['panel_width']['value'],
                 self.metric_data_dico['panel_nail_quantity']['value'],
                 self.metric_data_dico['in_line_nail_spacing']['value'],
                 self.metric_data_dico['nail_row_spacing']['value']) = (
                Rfu.PanelRepair(
                    self.metric_data_dico['tension_force_perpendicular']['value'],
                    self.metric_data_dico['Nordic_Lam_ply_quantity']['value'],
                    self.metric_data_dico['Nordic_Lam_depth']['value'],
                    self.metric_data_dico['hole_diameter']['value']))
            except:
                print('Panel repair function failed to load:\n'\
                      f"{self.metric_data_dico['tension_force_perpendicular']['value']}, "\
                      f"{self.metric_data_dico['Nordic_Lam_ply_quantity']['value']}, "\
                      f"{self.metric_data_dico['Nordic_Lam_depth']['value']}, "\
                      f"{self.metric_data_dico['hole_diameter']['value']}")

            self.pass_data_from_metric_to_imperial_dico()

#            for selected_key in self.metric_data_dico:
#                self.imperial_data_dico[selected_key]['value'] = (
#                        convert(
#                                self.metric_data_dico[selected_key]['value'],
#                                self.metric_data_dico[selected_key]['unit'],
#                                self.imperial_data_dico[selected_key]['unit']))            
            
            data_dico  = self.dico_in_use()
            round_at = self.round_factor_in_use(2,1)
#            if self.selected_unit.get().lower() == 'imperial':
#                data_dico = self.imperial_data_dico
#                round_at = 2
#                
#            elif self.selected_unit.get().lower() == 'metric':
#                data_dico = self.metric_data_dico
#                round_at = 1
            
        #Column 0 Input values & buttons
            Col = 0
            Ro = 0
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
            
            TextOut1Var = ('Repair detail based on two, full beam height, 3/4''"'' CSP'\
                           '\nplywood installed grain perpendicular to beam span'\
                           '\non either sides'\
                           '\n\nPanel fastened to beam with PL Premium or better structural'\
                           '\nadhesive and 0.131''"'' x 3.25''"'' round head driven nails.'\
                           f'\nIn line spacing: {try_round_NA(data_dico["in_line_nail_spacing"]["value"],round_at=round_at)}'\
                           f' {data_dico["in_line_nail_spacing"]["unit"]}'\
                           f'\nRow spacing: {try_round_NA(data_dico["nail_row_spacing"]["value"],round_at=round_at)}'\
                           f' {data_dico["nail_row_spacing"]["unit"]}')
                           
            TextOut2Var = (f'Panel dimensions: {try_round_NA(data_dico["panel_height"]["value"],round_at=round_at)}'\
                            f' {data_dico["panel_height"]["unit"]} x '\
                            f'{try_round_NA(data_dico["panel_width"]["value"],round_at=round_at)}'\
                            f' {data_dico["panel_width"]["unit"]}'\
                            '\nNail minimum quantity: '\
                            f'{data_dico["panel_nail_quantity"]["value"]}'\
                            '\nTension perpendicular to grain: '\
                            f'{try_round_NA(data_dico["tension_force_perpendicular"]["value"])}'\
                            f' {data_dico["tension_force_perpendicular"]["unit"]}')
                
            self.PanelOut1Var.set(TextOut1Var)
            self.PanelOut2Var.set(TextOut2Var)
            
        def ReportPanel(self):
            try:
                ReinforceMethod = self.Reinforce.get()
                Input = self.InputProcess()
                #### Update with report based on unit type
                PanelData = (ReinforceMethod, 
                             self.metric_data_dico['panel_height']['Value'],
                             self.metric_data_dico['panel_width']['value'],
                             self.metric_data_dico['panel_nail_quantity']['value'])
                Input += PanelData
            except:
                print('invalid data')
            
            Re.Report(Input)
            self.Glued_PanelWindow.destroy()            

#Generate an analysis and a report based on the inputed data.        
        def general_report(self):
            self.InputProcess()
            selected_data_dico =  self.dico_in_use()
            Re.Report(selected_data_dico)

#Convert input between imperial and international systems.
        def Radio(self):
            if self.Units[self.selected_unit_index] != self.selected_unit.get():
                self.selected_unit_index= 0 if self.selected_unit.get() == 'Imperial' else 1
                if self.selected_unit.get() == 'Imperial':
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
                self.LblInput2Txt.set(self.Lbl2+self.UnitDict[self.ID2][self.selected_unit_index])            
                self.Lbl_input_Nordic_Lam_depth_3_Txt.set(self.Lbl3+self.UnitDict[self.ID3][self.selected_unit_index])
                self.label_input_hole_diameter_4_text.set(self.Lbl4+self.UnitDict[self.ID4][self.selected_unit_index])
                self.label_input_bending_force_Mf_5_text.set(self.Lbl5+self.UnitDict[self.ID5][self.selected_unit_index])
                self.label_input_bending_resistance_Mr_6_text.set(self.Lbl6+self.UnitDict[self.ID6][self.selected_unit_index])
                self.label_input_shear_force_Vf_7_text.set(self.Lbl7+self.UnitDict[self.ID7][self.selected_unit_index])
                self.LblOutShearMaxLengthTxt.set(self.LblShearMaxLength+self.UnitDict[self.IDOutShearMaxLength][self.selected_unit_index])
                self.label_input_shear_force_Wf_10_text.set(self.Lbl10+self.UnitDict[self.ID10][self.selected_unit_index])
                self.label_input_shear_resistance_Wr_11_text.set(self.Lbl11+self.UnitDict[self.ID11][self.selected_unit_index])
                
                if self.selected_unit.get() == "Imperial":
                    self.Nordic_Lam_thickness_List = self.Nordic_Lam_imperial_thickness_List
                    self.Nordic_Lam_depth_List = self.Nordic_Lam_imperial_depth_List
                else:
                    self.Nordic_Lam_thickness_List = self.Nordic_Lam_metric_thickness_List
                    self.Nordic_Lam_depth_List = self.Nordic_Lam_metric_depth_List

                self.input_Nordic_Lam_thickness_2['values']=self.Nordic_Lam_thickness_List
                self.input_Nordic_Lam_depth_3['values']=self.Nordic_Lam_depth_List

#-_- #Create a dico with inputs and units.        
                InputList = (self.input_Nordic_Lam_thickness_2,self.input_Nordic_Lam_depth_3,self.input_hole_diameter_4)
                
                for i in InputList:
                    try:
                        self.ConvVal = convert(float(i.get()),FUni,SUni)
                        i.delete(0,END)
                        i.insert(0,round(self.ConvVal,2))
                    except:
                        i.delete(0,END)
                
                InputList = (self.input_bending_force_Mf_5,self.input_bending_resistance_Mr_6)
                for i in InputList:
                    try:
                        self.ConvVal = convert(float(i.get()),FMom,SMom)
                        i.delete(0,END)
                        i.insert(0,round(self.ConvVal,3))
                    except:
                        i.delete(0,END)

                InputList = (self.input_shear_force_Vf_7,self.input_shear_force_Wf_10,self.input_shear_resistance_Wr_11)
                for i in InputList:
                    try:
#                        print(i.get())
                        self.ConvVal = convert(float(i.get()),FFor,SFor)
                        i.delete(0,END)
                        i.insert(0,round(self.ConvVal,3))
                    except:
                        i.delete(0,END)

                # Special code for variable label
                InputList = [self.max_length_to_use_shear_resistance_b]
                for i in InputList:
                    try:
#                        print('This is SMLength: ',self.max_length_to_use_shear_resistance_b)
                        self.max_length_to_use_shear_resistance_b = convert(float(self.max_length_to_use_shear_resistance_b),FUni2,SUni2)
                        self.ShearMaxLength.set(str(round(self.max_length_to_use_shear_resistance_b,0)))
                    except:
                        self.ShearMaxLength.set("")
                        print('Invalid data to convert the max length to use equation 7.5.7.2 (b)')
            return

root = Tk()
GUI = MainInt(root)
root.iconbitmap(icon)
root.mainloop()            