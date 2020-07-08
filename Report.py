# -*- coding: utf-8 -*-
"""
Created on Fri Feb  1 08:54:19 2019
V 6.04 (2019-04-04)
@author: ludovicraymond
"""
import datetime
import time
import os
import pandas as pd
import NordicLamSplitAnalysisFunctions as fu
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from tkinter import messagebox

SheetNRep = 'Report'

###-_-
### Update: Add to selected deco the V or W values based on the Shear equation value. 
# special Vf or Wf new var = 'shear_force', special Vr or Wr new var = 'shear_resistance', 
### ADD shear_resistance_Vf

report_file_path = 'reports\\'
report_sheet_name = 'report'

def round_report_data(value):
    '''
        Converts numerical data with the appropriate precision.
        Let non numerical data pass by unchanged.
        converts the result to string.
    '''
    if type(value) == type(int()):
        value_out = value
    else:
        try:
            float_value = float(value)
            value_out = round(float_value, 2)
            if len(str(value_out)) <= 4:
                value_out = round(float_value, 4)
            elif len(str(value_out)) >= 6:
                value_out = int(round(float_value, 0))
        except:
            value_out = value
    value_out = str(value_out)
    return value_out

def Report(selected_data_dico, template_file_path):
    '''
        Organizes the data dico information in a formated excel sheet.
    '''
    if selected_data_dico['shear_method_a_used']['value'] == 1:
        selected_data_dico['shear_force']['value'] = selected_data_dico['shear_force_Wf']['value']
        selected_data_dico['shear_resistance']['value'] = selected_data_dico['shear_resistance_Wf']['value']
        selected_data_dico['shear_method_a_used_txt']['value'] = '(a) Wr'
    elif selected_data_dico['shear_method_a_used']['value'] == 0:
        selected_data_dico['shear_force']['value'] = selected_data_dico['shear_force_Vf']['value']
        selected_data_dico['shear_resistance']['value'] = selected_data_dico['shear_resistance_Vf']['value']    
        selected_data_dico['shear_method_a_used_txt']['value'] = '(b) Vr'

    Nordic_Lam_split_analyser_report_template = (
            pd.read_excel(template_file_path, sheet_name=0, header=0, index_col=0))    
    lookup_list = Nordic_Lam_split_analyser_report_template.index

    for selected_index in lookup_list:
        try:
            Nordic_Lam_split_analyser_report_template.at[selected_index,'value'] = round_report_data(
                    selected_data_dico[selected_index]['value'])
        except:
            pass
        try:
            Nordic_Lam_split_analyser_report_template.at[selected_index,'unit'] = str(
                    selected_data_dico[selected_index]['unit'])
        except:
            pass
    Nordic_Lam_split_analyser_report_template = (
            Nordic_Lam_split_analyser_report_template.replace(to_replace='None',value=''))

    Nordic_Lam_split_analyser_report_template = (
            Nordic_Lam_split_analyser_report_template [['label','value','unit','description']])
    
    # Build first set of data
    df = Nordic_Lam_split_analyser_report_template

    #Generate File name with a time stamp
    time_stamp = str(datetime.datetime.fromtimestamp(time.time()).strftime('%Y%m%d_%H_%M_%S'))
    file_name = report_file_path + df.at['report_type','label'] + '_' + time_stamp + '.xlsx'
    writer = pd.ExcelWriter(file_name)
    df.to_excel(writer, header=False, index=False, sheet_name=report_sheet_name)

    #Add inforation
    if os.getlogin() == 'ludovicraymond':
        designer = 'Ludovic Raymond'
    else:
        designer = os.getlogin()
    Notes = [('Project: ',selected_data_dico['project_name']['value']),
             ('Notes: ',selected_data_dico['Notes']['value']),
             ('Date: ', str(datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d'))),
             ('Designer: ', designer)] 
    df1 = pd.DataFrame(data=Notes)
    df1.to_excel(writer, index=False, header=False, sheet_name=report_sheet_name, startrow = df.shape[0] + 1)

# Add the references
#    df3 = pd.DataFrame(data=References)
#    df3.to_excel(writer, index=False, header=False, sheet_name=report_sheet_name, startcol = 1, startrow = df.shape[0] + df1.shape[0] + 2)

    #Save all the inputed data
    writer.save()

#   Create a workbook object.
    report_workbook = Workbook()
    report_workbook = load_workbook(file_name)

#   Open up the last edited workbook sheet.    
    sheet=report_workbook.active

    # Set column width
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 5
    sheet.column_dimensions['D'].width = 60

    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    sheet.oddHeader.left.text = f"Project: {selected_data_dico['project_name']['value']}"

    font_header_0 = Font(name='Calibri',
                  size=12,
                  bold=True,
                  italic=False,
                  vertAlign=None,
                  underline='none',
                  strike=False,
                  color='FF000000')
    
    heading_list = ('Nordic Lam specifications:', 'Analysis:',
                    'Splitting Analysis:', 'Reinforcement Analysis:')
    
    counter = 0
    while counter < df.shape[0]:
        for selected_heading in heading_list:
            if df.iloc[counter,0] == selected_heading:
                cell_name = 'A' + str(counter+1)
                cell_target = sheet[cell_name]
                cell_target.font = font_header_0
                sheet.row_dimensions[cell_target.row].height = 16
        counter += 1    
        
    font_header_1 = Font(name='Calibri',
                  size=24,
                  bold=True,
                  italic=False,
                  vertAlign=None,
                  underline='none',
                  strike=False,
                  color='FF000000')

    cell_report_title = sheet['A1']
    cell_report_title.font = font_header_1
    sheet.row_dimensions[cell_report_title.row].height = 36        

    align_top_left_wrap = Alignment(horizontal='left',
               vertical='top',
               wrap_text=True)

    align_top_right_wrap = Alignment(horizontal='right',
                   vertical='top',
                   wrap_text=True)

    merge_row_list = [df.shape[0] + 2, df.shape[0] + 3, df.shape[0] + 4, df.shape[0] + 5]
    deep_row_list = [df.shape[0] + 3,]

    for row_index in merge_row_list:
        target_cell_name = 'B' + str(row_index)
        if target_cell_name == 'B1':
            target_cell_name = 'A1'
        target_cell = sheet[target_cell_name]
        target_cells_merge = target_cell_name + ':D' + str(target_cell.row)
        target_cell.alignment = align_top_left_wrap 


        for selected_row in deep_row_list:
            cell_to_deepen = 'B' + str(selected_row)
            if target_cell_name == cell_to_deepen:
                sheet.row_dimensions[target_cell.row].height =104    

        sheet.merge_cells(target_cells_merge)

    selected_row = 4
    while selected_row <= df.shape[0]:
        selected_cell_name = 'A' + str(selected_row)
        selected_cell = sheet[selected_cell_name]
        selected_cell.alignment = align_top_right_wrap
        selected_row += 1

    selected_row = df.shape[0]+1
    while selected_row <= df.shape[0] + df1.shape[0]:
        selected_cell_name = 'A' + str(selected_row)
        selected_cell = sheet[selected_cell_name]
        selected_cell.alignment = align_top_left_wrap
        selected_row += 1
        
    counter = 0
    while counter < df.shape[0]:
        for selected_heading in heading_list:
            if df.iloc[counter,0] == selected_heading:
                cell_name = 'A' + str(counter+1)
                cell_target = sheet[cell_name]
                cell_target.alignment = align_top_left_wrap
                sheet.row_dimensions[cell_target.row].height = 16
        counter += 1    

    selected_row = 4
    while selected_row <= df.shape[0]:
        selected_cell_name = 'B' + str(selected_row)
        selected_cell = sheet[selected_cell_name]
        selected_cell.alignment = align_top_right_wrap
        selected_row += 1
    
    report_workbook.save(file_name) 
    
    pass


