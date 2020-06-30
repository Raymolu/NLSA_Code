# -*- coding: utf-8 -*-
"""
Created on Wed Dec 26 15:06:07 2018
V 6.04 (2019-04-04)
@author: ludovicraymond
"""
# Panda for advanced data handling. First 3 rows of excel sheet are for Title and notes. These will not be condidered by the program

import pandas as pd

class XLTable:
    
    def __init__(self,KeyQty,FileName):
        #Table caracteristics (True or False)
        self.KeyQty=KeyQty  
        self.FileName=FileName  

    #Transfer Excel file to pandas dataframe. Note: xls file and python or exe file need to be in the same folder.
    def XLToTable(self):  
        if self.KeyQty is 1:
            TbDataFrame = pd.read_excel(self.FileName, sheet_name=0, header=3, index_col=0)
        elif self.KeyQty is 2:
            TbDataFrame = pd.read_excel(self.FileName, sheet_name=0, header=3, index_col=[0,1])
        else:
            print("Enter valid quantity of row index")
        return TbDataFrame

    #Create Series Name List. If Index composite Series name will be a list of tuples.
    def SeriesName(self):
        RawSeriesName = self.XLToTable().index.tolist()[0:-1]
        SeriesName=[]
        if self.KeyQty is 2:
            for i in RawSeriesName:
                SeriesName+=[str(i[0])+" "+str(i[1])]  
        elif self.KeyQty is 1:
            SeriesName = self.XLToTable().index.tolist()[0:-1]
        return SeriesName
    
    #Return the Table units List
    def ThisTbUnits(self):
        Units = self.XLToTable().loc[self.XLToTable().index.tolist()[-1],self.XLToTable().columns.tolist()].tolist()
        return Units
    
    #Return the Table Labels or column headers
    def ThisTbLabels(self):
        Labels=self.XLToTable().columns.tolist()
        return Labels
    
    #Create dictionary with Series Name List and spec dictionary
    #Two first column are the key and a spec dictionary is the returned entry

    def GenDict(self):
        SpecDict = {}
        index = 0
        for i in self.XLToTable().index.tolist()[0:-1]:
            DataDict = {}
            for ii in self.ThisTbLabels():
                DataDict[ii]=self.XLToTable().at[i,ii]
            SpecDict[self.SeriesName()[index]]=DataDict
            index+=1
        return SpecDict

    def UnitDict(self):
        UnitDict={}
        Index=0
        for i in self.ThisTbLabels():
            UnitDict[i]=self.ThisTbUnits()[Index]
            Index+=1
        return UnitDict

#Example For testing:
#________________________________________________
#tb1 = XLTable(1, "NordicLamSpecs.xlsx")
#tb2 = XLTable(2, "NordicJoistSpecs.xlsx")

#print(tb1.UnitDict())
#print(tb1.ThisTbUnits())
#print(tb2.ThisTbUnits())

#print(tb1.GenDict())
#print(tb2.GenDict())

#print(tb1.GenDict()["NPG"])
#print(tb2.GenDict()["NI-20 9 1/2"])
#print(tb1.GenDict()["NPG"]["Fvx"])
#print(tb2.GenDict()["NI-40x 11 7/8"]["Vr"])

#print(tb1.UnitDict())
#print(tb2.UnitDict())
#print(tb1.UnitDict()["Fby"])
#print(tb2.UnitDict()["Mr"])