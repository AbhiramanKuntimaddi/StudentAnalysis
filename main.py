#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun  5 19:20:31 2020

@author: Madhav
"""

import pandas as pd
import numpy as np
import os

def readData(excelSheet):
    csv_file_name = os.path.splitext(excelSheet)[0]
    data = pd.read_excel(excelSheet,index_col=None, dtype=str)
    data.to_csv(csv_file_name+".csv", header=True)
    return data



def QuestionsAnalysis(studentRowData):
    # print("_"*20)
    # for i in range(len(studentRowData)):
    #     print(studentRowData[i])
    
    print("Phy:")
    physics_slice = studentRowData.iloc[4:32]
    print(physics_slice)

    print("Chem:")
    chemistry_slice = studentRowData.iloc[32:60]
    print(chemistry_slice)    
    
    print("Math:")
    mathematics_slice = studentRowData.iloc[61:90]
    print(mathematics_slice)    

    
   
    return None

def StudentNullAnalysis(studentRowData):
    physics_slice = studentRowData.iloc[4:32]
    print("Number of Unanswered Questions in Physics by "+studentRowData[2]+" is :", physics_slice.isna().sum())

    chemistry_slice = studentRowData.iloc[32:60]
    print("Number of Unanswered Questions in Chemistry by "+studentRowData[2]+" is :", chemistry_slice.isna().sum())
    
    mathematics_slice = studentRowData.iloc[60:90]
    print("Number of Unanswered Questions in Mathematics by "+studentRowData[2]+" is :", mathematics_slice.isna().sum())
    
    print("Total number of unanswered questions by "+ studentRowData[2]+" are ", studentRowData[4:90].isna().sum())
    
    return None

def main():
    excelSheet = "data/sheet_1_resp.xlsx"
    sheet1_key = "data/sheet_1_key.xlsx"
    dataFromExcel = readData(excelSheet)
    sheet1KeyFromExcel = readData(sheet1_key)
    dataFromExcel = dataFromExcel.sort_values('Student Name',na_position='first')
    dataFromExcel = dataFromExcel.drop('Timestamp',axis=1)
    
    number_colums = len(dataFromExcel)
    
    print(dataFromExcel.info)
    
    print(number_colums)
    
    # rowData = dataFromExcel.iloc[0]
    
    
    # #TODO: Compare Key and Responses 
    # #QuestionsAnalysis(rowData)
    
    # StudentNullAnalysis(rowData)
    
    
    #TODO : Append 
    
    for i in range(number_colums):
        rowData = dataFromExcel.iloc[i]
        StudentNullAnalysis(rowData)
        
    
    
    # print("\n","-"*20,"\n")
    # print(sheet1KeyFromExcel)
    
if __name__ == "__main__":
    main()