#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun  5 19:20:31 2020

@author: Madhav
"""

import pandas as pd
import numpy as np
import os

def getFilename(filename):
    #Get the file name to append for .csv
    csv_filename = os.path.splitext(filename)[0]
    return csv_filename

def readData(response_sheet, key_sheet):
    #Read responses sheet and convert it to .csv
    response_data = pd.read_excel(response_sheet, index_col=None, dtype=str)
    response_data.to_csv(getFilename(response_sheet)+".csv", header=True)
    
    #Read key sheet and convert it to .csv and drop index
    key_data = pd.read_excel(key_sheet, index_col=None, dtype=str)
    key_data.to_csv(getFilename(key_sheet)+".csv")
    
    #Returning the read data
    return response_data,key_data

def QuestionsAnalysis(studentRowData,key_data):
    # print("_"*20)
    # for i in range(len(studentRowData)):
    #     print(studentRowData[i])    
    
    studentRowData.to_frame()
    
    # print("Phy:")
    physics_slice = studentRowData.iloc[4:32]
    # print(physics_slice)
    # print("Chem:")
    chemistry_slice = studentRowData.iloc[32:60]
    # print(chemistry_slice)        
    # print("Math:")
    mathematics_slice = studentRowData.iloc[61:90]
    # print(mathematics_slice)  
    
    single_correct_phy_resp = studentRowData.iloc[4:12]
    oorm_correct_phy_resp = studentRowData.iloc[12:17]
    para_correct_phy_resp = studentRowData.iloc[17:22]
    
    single_correct_phy_key = key_data.iloc[0:8]
    oorm_correct_phy_key = key_data.iloc[8:13]
    para_correct_phy_key = key_data.iloc[13:18]
        
    oorm_correct_phy_resp_list = []
    oorm_correct_phy_key_list = []
    
    for i in range(len(oorm_correct_phy_key)):
        oorm_correct_phy_key_list.append(oorm_correct_phy_key.iloc[i][1])
        oorm_correct_phy_resp_list.append(oorm_correct_phy_resp.iloc[i])
        
    # print(oorm_correct_phy_key_list)
    # print(oorm_correct_phy_resp_list)
    
    # print("_"*5+"Response"+"_"*5)
    # for i in range(len(single_correct_phy_resp)):
    #     print(single_correct_phy_resp.iloc[i])
    # print("_"*5+"Key"+"_"*5)
    # for i in range(len(single_correct_phy_key)):
    #     print(single_correct_phy_key.iloc[i][1])
    

    phy_single_correct = 0
    for i in range(len(single_correct_phy_key)):
        if single_correct_phy_resp.iloc[i] == single_correct_phy_key.iloc[i][1]:
                phy_single_correct += 1
        else:
            pass
    print("Physics Single Result : ", phy_single_correct)
    
    

    phy_oorm_correct = 0
    for i in range(len(oorm_correct_phy_key_list)):
        if oorm_correct_phy_resp_list[i] == oorm_correct_phy_key_list[i]: 
            phy_oorm_correct += 1
        else : 
            pass
    print("Physics One or More Result : ",phy_oorm_correct)
    
    
    # print("_"*5+"Response"+"_"*5)
    # for i in range(len(para_correct_phy_resp)):
    #     print(para_correct_phy_resp.iloc[i])
    # print("_"*5+"Key"+"_"*5)
    # for i in range(len(para_correct_phy_key)):
    #     print(para_correct_phy_key.iloc[i][1])
    # return None

    phy_para_correct = 0
    for i in range(len(para_correct_phy_key)):
        if para_correct_phy_resp.iloc[i] == para_correct_phy_key.iloc[i][1]:
                phy_para_correct += 1
        else:
            pass
    print("Physics Para Result : ",phy_para_correct)
    
    
    print("Total Physics marks scored by "+ studentRowData[2] +" : ", phy_single_correct+phy_oorm_correct+phy_para_correct)
    
    
    

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
    response_sheet = "data/sheet_1_resp.xlsx"
    key_sheet = "data/sheet_1_key.xlsx"
    
    
    response_data, key_data = readData(response_sheet, key_sheet)
    
    
    response_data = response_data.sort_values('Student Name',na_position='first')
    response_data = response_data.drop('Timestamp',axis=1)
    
    number_colums = len(response_data)
    
    #TEST Col            
    rowData = response_data.iloc[0]  
       
    #TODO: Compare Key and Responses 
    print("_"*40)
    QuestionsAnalysis(rowData, key_data)
    print("_"*40)
    StudentNullAnalysis(rowData)
    
    #TODO : Append 
    
    for i in range(number_colums):
        rowData = response_data.iloc[i]
        QuestionsAnalysis(rowData, key_data)
        StudentNullAnalysis(rowData)
        
    
    
    # print("\n","-"*20,"\n")
    # print(sheet1KeyFromExcel)
    
if __name__ == "__main__":
    main()