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

def PhysicsAnalysis(single_phy_resp,oorm_phy_resp,para_phy_resp,int_phy_resp,
                    single_phy_key,oorm_phy_key,para_phy_key,int_phy_key, studentName):
    
    oorm_phy_resp_list = []
    oorm_phy_key_list = []
    
    for i in range(len(oorm_phy_key)):
        oorm_phy_key_list.append(oorm_phy_key.iloc[i][1])
        oorm_phy_resp_list.append(oorm_phy_resp.iloc[i])
        
    phy_single_correct = 0
    for i in range(len(single_phy_key)):
        if single_phy_resp.iloc[i] == single_phy_key.iloc[i][1]:
                phy_single_correct += 4
        else:
                phy_single_correct = -1
    print("Physics Single Result : ", phy_single_correct)
    
    
    phy_oorm_correct = 0
    for i in range(len(oorm_phy_key_list)):
        if oorm_phy_resp_list[i] == oorm_phy_key_list[i]: 
            phy_oorm_correct += 4
        else : 
            phy_single_correct = -1
    print("Physics One or More Result : ",phy_oorm_correct)
    
    
    # print("_"*5+"Response"+"_"*5)
    # for i in range(len(para_correct_phy_resp)):
    #     print(para_correct_phy_resp.iloc[i])
    # print("_"*5+"Key"+"_"*5)
    # for i in range(len(para_correct_phy_key)):
    #     print(para_correct_phy_key.iloc[i][1])
    # return None

    phy_para_correct = 0
    for i in range(len(para_phy_key)):
        if para_phy_resp.iloc[i] == para_phy_key.iloc[i][1]:
                phy_para_correct += 4
        else:
                 phy_para_correct -= 1
    print("Physics Para Result : ",phy_para_correct)
    
    phy_int_correct = 0
    for i in range(len(para_phy_key)):
        if int_phy_resp.iloc[i] == int_phy_key.iloc[i][1]:
                phy_int_correct += 4
        else:
                phy_int_correct -= 1
    print("Physics Integer Result : ",phy_int_correct)
    
    physics_total = phy_single_correct + phy_oorm_correct + phy_para_correct + phy_int_correct
    return physics_total

def MathAnalysis(single_math_resp,oorm_math_resp,para_math_resp,int_math_resp,
                    single_math_key,oorm_math_key,para_math_key,int_math_key
                    ,studentName):
    
    oorm_math_resp_list = []
    oorm_math_key_list = []
    
    for i in range(len(oorm_math_key)):
        oorm_math_key_list.append(oorm_math_key.iloc[i][1])
        oorm_math_resp_list.append(oorm_math_resp.iloc[i])
        
    math_single_correct = 0
    for i in range(len(single_math_key)):
        if single_math_resp.iloc[i] == single_math_key.iloc[i][1]:
                math_single_correct += 4
        else:
                math_single_correct = -1
    print("Math Single Result : ", math_single_correct)
    
    
    math_oorm_correct = 0
    for i in range(len(oorm_math_key_list)):
        if oorm_math_resp_list[i] == oorm_math_key_list[i]: 
            math_oorm_correct += 4
        else : 
            math_single_correct = -1
    print("Math One or More Result : ",math_oorm_correct)
    
    
    # print("_"*5+"Response"+"_"*5)
    # for i in range(len(para_correct_phy_resp)):
    #     print(para_correct_phy_resp.iloc[i])
    # print("_"*5+"Key"+"_"*5)
    # for i in range(len(para_correct_phy_key)):
    #     print(para_correct_phy_key.iloc[i][1])
    # return None

    math_para_correct = 0
    for i in range(len(para_math_key)):
        if para_math_resp.iloc[i] == para_math_key.iloc[i][1]:
                math_para_correct += 4
        else:
                 math_para_correct -= 1
    print("Math Para Result : ",math_para_correct)
    
    math_int_correct = 0
    for i in range(len(para_math_key)):
        if int_math_resp.iloc[i] == int_math_key.iloc[i][1]:
            math_int_correct += 4
        else:
            math_int_correct -= 1
    math_total = math_single_correct + math_oorm_correct + math_para_correct + math_int_correct     
    return math_total

def ChemAnalysis(single_chem_resp,oorm_chem_resp,para_chem_resp,int_chem_resp,
                single_chem_key,oorm_chem_key,para_chem_key,int_chem_key
                ,studentName):
    oorm_chem_resp_list = []
    oorm_chem_key_list = []
    
    for i in range(len(oorm_chem_key)):
        oorm_chem_key_list.append(oorm_chem_key.iloc[i][1])
        oorm_chem_resp_list.append(oorm_chem_resp.iloc[i])
        
    chem_single_correct = 0
    for i in range(len(single_chem_key)):
        if single_chem_resp.iloc[i] == single_chem_key.iloc[i][1]:
                chem_single_correct += 4
        else:
                chem_single_correct = -1
    print("chem Single Result : ", chem_single_correct)
    
    
    chem_oorm_correct = 0
    for i in range(len(oorm_chem_key_list)):
        if oorm_chem_resp_list[i] == oorm_chem_key_list[i]: 
            chem_oorm_correct += 4
        else : 
            chem_single_correct = -1
    print("chem One or More Result : ",chem_oorm_correct)
    
    
    # print("_"*5+"Response"+"_"*5)
    # for i in range(len(para_correct_phy_resp)):
    #     print(para_correct_phy_resp.iloc[i])
    # print("_"*5+"Key"+"_"*5)
    # for i in range(len(para_correct_phy_key)):
    #     print(para_correct_phy_key.iloc[i][1])
    # return None

    chem_para_correct = 0
    for i in range(len(para_chem_key)):
        if para_chem_resp.iloc[i] == para_chem_key.iloc[i][1]:
                chem_para_correct += 4
        else:
                 chem_para_correct -= 1
    print("chem Para Result : ",chem_para_correct)
    
    chem_int_correct = 0
    for i in range(len(para_chem_key)):
        if int_chem_resp.iloc[i] == int_chem_key.iloc[i][1]:
            chem_int_correct += 4
        else:
            chem_int_correct -= 1
    chem_total = chem_single_correct + chem_oorm_correct + chem_para_correct + chem_int_correct
    return chem_total

def QuestionsAnalysis(studentRowData,key_data):
    # print("_"*20)
    # for i in range(len(studentRowData)):
    #     print(studentRowData[i])    
    
    studentRowData.to_frame()
 
    
    single_phy_resp, oorm_phy_resp, para_phy_resp, int_phy_resp  = studentRowData.iloc[4:12], studentRowData.iloc[12:17], studentRowData.iloc[17:22],studentRowData.iloc[22:32]        
    single_math_resp, oorm_math_resp, para_math_resp, int_math_resp = studentRowData.iloc[32:40], studentRowData.iloc[40:45], studentRowData.iloc[45:50],studentRowData.iloc[50:60]  
    single_chem_resp, oorm_chem_resp, para_chem_resp, int_chem_resp = studentRowData.iloc[32:40], studentRowData.iloc[40:45], studentRowData.iloc[45:50],studentRowData.iloc[50:60]  

    single_phy_key, oorm_phy_key, para_phy_key, int_phy_key  = key_data.iloc[0:8], key_data.iloc[8:13], key_data.iloc[13:18], key_data.iloc[18:28]
    single_math_key, oorm_math_key, para_math_key, int_math_key  = key_data.iloc[28:36], key_data.iloc[36:41], key_data.iloc[41:46], key_data.iloc[46:56]
    single_chem_key, oorm_chem_key, para_chem_key, int_chem_key  = key_data.iloc[56:64], key_data.iloc[64:69], key_data.iloc[69:74], key_data.iloc[74:84]
    
    studentName = studentRowData.iloc[2]
        

    physics_total = PhysicsAnalysis(single_phy_resp,oorm_phy_resp,para_phy_resp,int_phy_resp,
                    single_phy_key,oorm_phy_key,para_phy_key,int_phy_key
                    ,studentName)
    print("Total chem marks scored by "+ studentName +" : ", physics_total)
    print("_"*50)    
    
    math_total = MathAnalysis(single_math_resp,oorm_math_resp,para_math_resp,int_math_resp,
                    single_math_key,oorm_math_key,para_math_key,int_math_key
                    ,studentName)
    print("Total Math marks scored by "+ studentName +" : ", math_total)
    print("_"*50)    

    chem_total = ChemAnalysis(single_chem_resp,oorm_chem_resp,para_chem_resp,int_chem_resp,
                single_chem_key,oorm_chem_key,para_chem_key,int_chem_key
                ,studentName)
    print("Total chem marks scored by "+ studentName +" : ", chem_total)
    print("_"*50)    
    
    total = physics_total+math_total+chem_total
    
    print("Total Marks scored by "+studentName+" : ",total)
    print("_"*50)
    # print(oorm_correct_phy_key_list)
    # print(oorm_correct_phy_resp_list)
    
    # print("_"*5+"Response"+"_"*5)
    # for i in range(len(single_phy_resp)):
    #     print(single_phy_resp.iloc[i])
    # print("_"*5+"Key"+"_"*5)
    # for i in range(len(single_phy_key)):
    #     print(single_phy_key.iloc[i][1])
    
    return None    

def StudentNullAnalysis(studentRowData):
    physics_slice = studentRowData.iloc[4:32]
    print("Number of Unanswered Questions in Physics by "+studentRowData[2]+" is :", physics_slice.isna().sum())
    chemistry_slice = studentRowData.iloc[32:60]
    print("Number of Unanswered Questions in Chemistry by "+studentRowData[2]+" is :", chemistry_slice.isna().sum())    
    mathematics_slice = studentRowData.iloc[60:90]
    print("Number of Unanswered Questions in Mathematics by "+studentRowData[2]+" is :", mathematics_slice.isna().sum())    
    print("_"*50)
    print("Total number of unanswered questions by "+ studentRowData[2]+" are ", studentRowData[4:90].isna().sum())    
    return None

def main():
    response_sheet = "data/sheet_1_resp.xlsx"
    key_sheet = "data/sheet_1_key.xlsx"
    
    
    response_data, key_data = readData(response_sheet, key_sheet)
    
    
    response_data = response_data.sort_values('OMR No.',ascending=False ,na_position='first')
    response_data = response_data.drop('Timestamp',axis=1)
    
    number_colums = len(response_data)
    
    #TEST Col            
    rowData = response_data.iloc[0]  
       
    QuestionsAnalysis(rowData, key_data)
    StudentNullAnalysis(rowData)
    
    
    #TODO : Append 
    
    # for i in range(number_colums):
    #     rowData = response_data.iloc[i]
    #     QuestionsAnalysis(rowData, key_data)
    #     StudentNullAnalysis(rowData)
        
    
    
    # print("\n","-"*20,"\n")
    # print(sheet1KeyFromExcel)
    
if __name__ == "__main__":
    main()