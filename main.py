#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun  5 19:20:31 2020

@author: Madhav
"""

from styleframe import StyleFrame, Styler
import pandas as pd
import os

def getFilename(filename):
    #Get the file name to append for .csv
    csv_filename = os.path.splitext(filename)[0]
    return csv_filename

def readData(response_sheet, key_sheet):
    #Read responses sheet and convert it to .csv
    response_data = pd.read_excel(response_sheet, index_col=None,  dtype=str)  
    response_data.fillna('Nan', inplace=True)
    response_data.to_csv(getFilename(response_sheet)+".csv", header=True, mode='w',sep=',',encoding='utf-8')
     
    #Read key sheet and convert it to .csv and drop index
    key_data = pd.read_excel(key_sheet, index_col=None)
    key_data.to_csv(getFilename(key_sheet)+".csv")
    
    #Returning the read data
    return response_data,key_data

def PhysicsAnalysis(single_phy_resp,oorm_phy_resp,para_phy_resp,int_phy_resp,
                    single_phy_key,oorm_phy_key,para_phy_key,int_phy_key,unanswered_phy,studentName):
    
    oorm_phy_resp_list = []
    oorm_phy_key_list = []
    
    for i in range(len(oorm_phy_key)):
        oorm_phy_key_list.append(oorm_phy_key.iloc[i][1])
        oorm_phy_resp_list.append(oorm_phy_resp.iloc[i])
            
    phy_single_correct = 0
    for i in range(len(single_phy_key)):
        if single_phy_resp.iloc[i] == 'Nan':
            pass
        else:
            if single_phy_resp.iloc[i] == single_phy_key.iloc[i][1]:
                    phy_single_correct += 4
            else:
                    phy_single_correct = -1
    # print("Physics Single Result : ", phy_single_correct)
    
    
    phy_oorm_correct = 0
    for i in range(len(oorm_phy_key_list)):
        if oorm_phy_key_list[i] == 'Nan':
            pass
        else:
            if oorm_phy_resp_list[i] == oorm_phy_key_list[i]: 
                phy_oorm_correct += 4
            else: 
                phy_oorm_correct = -1
    # print("Physics One or More Result : ",phy_oorm_correct)
    
    phy_para_correct = 0
    for i in range(len(para_phy_key)):
        if phy_para_correct.iloc[i] == 'Nan':
            pass
        else:
            if para_phy_resp.iloc[i] == para_phy_key.iloc[i][1]:
                phy_para_correct += 4
            else:
                phy_para_correct -= 1
    # print("Physics Para Result : ",phy_para_correct)
    
    phy_int_correct = 0
    for i in range(len(para_phy_key)):
        if int_phy_resp.iloc[i] == 'Nan':
            pass
        else:
            if int_phy_resp.iloc[i] == int_phy_key.iloc[i][1]:
                    phy_int_correct += 4
            else:
                    phy_int_correct -= 1
    #print("Physics Integer Result : ",phy_int_correct)
    
    physics_total = phy_single_correct + phy_oorm_correct + phy_para_correct + phy_int_correct 
    return physics_total

def MathAnalysis(single_math_resp,oorm_math_resp,para_math_resp,int_math_resp,
                    single_math_key,oorm_math_key,para_math_key,int_math_key, unanswered_math
                    ,studentName):
    
    oorm_math_resp_list = []
    oorm_math_key_list = []
    
    for i in range(len(oorm_math_key)):
        oorm_math_key_list.append(oorm_math_key.iloc[i][1])
        oorm_math_resp_list.append(oorm_math_resp.iloc[i])
        
    math_single_correct = 0
    for i in range(len(single_math_key)):
        if single_math_resp.iloc[i] == 'Nan':
            pass
        else:
            if single_math_resp.iloc[i] == single_math_key.iloc[i][1]:
                    math_single_correct += 4
            else:
                    math_single_correct = -1
    # print("Math Single Result : ", math_single_correct)
    
    
    math_oorm_correct = 0
    for i in range(len(oorm_math_key_list)):
        if oorm_math_resp_list[i] == 'Nan':
            pass
        else:
            if oorm_math_resp_list[i] == oorm_math_key_list[i]: 
                math_oorm_correct += 4
            else : 
                math_oorm_correct = -1
    # print("Math One or More Result : ",math_oorm_correct)


    math_para_correct = 0
    for i in range(len(para_math_key)):
        if math_para_correct.iloc[i] == 'Nan':
            pass
        else:
            if para_math_resp.iloc[i] == para_math_key.iloc[i][1]:
                    math_para_correct += 4
            else:
                    math_para_correct -= 1
    # print("Math Para Result : ",math_para_correct)
    
    math_int_correct = 0
    for i in range(len(para_math_key)):
        if para_math_resp.iloc[i] == 'Nan':
            pass
        else:
            if int_math_resp.iloc[i] == int_math_key.iloc[i][1]:
                math_int_correct += 4
            else:
                math_int_correct -= 1
                
    math_total = math_single_correct + math_oorm_correct + math_para_correct + math_int_correct 
    return math_total

def ChemAnalysis(single_chem_resp,oorm_chem_resp,para_chem_resp,int_chem_resp,
                single_chem_key,oorm_chem_key,para_chem_key,int_chem_key, unanswered_chem
                ,studentName):
    oorm_chem_resp_list = []
    oorm_chem_key_list = []
    
    for i in range(len(oorm_chem_key)):
        oorm_chem_key_list.append(oorm_chem_key.iloc[i][1])
        oorm_chem_resp_list.append(oorm_chem_resp.iloc[i])
        
    chem_single_correct = 0
    for i in range(len(single_chem_key)):
        if single_chem_resp.iloc[i] == 'Nan':
            pass
        else:
            if single_chem_resp.iloc[i] == single_chem_key.iloc[i][1]:
                chem_single_correct += 4
            else:
                chem_single_correct = -1
    # print("chem Single Result : ", chem_single_correct)
    
    
    chem_oorm_correct = 0
    for i in range(len(oorm_chem_key_list)):
        if oorm_chem_resp_list[i] == 'Nan':
            pass
        else:
            if oorm_chem_resp_list[i] == oorm_chem_key_list[i]: 
                chem_oorm_correct += 4
            else : 
                chem_oorm_correct = -1
    # print("chem One or More Result : ",chem_oorm_correct)


    chem_para_correct = 0
    for i in range(len(para_chem_key)):
        if para_chem_resp.iloc[i] == 'Nan':
            pass
        else:
            if para_chem_resp.iloc[i] == para_chem_key.iloc[i][1]:
                chem_para_correct += 4
            else:
                chem_para_correct -= 1
    # print("chem Para Result : ",chem_para_correct)
    
    chem_int_correct = 0
    for i in range(len(para_chem_key)):
        if int_chem_resp.iloc[i] == 'Nan':
            pass
        else:
            if int_chem_resp.iloc[i] == int_chem_key.iloc[i][1]:
                chem_int_correct += 4
            else:
                chem_int_correct -= 1
    chem_total = chem_single_correct + chem_oorm_correct + chem_para_correct + chem_int_correct 
    return chem_total

def PhysicsAnalysis_sheet2(single_phy_resp,int_phy_resp,para_phy_resp,mat_phy_resp,
                    single_phy_key,int_phy_key,para_phy_key,mat_phy_key, unanswered_phy_sheet2
                    ,studentName):
    
    mat_phy_resp_list = []
    mat_phy_key_list = []
    
    
    for i in range(len(mat_phy_key)):
        mat_phy_key_list.append(mat_phy_key.iloc[i][1])
        mat_phy_resp_list.append(mat_phy_resp.iloc[i])
        
    phy_single_correct = 0 
    
    for i in range(len(single_phy_key)):
        if single_phy_resp.iloc[i] == 'Nan':
            pass
        else:
            if single_phy_resp.iloc[i] == single_phy_key.iloc[i][1]:
                phy_single_correct += 4
            else:
                phy_single_correct = -1
    # print("Physics Single Result : ", phy_single_correct)
        
    phy_int_correct = 0
    
    for i in range(len(int_phy_key)):
        if int_phy_resp.iloc[i] == 'Nan':
            pass
        else:
            if int_phy_resp.iloc[i] == int_phy_key.iloc[i][1]:
                phy_int_correct += 4
            else:
                phy_int_correct -= 1
            
    # print("Physics Int Result : ",phy_int_correct)
    
    phy_para_correct = 0
    
    for i in range(len(para_phy_resp)):
        if para_phy_resp.iloc[i] == 'Nan':
            pass
        else:
            if para_phy_resp.iloc[i] == para_phy_key.iloc[i][1]:
                phy_para_correct += 4
            else:
                phy_para_correct -= 1
            
    # print("Physics Para Result : ",phy_para_correct)
    
    phy_mat_correct = 0
    
    for i in range(len(mat_phy_key_list)):
        if mat_phy_resp_list[i] == 'Nan':
            pass
        else:
            if mat_phy_resp_list[i] == mat_phy_key_list[i]:
                phy_mat_correct += 4
            else:
                phy_mat_correct -= 1
    
    # print("Physics Matrix Result :", phy_mat_correct)
    

    physics_total = phy_single_correct + phy_int_correct + phy_para_correct + phy_mat_correct 
    return physics_total


def MathAnalysis_sheet2(single_math_resp,int_math_resp,para_math_resp,mat_math_resp,
                    single_math_key,int_math_key,para_math_key,mat_math_key, unanswered_math_sheet2
                    ,studentName):
    
    mat_math_resp_list = []
    mat_math_key_list = []
    
    for i in range(len(mat_math_key)):
        mat_math_key_list.append(mat_math_key.iloc[i][1])
        mat_math_resp_list.append(mat_math_resp.iloc[i])
        
    math_single_correct = 0
    for i in range(len(single_math_key)):
        if single_math_resp.iloc[i] == 'Nan':
            pass
        else:
            if single_math_resp.iloc[i] == single_math_key.iloc[i][1]:
                    math_single_correct += 4
            else:
                    math_single_correct = -1
    # print("Math Single Result : ", math_single_correct)
    
    
    math_mat_correct = 0
    for i in range(len(mat_math_key_list)):
        if mat_math_resp_list[i] == 'Nan':
            pass
        else:
            if mat_math_resp_list[i] == mat_math_key_list[i]: 
                math_mat_correct += 4
            else : 
                math_mat_correct = -1
    # print("Math One or More Result : ",math_oorm_correct)
    

    math_para_correct = 0
    for i in range(len(para_math_key)):
        if para_math_resp.iloc[i] == 'Nan':
            pass
        else:
            if para_math_resp.iloc[i] == para_math_key.iloc[i][1]:
                math_para_correct += 4
            else:
                math_para_correct -= 1
    # print("Math Para Result : ",math_para_correct)
    
    math_int_correct = 0
    for i in range(len(int_math_key)):
        if int_math_resp.iloc[i] == 'Nan':
            pass
        else:
            if int_math_resp.iloc[i] == int_math_key.iloc[i][1]:
                math_int_correct += 4
            else:
                math_int_correct -= 1
    math_total = math_single_correct + math_mat_correct + math_para_correct + math_int_correct    
    return math_total

def ChemAnalysis_sheet2(single_chem_resp,int_chem_resp,para_chem_resp,mat_chem_resp,
                single_chem_key,int_chem_key,para_chem_key,mat_chem_key, unanswered_chem_sheet2
                ,studentName):
    mat_chem_resp_list = []
    mat_chem_key_list = []
    
    for i in range(len(mat_chem_key)):
        mat_chem_key_list.append(mat_chem_key.iloc[i][1])
        mat_chem_resp_list.append(mat_chem_resp.iloc[i])
        
    chem_single_correct = 0
    for i in range(len(single_chem_key)):
        if single_chem_resp.iloc[i] == 'Nan':
            pass
        else:
            if single_chem_resp.iloc[i] == single_chem_key.iloc[i][1]:
                    chem_single_correct += 4
            else:
                    chem_single_correct = -1
    # print("chem Single Result : ", chem_single_correct)
    
    
    chem_mat_correct = 0
    for i in range(len(mat_chem_key_list)):
        if mat_chem_resp_list[i] == 'Nan':
            pass
        else:
            if mat_chem_resp_list[i] == mat_chem_key_list[i]: 
                chem_mat_correct += 4
            else : 
                chem_mat_correct = -1
    # print("chem One or More Result : ",chem_oorm_correct)
    
    
    chem_para_correct = 0
    for i in range(len(para_chem_key)):
        if para_chem_resp.iloc[i] == 'Nan':
            pass
        else:
            if para_chem_resp.iloc[i] == para_chem_key.iloc[i][1]:
                    chem_para_correct += 4
            else:
                    chem_para_correct -= 1
    # print("chem Para Result : ",chem_para_correct)
    
    chem_int_correct = 0
    for i in range(len(int_chem_key)):
        if int_chem_resp.iloc[i] == 'Nan':
            pass
        else:
            if int_chem_resp.iloc[i] == int_chem_key.iloc[i][1]:
                chem_int_correct += 4
            else:
                chem_int_correct -= 1
    chem_total = chem_single_correct + chem_mat_correct + chem_para_correct + chem_int_correct
    return chem_total


def QuestionsAnalysis_sheet2(studentRowData,key_data,unanswered_phy, unanswered_chem, unanswered_math, total_unanswered):
    # print("_"*20)
    # for i in range(len(studentRowData)):
    #     print(studentRowData[i])   
    
    studentRowData.to_frame()
    
    
    single_phy_resp, int_phy_resp, para_phy_resp, mat_phy_resp  = studentRowData.iloc[4:10], studentRowData.iloc[10:15], studentRowData.iloc[15:20],studentRowData.iloc[20:29]      
    single_chem_resp, int_chem_resp, para_chem_resp, mat_chem_resp = studentRowData.iloc[29:35], studentRowData.iloc[35:40], studentRowData.iloc[40:46],studentRowData.iloc[46:54]
    single_math_resp, int_math_resp, para_math_resp, mat_math_resp = studentRowData.iloc[54:60], studentRowData.iloc[60:65], studentRowData.iloc[65:71],studentRowData.iloc[71:79]      
         
    single_phy_key, int_phy_key, para_phy_key, mat_phy_key  = key_data.iloc[0:6], key_data.iloc[6:11], key_data.iloc[11:17], key_data.iloc[17:25]
    single_chem_key, int_chem_key, para_chem_key, mat_chem_key = key_data.iloc[25:31], key_data.iloc[31:36], key_data.iloc[36:42], key_data.iloc[42:50]
    single_math_key, int_math_key, para_math_key, mat_math_key = key_data.iloc[50:56], key_data.iloc[56:61], key_data.iloc[61:67], key_data.iloc[67:75]
        
    studentName = studentRowData.iloc[3]
    
    physics_total = PhysicsAnalysis_sheet2(single_phy_resp,int_phy_resp,para_phy_resp,mat_phy_resp,
                    single_phy_key,int_phy_key,para_phy_key,mat_phy_key,unanswered_phy,studentName)
    # print("Total chem marks scored by "+ studentName +" : ", physics_total)
    # print("_"*50)    
    
    math_total = MathAnalysis_sheet2(single_math_resp,int_math_resp,para_math_resp,mat_math_resp,
                    single_math_key,int_math_key,para_math_key,mat_math_key,unanswered_math,studentName)
    # print("Total Math marks scored by "+ studentName +" : ", math_total)
    # print("_"*50)    

    chem_total = ChemAnalysis_sheet2(single_chem_resp,int_chem_resp,para_chem_resp,mat_chem_resp,
                single_chem_key,int_chem_key,para_chem_key,mat_chem_key,unanswered_chem,studentName)
    # print("Total chem marks scored by "+ studentName +" : ", chem_total)
    # print("_"*50)    
    
    total = physics_total+math_total+chem_total    
    
    return studentName, physics_total, chem_total, math_total, total 


def QuestionsAnalysis_sheet1(studentRowData,key_data,unanswered_phy, unanswered_chem, unanswered_math, total_unanswered):
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
                    single_phy_key,oorm_phy_key,para_phy_key,int_phy_key, unanswered_phy
                    ,studentName)
    # print("Total chem marks scored by "+ studentName +" : ", physics_total)
    # print("_"*50)    
    
    math_total = MathAnalysis(single_math_resp,oorm_math_resp,para_math_resp,int_math_resp,
                    single_math_key,oorm_math_key,para_math_key,int_math_key, unanswered_math
                    ,studentName)
    # print("Total Math marks scored by "+ studentName +" : ", math_total)
    # print("_"*50)    

    chem_total = ChemAnalysis(single_chem_resp,oorm_chem_resp,para_chem_resp,int_chem_resp,
                single_chem_key,oorm_chem_key,para_chem_key,int_chem_key, unanswered_chem
                ,studentName)
    # print("Total chem marks scored by "+ studentName +" : ", chem_total)
    # print("_"*50)    
    
    total = physics_total+math_total+chem_total
        
    return studentName, physics_total, chem_total, math_total, total 

def StudentNullAnalysis_sheet1(studentRowData):
    physics_slice = studentRowData.iloc[4:32]
    unanswered_phy = physics_slice.isna().sum()
    # print("Number of Unanswered Questions in Physics by "+studentRowData[2]+" is :", unanswered_phy)
    chemistry_slice = studentRowData.iloc[32:60]
    unanswered_chem = chemistry_slice.isna().sum()
    # print("Number of Unanswered Questions in Chemistry by "+studentRowData[2]+" is :", unanswered_chem)    
    mathematics_slice = studentRowData.iloc[60:90]
    unanswered_math = mathematics_slice.isna().sum()
    # print("Number of Unanswered Questions in Mathematics by "+studentRowData[2]+" is :", unanswered_math)    
    # print("_"*50)
    total_unanswered = studentRowData[4:90].isna().sum()
    # print("Total number of unanswered questions by "+ studentRowData[2]+" are ", total_unanswered)    
    return unanswered_phy, unanswered_chem, unanswered_math, total_unanswered

def StudentNullAnalysis_sheet2(studentRowData):
    physics_slice = studentRowData.iloc[4:29]
    chemistry_slice =  studentRowData.iloc[29:54]
    mathematics_slice = studentRowData.iloc[54:79]
    unanswered_phy_sheet2 = physics_slice.isna().sum()
    unanswered_chem_sheet2 = chemistry_slice.isna().sum()
    unanswered_math_sheet2 = mathematics_slice.isna().sum()
    total_unanswered_sheet2 = studentRowData[4:79].isna().sum()
    
    return unanswered_phy_sheet2,unanswered_chem_sheet2, unanswered_math_sheet2, total_unanswered_sheet2
    

def main():
    response_sheet_1 = "data/sheet_1_resp.xlsx"
    key_sheet_1 = "data/sheet_1_key.xlsx"
    
    response_sheet_2 = "data/sheet_2_resp.xlsx"
    key_sheet_2 = "data/sheet_2_key.xlsx"
    
    response_data_1, key_data_1 = readData(response_sheet_1, key_sheet_1)
    response_data_2, key_data_2 = readData(response_sheet_2, key_sheet_2)
    
    
    response_data_1 = response_data_1.sort_values('OMR No.',ascending=True ,na_position='first')
    response_data_1 = response_data_1.drop('Timestamp',axis=1)
    
    
    response_data_2 = response_data_2.sort_values('OMR No.',ascending=True, na_position='first')
    response_data_2 = response_data_2.drop(['Timestamp','NAME OF THE STUDENT'],axis=1)

    
    number_colums_sheet_1 = len(response_data_1)
    
    number_colums_sheet_2 = len(response_data_2)
    
    # print(type(key_data_1))
    # print(type(key_data_2))
    
    
    # #TEST COL Sheet 2 
    
    sheet2_data = response_data_2.iloc[0]
    unanswered_phy_sheet2, unanswered_chem_sheet2, unanswered_math_sheet2, total_unanswered_sheet2 = StudentNullAnalysis_sheet2(sheet2_data)
    QuestionsAnalysis_sheet2(sheet2_data, key_data_2, unanswered_phy_sheet2, unanswered_chem_sheet2, unanswered_math_sheet2, total_unanswered_sheet2)
    
    
    #TEST Col            
    rowData = response_data_1.iloc[1]
    unanswered_phy, unanswered_chem, unanswered_math, total_unanswered = StudentNullAnalysis_sheet1(rowData)           
    studentName, physics_total, chem_total, math_total, total = QuestionsAnalysis_sheet1(rowData, key_data_1,unanswered_phy, unanswered_chem, unanswered_math, total_unanswered)    
    # print(studentName)
    
    result_sheet_1 = []
    result_sheet_2 = []
    for i in range(number_colums_sheet_1):
        rowData = response_data_1.iloc[i]
        unanswered_phy, unanswered_chem, unanswered_math, total_unanswered = StudentNullAnalysis_sheet1(rowData)
        studentName, physics_total, chem_total, math_total, total = QuestionsAnalysis_sheet1(rowData, key_data_1,unanswered_phy, unanswered_chem, unanswered_math, total_unanswered)
        result_sheet_1.append([studentName, physics_total, chem_total, math_total, total, unanswered_phy, unanswered_chem, unanswered_math, total_unanswered])    
        resultdf = pd.DataFrame(result_sheet_1, columns=['Student Name', 'Physics Total', 'Chemistry Total', 'Mathematics Total', 'Total', 'Unanswered Physics', 'Unanswered Chemistry', 'Unanswered Math', 'Total Unanswered Questions'])  
        
    for i in range(number_colums_sheet_2):
        sheet2_data = response_data_2.iloc[i]
        unanswered_phy_sheet2, unanswered_chem_sheet2, unanswered_math_sheet2, total_unanswered_sheet2 = StudentNullAnalysis_sheet2(sheet2_data)
        studentName_2, physics_total_2, chem_total_2, math_total_2, total_2 = QuestionsAnalysis_sheet2(sheet2_data, key_data_2, unanswered_phy_sheet2, unanswered_chem_sheet2, unanswered_math_sheet2, total_unanswered_sheet2)
        result_sheet_2.append([studentName_2, physics_total_2, chem_total_2, math_total_2, total_2, unanswered_phy_sheet2, unanswered_chem_sheet2, unanswered_math_sheet2, total_unanswered_sheet2])    
        resultdf_2 = pd.DataFrame(result_sheet_2, columns=['Student Name', 'Physics Total', 'Chemistry Total', 'Mathematics Total', 'Total', 'Unanswered Physics', 'Unanswered Chemistry', 'Unanswered Math', 'Total Unanswered Questions'])  
        
    print(resultdf)
    print(resultdf_2)
    
    resultdf.to_csv(index=False) 
    resultdf.to_excel("Result_sheet1.xlsx", sheet_name='Result_sheet1', engine='openpyxl', index=False)
    
    resultdf_2.to_csv(index=False)
    resultdf_2.to_excel("Result_sheet2.xlsx", sheet_name='Result_sheet2', engine='openpyxl', index=False)
    
    
    excel_writer = StyleFrame.ExcelWriter('Result_sheet1.xlsx')
    sf = StyleFrame(resultdf)
    styler = Styler(font_size=15)
    sf.apply_headers_style(styler,style_index_header=True)
    sf.set_column_width('Student Name', 40)
    sf.set_column_width(['Physics Total', 'Chemistry Total', 'Mathematics Total','Total'], 20)
    sf.set_column_width(['Unanswered Physics','Unanswered Chemistry', 'Unanswered Math' ,'Total Unanswered Questions'], 25)
    sf.to_excel(excel_writer=excel_writer)
    excel_writer.save()
    
    excel_writer = StyleFrame.ExcelWriter('Result_sheet2.xlsx')
    sf = StyleFrame(resultdf_2)
    styler = Styler(font_size=15)
    sf.apply_headers_style(styler,style_index_header=True)
    sf.set_column_width('Student Name', 40)
    sf.set_column_width(['Physics Total', 'Chemistry Total', 'Mathematics Total','Total'], 20)
    sf.set_column_width(['Unanswered Physics','Unanswered Chemistry', 'Unanswered Math' ,'Total Unanswered Questions'], 25)
    sf.to_excel(excel_writer=excel_writer)
    excel_writer.save()
    
    
if __name__ == "__main__":
    main()