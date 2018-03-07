#!/usr/bin/env python
# coding=utf-8
import pandas as pd #导入处理数据的第三方模块
import numpy as np
import xlrd 
import pymysql
pymysql.install_as_MySQLdb()
import MySQLdb
import easygui as g
import sys

#创建文件交互界面模块
g.msgbox("欢迎使用BOL香料数据导入工具")
g.msgbox("请选择需要导入香料数据的Excel表")
str1 = g.fileopenbox('选择文件','提示','C:/User/Administrator/Desktop/__pycache__')

#读入excel文件模块
ObjExcel = pd.read_excel(str1)
ObjExcel.head()
Excel = xlrd.open_workbook(str1)
sheet = Excel.sheet_by_name("source")

#检查数据缺失模块
CheckResult = str(any(ObjExcel.isnull()))#检查总的香料数集是否齐全
if CheckResult == 'False':#数据齐备就能导入数据库
    print("数据没缺失，可以写入数据库")

else:
    #检查每一列是否有缺失
    IsNull = []
    for col in ObjExcel.columns:
        IsNull.append(any(pd.isnull(ObjExcel[col])))#检查每一列是否缺失，将结果返回到is_null的列表中
    #NoCol=1
    #for ColResult in IsNull :   #遍历列表中的结果，然后输出缺失的列数
        #ColResult = str(ColResult)   
        #if ColResult  == 'True': #检查到该列为有缺失的数据，进行行数的检测
            #ColExcel  =  ObjExcel.iloc[:,[NoCol-1]] #取该列进行行检测
            #is_null = []
            #NoIndex = 1
            #for index in list(ColExcel.index):
                #is_null.append(any(pd.isnull(ColExcel.iloc[index,:])))
            #for IndexResult  in is_null:
                #IndexResult =  str(IndexResult)
                #if IndexResult == 'True':
                    #print("第%d行第%d列数据缺失请检查后添加"%(NoIndex,NoCol))
                #NoIndex = NoIndex + 1
           # NoIndex = 1
        #NoCol = NoCol + 1 
    ColResult_1 = str(IsNull[0])#提取香料名称检查结果
    ColResult_6 = str(IsNull[5])#提取SMLES检查结果
    ColResult_7 = str(IsNull[6])#提取分子式检查结果
    if ColResult_1 == 'False':
        if ColResult_6 == 'False':
            if ColResult_7 == 'False':
                g.msgbox("主要数据没有缺失可导入数据库")
                g.msgbox("正在导入数据库")
                try:
                    database = MySQLdb.connect (host="192.168.111.190",user = "root",passwd="YES",db = "Prefume",use_unicode=True,port = 3306,charset="utf8")#建立数据库链接，配置这里即可
                    cursor = database.cursor()
                    cursor.execute("set names utf8")
                    query = """INSERT INTO prefumes(name,english_name,fema_number,cas_number,SMILES,molecular_formula,molecular_weight,
                    molecular_size,essential_aroma,basic_flavor,flavor_descriptors,smell,boiling_point,threshold,aroma_relative_value,main_aroma,
                    natura,functional_group,toxicity)VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                    for r in range(1,sheet.nrows):
                        name = sheet.cell(r,0).value
                        english_name = sheet.cell(r,1).value
                        fema_number = sheet.cell(r,2).value
                        cas_number = sheet.cell(r,3).value
                        SMILES = sheet.cell(r,5).value
                        molecular_formula = sheet.cell(r,6).value
                        molecular_weight = sheet.cell(r,7).value
                        molecular_size = sheet.cell(r,8).value
                        essential_aroma = sheet.cell(r,9).value
                        basic_flavor = sheet.cell(r,10).value
                        flavor_descriptors = sheet.cell(r,11).value
                        smell = sheet.cell(r,12).value
                        boiling_point = sheet.cell(r,13).value
                        threshold = sheet.cell(r,14).value
                        aroma_relative_value = sheet.cell(r,15).value
                        main_aroma = sheet.cell(r,16).value
                        natura = sheet.cell(r,17).value
                        functional_group = sheet.cell(r,18).value
                        toxicity = sheet.cell(r,19).value

                        values =(name,english_name,fema_number,cas_number,SMILES,molecular_formula,molecular_weight,
                        molecular_size,essential_aroma,basic_flavor,flavor_descriptors,smell,boiling_point,threshold,aroma_relative_value,main_aroma,
                        natura,functional_group,toxicity)

                        cursor.execute(query,values)
                    cursor.close()
                    database.commit()
                    database.close()
                    g.msgbox("导入数据成功!")
                except:
                    g.msgbox("数据库写入不成功，请联系数据库管理员")
            else: #检查分子式那些数据缺失
                Col_6 = ObjExcel.iloc[:,[6]]
                Line_6IsNull = []
                for index in list(Col_6.index):
                    Line_6IsNull.append(any(pd.isnull(Col_6.iloc[index,:])))
                LineNo = 1
                for LineResult in Line_6IsNull:
                    LineResult = str(LineResult)
                    if LineResult == 'True':
                        g.msgbox("分子式第%d个缺失，请检查后重新录入数据库"%LineNo)
                    LineNo = LineNo + 1
        else: #检查SIMELS数据那些缺失
            Col_5 = ObjExcel.iloc[:,[5]]
            Line_5IsNull = []
            for index in list(Col_5.index):
                Line_5IsNull.append(any(pd.isnull(Col_5.iloc[index,:])))
            LineNo = 1
            for LineResult in Line_5IsNull:
                LineResult = str(LineResult)
                if LineResult == 'True':
                    g.msgbox("SMILES第%d个缺失，请检查后重新录入数据库"%LineNo)
                LineNo = LineNo + 1
    else:#检查香料名称那些数据缺失
        ColExcel = ObjExcel.iloc[:,[0]]
        LineIsNull = []
        for index in list(ColExcel.index):
            LineIsNull.append(any(pd.isnull(ColExcel.iloc[index,:])))
        LineNo = 1
        for LineResult in LineIsNull:
            LineResult = str(LineResult)
            if LineResult == 'True':
                g.msgbox("香料名称第%d个缺失，请检查后重新录入数据库"%LineNo)
            LineNo = LineNo + 1
