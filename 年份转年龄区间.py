# -*- coding: utf-8 -*-
"""
Created on Mon Dec 27 15:14:43 2021

@author: RF.Backspace
@email: rf_backspace@Outlook.com
"""
import datetime
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog

file_type_check = str(input("请输入文件格式（Excel/CSV): "))

if str(file_type_check) == "Excel" or str(file_type_check) == "excel":
    print ("1",file_type_check)
    file_name1 = filedialog.askopenfilename()
    origin_dataframe = pd.read_excel(file_name1).fillna(0)
if str(file_type_check) == 'CSV' or str(file_type_check) == 'csv':
    print ('2',file_type_check)
    file_name1 = filedialog.askopenfilename()
    origin_dataframe = pd.read_csv(file_name1).fillna(0)
#修改值 - 设置区间段节点/设置区间段标注
age_bins = [-1,20,25,30,35,40,45,50,55,60,150]
age_labels = ['小于20岁','21岁-25岁', '26岁-30岁', '31岁-35岁', '36岁-40岁','41岁-45岁','46岁-50岁','51岁-55岁','56岁-60岁','大于60岁']
age_list = []

#获取当前年份（无需修改）
year = datetime.date.today().year

for age_index in origin_dataframe['year']:
    age = int(year) - int(age_index)
    age_list.append(age)

origin_dataframe['age'] = age_list

age_group = pd.cut(x=origin_dataframe['age'], bins=age_bins,labels=age_labels,right=True)
origin_dataframe['age_range'] = age_group

pd1 = origin_dataframe

writer = pd.ExcelWriter(r'增加年龄区间.xls')

#设置Sheet名称
pd1.to_excel(writer,sheet_name='sheet1')

writer.save()
writer.close()
