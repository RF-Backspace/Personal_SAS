# -*- coding: utf-8 -*-
"""
Created on Mon Dec 13 14:31:15 2021

@author: RF.Backspace
@email: rf_backspace@Outlook.com
"""
import pandas as pd
import tkinter as tk
from tkinter import filedialog

run_time = 1

fdf1 = pd.DataFrame()
fdf2 = pd.DataFrame()
fdf3 = pd.DataFrame()

file_type_check = str(input("请输入文件格式（Excel/CSV): "))

if str(file_type_check) == "Excel" or str(file_type_check) == "excel":
    print ("1",file_type_check)
    file_name1 = filedialog.askopenfilename()
    df = pd.read_excel(file_name1).fillna(0)
if str(file_type_check) == 'CSV' or str(file_type_check) == 'csv':
    print ('2',file_type_check)
    file_name1 = filedialog.askopenfilename()
    df = pd.read_csv(file_name1).fillna(0)

while run_time != 0:
    run_time = int(input("是否执行多选统计？(0.不执行，1.执行) "))

    if run_time != 0:
        run_time = 1
        columns_name = str(input('请输入多选题列名：'))

        temdf = df[columns_name].str.split(',', expand=True).fillna(0)
        temdf2 = df[columns_name].dropna()
        dic1 = {}
        dic2 = {}

        for i1 in temdf:
            for i2 in temdf[i1]:
                if i2 != 0:
                    if str(i2) not in dic1.keys():
                        dic1[str(i2)] = 0
                    if str(i2) in dic1.keys():
                        dic1[str(i2)] = int(dic1[str(i2)]) + 1
        for m in dic1.keys():
            dic2[m] = int(dic1[m])/temdf2.count()

        tem_range = str(columns_name+'的数量')
        dic1[tem_range] = temdf2.count()
        print(temdf2.count())
        print(len(df[columns_name]))

        df2 = pd.DataFrame(pd.Series(dic1), columns=['数量'])
        df3 = pd.DataFrame(pd.Series(dic2), columns=['比例'])
        df4 = df3.applymap(lambda x: '%.2f%%' % (x*100))
        series = pd.Series({"比例": temdf2.count()}, name=tem_range)
        df4 = df4.append(series)

        fdf1 = pd.concat([fdf1, df2], axis=0)
        fdf2 = pd.concat([fdf2, df3], axis=0)
        fdf3 = pd.concat([fdf3, df4], axis=0)
        print(df3)

    if run_time == 0:

        fidf = pd.concat([fdf1, fdf3], axis=1)
        writer = pd.ExcelWriter(r'多选题.xls')

        # 设置Sheet名称
        fidf.to_excel(writer,sheet_name = 'sheet1')
        '''
        fdf1.to_excel(writer, sheet_name='计数')
        fdf2.to_excel(writer, sheet_name='标准化')
        fdf3.to_excel(writer, sheet_name='百分比')
'''
        writer.save()
        writer.close()

        print("感谢使用，再见")
        run_time = 0
