# -*- coding: utf-8 -*-
"""
Created on Mon Nov  8 13:48:50 2021

@author: RF.Backspace
@email: rf_backspace@Outlook.com
"""
import datetime
import pandas as pd
import numpy as np
import sas.tabulate.pv as pv
import os

#------------------------------------------------------------------------------
#修改值 - 文件名/文件路径 （示例：“C:\\Users\\xiaobin.ma\\Desktop\\10.28 - xx.xx\\test_excel.xlsx”）
file_name = "C:\\Users\\xiaobin.ma\\Downloads\\副驾体验问卷调研code(1).xlsx"
index = '提交编号'
columns = ['A7','A8','A10','A11','A12','A13','A15','A19','A20','A23','A24','A27','A29','A31','A35','A32','A33']
rows = ["car_type",'gender','age_range']
#修改值 - 设置区间段节点/设置区间段标注
age_bins = [-1,20,25,30,35,40,45,100]
age_labels = ['小于20岁','21岁-25岁', '26岁-30岁', '31岁-35岁', '36岁-40岁','41岁-45岁','大于45岁']
age_list = []

#获取当前年份（无需修改）
year = datetime.date.today().year

#可修改值 - 若使用csv使用pd.read_csv,若为excel使用pd.read_excel
origin_dataframe = pd.read_excel(file_name)
#origin_dataframe = pd.read_csv(file_name)

include_columns_name=[]
without_columns_name=[]

source={}
#-----------------文件处理-增加年龄/按照年龄划分年龄段/更新性别显示----------------
for age_index in origin_dataframe['year']:
    age = int(year) - int(age_index)
    age_list.append(age)

origin_dataframe['age'] = age_list

age_group = pd.cut(x=origin_dataframe['age'], bins=age_bins,labels=age_labels,right=True)
origin_dataframe['age_range'] = age_group

'''
#-----------------------------------性别转换------------------------------------
origin_dataframe.gender[origin_dataframe['gender'] ==1] = '男'
origin_dataframe.gender[origin_dataframe['gender'] ==2] = '女'
'''
include_not_following_dataframe = origin_dataframe.reindex()
without_not_following_dataframe = origin_dataframe.reindex()
#------------------------------------------------------------------------------

for columns_index in columns:

    tem_include = include_not_following_dataframe[columns_index].dropna(axis = 0)
    tem_without = without_not_following_dataframe[columns_index].dropna(axis = 0)

    for columns_values in tem_include:
        
        if columns_values not in range (1,6):
            if str(columns_values) == "非常满意":
                tem_include.replace(columns_values,int(5),inplace=True)
                tem_without.replace(columns_values,int(5),inplace=True)
            if str(columns_values) == "满意":
                tem_include.replace(columns_values,int(4),inplace=True)
                tem_without.replace(columns_values,int(4),inplace=True)
            if str(columns_values) == "一般":
                tem_include.replace(columns_values,int(3),inplace=True)
                tem_without.replace(columns_values,int(3),inplace=True)
            if str(columns_values) == "不满意":
                tem_include.replace(columns_values,int(2),inplace=True)
                tem_without.replace(columns_values,int(2),inplace=True)
            if str(columns_values) == "非常不满意":
                tem_include.replace(columns_values,int(1),inplace=True)
                tem_without.replace(columns_values,int(1),inplace=True)
            else:
                tem_include.replace(columns_values,int(0),inplace=True)
                tem_without.replace(columns_values,np.nan,inplace = True)
    include_not_following_dataframe[columns_index] = tem_include
    without_not_following_dataframe[columns_index] = tem_without
#------------------------------------------------------------------------------
include_crosstab_count = pd.DataFrame()
with_tem_include_crosstab_count = pd.DataFrame()

include_crosstab_normalize = pd.DataFrame()
with_tem_include_crosstab_normalize = pd.DataFrame()

include_crosstab_percentage = pd.DataFrame()
with_tem_include_crosstab_percentage = pd.DataFrame()

def crosstab_include(index, columns,rows):

    global with_tem_include_crosstab_count
    global with_tem_include_crosstab_normalize
    global with_tem_include_crosstab_percentage

    part_count = pv.enhance_pivot_table(include_not_following_dataframe,row=[rows],columns=[(columns)],aggfunc={index:len},agg_axis=all,column_margin=True,row_margin=True).fillna(0)
    with_tem_include_crosstab_count = pd.concat([with_tem_include_crosstab_count,part_count],axis=1)

    part_normalize = pv.enhance_pivot_table(include_not_following_dataframe,row=[rows],columns=[(columns)],aggfunc={index:len},agg_axis=all,column_margin=True,row_margin=True,normalize='column')
    with_tem_include_crosstab_normalize = pd.concat([with_tem_include_crosstab_normalize,part_normalize],axis=1).fillna(0)
    

    part_percentage = with_tem_include_crosstab_normalize.applymap(lambda x: '%.2f%%' % (x*100))
    with_tem_include_crosstab_percentage = part_percentage
#------------------------------------------------------------------------------
without_crosstab_count = pd.DataFrame()
without_tem_include_crosstab_count = pd.DataFrame()

without_crosstab_normalize = pd.DataFrame()
without_tem_include_crosstab_normalize = pd.DataFrame()

without_crosstab_percentage = pd.DataFrame()
without_tem_include_crosstab_percentage = pd.DataFrame()

def crosstab_without(index, columns,rows):
    
    global without_tem_include_crosstab_count
    global without_tem_include_crosstab_normalize
    global without_tem_include_crosstab_percentage
    
    part_count = pv.enhance_pivot_table(without_not_following_dataframe,row=[rows],columns=[(columns)],aggfunc={index:len},agg_axis=all,column_margin=True,row_margin=True).fillna(0)
    without_tem_include_crosstab_count = pd.concat([without_tem_include_crosstab_count,part_count],axis=1)

    part_normalize = pv.enhance_pivot_table(without_not_following_dataframe,row=[rows],columns=[(columns)],aggfunc={index:len},agg_axis=all,column_margin=True,row_margin=True,normalize='column')
    without_tem_include_crosstab_normalize = pd.concat([without_tem_include_crosstab_normalize,part_normalize],axis=1).fillna(0)

    part_percentage = without_tem_include_crosstab_normalize.applymap(lambda x: '%.2f%%' % (x*100))
    without_tem_include_crosstab_percentage =part_percentage
#------------------------------------------------------------------------------
for columns_index in columns:
    with_tem_include_crosstab_count =  pd.DataFrame()
    with_tem_include_crosstab_normalize =  pd.DataFrame()
    with_tem_include_crosstab_percentage =  pd.DataFrame()
    
    without_tem_include_crosstab_count =  pd.DataFrame()
    without_tem_include_crosstab_normalize =  pd.DataFrame()
    without_tem_include_crosstab_percentage =  pd.DataFrame()
    
    for rows_index in rows:
        crosstab_include(index,rows_index,columns_index)
        crosstab_without(index,rows_index,columns_index)

    add_include_columns_name_run_time = 0
    while add_include_columns_name_run_time <int(len(with_tem_include_crosstab_count)):
        include_columns_name.insert(0, str(columns_index))
        add_include_columns_name_run_time = int(add_include_columns_name_run_time) + 1
    
    add_without_columns_name_run_time = 0
    while add_without_columns_name_run_time <int(len(without_tem_include_crosstab_count)):
        without_columns_name.insert(0, str(columns_index))
        add_without_columns_name_run_time = int(add_without_columns_name_run_time) + 1

    include_crosstab_count = pd.concat([with_tem_include_crosstab_count,include_crosstab_count],axis=0).fillna(0)
    include_crosstab_normalize = pd.concat([with_tem_include_crosstab_normalize,include_crosstab_normalize],axis=0).fillna(0)
    include_crosstab_percentage = pd.concat([with_tem_include_crosstab_percentage,include_crosstab_percentage],axis=0).fillna(0)

    without_crosstab_count = pd.concat([without_tem_include_crosstab_count,without_crosstab_count],axis=0).fillna(0)
    without_crosstab_normalize = pd.concat([without_tem_include_crosstab_normalize,without_crosstab_normalize],axis=0).fillna(0)
    without_crosstab_percentage = pd.concat([without_tem_include_crosstab_percentage,without_crosstab_percentage],axis=0).fillna(0)

include_crosstab_count.insert(0,'columns_name',include_columns_name,allow_duplicates=False)
include_crosstab_normalize.insert(0,'columns_name',include_columns_name,allow_duplicates=False)
include_crosstab_percentage.insert(0,'columns_name',include_columns_name,allow_duplicates=False)

without_crosstab_count.insert(0,'columns_name',without_columns_name,allow_duplicates=False)
without_crosstab_normalize.insert(0,'columns_name',without_columns_name,allow_duplicates=False)
without_crosstab_percentage.insert(0,'columns_name',without_columns_name,allow_duplicates=False)

#------------------------------------分数计算-----------------------------------
'''
for source_index in source_columns:
    tem_sourcedf = without_not_following_dataframe[source_index]

    tem_sourcedf = pd.DataFrame(tem_sourcedf)
    tem_sourcedf = tem_sourcedf.dropna(axis=0,how='all')
    
    module_source = tem_sourcedf .value_counts(normalize = True).sort_index(ascending=True)
    
    final_source = 0
    
    if len(module_source) == 5:
        value = 1
    elif len(module_source) == 6:
        value = 0
        
    for multiplier_for_source in module_source:
        final_source = final_source + multiplier_for_source * value
        final_source_tem = "%.1f" % final_source
        value = value + 1
    source[source_index] = final_source_tem
    print (source_index,final_source)
sourcedb = pd.DataFrame(source,index = ['得分'])
crosstab_source = sourcedb.sort_values(by='得分',ascending=False,axis =1)
'''
#----------------------------------存储为Excel----------------------------------
#设置Excel文件名
writer = pd.ExcelWriter(r'modified data.xls')

#设置Sheet名称
include_crosstab_count.to_excel(writer,sheet_name='Crosstab - Include 0')
include_crosstab_normalize.to_excel(writer,sheet_name='Crosstab_Normalize - Include 0')
include_crosstab_percentage.to_excel(writer,sheet_name='Crosstab_Percentage - Include 0')
without_crosstab_count.to_excel(writer,sheet_name='Crosstab - without 0')
without_crosstab_normalize.to_excel(writer,sheet_name='Crosstab_Normalize - without 0')
without_crosstab_percentage.to_excel(writer,sheet_name='Crosstab_Percentage - without 0')
#crosstab_source.to_excel(writer,sheet_name='Crosstab_Source')
writer.save()
writer.close()
