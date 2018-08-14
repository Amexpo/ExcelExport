# -*- coding: utf-8 -*-
"""
Created on Tue Aug 14 10:59:00 2018

@author: yangyanhao
"""

import os
import time

'''
path_set=[]
name_set=[]
path='2018年\\'
list_dir=os.listdir(path)
path1=path+list_dir[0]
list_dir1=os.listdir(path1)
name=list_dir1[0][9:]
path2=path+list_dir[0]+'\\'+list_dir1[0]
list_dir2=os.listdir(path2)
path3=path+list_dir[0]+'\\'+list_dir1[0]+'\\'+list_dir2[0]
list_dir3=os.listdir(path3)
path4=path+list_dir[0]+'\\'+list_dir1[0]+'\\'+list_dir2[0]+'\\'+list_dir3[0]
path_set.append(path4)
'''

import os
import time
import xlwt
import xlrd
import win32com.client
import requests

def GetTime():
    return time.strftime("%Y-%m-%d",time.localtime(time.time()))

name='企业名称'

def xls_write(file_dir,name):
    workbook = xlrd.open_workbook(file_dir)
    sheet1 = workbook.sheet_by_index(1)
    write_file_name =GetTime()+'符合性确认数据.xls'
    workbook = xlwt.Workbook(write_file_name)
    worksheet = workbook.add_sheet('统计数据')
    for i in range(3,sheet1.nrows):
        row = sheet1.row_values(i)
        for j in range(7):
            worksheet.write(i-3, j, label = row[j])
        worksheet.write(i-3, 7, label = name)
    workbook.save(write_file_name)
    return()

xls_write('201808001-1深圳华润九新药业有限公司.xls',name)