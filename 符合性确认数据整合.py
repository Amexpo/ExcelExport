# -*- coding: utf-8 -*-
"""
Created on Tue Aug 14 08:56:47 2018

@author: yangyanhao
"""

import os
import time
import xlwt
import win32com.client
import requests

def time_compare(file_dir):
    ft=time.gmtime(os.stat(file_dir).st_mtime)
    nt=time.localtime(time.time())
    x=((ft[0]*12+ft[1]*30+ft[2])>
       (nt[0]*12+nt[1]*30+nt[2]-90)
       )
    return(x)

def xls_write(file_dir,name,worksheet,k):
    workbook = xlrd.open_workbook(file_dir)
    sheet1 = workbook.sheet_by_index(1)
    for i in range(3,sheet1.nrows):
        row = sheet1.row_values(i)
        for j in range(7):
            worksheet.write(i+k, j, label = row[j])
        worksheet.write(i+k, 7, label = name)
    workbook.save(write_file_name)
    return()

def write_file():
    return()

write_file_name =GetTime()+'.xls'
workbook = xlwt.Workbook(write_file_name)
worksheet = workbook.add_sheet('服务性检测')


path='2018年\\'
def byyyh(path):
    path_set=[]
    name_set=[]
    list_dir=os.listdir(path)
    for i in range(len(list_dir)):
        path1=path+list_dir[i]
        list_dir1=os.listdir(path1)
        for j in range(len(list_dir1)):
            name=list_dir1[j][9:]
            name_set.append(name)
            path2=path+list_dir[i]+'\\'+list_dir1[j]
            list_dir2=os.listdir(path2)
            for k in range(len(list_dir2)-1):
                path3=path+list_dir[i]+'\\'+list_dir1[j]+'\\'+list_dir2[k]
                list_dir3=os.listdir(path3)
                for l in range(len(list_dir3)-1):
                    path4=path+list_dir[i]+'\\'+list_dir1[j]+'\\'+list_dir2[k]+'\\'+list_dir3[l]
                    path_set.append(path4)
    return(path_set,name_set)
    
write_file_name =time.strftime("%Y-%m-%d",time.localtime(time.time()))+'符合性确认数据.xls'
workbook = xlwt.Workbook(write_file_name)
worksheet = workbook.add_sheet('统计数据') 