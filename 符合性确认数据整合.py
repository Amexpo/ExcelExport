# -*- coding: utf-8 -*-
"""
Created on Tue Aug 14 08:56:47 2018

@author: yangyanhao
"""

import os
import time
import xlwt
import xlrd
import win32com.client
import requests

def time_compare(file_dir):
    ft=time.gmtime(os.stat(file_dir).st_mtime)
    nt=time.localtime(time.time())
    x=((ft[0]*12+ft[1]*30+ft[2])>
       (nt[0]*12+nt[1]*30+nt[2]-90)
       )
    return(x)


    
def xls_write(dir_list,worksheet,
              #file_name,
              init):#file_dir,name):
    workbook = xlrd.open_workbook(dir_list[0])
    sheet1 = workbook.sheet_by_index(1)
    #write_file_name =GetTime()+'符合性确认数据.xls'
    #workbook = xlwt.Workbook(file_name)#write_file_name)
    #worksheet = workbook.add_sheet('统计数据')
    for h in range(3,sheet1.nrows):
        row = sheet1.row_values(h)
        for j in range(7):
            worksheet.write(h+init-3, j, label = row[j])
        worksheet.write(h+init-3, 7, label = dir_list[1])
    #workbook.save(file_name)
    if ((sheet1.nrows-3)<0):
        rows=0
    else:
        rows=sheet1.nrows-3
    return(rows)


def find_dir(path):
    write_file_name =time.strftime("%Y-%m-%d",time.localtime(time.time()))+'符合性确认数据.xls'
    workbook = xlwt.Workbook(write_file_name)
    worksheet = workbook.add_sheet('统计数据') 
    path_set=[]
    name_set=[]
    x=0
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
                    if list_dir3[l][-3:]=='xls':
                        path4=path+list_dir[i]+'\\'+list_dir1[j]+'\\'+list_dir2[k]+'\\'+list_dir3[l]
                        if (time_compare(path4)==1):
                            path_set.append([path4,name])
                        else:
                            pass
                    else:
                        pass
    return(path_set)

write_file_name =time.strftime("%Y-%m-%d",time.localtime(time.time()))+'符合性确认数据.xls'
y=find_dir('2018年\\')
init=0

workbook = xlwt.Workbook(write_file_name)#write_file_name)
worksheet = workbook.add_sheet('统计数据')
for o in range(len(y)):
    x=xls_write(y[o],worksheet,init)
    init+=(x)
workbook.save(write_file_name)
