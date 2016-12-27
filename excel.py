# -*- coding: utf-8 -*-
"""
Created on Mon Dec 26 11:38:06 2016

@author: Link
"""
import sys
from itertools import islice
from openpyxl.styles import Border, Side, Font, Alignment
from pandas import DataFrame
from openpyxl import load_workbook
from openpyxl import Workbook
#from openpyxl.utils.dataframe import dataframe_to_rows
from fontAndAlign import set_align
from drawHead import drawHead
from drawTail import drawTail
from appendRowsViaDataFrame import appendRowsViaDataFrame
#import numpy as np
if len(sys.argv)>1:
    fileName = sys.argv[1]
    print sys.argv[1]
else:
    fileName = '20161205.xlsx'
wb2 = load_workbook(fileName)
print wb2.get_sheet_names()
ws1 = wb2[u'20161204']
data = ws1.values
cols = next(data)[1:]
data = list(data)
idx = [r[0] for r in data]
data = (islice(r, 1, None) for r in data)
df = DataFrame(data, index=idx, columns=cols)
#df = DataFrame(data)
#格式
border = Border(left=Side(style='medium',color='FF000000'),right=Side(style='medium',color='FF000000'),top=Side(style='medium',color='FF000000'),bottom=Side(style='medium',color='FF000000'),diagonal=Side(style='medium',color='FF000000'),diagonal_direction=0,outline=Side(style='medium',color='FF000000'),vertical=Side(style='medium',color='FF000000'),horizontal=Side(style='medium',color='FF000000'))
fontObj2 = Font(size=9, italic=False)
align = Alignment(horizontal='center', vertical='center', wrapText = True) 

# new workBook
wb = Workbook()
ws = []

group=[]
j = 0
flag = 0
for i in df.iloc[:,2]:
    if i == 1 and flag == 0:
        group.append(j)
        flag = 1
    elif i != 1:
        flag = 0
    j = j+1                
groupNum = len(group)    

#split the dataframe
groupInList = []

if groupNum != 1: # Need to be split
    for i in range(0,groupNum):
        if i != groupNum-1:
            groupInList.append(df.iloc[group[i]:group[i+1]])
        else:
            groupInList.append(df.iloc[group[i]:len(df)])
else:
    groupInList.append(df)
#data process in each group
j = 0
for groupDf in groupInList:
    if j ==0:
        ws.append(wb.active)
    else:
        ws.append(wb.create_sheet("Sheet"))
    drawHead(ws[j])
    
    
    appendRowsViaDataFrame(ws[j], groupDf)
              
    drawTail(ws[j])  
    ws[j].row_dimensions[ws[j].max_row-1].height = 25
    ws[j].row_dimensions[6].height = 25
    set_align(ws[j],'A1:'+'G'+str(ws[j].max_row),align,fontObj2,border)
    j=j+1     

############# process data
#for r in dataframe_to_rows(df, index=False, header=False):
#    ws1.append(r)
#str()
#for cell in ws1['A']:
#    cell.style = 'Pandas'
################    

#表尾
#drawTail(ws1)

#设置格式与对齐方式    
fileNameSave=fileName[:-5]
wb.save(fileNameSave+'New'+'.xlsx')