# -*- coding: utf-8 -*-
"""
Created on Mon Dec 26 20:11:45 2016

@author: Link
"""
import numpy as np
def appendRowsViaDataFrame(ws, df):   
    group = []
    #按层数分割
    minLayer = df[u'层数'].min()
    for i in range(minLayer, len(df.groupby(u'层数'))+minLayer):
        group.append(df[df[u'层数']==i])
#    for i in range(df(u'层数').min(), len(df.groupby(u'层数'))+1):
#        group.append(df[df[u'层数']==i])
        
    j=1
    #往ws添加行
    for item in group:
        sheetRow = []
        sheetRow.append(item[u'层数'].min())                    #1. 写层数      
        speedMax = item[u'转速'].max()
        forceMax = item[u'张力值'].max()
        sheetRow.append(item[u'层圈数'].max()) #2. 写圈层数     
        sheetRow.append(' ')                  #3. 瑕疵不填  
        if len(item)==1:
            pass
        
        zeroSpeedFlag = False
        for i in range(0, len(item)):
            if item.iloc[i,5]==0:   
                item.iloc[i,4] = np.inf #把速度为0的张力写成inf 寻找最小值的时候可以直接忽略该项
                if i == len(item)-1:
                    zeroSpeedFlag = True
            else:                 #速度不为0  #4. 写时间
#                sheetRow.append(item.iloc[i,0])
                sheetRow.append(str(item.index[i])[11:19])
                break
        forceMin = item[u'张力值'].min()       
        if zeroSpeedFlag==True:  #该层的速度全为0
            sheetRow.append('0')
            forceMin = 0
            forceMax = 0
                           
        sheetRow.append(forceMax)           #5. 写张力最大值
        sheetRow.append(forceMin)           #6. 写张力最小值
        sheetRow.append(speedMax)           #7  写速度最大值
        ws.append(sheetRow)                 #   把行写到excel sheet
        j=j+1
