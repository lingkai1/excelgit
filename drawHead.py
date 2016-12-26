# -*- coding: utf-8 -*-
"""
Created on Mon Dec 26 11:27:26 2016

@author: Link
"""
from fontAndAlign import set_border
def drawHead(ws1):

    #设置开头表项
################################
################################
#第一行   
    ws1.merge_cells('A1:G1')
    set_border(ws1, 'A1:G1')
    ws1['A1']='线管绕制记录' 
#ws1['A1'].font=font
#    ws1['A1'].border=border
    #第二行
    ws1.merge_cells('A2:G2')    #日期
#    ws1['A2'].border=border
    #第三行
    ws1.merge_cells('A3:B3')
    ws1['A3']='线盘号'
    ws1['C3']                   #线盘号
    ws1.merge_cells('D3:E3')
    ws1['D3']='温度（℃）'
    ws1.merge_cells('F3:G3')
    ws1['F3']                   #s温度 
    #第四行
    ws1.merge_cells('A4:B4')
    ws1['A4']='线盘总质量(g)'
    ws1['C4']                   #输入总重量
    ws1.merge_cells('D4:E4')
    ws1.merge_cells('F4:G4')
    ws1['D4']='湿度（%）'
    ws1['F4']                    #输入湿度
    #第五行
    ws1.merge_cells('A5:B5')
    ws1['A5']='线盘重量（克）'
    ws1['C5']                   #输入
    ws1.merge_cells('D5:E5')
    ws1.merge_cells('F5:G5')
    ws1['D5']='翼筒线管体编号（%）'
    ws1['F5']                    #输入
    #第六行
    ws1.merge_cells('A6:B6')
    ws1['A6']='制导线重（克）'
    ws1['C6']                   #输入
    ws1.merge_cells('D6:E6')
    ws1.merge_cells('F6:G6')
    ws1['D6']='翼筒线管体重量（克）（%）'
    ws1['F6']                    #输入
    #第七行
    ws1.merge_cells('A7:A8')
    ws1['A7']='层数'
    ws1.merge_cells('B7:B8')
    ws1['B7']='圈数'
    ws1.merge_cells('C7:C8')
    ws1['C7']='导线瑕疵'
    ws1.merge_cells('D7:D8')
    ws1['D7']='时间'
    ws1.merge_cells('e7:f7')
    ws1['E7']='张力'
    ws1['E8']='最大值'
    ws1['F8']='最小值'
    ws1.merge_cells('G7:G8')
    ws1['G7']='最高转速\n(r/min)'
    #ws1['A1'].font=fontObj2
