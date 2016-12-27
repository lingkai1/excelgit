# -*- coding: utf-8 -*-
"""
Created on Mon Dec 26 11:38:06 2016

@author: Link
"""
def drawTail(ws1):
    length = ws1.max_row+1
    ws1.merge_cells('A'+str(length)+':C'+str(length))
    ws1['A'+str(length)]='进入加强段圈数（圈）'
    ws1.merge_cells('D'+str(length)+':G'+str(length))
    ws1['D'+str(length)]               #输入进入加强段圈数（圈）
    #第二行
    ws1.merge_cells('A'+str(length+1)+':C'+str(length+1))
    ws1['A'+str(length+1)]='加强段长度（米）'
    ws1.merge_cells('D'+str(length+1)+':G'+str(length+1))
    ws1['D'+str(length+1)]               #加强段长度（米）
    #第三行
    ws1.merge_cells('A'+str(length+2)+':C'+str(length+2))
    ws1['A'+str(length+2)]='导线总长度（米）'
    ws1.merge_cells('D'+str(length+2)+':G'+str(length+2))
    ws1['D'+str(length+2)]               #导线总长度（米）
    #第四行
    ws1.merge_cells('B'+str(length+3)+':C'+str(length+3))
    ws1['A'+str(length+3)]='绕线者：'
    ws1['B'+str(length+3)]             #输入绕线者
    
    ws1.merge_cells('D'+str(length+3)+':E'+str(length+3))    
    ws1.merge_cells('F'+str(length+3)+':G'+str(length+3))

    ws1['D'+str(length+3)]='检验人员： '   
    ws1['F'+str(length+3)]             #输入检验人员      
    
    #第五行
    ws1.merge_cells('A'+str(length+4)+':B'+str(length+5))
    ws1['A'+str(length+4)]='备注'
    ws1.merge_cells('C'+str(length+4)+':G'+str(length+4))
    ws1['C'+str(length+4)]='制导线生产厂家：江苏泰兴电工厂  □    ***       湖南华菱线缆股份有限公司 □   ***'              #导线总长度（米）
    ws1.merge_cells('C'+str(length+5)+':G'+str(length+5))
    ws1['C'+str(length+5)]='4圈=1米；加强段长度（米）=（总圈数-进入加强段圈数）/4'