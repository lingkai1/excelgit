from itertools import islice
from openpyxl.styles import Border, Side, Font, Alignment
from pandas import DataFrame
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
def set_align(ws, cell_range,align,font,border):
    for cells in ws[cell_range]:
        for cell in cells:
            cell.alignment = align
            cell.font = font
            cell.border=border

    
def set_border(ws, cell_range):
    rows = list(ws.iter_rows(cell_range))
    side = Side(border_style='medium', color="FF000000")

    rows = list(rows)  # we convert iterator to list for simplicity, but it's not memory efficient solution
    max_y = len(rows) - 1  # index of the last row
    for pos_y, cells in enumerate(rows):
        max_x = len(cells) - 1  # index of the last cell
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            if pos_x == 0:
                border.left = side
            if pos_x == max_x:
                border.right = side
            if pos_y == 0:
                border.top = side
            if pos_y == max_y:
                border.bottom = side

            # set new border only if it's one of the edge cells
            if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                cell.border = border

wb2 = load_workbook('20161204.xlsx')
print wb2.get_sheet_names()
ws1 = wb2[u'20161204']
data = ws1.values
cols = next(data)[1:]
data = list(data)
idx = [r[0] for r in data]
data = (islice(r, 1, None) for r in data)
df = DataFrame(data, index=idx, columns=cols)
#df = DataFrame(data)
border = Border(left=Side(style='medium',color='FF000000'),right=Side(style='medium',color='FF000000'),top=Side(style='medium',color='FF000000'),bottom=Side(style='medium',color='FF000000'),diagonal=Side(style='medium',color='FF000000'),diagonal_direction=0,outline=Side(style='medium',color='FF000000'),vertical=Side(style='medium',color='FF000000'),horizontal=Side(style='medium',color='FF000000'))
fontObj2 = Font(size=9, italic=False)
align = Alignment(horizontal='center', vertical='center', wrapText = True) 

wb = Workbook()
ws1=wb.active
#设置开头表项
################################
################################
#第一行   
ws1.merge_cells('A1:G1')
set_border(ws1, 'A1:G1')
ws1['A1']='线管绕制记录' 
#ws1['A1'].font=font
ws1['A1'].border=border
#第二行
ws1.merge_cells('A2:G2')
ws1['A2'].border=border
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
ws1['A1'].font=fontObj2



for r in dataframe_to_rows(df, index=False, header=False):
    ws1.append(r)
str()
for cell in ws1['A']:
    cell.style = 'Pandas'
#设置末尾表项
########################
#######################
#第一行
length = ws1.max_row
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

#设置格式与对齐方式    
ws1.row_dimensions[ws1.max_row-1].height = 25
ws1.row_dimensions[6].height = 25
set_align(ws1,'A2:'+'G'+str(ws1.max_row),align,fontObj2,border)   
str(ws1.max_row)
wb.save('pandas_openpyxl.xlsx')