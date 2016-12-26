# -*- coding: utf-8 -*-
"""
Created on Mon Dec 26 08:45:47 2016

@author: Link
"""
from openpyxl.styles import Border, Side
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