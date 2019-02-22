from openpyxl import load_workbook, Workbook
import time
from myFunctions import *

workbook1 = ('example3.xlsx')
workbook2 = ('Results - Initinum - Feb 15-Canada-Results.xlsx')

wb1 = load_workbook(workbook1)
ws1 = wb1.active

wb2 = load_workbook(workbook2)
ws2 = wb2["Exact"]


for rows in ws1.rows://
    for cell in rows:
        cell.fill = PatternFill(fgColor='FAFAD2', fill_type='solid')
for cell in ws1[1]:
    cell.fill=PatternFill(fgColor='FFFFFF', fill_type='solid')
    cell.font = Font(bold=True)

for col in ws2.iter_cols(min_row=3):
        for cell in col:
            initium_info = [ws2.cell(row=cell.row, column=1).value, ws2.cell(row=cell.row, column=10).value,
                            ws2.cell(row=cell.row, column=11).value,ws2.cell(row=cell.row, column=12).value,
                            ws2.cell(row=cell.row, column=13).value,ws2.cell(row=cell.row, column=15).value,
                            ws2.cell(row=cell.row, column=16).value,ws2.cell(row=cell.row, column=17).value]

            replace_info(ws1, initium_info)


wb1.save('example3.xlsx')

