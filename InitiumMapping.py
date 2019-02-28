from openpyxl import load_workbook, Workbook
import time
import glob
from myFunctions import *
import os, sys


os.chdir("../../../sakiikas/Documents/ScriptFiles_TEST/Folder1")

initium_results_file = (glob.glob("*Results*")[0])


workbook1 = ('combined_workbook.xlsx')
workbook2 = (initium_results_file)

wb1 = load_workbook(workbook1)
ws1 = wb1.active

wb2 = load_workbook(workbook2)
ws2 = wb2["Exact"]


colour_worksheet(ws1)

for col in ws2.iter_cols(min_row=3):
        for cell in col:
            initium_info = [ws2.cell(row=cell.row, column=1).value, ws2.cell(row=cell.row, column=10).value,
                            ws2.cell(row=cell.row, column=11).value, ws2.cell(row=cell.row, column=12).value,
                            ws2.cell(row=cell.row, column=13).value, ws2.cell(row=cell.row, column=15).value,
                            ws2.cell(row=cell.row, column=16).value, ws2.cell(row=cell.row, column=17).value]

            replace_info(ws1, initium_info)


wb1.save('combined_workbook.xlsx')

