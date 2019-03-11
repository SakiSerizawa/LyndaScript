from openpyxl import load_workbook, Workbook
import time
import glob
from myFunctions import *
import os, sys


os.chdir("C:/Users/sakiikas/Documents/LyndaScript/FromRecordsFolder/files")

Campaigner_download_file = (glob.glob("*Campaigner*")[0])


workbook1 = Campaigner_download_file

wb1 = load_workbook(workbook1)
ws1 = wb1.active

"""Moves values over from K column and places them into Q column"""
for row in ws1.iter_rows(min_col=column_index_from_string('K'),max_col=column_index_from_string('K'), min_row=2):
    for cell in row:
        if cell.value is not None:
            cell.offset(row=0, column=6).value = cell.value
            cell.value = None

"""Moves values over from P column and places them into Q column"""
for row in ws1.iter_rows(min_col=column_index_from_string('P'),max_col=column_index_from_string('P'), min_row=2):
    for cell in row:
        if cell.value is not None:
            cell.offset(row=0, column=1).value = cell.value
            cell.value = None

"""Moves values over from R column and places them into Q column"""
for row in ws1.iter_rows(min_col=column_index_from_string('R'),max_col=column_index_from_string('R'), min_row=2):
    for cell in row:
        if cell.value is not None:
            cell.offset(row=0, column=-1).value = cell.value
            cell.value = None

"""Moves values over from T U V W X Y  columns and places them into Q column"""
for row in ws1.iter_rows(min_col=column_index_from_string('T'),max_col=column_index_from_string('Y'), min_row=2):
    for cell in row:
        if cell.value is not None:
            ws1.cell(row=cell.row, column=column_index_from_string('Q')).value = cell.value
            cell.value = None

"""Moves values over from AB and AC columns and places them into AA column"""
for row in ws1.iter_rows(min_col=column_index_from_string('AB'),max_col=column_index_from_string('AC'), min_row=2):
    for cell in row:
        if cell.value is not None:
            ws1.cell(row=cell.row, column=column_index_from_string('AA')).value = cell.value
            cell.value = None

wb1.save("C:/Users/sakiikas/Documents/ScriptFiles_TEST/Folder1/Campaigner_workbook.xlsx")



os.chdir("../../../sakiikas/Documents/ScriptFiles_TEST/Folder1")


workbook2 = ('Campainger_workbook.xlsx')
wb2 = load_workbook(workbook2)
ws2 = wb2.active



"""Formats phone number by removing extra spaces and unnecessary characters"""
for row in ws2.iter_rows(min_row=2, min_col=column_index_from_string('AD'), max_col=column_index_from_string('AD')):
    for cell in row:
        phone = str(cell.value)
        cell.value = phone.replace('-', '').replace('(', '').replace(')', '').replace(' ', '').replace('None', '')
        if cell.value is None or cell.value == '':
            continue
        else: cell.value = int(cell.value)

