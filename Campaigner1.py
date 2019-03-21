from openpyxl import load_workbook, Workbook
import glob
from myFunctions import *
from stateAbbreviations import *
import os, sys
import unicodedata

"""Campainger1 will move the necessary cells in the large Campaigner file to something more succient and somewhat
smaller. It also creates a file (cmt.xlsx) which is ready for the Link Constituent Matching Tool"""

os.chdir("C:/Users/sakiikas/Documents/ScriptFiles_TEST/Campaigner")

#os.chdir("W:/Records/LyndaScript-master/Campaigner Files")


Campaigner_download_file = (glob.glob("*Download*")[0])

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

"""Moves values over from R S T U V W X Y Z columns and places them into Q column"""
for row in ws1.iter_rows(min_col=column_index_from_string('R'),max_col=column_index_from_string('Z'), min_row=2):
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

"""Moves values over from AS,AT,AU,AV,AW,AX,AY,AZ,BA and AC and places them into AR column"""
for row in ws1.iter_rows(min_col=column_index_from_string('AS'),max_col=column_index_from_string('BA'), min_row=2):
    for cell in row:
        if cell.value is not None:
            ws1.cell(row=cell.row, column=column_index_from_string('AR')).value = cell.value
            cell.value = None

"""Moves values over from AQ and places them into AR column"""
for row in ws1.iter_rows(min_col=column_index_from_string('AQ'),max_col=column_index_from_string('AQ'), min_row=2):
    for cell in row:
        if cell.value is not None:
            ws1.cell(row=cell.row, column=column_index_from_string('AR')).value = cell.value
            cell.value = None

"""Moves values over from AK and places them into AR column"""
for row in ws1.iter_rows(min_col=column_index_from_string('AK'),max_col=column_index_from_string('AK'), min_row=2):
    for cell in row:
        if cell.value is not None:
            ws1.cell(row=cell.row, column=column_index_from_string('AR')).value = cell.value
            cell.value = None

"""Moves values over from BC,BD  and places them into AR column"""
for row in ws1.iter_rows(min_col=column_index_from_string('BC'),max_col=column_index_from_string('BD'), min_row=2):
    for cell in row:
        if cell.value is not None:
            ws1.cell(row=cell.row, column=column_index_from_string('BB')).value = cell.value
            cell.value = None



wb1.save("C:/Users/sakiikas/Documents/ScriptFiles_TEST/Campaigner/Campaigner_workbook.xlsx")

# wb1.save("W:/Records/LyndaScript-master/Campaigner Files/Campaigner_workbook.xlsx")


workbook2 = ('Campaigner_workbook.xlsx')
wb2 = load_workbook(workbook2)
ws2 = wb2.active


"""Formats phone number by removing extra spaces and unnecessary characters"""
for row in ws2.iter_rows(min_row=2, min_col=column_index_from_string('AD'), max_col=column_index_from_string('AD')):
    for cell in row:
        # cell.value = "HIT"
        phone = str(cell.value)
        cell.value = phone.replace('-', '').replace('(', '').replace(')', '').replace(' ', '').replace('None', '')
        if cell.value is None or cell.value == '':
            continue
        else:
            try:
                cell.value = int(cell.value)
            except:
                continue

"""Formats the states of both popular countries, USA, and Canada and reformats to their matching state"""
for cell in ws2['I']:
    if cell.value in popular_countries and cell.offset(row=0, column=8).value is not None:
        try:
            cell.offset(row=0, column=8).value = popular_countries[cell.value][cell.offset(row=0, column=8).value]
        except:
            cell.offset(row=0, column=8).font = Font(color='F9f631')
    elif cell.value in usa_canada and cell.offset(row=0, column=8).value is not None:
        try:
            cell.offset(row=0, column=8).value = usa_canada[cell.value][cell.offset(row=0, column=8).value]
        except:
            cell.offset(row=0, column=8).font = Font(color='F9f631')



wb2.save("Campaigner_workbook.xlsx")


wb3 = Workbook()
ws3 = wb3.active


column_list = ['D', 'E', 'F', 'J', 'L', 'M', 'N', 'O', 'Q', 'AA', 'I', 'AF', 'AD', 'G']
x = make_column_list(ws2, column_list)
information_from_excel = list(x)
maximum_rows = len(information_from_excel)
maximum_col = len(information_from_excel[0])

i = 0
for rows in ws3.iter_rows(max_row=maximum_rows, max_col=maximum_col):
    j = 0
    for cell in rows:
        cell.value = information_from_excel[i][j]
        j = j+1
    i = i + 1

ws3.insert_cols(1, 1)
ws3.insert_cols(5, 1)


"""Combines street address 1 and 2 into one"""
for cell in ws3['G']:
    if cell.offset(row=0, column=-1).value is not None:
        cell_tuple = (str(cell.offset(row=0, column=-1).value), str(cell.value))
        x = "-".join(cell_tuple)
        cell.value = x
        cell.offset(row=0, column=-1).value = ""
ws3.delete_cols(column_index_from_string('F'),1)
ws3.insert_cols(column_index_from_string('I'), 1)



cmt_title_row = ["Title", "First name","Middle","Last name", "Suffix", "Address 1", "Address 2",	"Address 3",
                     "Address 4","City","Province","Postal code",	"Country","Email","Phone",	"Degree Info",
                     "Alternate ID type","Alternate ID"]
i = 0
for row in ws3.iter_rows(min_row=1, max_row=1, max_col=len(cmt_title_row)):
    for cell in row:
        cell.value = cmt_title_row[i]
        cell.font = Font(bold=True, color='FF0000')
        i = i + 1




wb3.save("cmt.xlsx")
wb2.save("Campaigner_workbook.xlsx")
