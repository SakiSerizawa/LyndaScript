from openpyxl import load_workbook, Workbook
import glob
from myFunctions import *
import os, sys

"""Campainger1 will move the necessary cells in the large Campaigner file to something more succient and somewhat
smaller. It also creates a file (cmt.xlsx) which is ready for the Link Constituent Matching Tool"""

os.chdir("C:/Users/sakiikas/Documents/LyndaScript/FromRecordsFolder/files")

#os.chdir("W:/Records/LyndaScript-master/Folder2-Input_Files")


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

wb1.save("C:/Users/sakiikas/Documents/ScriptFiles_TEST/Folder1/Campaigner/Campaigner_workbook.xlsx")

#wb1.save("W:/Records/LyndaScript-master/Folder1-Output_Files/CampaignerFiles/Campaigner_workbook.xlsx")


os.chdir("C:/Users/sakiikas/Documents/ScriptFiles_TEST/Folder1/Campaigner")
# os.chdir("W:/Records/LyndaScript-master/Folder1-Output_Files/CampaignerFiles")


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
            cell.value = int(cell.value)


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


initium_title_row = ["Title", "First name","Middle","Last name", "Suffix", "Address 1", "Address 2",	"Address 3",
                     "Address 4","City","Province","Postal code",	"Country","Email","Phone",	"Degree Info",
                     "Alternate ID type","Alternate ID"]
i = 0
for row in ws3.iter_rows(min_row=1, max_row=1, max_col=len(initium_title_row)):
    for cell in row:
        cell.value = initium_title_row[i]
        cell.font = Font(bold=True, color='FF0000')
        i = i + 1


wb3.save("cmt.xlsx")
wb2.save("Campaigner_workbook.xlsx")
