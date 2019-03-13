from openpyxl import load_workbook, Workbook
import glob
from myFunctions import *
import os, sys

# This script is to be run after the cmt file has run through the LINKS Constituent Matching Tool
# It is assumed the user has copied back the results back into the Campainger_workbook.xlsx file

#
# The purpose of this script is 1) Create a CommMailPreferences file based off the updated Campaigner_workbook.xlsx
# 2) to delete any rows of data that could not be matched (through matching tool or manually
# 3) Prepare the file for LINKS

os.chdir("C:/Users/sakiikas/Documents/ScriptFiles_TEST/Folder1")


workbook1 = ('Campaigner_workbook.xlsx')
wb1 = load_workbook(workbook1)
ws1 = wb1.active

wb2 = Workbook()
ws2 = wb2.active

"""Creates LINKS ready file of those who would only like to receive TREK online"""


"""Extracts LOOKUPID, FIRST/MIDDLE/LAST NAME, EMAIL from Campaigner_workbook and puts them in new CommPrefUpdate file"""
extra_row_info = ["General Correspondence", " ", "AA - TREK Magazine"," ", "No", "M", "Alumni Affairs",
                  "Requested by constituent", "Last_UPDT", "Alumni Association"]
for cellz in ws1['BH']:
    row_info = []
    if "online" in cellz.value:
        if "unable to locate" not in str(ws1.cell(row=cellz.row, column=2).value):
            row_info.append(ws1.cell(row=cellz.row, column=2).value)
            row_info.append(ws1.cell(row=cellz.row, column=4).value)
            row_info.append(ws1.cell(row=cellz.row, column=5).value)
            row_info.append(ws1.cell(row=cellz.row, column=6).value)
            row_info.append(ws1.cell(row=cellz.row, column=column_index_from_string('AF')).value)
            for item in extra_row_info:
                row_info.append(item)
            ws2.append(row_info)


"""Converts most LOOKUPID's back to integers to prevent warnings in excel (ex "this number is stored as string"""
for cell in ws2['A']:
    try:
        cell.value = int(cell.value)
    except:
        continue


"""Creates a properly formatted Title Row in the CommMailPreference workbook"""
comm_mail_preferences_title_row = ["LOOKUP ID", "EMAIL", "FIRST_NAME", "MIDDLE_NAME", "LAST_NAME", "MAIL Type", "Site",
                                   "Correspondence code", "Category", "Send Yes/No", "Send by", "Business Unit",
                                   "Comments", "Last_UPDT", "Source", "Sequence"]

ws2.insert_rows(1,1)
i=0
for row in ws2.iter_rows(min_row=1, max_row=1):
    for cell in row:
        cell.value = comm_mail_preferences_title_row[i]
        cell.font = Font(bold=True, color='FF0000')
        i = i + 1

wb2.save("CommMailPreferences.xlsx")
