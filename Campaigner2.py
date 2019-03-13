from openpyxl import load_workbook, Workbook
import glob
from myFunctions import *
import os, sys
from openpyxl.worksheet.datavalidation import DataValidation

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

"""New workbook: Contact Update Template"""
wb3= Workbook()
ws3 = wb3.active

"""New workbook: Campaigner Initium Ready file"""
wb4= Workbook()
ws4 = wb4.active

column_list = ['C', 'D', 'E', 'F', 'J', 'L', 'M', 'N', 'O', 'Q', 'AA', 'I', 'AE', 'AD', 'AF']

x = make_column_list(ws1, column_list)

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

ws3.insert_cols(column_index_from_string('M'), 2)
"""Sets the Address Type and Address is Primary option to newly created column"""
for cell in ws3['M']:
    cell.value = 'H'
    (cell.offset(row=0, column=1).value) = 0
ws3.insert_cols(column_index_from_string('Q'), 2)
"""Sets the Phone Type and Phone is Primary option to newly created column"""
for cell in ws3['Q']:
    cell.value = 'H'
    (cell.offset(row=0, column=1).value) = 0

"""Sets the Email Type and Email is Primary option to newly created column"""
for cell in ws3['T']:
    cell.value = 'H'
    (cell.offset(row=0, column=1).value) = 0

"""Sets the Source column of the worksheet to Online Contact Update"""
for cell in ws3['W']:
    cell.value = 'Online Contact Update'


title_row = ["LOOKUP ID", "FIRST_NAME", "MIDDLE_NAME",	 "LAST_NAME",	"Street1", "Street2", "Street3",	"Street4",
            "CITY", "STATE","Postal_Code", "COUNTRY", "Address Type", "Address is Primary",  "Mystery Column" , "Phone",
            "Phone Type", "Phone is Primary", "Email", "Email Type", "Email is Primary",	"Last_UPDT", "Source"]
i = 0
for row in ws3.iter_rows(min_row=1, max_row=1, max_col=len(title_row)):
    for cell in row:
        cell.value = title_row[i]
        cell.font = Font(bold=True)
        i = i + 1


dv1 = DataValidation(type="list", formula1='"H,B,A,O,P,S"', allow_blank=True)
create_data_validation(dv1, ws3, 'M')

dv2 = DataValidation(type="list", formula1='"0,1"', allow_blank=True)
create_data_validation(dv2, ws3, 'M')
create_data_validation(dv2, ws3, 'N')
create_data_validation(dv2, ws3, 'R')

dv3 = DataValidation(type="list", formula1='"H,B,C,F,0,S"', allow_blank=True)
create_data_validation(dv3, ws3, 'Q')

dv4 = DataValidation(type="list", formula1='"H,A,B,P"', allow_blank=True)
create_data_validation(dv4, ws3, 'T')

# This drop down is ready to go if Lynda wants it
# dv5 = DataValidation(type="list", formula1='"Online Contact Update"', allow_blank=True)
# create_date_validation(dv5, ws3, 'W')

values_for_initium = ['B','J', 'L', 'M', 'N', 'O', 'Q', 'I']
"""Copies out Canadian address from Campaigner_workbook and puts them into a Initium ready file"""
for cell in ws1['I']:
    if cell.value == 'Canada':
        all_info = []
        for key in values_for_initium:
            all_info.append(ws1.cell(row=cell.row, column=column_index_from_string(key)).value)
        ws4.append(all_info)



wb3.save("Campaigner - Contact_Update_Template.xlsx")
wb4.save("Campaigner - Initium Ready.xlsx")



