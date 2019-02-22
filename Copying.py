import argparse
from openpyxl import load_workbook, Workbook
import time

from argparse import ArgumentParser
import sys
from sys import argv

# from Copying2 import *
from myFunctions import *

# parser = argparse.ArgumentParser(description='Takes in the excel documents')
# parser.add_argument('workbooks', type=argparse.FileType('r'), nargs='+')
# args = parser.parse_args()

workbook1 = ('LyndasData_RC.xlsx')
workbook2 = ('LyndasData_SIS.xlsx')
workbook3 = ('example3.xlsx')

timestr = time.strftime("%Y%m%d")

workbook4 = ('Initium.xlsx', timestr)
# worksheet1 = args.workbooks[0].name
# worksheet2 = args.workbooks[1].name
# worksheet3 = args.workbooks[2].name


wb1 = load_workbook(workbook1)
ws1 = wb1.active

wb2 = load_workbook(workbook2)
ws2 = wb2.active

wb3 = Workbook()
ws3_all_info = wb3.active
wb4 = Workbook()
initium_Canada_info = wb4.active
initium_Canada_info.title = ("Canada")
initium_USA_info = wb4.create_sheet("USA")




copy_paste_lookupid(ws1, ws3_all_info)
copy_paste_initial_info(ws1, ws3_all_info)
copy_paste_other_info(ws1, ws3_all_info)
length = len(ws3_all_info['A']) + 1
append_second_worksheet_initial_info(ws2, ws3_all_info)
append_second_worksheet_other_info(ws2, ws3_all_info, length)
categorize_emails(ws3_all_info)
format_phone_number(ws3_all_info)
format_country(ws3_all_info)
format_postal_code(ws3_all_info)
format_first_row(ws3_all_info)
format_non_initium_address(ws3_all_info)
# If worksheet flags certain address, we don't want to make a second Initium file until those are fixed
# Have a second script ready once she's happy and fixed everything? Yes.
# copy_paste_to_initium_file(ws3_all_info, initium_Canada_info, "Canada")
# copy_paste_to_initium_file(ws3_all_info, initium_USA_info, "United States of America")


wb3.save('example3.xlsx')
# wb4.save('Initium Ready - ' + timestr + '.xlsx')
