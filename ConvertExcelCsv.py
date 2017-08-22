#! python3

import openpyxl, os, sys

wbpath=sys.argv[1]
currentworkdirectory=os.getcwd()

print(wbpath)
workbook=openpyxl.load_workbook(wbpath, data_only=True)
workbook.get_sheet_names()