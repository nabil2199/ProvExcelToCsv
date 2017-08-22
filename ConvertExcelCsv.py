#! python3

import openpyxl, os, sys

#wbpath=sys.argv[1]
#currentworkdirectory=os.getcwd()

#print(wbpath)
#workbook=openpyxl.load_workbook(wbpath, data_only=True)
#workbook.get_sheet_names()

workbook=openpyxl.load_workbook('C:\\Users\\NXLX8474\\Desktop\\todo20170824\\lotb5\\lotb5.xlsm', data_only=True)
workbook.get_sheet_names()
sheet=workbook.get_sheet_by_name('Lot A - User Information')
sheetsize=sheet.max_row

#def getUserInfoSheetSize(workbookfile):
#    size=1
#    cellcoordinate="F"+"size"
#    cell=sheet[cellcoordinate]
#    while (cell.value!=0 and cell.value!=None):
#        size = size + 1
#    return size

def loadColumn(columnindex):
    valuelist=[]
    i=1
    while (i < sheetsize):
        cellcoordinate=columnindex + str(i)
        cell = sheet[cellcoordinate]
        valuelist.insert(i,cell.value)
        i=i+1
    return valuelist

def loadUPN(sheetsize):
    upnlist = loadColumn("E")
    return upnlist

def loadSIP(sheetsize):
    upnlist = loadColumn("F")
    return upnlist

def loadFName(sheetsize):
    fnamelist = loadColumn("E")
    return fnamelist

def loadLName(sheetsize):
    lnamelist = loadColumn("F")
    return lnamelist

upns=loadUPN(sheetsize)
sips=loadSIP(sheetsize)


for i in range(0,sheetsize-1):
    if upns[i] != None and sips[i] != None:
        print(upns[i])
        print(sips[i])

def initFile():
    return 5

