#All Imports
import os
import xlsxwriter
import win32com.client
from os import listdir
from os.path import isfile, join

#get the path to your files.This is useful in counting number of files in the directory
path='E:\Mobile\Test\\'

#Use the win32com methods here
sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)

#Insert the same path here as that but in this format.
ns = sh.NameSpace(r'E:\Mobile\Test')
workbook=xlsxwriter.Workbook("MetaData.xlsx")
worksheet=workbook.add_worksheet("Metadata")

#Methods to get number of files in the folder specified.
def getFileCount():
    counter=0;
    files =[f for f in listdir(path) if isfile(join(path,f))]
    for file in files:
        print(file)
    for i in range(len(files)):
       counter+=1
    return counter

#Method to get the column titles to be entered in the first row of the .
def getColumntitleValues():
    columntitle=[]

    colnum = 0
    columns = []
    while True:
        colname = ns.GetDetailsOf(None, colnum)
        if not colname:
            break
        columns.append(colname)
        colnum += 1
        #columntitle+=columns[colnum]
    return columns

#method to get the actual metadata data to be inserted in the excel sheet.
def getColumnValues():
    columnValues=[]
    colnum = 0
    columns = []
    while True:
        colname = ns.GetDetailsOf(None, colnum)
        if not colname:
            break
        columns.append(colname)
        colnum += 1
        rowdata=[]
    for count in range(getFileCount()):
        for item in ns.Items():
            print(item.Path)
            rows = []
            for colnum in range(len(columns)):
                colval = ns.GetDetailsOf(item, colnum)
                rows.insert(colnum, colval)
                colnum+=1
            rowdata.append(rows)
        count += 1
    return rowdata









print(getColumnValues())



def setTitle():
    row = 0
    column = 0
    coltitle = getColumntitleValues()
    for i in range(0, len(coltitle)):
        worksheet.write(row, column+i, coltitle[i])
def populate():
    row = 1
    column = 0
    cellvalues = getColumnValues()
    for i in range(0, len(cellvalues)):
        for cellvalue in (cellvalues[i]):
            worksheet.write(row+i,column+(cellvalues[i].index(cellvalue)),cellvalue)
        i+=1


print(len(getColumnValues()))
setTitle()
populate()
workbook.close()

