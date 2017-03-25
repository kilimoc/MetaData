
import win32com.client
import xlsxwriter
from os import listdir
from os.path import isfile, join

#get the path to your files folder.This is only used  to get the number of files in a directory.
path='E:\Mobile\Test\\'
#create the excel workbook and include one workbook.
workbook=xlsxwriter.Workbook("MetaData.xlsx")
worksheet=workbook.add_worksheet("Metadata")
#Use the win32com methods to manipulate windows files.
sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
#get to the folder/directory where your files are located.
ns = sh.NameSpace(r'E:\Mobile\Muhammad')
#method to get all the column title for the metadata.
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

#Method to get the number of files located in the folder.You can see here the path is used.
def getFileCount():
    counter=0;
    files =[f for f in listdir(path) if isfile(join(path,f))]
    for file in files:
        print(file)
    for i in range(len(files)):
       counter+=1
    return counter

#get all the values corresponding to columntitlevalues and return a list.
def getColumnValues():
    columnValues=[]
    colnum = 0
    columns = []
    dataset=[]
    while True:
        colname = ns.GetDetailsOf(None, colnum)
        if not colname:
            break
        columns.append(colname)
        colnum += 1

        rowdata=[]
    #for count in range(getFileCount()):

        for item in ns.Items():
            rows = []
            for colnum in range(len(columns)):
                colval = ns.GetDetailsOf(item, colnum)
                rows.insert(colnum, colval)
            rowdata.append(rows)
    return rowdata

#Set the title describing the metadata.You can see the method called inside is that containing a list of all titles.
def setTitle():
    row = 0
    column = 0
    coltitle = getColumntitleValues()
    for i in range(0, len(coltitle)):
        worksheet.write(row, column+i, coltitle[i])

#populate each row of the excel with each set of metadata.
def populate():
    row = 1
    column = 0
    cellvalues = getColumnValues()
    for i in range(0, len(cellvalues)):
        for cellvalue in (cellvalues[i]):
            worksheet.write(row+i,column+(cellvalues[i].index(cellvalue)),cellvalue)
        i+=1



#call the two methods.Othere methods are called inside these two.
setTitle()
populate()
#Close the workbook.
workbook.close()
print("Data reading and Writing Successful.Check the same folder for the excel data.")







