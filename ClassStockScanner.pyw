import os
from clsDoExcel import DoExcel

StockFile="SampleData.xlsx"
foldPath=os.getcwd()
excel=DoExcel(foldPath,StockFile)
print(excel.getColumns(1))
lstCols=input("Please type consecutive column numbers separated by comma for Code columns to searched : ")
lstCols=excel.mapColsList(1,lstCols)
print(lstCols)
while True:
    code=input("Please scan the barcode : ")
    if excel.findOnce(lstCols,code,False)[0]:
        print(excel.findOnce(lstCols,code,False))
    else:
        print("Not Found")
