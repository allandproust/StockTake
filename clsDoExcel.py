from openpyxl import load_workbook
import os

class DoExcel:

    def __init__(self,folderPath,fileName): # Opens excel file and returns file obj with active sheet to work on.

        self.fPath=folderPath
        self.file=fileName
        self.book=load_workbook(os.path.join(self.fPath,self.file))
        self.sheet=self.book.active
        self.mappedCodes={}
        self.lastColumn=self.sheet.max_column+1
        print("{} Workbook loaded".format(self.file))
##----------------------------------------------------------------------------------------------------------------------------------
    def __mapColumn(self,num,row=None,blncolNum=False):
        # takes in number,optional row and boolean for column number - returns corresponding column name (Private method)
        _alpha={1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M',\
               14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z'}
        if blncolNum:
            colName=str(_alpha.get(int(num),0))+str(row)
        else:
            colName=str(_alpha.get(int(num),0))
        return colName
##----------------------------------------------------------------------------------------------------------------------------------
    def mapColsList(self,row=1,lstCols=[]): # Takes in numbers and returns consecutive Column alphabets
        lstMappedCols=[]
        for item in lstCols:
            if not item==",":
                item=int(item)
                lstMappedCols.append(self.__mapColumn(item,row))
        return lstMappedCols
##----------------------------------------------------------------------------------------------------------------------------------
    def getColumns(self,row=1,blncolNum=False): # Displays column names in with headers and returns a dictionary for the same
        for col in range(1,self.lastColumn):
            print("{} : {}".format(col,self.sheet.cell(row=row,column=col).value))
            self.mappedCodes[col]=self.sheet.cell(row=row,column=col).value
        return self.mappedCodes
##----------------------------------------------------------------------------------------------------------------------------------
    def deleteCellData(self, row,col):
        self.sheet.cell(row=row,column=col).value=None
        self.saveFile()
##----------------------------------------------------------------------------------------------------------------------------------      
    def writeToCell(self,row,col,data): # Method write to cell
        self.sheet.cell(row=row,column=col).value=data
        self.saveFile() # testing write process - TBC  
##----------------------------------------------------------------------------------------------------------------------------------
    def writeColumnsHeaders(self,row=1,*args):  # takes in row and headers and writes to the columns
        lstNewHeader=[]
        for header in args:
            print('printing {} at column={}'.format(header,self.sheet.max_column+1))
            self.writeToCell(row,self.sheet.max_column+1,header)
            lstNewHeader.append(self.sheet.max_column)
        return lstNewHeader
##----------------------------------------------------------------------------------------------------------------------------------                
    def findOnce(self,lstColumn,data,Multiple=False): # takes Column List,Data to be found, find multiple or one instance as boolean
        Found=False
        count=0
        if not Multiple:
            for col in lstColumn:
                for row in range(1,len(self.sheet[col])):
                    #print("column-{}:row-{} : {}".format(col,row,self.sheet[col][row].value))
                    if data==self.sheet[col][row].value:
                        Found=True
                        return Found,row
                    else:
                        Found=False
            return [Found]
        else:
            for col in lstColumn:
                for row in range(len(self.sheet[col])):
                    #print("searching row-{} and col - {}".format(row,col))
                    if data!=self.sheet[col][row].value:
                        Found=False
                        pass
                    else:
                        Found=True
                        count+=1
                        continue
                if count>0:
                    Found=True
            return [Found,count]
##----------------------------------------------------------------------------------------------------------------------------------
    def saveFile(self): # saveFile=takes path and saves file, if no file location overwrites the current file.
        self.book.save(os.path.join(self.fPath,self.file))
##----------------------------------------------------------------------------------------------------------------------------------

