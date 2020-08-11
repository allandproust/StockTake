from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import os

#---------------------------------------------------------------------------------------------------------------
fPath='StockTake.xlsx'
gblScanCount=0

###---Functions ----------------------------------------------------------------------------------------------------
##
def openFile():
   window.withdraw()
   address=filedialog.askopenfilename(initialdir=os.getcwd(),title='Select file for operations',filetypes=(('EXCEL FILES','*.xlsx'),('CSV FILES','*.csv')))
   return address

#-- Intialize --------------------------------------------------------------------------------------------------

from clsDoExcel import DoExcel

def findCode():

    txtCode.focus()
    code=txtCode.get()
    print("code : {}".format(code))
    result=excel.findValue(lstCodeCols,code,ALL=False)
    txtCode.delete(0, END)
    if not result[0]:
        lblScanCode['text']="Not Found"
        #messagebox.showinfo("Code labeltxt"," Code not found")
    else:
        excel.highlightRow(result[1])
        lblScanCode['text']=code
        global gblScanCount
        gblScanCount+=1
        lblScanCounter['text']=gblScanCount
        UpdateStatus(result[1],Found,'1')
        #messagebox.showinfo("Code labeltxt"," Code found at {}".format(labeltxt[1]))

def UpdateStatus(row,col,value):
   print(f' printing to row : {row}')
   print(f' printing to col : {col}')
   excel.writeToCell(row+1,col,value)
##    if excess:
##        update labelexcess
##        write to file no inward
##    else:
##        update label to short
   
def Calculate(row,total,found,colExcess,colShort):
   if total-found > 0:
      print('short')
   elif total-found < 0:
      print('short')
   else:
      print('equal')

def Save():
   excel.saveFile()

def Close():
   window.destroy()
## Interface---------------------------------------------------------------------------------------------------------

window=Tk()
window.geometry("500x400")

windowWidth = window.winfo_reqwidth()
windowHeight = window.winfo_reqheight()

# Gets both half the screen width/height and window width/height
positionRight = int(window.winfo_screenwidth()/2.5 - windowWidth)
positionDown = int(window.winfo_screenheight()/2.5 - windowHeight)

# Positions the window in the center of the screen.
window.geometry("+{}+{}".format(positionRight, positionDown))
window.config(bg='#364549')
#window.title("Stock Check : "+os.path.basename(fileAddress).split('.')[0])
window.title("Stock Check : ")

s=ttk.Style()
s.configure('My.TFrame',background='#364549')

content=ttk.Frame(window,padding=25,style='My.TFrame')
mainfrm=ttk.Frame(content,borderwidth=5,relief='sunken',width=400,height=500,style='My.TFrame')
content.grid(column=0,row=0)
mainfrm.grid(column=0,row=0)

txtCode=Entry(mainfrm,width=30)
txtCode.grid(column=0,row=0,padx=15)

btnNext=Button(mainfrm,text="Next barcode",width=20,height=2,command=lambda:findCode())
btnNext.config(font=('Arial',7,'bold'))
btnNext.grid(column=1,row=0,padx=10,pady=12)

lblTxtLastScan=Label(mainfrm,text="Last Bar code Scanned: ",bg='#364549',fg='#d8e3e3',width=22)
lblTxtLastScan.config(font=('Courier',10,'bold'))
lblTxtLastScan.grid(column=0,row=1,pady=10)

lblScanCode=Label(mainfrm,text='8907233000000',bg='#364549',fg='#d8e3e3')
lblScanCode.config(font=('Courier',10,'bold'))
lblScanCode.grid(column=1,row=1)

separate=ttk.Separator(mainfrm,orient='horizontal')
separate.grid(sticky=(E,W),row=2,columnspan=10,padx=2, pady=5)

lblTxtScanCount=Label(mainfrm,text="Scan Count : ",bg='#364549',fg='#d8e3e3')
lblTxtScanCount.config(font=('Courier',10,'bold'))
lblTxtScanCount.grid(column=0,row=4)

lblScanCounter=Label(mainfrm,text="0",bg='#364549',fg='#d8e3e3')
lblScanCounter.config(font=('Courier',10,'bold'))
lblScanCounter.grid(column=1,row=4)

lblTxtFound=Label(mainfrm,text="Found count : ",bg='#364549',fg='#d8e3e3')
lblTxtFound.config(font=('Courier',10,'bold'))
lblTxtFound.grid(column=0,row=5)

lblFoundCount=Label(mainfrm,text="-",bg='#364549',fg='#d8e3e3')
lblFoundCount.config(font=('Courier',10,'bold'))
lblFoundCount.grid(column=1,row=5)

##lblTxtPrice=Label(mainfrm,text="MRP : ",bg='#364549',fg='#d8e3e3')
##lblTxtPrice.config(font=('Courier',10,'bold'))
##lblTxtPrice.grid(column=0,row=6)

##lblDiscount=Label(mainfrm,text=" %",bg='#364549',fg='#d8e3e3')
##lblDiscount.config(font=('Courier',10,'bold'))
##lblDiscount.grid(column=1,row=6)

btnSave=Button(window,text="Save File",relief='raised',borderwidth=2,width=20,height=2,command=lambda:Save())
btnSave.config(font=('Arial',7,'bold'))
btnSave.grid(column=0,row=6)

btnQuit=Button(window,text="Quit",padx=5,width=20,height=2)
btnQuit.config(font=('Arial',7,'bold'))
btnQuit.grid(column=0,row=7,pady=10)

window.bind('<Return>', (lambda e, btnNext=btnNext: btnNext.invoke())) # mapping enter key to scan new barcode setting btnNext to default
#-Intializing ---------------------------------------------------------------------------------------------------------------------------

fileaddress=openFile()
window.deiconify()
excel=DoExcel(os.path.dirname(fileaddress))
print(excel.loadFile(os.path.basename(fileaddress)),' Loaded')
BarCode,Found,Excess,Short=excel.writeColumnHeaders(3,"Bar Code","Found","Excess","Short")
#excel.displayColumns(3)
lstCodeCols=excel.mapColsList(3,[8,9])
ItemTotal=excel.mapColsList(3,[11])
window.focus_force()
txtCode.focus()
#----------------------------------------------------------------------------------------------------------------------------------------
window.mainloop()
