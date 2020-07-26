from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import os

ColumnName={1:"A",2:"B",3:"C",4:"D",5:"E",6:"F",7:"G",8:"H",9:"I",10:"J",11:"K",12:"L",13:"M",14:"N",15:"O",16:"P",17:"Q",18:"R",19:"S",20:"T",21:"U",22:"V",23:"W",24:"X",25:"Y",26:"Z"}
#------------------------------------------------
# drawing window
window=Tk()
window.geometry("500x400")

windowWidth = window.winfo_reqwidth()
windowHeight = window.winfo_reqheight()

# Gets both half the screen width/height and window width/height
positionRight = int(window.winfo_screenwidth()/2.5 - windowWidth)
positionDown = int(window.winfo_screenheight()/2.5 - windowHeight)

# Positions the window in the center of the screen.
window.geometry("+{}+{}".format(positionRight, positionDown))
window.config(bg='lightcyan4')

window.title("Stock Check")
#window.wm_iconbitmap(os.path.join(foldPath,"FindIcon.ico"))
#--------- Code ------------------------------------------------------------------------------------------------
fPath=os.path.join(os.getcwd(),'Apparels EOSS.xlsx')
    
def LoadExcel(filePath):
    
    from openpyxl import load_workbook # importing workbook module
    wb=load_workbook(fPath) #loading workbook in memory
    ws=wb.active # setting active worksheet
    window.title("Stock Check : "+os.path.basename(fPath).split('.')[0])

LoadExcel(fPath)

#---------------------------------------------------------------------------------------------------------------
##      finding barcode in excel
##      first open excel
##      set code columns to item code and additional code
##      get last column number for heading insertion
##      insert heading to "Found" and "Balance"
        
        
#   setting heading for column
#   incrementing the total and balance with qty
#   displaying the discount
#   saving file

#---------------------------------------------------------------------------------------------------------------

# Interface
#---------------------------------------------------------------------------------------------------------------

s=ttk.Style()
s.configure('My.TFrame',background='lightcyan4')

content=ttk.Frame(window,padding=25,style='My.TFrame') #-- content for scan items (c-0,r-0)
mainfrm=ttk.Frame(content,borderwidth=5,relief='sunken',width=400,height=500,style='My.TFrame')
content.grid(column=0,row=0)
mainfrm.grid(column=0,row=0)

txtCode=Entry(mainfrm,width=30) #-- Text field to scan barcode (c-0,r-0)
txtCode.grid(column=0,row=0,padx=15)
txtCode.focus_set()

btnNext=Button(mainfrm,text="Next barcode",width=20,height=2)#-- default button for next barcode (c-1,r-0)
btnNext.config(font=('Arial',7,'bold'))
btnNext.grid(column=1,row=0,padx=10,pady=12)

lblLastScan=Label(mainfrm,text="Last Bar code Scanned: ",bg='lightcyan4',fg='white',width=22) #-- Label for sacnned barcode (c-1,r-1)
lblLastScan.config(font=('Courier',10,'bold'))
lblLastScan.grid(column=0,row=1,pady=10)

lblScanCode=Label(mainfrm,text="8907233000000",bg='lightcyan4',fg='white') #-- Last barcode scanned (c-1,r-1)
lblScanCode.config(font=('Courier',10,'bold'))
lblScanCode.grid(column=1,row=1)

separate=ttk.Separator(mainfrm,orient='horizontal')#-- separator between text field and data display (r-2)
separate.grid(sticky=(E,W),row=2,columnspan=10,padx=2, pady=5)

lblScanCount=Label(mainfrm,text="Scan Count : ",bg='lightcyan4',fg='white') #-- (c-0,r-4)
lblScanCount.config(font=('Courier',10,'bold'))
lblScanCount.grid(column=0,row=4)

lblScanCounter=Label(mainfrm,text="5",bg='lightcyan4',fg='white')#-- (c-1,r-4s)
lblScanCounter.config(font=('Courier',10,'bold'))
lblScanCounter.grid(column=1,row=4)

lblFound=Label(mainfrm,text="Found count : ",bg='lightcyan4',fg='white')#-- Label for found (c-0,r-5)
lblFound.config(font=('Courier',10,'bold'))
lblFound.grid(column=0,row=5)

lblFoundCount=Label(mainfrm,text="10",bg='lightcyan4',fg='white') #-- (c-1,r-5)
lblFoundCount.config(font=('Courier',10,'bold'))
lblFoundCount.grid(column=1,row=5)

lblPrice=Label(mainfrm,text="MRP : ",bg='silver',fg='black')#-- Label for found (c-0,r-6)
lblPrice.config(font=('Courier',10,'bold'))
lblPrice.grid(column=0,row=6)

lblDiscount=Label(mainfrm,text=" %",bg='silver',fg='black') #-- (c-1,r-6)
lblDiscount.config(font=('Courier',10,'bold'))
lblDiscount.grid(column=1,row=6)

btnSave=Button(window,text="Save File",relief='raised',borderwidth=2,width=20,height=2)#-- Button to save the excel sheet (c-0,r-1)
btnSave.config(font=('Arial',7,'bold'))
btnSave.grid(column=0,row=6)

btnQuit=Button(window,text="Quit",padx=5,width=20,height=2)#-- button to save file and quit app (c-0,r-3)
btnQuit.config(font=('Arial',7,'bold'))
btnQuit.grid(column=0,row=7,pady=10)

#---------------------------------------------------------------------------------------------------------------

#window.bind('<Return>', (lambda e, btnNext=btnNext: btnNext.invoke())) # mapping enter key to scan new barcode setting btnNext to default

window.mainloop()
