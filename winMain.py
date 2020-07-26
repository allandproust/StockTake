from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os

#-- INIT variables ----------------------------------------------------------------------------------------------
fpath=os.getcwd()

#---Functions ----------------------------------------------------------------------------------------------------

def openFile(self):
    file=filedialog.askopenfilename(initialdir=fpath,title='Select file for operations',filetypes=(('EXCEL FILES','*.xlsx'),('CSV FILES','*.csv')))
    openNext(file)
    
def openNext(self,fileName):
    status.set("File Loaded : " +os.path.basename(fileName))
    return fileName
    
def mainFunction():
    #-- Window configuration -----------------------------------------------------------------------------------------
    win=Tk()
    win.title('Load File')

    posX=round(win.winfo_screenwidth()/2)-win.winfo_reqwidth()
    posY=round(win.winfo_screenheight()/2)-win.winfo_reqheight()
    win.resizable(False,False)
    win.geometry('600x300+{}+{}'.format(posX,posY))
    win.grid_columnconfigure(0,weight=1)

    style=ttk.Style()
    style.configure('white.TLabel',font=('Arial',10,'bold'))
    style.configure('white.TLabel',background='white')
   
    #-- Interface ----------------------------------------------------------------------------------------------------

    status=StringVar()
    status.set(" File Loaded : ")
    lblFile=ttk.Label(win,textvariable=status,style='white.TLabel').grid(row=0,column=0,padx=2,pady=10,columnspan=2,sticky='EW')

    content=Frame(win,borderwidth=2).grid(column=0,row=1,columnspan=6)

    frmChoice=ttk.LabelFrame(content,text=' Choose a option ',borderwidth=2,relief='groove',width=300,height=100,)
    frmChoice.grid(column=0,row=1,columnspan=6,rowspan=5,padx=20,pady=20,sticky=(N,S,E,W))#

    choice=StringVar()

    rStockTake=ttk.Radiobutton(frmChoice,text='Stock Take',variable=choice,width=50,value=1).grid(column=0,padx=10)
    rStockTake=ttk.Radiobutton(frmChoice,text='Check Discount',variable=choice,width=50,value=2).grid(column=0,padx=10)
    rStockTake=ttk.Radiobutton(frmChoice,text='Fetch Stock',variable=choice,width=50,value=3).grid(column=0,padx=10)
    choice.set(1)

    divider=ttk.Separator(win,orient=HORIZONTAL).grid(column=0,columnspan=6,sticky="ew")

    btnOpen=ttk.Button(win,text="  Open File  ",command=openFile).grid(column=2)
    #-------------------------------------------------------------------------------------------------------------------
    win.mainloop()

if __name__ == "__main__": 
    mainFunction() 
