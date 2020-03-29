# -*- coding: utf-8 -*-
"""
Created on Thu Oct  3 08:26:05 2019

@author: A328754
"""

import os
import pandas as pd
import tkinter as tk
import tkinter.messagebox
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter.messagebox

# =============================================================================
# GUI design
# =============================================================================

root = tk.Tk()
root.title('Forecast Adjustment Tool')
root.geometry('800x350+320+50')

f1 = tk.Frame(root, width = 700, height = 100)
f1.grid(row = 0, column = 0,columnspan = 3)
 
f2 = tk.Frame(root, width = 450, height = 100)
f2.grid(row = 1,column = 1,columnspan = 2)

f3 = tk.LabelFrame(root, width = 200, height = 150,text = 'Options')
f3.grid(row=1,column = 0,padx = 10)



#---------------------------------name-----------------------------------------
label_f1 = tk.Label(f1,text='Forecast analyser', font=('Verdana 20 bold'))
label_f1.grid(pady = 40)



#-------------------------------select data set---------------------------------

label_query = tk.Label(f2, text = 'Please upload first transaction demand data',font = ('Verdana 10 bold'), width = 50)
label_query.grid(row = 0, column = 1,pady = 5)

label_query2 = tk.Label(f2, text = 'Please upload second transaction demand data',font = ('Verdana 10 bold'), width = 50)
label_query2.grid(row = 1, column = 1,pady = 5)

label_query3 = tk.Label(f2, text = 'Please upload MMI Forecast data',font = ('Verdana 10 bold'), width = 50)
label_query3.grid(row = 2, column = 1,pady = 5)

label_query4 = tk.Label(f2, text = 'Make sure there exists item in PC and VAU class',font = ('Verdana 10 bold'), bg=  'yellow', width = 50)
label_query4.grid(row = 5, column = 1,pady = 5)

#------------------------------upload first transdemand------------------------
transdemand_file = pd.DataFrame()

def upload_first_trandsdemand_file():
    
    #build window
    window=tk.Tk()

    #set title
    window.title('select file')

    window.geometry('600x300+200+200')


    label1 = tk.Label(window,text ='Please select the file:')
    text1 = tk.Entry(window, bg = 'white', width = 55)

    
    def selectExcelfile():
        global filename1
        filename1 = filedialog.askopenfilename(title = 'select Excel file', filetypes = [('Excel', '*.xlsx'), ('All Files', '*')])
        text1.insert(INSERT, filename1)

    def closeThisWindow():
        window.destroy()
    
    def makesure():
        window.destroy()
        

    
    
    button1 = tk.Button(window, text = 'overview',width = 8, command = selectExcelfile)

    button2 = tk.Button(window, text = 'No',width = 8, command = closeThisWindow)

    button3 = tk.Button(window, text = 'Yes',width = 8, command = makesure)

    label1.place(x = 30,y = 30)
    text1.place(x = 100,y = 30)
    button1.place(x = 500,y = 26)
    button2.place(x = 160,y = 80)
    button3.place(x = 300, y = 80)
 
    window.mainloop() 

#-------------------upload second transdemand window---------------------------
transdemand_file = pd.DataFrame()

def upload_second_trandsdemand_file():
    
    #build window
    window=tk.Tk()

    #set title
    window.title('select file')

    window.geometry('600x300+200+200')


    label1 = tk.Label(window,text ='Please select the file:')
    text1 = tk.Entry(window, bg = 'white', width = 55)

    
    def selectExcelfile():
        global filename2
        filename2 = filedialog.askopenfilename(title = 'select Excel file', filetypes = [('Excel', '*.xlsx'), ('All Files', '*')])
        text1.insert(INSERT, filename2)

    def closeThisWindow():
        window.destroy()
    
    def makesure():
        window.destroy()
        

    
    
    button1 = tk.Button(window, text = 'overview',width = 8, command = selectExcelfile)

    button2 = tk.Button(window, text = 'No',width = 8, command = closeThisWindow)

    button3 = tk.Button(window, text = 'Yes',width = 8, command = makesure)

    label1.place(x = 30,y = 30)
    text1.place(x = 100,y = 30)
    button1.place(x = 500,y = 26)
    button2.place(x = 160,y = 80)
    button3.place(x = 300, y = 80)
 
    window.mainloop() 

#-------------------------------upload MMI window------------------------------
MMI_file = pd.DataFrame()

def upload_MMI_file():
    
    
    #build window
    window=tk.Tk()

    #set title
    window.title('select file')

    window.geometry('600x300+200+200')


    label1 = tk.Label(window,text ='Please select the file:')
    text1 = tk.Entry(window, bg = 'white', width = 55)

    
    def selectExcelfile():
        global filename_MMI
        filename_MMI = filedialog.askopenfilename(title = 'select txt file', filetypes = [('txt', '*.txt'), ('All Files', '*')])
        text1.insert(INSERT, filename_MMI)

    def closeThisWindow():
        window.destroy()
    
    def makesure():
        window.destroy()
        

    
    
    button1 = tk.Button(window, text = 'overview',width = 8, command = selectExcelfile)

    button2 = tk.Button(window, text = 'No',width = 8, command = closeThisWindow)

    button3 = tk.Button(window, text = 'Yes',width = 8, command = makesure)

    label1.place(x = 30,y = 30)
    text1.place(x = 100,y = 30)
    button1.place(x = 500,y = 26)
    button2.place(x = 160,y = 80)
    button3.place(x = 300, y = 80)
 
    window.mainloop() 


button1 = tk.Button(f2, text = 'upload', width = 8, command = upload_first_trandsdemand_file)
button1.grid(row = 0, column = 3, pady = 5, padx = 10)

button2 = tk.Button(f2, text = 'upload', width = 8, command = upload_second_trandsdemand_file)
button2.grid(row = 1, column = 3, pady = 5, padx = 10)

button3 = tk.Button(f2, text = 'upload', width = 8, command = upload_MMI_file)
button3.grid(row = 2, column = 3, pady = 5, padx = 10)


# =============================================================================
# option buttom area
# =============================================================================
# picks class
label_lsp = tk.Label(f3,text = 'Picks class：')
label_lsp.grid(row = 0,column = 0,padx = 10)

var_lsp = tk.IntVar()
com1 = ttk.Combobox(f3, textvariable = var_lsp, width = 7, state = "readonly") 
com1.grid(row = 0,column = 1,padx = 10,pady = 10) 
com1["value"] = ('0', '1', '2', '3', '4', '5', '6', '7', '8')
com1.current(0)

# value class
label_pro = tk.Label(f3,text = 'Value class：')
label_pro.grid(row = 1,column = 0,padx = 10)

var_pro = tk.StringVar()
com2 = ttk.Combobox(f3, textvariable = var_pro, width = 7, state = "readonly") 
com2.grid(row = 1, column = 1, padx = 10, pady = 10) 
com2["value"] = ('A1', 'A2', 'A3', 'A4', 'B1', 'B2', 'B3', 'C1', 'C2')
com2.current(0)

# warehouse code
label_wareh = tk.Label(f3,text = 'Warehouse code：')
label_wareh.grid(row = 2,column = 0,padx = 10)

var_wareh = tk.StringVar()
com3 = ttk.Combobox(f3, textvariable = var_wareh, width = 7, state = "readonly") 
com3.grid(row = 2, column = 1, padx = 10, pady = 10) 
com3["value"] = ('VOLCN', 'VIPIN', 'VOLEA', 'VOLKR', 'VOLPE', 'PDCZA', 'VOLME', 'VLABR')
com3.current(0)

def get_gridsearch_result():
    global res   
    mmi_file = filename_MMI 
    transdemand_file1 = filename1
    transdemand_file2 = filename2
    warehouse = var_wareh.get()
    pclass = var_lsp.get()
    value = var_pro.get()
    res = grid_search(mmi_file, warehouse, transdemand_file1, transdemand_file2, pclass, value)
    res.to_excel('res.xlsx')
    tk.messagebox.showinfo('Sweet notice','The file is downloaded')

button3 = tk.Button(f2, text = 'I want result', width = 10, command = get_gridsearch_result)
button3.grid(row = 5, column = 3, pady = 5, padx = 10)

root.lift()
root.attributes('-topmost',True)
root.after_idle(root.attributes,'-topmost',False)
root.mainloop()
















