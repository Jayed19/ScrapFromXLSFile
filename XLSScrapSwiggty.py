import pandas as pd
import re
from tkinter import *
from tkinter import messagebox
from tkinter import messagebox
import os

from pyparsing import empty

#GUI Window
root = Tk()
root.title("XLXS File TO SQL Converter")
root.geometry('700x650')
root.option_add("*Font", ('Arial', 10))



#GUI Browse XLX File
from tkinter import filedialog
filedir=''
def browseFiles():
    global filedir
    filedir = filedialog.askopenfilename(
                                          title = "Select XLXS File",
                                          filetypes = (("all files","*.*"),
                                                        ("XLXS files","*.xlxs*"),
                                                        ("XLX files","*.xlx*")
                                                       ))
    

 
    print(filedir)
    return filedir

browsebtn_style={'padx':5, 'pady':5,'borderwidth':5, 'relief':'raised','font':'Arial 9'}
BrowseFilebtn = Button(root, text = "Browse XLX File",command=browseFiles,**browsebtn_style)
BrowseFilebtn.grid(row=3,column=2)



#Run SQL Generation
def GenerateSQL():
    df = pd.read_excel(filedir,keep_default_na=False)
    

    f = open("test1.txt", "a",encoding='utf-8')
        #Row Finding
    for row in df.itertuples(index=False,name='eachrow'):
        #print(row)
        tuple1=row
        #print(tuple1)
        ii = '[rowstart]'

# Use for loop to convert tuple to string.
        for item in tuple1:
            it=item
            ii = ii +'[join]'+str(it)

            
        f.write(ii)
        f.write("\n")
    
    f.close()


        


#GUI Generate SQL Button

GenerateSQLFilebtn = Button(root, text = "Generate",command=GenerateSQL,**browsebtn_style)
GenerateSQLFilebtn.grid(row=4,column=2)


root.mainloop()











