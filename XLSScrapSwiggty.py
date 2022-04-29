import pandas as pd
import re
from tkinter import *
from tkinter import messagebox
from tkinter import messagebox
import os
from os.path import exists




#from pyparsing import empty

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
    
    

    f = open("Email.csv", "a",encoding='utf-8')
    
    for row in df.itertuples(index=False,name='eachrow'):
        #print(row)
        tuple1=row
        #print(tuple1)
        ii = '[]'
        
            
        for item in tuple1:
            it=item
            ii = ii +"[]"+str(it)
        
        #ii is now row
        x=re.search(".*POC:", ii)
        y=re.search(".*Phone:", ii)
        z=re.search(".*E-mail:",ii)

        
        '''if x:
            rowsting=x.string
            splt=rowsting.split(",")
            #print(splt)
            for i in splt:
                x1=re.search(".*POC:", i)
                if x1 is None:
                    splt.remove(i)

            if len(splt)==1:      
                for i in splt:
                    if len(re.findall("POC:",i))>1:
                        i=str(i).split("POC:")
                        person1=i[1]
                        person2=i[2]
                        
                        #person1=person1.replace("[]"," ")
                        person1=re.sub("[][]+", "[]", person1)
                        person1=person1.split("[]")
                        if len(person1)>1:
                            person1=person1[1]
                            f.write(str(person1)+",")
                        else:
                            person1=""
                            f.write(str(person1)+",")

                        
                        

                        

                        
                        person2=re.sub("[][]+", "[]", person2)
                        person2=person2.split("[]")
                        if len(person2)>1:
                            person2=person2[1]
                            f.write(str(person2))
                        else:
                            person2=""
                            f.write(str(person2))


                        
                        
                        
                        
                        
                        

                                



            else:
                if splt!=None:
                    person1=splt[0]
                    person1=str(person1).split("POC:")
                    person1=person1[1]
                    person1=re.sub("[][]+", "[]", person1)
                    person1=person1.replace("[]","")
                    #print(person1)
                    
                    

                    person2=splt[1]
                    person2=str(person2).split("POC:")
                    person2=person2[1]
                    person2=re.sub("[][]+", "[]", person2)
                    person2=person2.replace("[]","")

                    
                    f.write(str(person1)+","+str(person2))
                    
            f.write("\n")'''
                    



    
        '''if y:
            rowsting_phone=y.string
            splt_phone=rowsting_phone.split(":")
            person1_phone=splt_phone[1]
            person1_phone=re.sub("[][]+", "[]", person1_phone)
            person1_phone=person1_phone.split('[]')
            if len(person1_phone)>1:
                person1_phone=person1_phone[1]
                
            else:
                person1_phone=" "

            #print(str(person1_phone))

            if len(splt_phone)>2:
                person2_phone=splt_phone[2]
                person2_phone=re.sub("[][]+", "[]", person2_phone)
                person2_phone=person2_phone.split('[]')
                if len(person2_phone)>1:
                    person2_phone=person2_phone[1]
                else:
                    person2_phone=" "

            else:
                person2_phone=" "

            
            f.write(str(person1_phone)+","+str(person2_phone))
            f.write("\n")'''
            

        
                    



        
        if z:
            rowsting_email=z.string
            splt_email=rowsting_email.split(":")
            person1_email=splt_email[1]
            person1_email=re.sub("[][]+", "[]", person1_email)
            person1_email=person1_email.split('[]')
            if len(person1_email)>1:
                person1_email=person1_email[1]
                
            else:
                person1_email=" "

            #print(str(person1_email))

            if len(splt_email)>2:
                person2_email=splt_email[2]
                person2_email=re.sub("[][]+", "[]", person2_email)
                person2_email=person2_email.split('[]')
                if len(person2_email)>1:
                    person2_email=person2_email[1]
                else:
                    person2_email=" "

            else:
                person2_email=" "
            
            f.write(str(person1_email)+","+str(person2_email))
            f.write("\n")

            
               

            

                
                

                    
    
    f.close()
    



        


#GUI Generate SQL Button

GenerateSQLFilebtn = Button(root, text = "Generate",command=GenerateSQL,**browsebtn_style)
GenerateSQLFilebtn.grid(row=4,column=2)


root.mainloop()











