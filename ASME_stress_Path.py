###------- this macro works for the files inloacation C:\Users\stephen.arput\Documents\RESULTS\LOC222\LOC2--
###-------file path customization needs to be done...in later version...
####------- Error handling has to be taken in the later version.....

from tkinter import filedialog
from PIL import ImageTk,Image
import PIL.Image
import os
import io
import pandas as pd
from tqdm import tqdm
from tkinter import *
from tkinter import ttk
from execute_macro import execute
root = Tk()
root.geometry("1600x1600+20+40")
root['bg']='SlateGray3'
root['bd']= 3
# image resizing


##img = Image.open("C:/Users/arpuste/Downloads/Post-processing-Tool-main/path.png")
img = PIL.Image.open("C:/Users/arpuste/Downloads/Post-processing-Tool-main/path.png")
resize=img.resize((400,300))

new_img = ImageTk.PhotoImage(resize)

#panel = Label(root, image = new_img, height = 400, width = 300)
panel = Label(root, image = new_img)
panel.place(x=1000, y=350)
#oldlace'azure'
root.title('ASME DIV 2 Based Verification Criteria  ')
Label(root,text = "Path Operation Postprocessing Tool v1.0 ", bg="#116562",height ="2",\
      width = "800", fg ="white",
      font = ("Calibri",40)).pack()
Label(root,text = "Note: This tool is specific to project requirement, read the requirement\
  of this tool for process and input data. Files with binary excel format not supported",
      height ="3",
      width = "400",
      font = ("Calibri",12)).pack()
print("entering post processing module")

##---progress bar
#progbar = ttk.Progressbar(root, orient= HORIZONTAL,length= 220, mode="indeterminate")
#progbar.place(x=500, y=340)
##progbar.pack(pady =20)
#progbar.start()
Label(root, text="Enter details :",
      height ="3",
      width = "400",
      font = ("Calibri",12)).pack()
#label1.grid(row=1)

def validate(U_input):
    st=U_input.isdigit()
    if(not st):
        button1.config(state='disabled')
    else:
        button1.config(state='normal')
    return st

my_valid=root.register(validate)
e1=Entry(root,validate='focus',validatecommand=(my_valid,'%P'), width =20, font=('Arial 14'),borderwidth = 5)
e1.pack(padx=10, pady=10)
Label(root, text="Yield Strength :",font = ("Calibri",13)).place(x = 510, y = 275)

e2=Entry(root, width =20, font=('Arial 14'),borderwidth = 5)
e2.pack(padx=10, pady=10)
Label(root, text="Ultimate Strength :",font = ("Calibri",13)).place(x = 510, y = 330)

e3=Entry(root, width =20, font=('Arial 14'),borderwidth = 5)
e3.pack(padx=10, pady=10)
Label(root, text="  Sm   :",font = ("Calibri",13)).place(x = 510, y = 380)


print("getting path from string")


global frame
def clearframe():
    print("entering clear function")
        #my_label1.pack_forget()
    root.my_label1.destroy()

def sel():
   selection = "You selected the option " + str(var.get())
   
   label.config(text = selection)
   YS= e1.get()
   UTS= e2.get()
   SM= e3.get()
   print(YS)
   print(UTS)
   print(SM)
   return(YS,UTS,SM)


var = IntVar()
R1 = Radiobutton(root, text="Hydro-Test condition",font = ("Calibri",13), variable=var, value=1,command=sel,bg='lightgrey')
#bg='#0052cc'

R1.pack(side="top", anchor = N ,pady = 10)

R2 = Radiobutton(root, text="   Design condition    ",font = ("Calibri",13), variable=var, value=2,command=sel,bg='lightgrey')
R2.pack(side="top", anchor = N, pady = 10 )

R3 = Radiobutton(root, text="          Option 3          ",font = ("Calibri",13), variable=var, value=3,command=sel,bg='lightgrey')
R3.pack(side="top", anchor = N, pady = 10)

def openfile():
    global fpath,YS,UTS,SM
    if True:
      print("Enter the values and open file")
       #break;
    fpath = filedialog.askopenfilename()
    location=fpath
    frame = Frame(root, width=300, height=50,highlightthickness=2)
    frame.place(x=60, y=375)
    my_label1 = Label(frame, text=location,font=("Arial", 12)).pack()
    print("file read from the path")
    return(fpath)

global fpath,YS,UTS,SM
def execute():
    global my_label1
    global fpath,YS,UTS,SM
    greet=openfile
    frame = Frame(root, width=300, height=150,highlightthickness=2)
    frame.place(x=300, y=525)
    my_label1 = Label(frame, text="Output File saved to path",font=("Arial", 12)).pack()
    df2 = pd.read_csv(fpath, skiprows=1)
    df2 = df2[df2.filter(regex='^(?!Unnamed)').columns]
    print(df2.shape)
    print(df2.columns)
    #execute(fpath)
    print(fpath)
    print("entered EXECUTION module")


#!!!!! df2 is dataframe for Hydro test conditions that needs to be evaluated based on Hydro testloads obtained from ansys output!!!!
    workingdata=df2.copy()
    df3=df2.copy()
    workingdata.drop(workingdata.columns[[0, 1]], axis = 1, inplace = True)

#!!!!!  max(axis=1) method searches column-wise and returns the maximum value for each row.

    max_element = workingdata.max(axis=1)
    print(max_element)

    df2['max value']=max_element
    print(df2.shape)
    YS,UTS,SM =sel()
    print("Yield strength value =",YS)
    print("Ultimate Strength value=",UTS)
    print("mean stress value sm=",SM)
#!! yield strength value is 600 Mpa !!!

#!!!!!!!!df2["Hydro_mem"]=df2[' Membrane ']/(0.67*600) 


#!!!!!!!!!! condition to meet the Hydro-test condition as per ASME VIII DIV2 part5 for any path opeartions!!!
    dff2=[]
    for i in range (0,df2[' Membrane '].count()):
      if (df2[' Membrane '][i])<=(0.67*float(YS)):
        dff2.append(df2["max value"][i]/(1.4*float(YS)))
      elif(df2[' Membrane '][i])<=(0.95*float(YS)):
        dff2.append(df2["max value"][i]/(2.43*float(YS)-(df2[' Membrane '][i]*1.5)))
      else:
        dff2.append("Failed")

    HTresult=pd.DataFrame(dff2)

    df2["HT result"]=HTresult
    df2.loc[df2['HT result'] <= 1, 'Test case'] = '< 0.67*Sy'
    df2.loc[df2['HT result'] > 1, 'Test case'] = ' Between 0.67*Sy & 0.95*Sy'
    df2.loc[df2['HT result'] > 1.46, 'Test case'] = ' Failed'
    print(" Hydro test case calculation complete !!")
    df2.to_excel('HydroTest_Check1_output.xlsx')

def prbar():
    progbar = ttk.Progressbar(root, orient=HORIZONTAL, length=220, mode="indeterminate")
    progbar.place(x=500, y=340)
    progbar.start()
    #progbar.stop()

def group():

    grouped=collect(e.get())
    frame = Frame(root, width=300, height=150,highlightbackground="blue",highlightthickness=2)
    frame.place(x=500, y=375)
    my_label1 = Label(frame, text=grouped,font=("Arial", 12)).pack()
    #folderpath= "\n Entered path : "+ fpath
    #my_label1=Label(root,text=grouped).pack()
    #e.delete(0, END)

def prin_stress():
    global my_label1

    stress_results=principal(e.get())
    frame = Frame(root, width=300, height=150,highlightbackground="blue",highlightthickness=2)
    frame.place(x=500, y=525)

    my_label1 = Label(frame, text=stress_results,font=("Arial", 12)).pack()
    #folderpath= "\n Entered path : "+ fpath
    #my_label1=Label(root,text=stress_results).pack()
    #e.delete(0, END)

#filepath = concat(fpath)



def close():
    root.destroy()

print("entering button controls")
#---------------------------------

button1 = Button(root,text = "  Open File", height ="2", width = "25",\
                 font = ("Calibri",13),bg="light grey",fg ="black", command = openfile,state= DISABLED)
button1.place(x = 60, y = 307)

#---------------------------------
button2 = Button(root,text = "  Execute ",height ="2", width = "25",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white", command =execute)
button2.place(x = 60, y = 520)
#--------------------------------
#--------------------------------
#---------------------------------
button6 = Button(root,text = " Close ",height ="2", width = "25",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white", command = close)
button6.place(x = 60, y = 600)
#--------------------------------
label = Label(root)
label.pack()
root.mainloop()

print("completed button controls")
root.state("zoomed")
root.mainloop()