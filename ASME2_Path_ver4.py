
###------- this macro works for the files inloacation C:\Users\stephen.arput\Documents\RESULTS\LOC222\LOC2--
###-------file path customization needs to be done...in later version...
####------- Error handling has to be taken in the later version.....
###### macro created by stephen.arputharajgerald@bakerhughes.com ####

from tkinter import filedialog
from PIL import ImageTk,Image
import PIL.Image
import os
import io
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl
import pandas as pd
from tkinter import *
from tkinter import ttk
root = Tk()
root.geometry("1600x1600+20+40")
#root.attributes('-fullscreen', True)
root['bg']='SlateGray3'
root['bd']= 3
# image resizing


## image for the front page in the tool ,img = Image.open("C:/Users/cprav/Projects/Macros/path.png")
img = PIL.Image.open("C:/Users/arpuste/Downloads/Post-processing-Tool-main/path.png")
resize=img.resize((400,300))

new_img = ImageTk.PhotoImage(resize)

#panel = Label(root, image = new_img, height = 400, width = 300)
panel = Label(root, image = new_img)
panel.place(x=1000, y=350)
#front page title
root.title('ASME DIV 2 Based Verification Criteria  ')
Label(root,text = "Path Operation Postprocessing Tool v2.0 ", bg="#116562",height ="2",\
      width = "800", fg ="white",
      font = ("Calibri",40)).pack()
Label(root,text = "Note: This tool is specific to project requirement, read the requirement\
  of this tool for process and input data.",
      height ="3",
      width = "400",
      font = ("Calibri",12)).pack()
print("entering post processing module")

##---progress bar

Label(root, text="",
      height ="3",
      width = "400",
      font = ("Calibri",12)).pack()

##---validating for the input in the entry box
def validate(U_input):
    #st=U_input.isalpha()
    st=isinstance(U_input, float)
    print(st)
    if(not st):
        button1.config(state='normal')
    else:
        button1.config(state='disabled')
    return st

my_valid=root.register(validate)

e1=Entry(root,validate='focus',validatecommand=(my_valid,'%P'), width =20, font=('Arial 14'),borderwidth = 5)
e1.pack(padx=10, pady=10)
Label(root, text="Yield Strength :",bg='SlateGray3',font = ("Calibri",14)).place(x = 460, y = 275)

#e2=Entry(root,validate='focus',validatecommand=(my_valid,'%P'), width =20, font=('Arial 14'),borderwidth = 5)
#e2.pack(padx=10, pady=10)
#Label(root, text="Ultimate Strength :",font = ("Calibri",13)).place(x = 460, y = 330)

e3=Entry(root,validate='focus',validatecommand=(my_valid,'%P'), width =20, font=('Arial 14'),borderwidth = 5)
e3.pack(padx=10, pady=10)
Label(root, text="Allowable Stress Sm:",bg='SlateGray3',font = ("Calibri",14)).place(x = 460, y = 330)

print("getting path from string")
global frame
def toplevel2():
    #path= 'C:/Users/arpuste/PycharmProjects/pythonProject1'
    #path = " C:\Users\ arpuste\PycharmProjects\pythonProject1 "
    ls = 'ASME2_readme.py'
    os.system(ls)
def clearframe():
    print("entering clear function")
        #my_label1.pack_forget()
    root.my_label1.destroy()

def sel():
   selection = "You selected the option " + str(var.get())
   
   label.config(text = selection)
   YS= float(e1.get())
   #UTS= e2.get()
   SM= float(e3.get())
   print(YS)
   #print(UTS)
   print(SM)
   return(YS,SM)


#var = StringVar()
var = IntVar()
R1 = Radiobutton(root, text="Hydro-Test condition: criteria 1",font = ("Calibri",14), variable=var, value=1,command=sel,bg='lightgrey')
#bg='#0052cc'

R1.pack(side="top", anchor = N ,pady = 10)

R2 = Radiobutton(root, text="   Design condition: criteria 1    ",font = ("Calibri",14), variable=var, value=2,command=sel,bg='lightgrey')
R2.pack(side="top", anchor = N, pady = 10 )

R3 = Radiobutton(root, text="   Design condition: criteria 2    ",font = ("Calibri",14), variable=var, value=3,command=sel,bg='lightgrey')
R3.pack(side="top", anchor = N, pady = 10)
global frame1
def openfile():
    global fpath,YS, frame1,SM

       #break;
    fpath = filedialog.askopenfilename()

    if os.path.isfile(fpath)==True:
      location= os.path.basename(fpath)
      print("File selected")
    else:
      print("File not selected        ")
      location= "File not selected"
    frame1 = Frame(root, width=300, height=50,bg='SlateGray3')
    frame1.place(x=60, y=400)
    my_label1 = Label(frame1, text=location,font=("Arial", 12)).pack()
    print("file OPEN Process completed")
    #my_label1=""
    return(fpath)
    

global fpath,YS,UTS,SM,dfr,filename1, sheetname
def execute():
    global my_label1,dfr,filename1, sheetname
    global fpath,YS,UTS,SM,frame
    greet=openfile

    df2 = pd.read_csv(fpath, skiprows=1)
    # skipping 1st row from the input sheet which has heading"Stress Linearisation results for all the paths"
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
    YS,SM = sel()
    print("Yield strength value =",YS)
    #print("Ultimate Strength value=",UTS)
    print("Allowable stress value sm=",SM)
    print(df2)
#!! yield strength value is 600 Mpa !!!
#allowable stress@ design temperature = minimum value of YS/1.5, UTS/2.4

#!!!!!!!!df2["Hydro_mem"]=df2[' Membrane ']/(0.67*600) 

    if var.get()==1:
      print( "Hydro Test condition calculation")

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
      frame = Frame(root, width=300, height=150,bg='SlateGray3')
      frame.place(x=300, y=525)
      my_label1 = Label(frame, text="Output File saved to path",font=("Arial", 12)).pack()
      
      print(" Hydro test case calculation complete !!")
      dfr=df2.copy()
      dfr=dfr[['Path Name',' Membrane ','max value','HT result']]
      dfr.rename(columns={'Path Name':'Path #',' Membrane ':'PM','max value':'PM+B','HT result':'HT Check'},inplace = True)
      dfr.to_excel('HydroTest_calculation.xlsx')

# Highlight the Max values in each column
      print("\nModified Stlying DataFrame:")
      def highlight_rows(row):
          value = row.loc['HT Check']
          if value <1 :
              color = '#BAFFC9'  # Green
          elif value > 1:
              color = '#FFB3BA' # Red
          else:
              color = '#FCFCFA' # Grey
          return ['background-color: {}'.format(color) for r in row]

      DF1=dfr.style.apply(highlight_rows,axis=1)

      #DF1.to_excel('hydro results.xlsx')
      n1 = len(dfr.index)+2
      sheetname = 'Hydro Test_result'
      title=" Hydraulic Test "
      Merge_Range='A1:D1'
      filename1 = "Hydro_condition_result.xlsx"


#!!!!!!!!!! condition to meet the design-test condition as per ASME VIII DIV2 part5 for any path opeartions!!!
    elif var.get()==2:
      print( "design Condition calculation")
      df4 =df2.copy()
      print(df4.shape)
      print(df4.columns)
      check1=[]
      for i in range (0,df4[' Membrane '].count()):
        check1.append((df4[' Membrane '][i])/float(SM))

      DesignCheck1=pd.DataFrame(check1)
      print("exporting file:")
      DesignCheck1.to_excel('design_check1.xlsx')
      check2=[]
      for i in range (0,df4['max value'].count()):
        check2.append(df4['max value'][i]/(1.5*float(SM)))
      DesignCheck2=pd.DataFrame(check2)
      df4["check #1"]=DesignCheck1
      df4["check #2"]=DesignCheck2
      dfr=df4.copy()
#### Same code from case1 copied below for formating ~~~~~~~~
      frame = Frame(root, width=300, height=150,bg='SlateGray3')
      frame.place(x=300, y=525)
      my_label1 = Label(frame, text="Output File saved to path",font=("Arial", 12)).pack()

# Highlight the Max values in each column
      print("\nModified Stlying DataFrame:")
      df4.style.highlight_max(axis=0)

      df4.to_excel('DesignTest_Check_output.xlsx')
      print(" Designtest case calculation complete !!")
#### Same code from case1  formating ~~~~~~~
      print(dfr.columns)
      dfr=dfr[['Path Name',' Membrane ','max value','check #1','check #2']]
      dfr.rename(columns={'Path Name':'Path #',' Membrane ':'PM','max value':'PM+B','check #1':'check #1 \n PM/SM<1','check #2':'check #2 \n (PM+B)/1,5*SM<1'},inplace = True)
      print("working in design calc")
      dfr.to_excel('design_results.xlsx')
# parameters for xlsx writer which is used for saving excel variables
      n1 = len(dfr.index)+2
      sheetname = 'design_result'
      title=" Design Cond. "
      Merge_Range='A1:E1'	
      filename1 = "Design_condition_result.xlsx"

    elif var.get()==3:
      print("entering principal stress calculation")
#!!!!!!!!!! condition to meet the Design -principal stress condition as per ASME VIII DIV2 part5 for any path opeartions!!!


      dfp=df2.copy()


      frame = Frame(root, width=300, height=150,bg='SlateGray3')
      frame.place(x=300, y=525)
      my_label1 = Label(frame, text="Output File saved to path",font=("Arial", 12)).pack()

      dfp.drop(dfp.columns[[0,4]], axis = 1, inplace = True)# !! removing 1st and 5th columns to get only numerical values
      df4=dfp.values.tolist()  #!! this gets the list of I,C, O values
      ser= pd.Series(df4)
      #ser.columns = 'values'
      dfff=pd.DataFrame(ser)
      dfff.columns =['stress_values']# this has no dataframe column name. now assigning it to this.
      sort_df =dfff.explode('stress_values') #!!!!! converts all values I,C,O to single column dataframe
      dfindex=df2['Path Name']
      sort_df['Path_Name']=dfindex
      print(sort_df)
      columns_titles = ['Path_Name','stress_values']
      df_reorder=sort_df.reindex(columns=columns_titles) # reordering the column names
      df_reorder.reset_index(inplace = True, drop = True)
      df_new=df_reorder.copy()
   
      df_new[['Principle Stress','Section','Path']] = df_new.Path_Name.str.split("_", expand=True)
      print(df_new)
      print(df_reorder)
      df_heading = df_new[['Path_Name']]

      ##!!!! slicing into different sections
      for i in range(1,4):
        globals()["df_"+str(i)] = df_reorder[df_reorder['Path_Name'].str.contains('S'+str(i))]
        globals()["df_"+str(i)].reset_index(inplace = True, drop = True)
      dfadd = pd.concat([df_1, df_2, df_3], axis = 'columns')
      print(dfadd)


      dfdata=dfadd.drop(['Path_Name'], axis=1)
      dfdata.set_axis(['S1', 'S2', 'S3'], axis='columns', inplace=True)
      print(dfdata)

      dfadd2 = pd.concat([df_heading,dfdata], axis = 'columns')
      dfadd3 =dfadd2[dfadd2['Path_Name'].str.contains('S1')]
      dfadd3 = dfadd3.Path_Name.str.split("_", expand=True)
      dfadd3.set_axis(['Principle Stress','Section','Path'], axis='columns', inplace=True)
      #dfadd3.drop(['Path_Name'], axis=1, inplace=True)
      print(dfadd3)

      n2 = len(dfdata.index)/3
      n2=int(n2+1)
      print(n2)
      #### index for the new dataframe"
      df = pd.DataFrame()
      c = []
      for i in range(1, n2):

          chars = ['I', 'C', 'O']
          for index, value in enumerate(chars):
              c.append(value)
      dfdata["position"] = c
      dfdata['Section'] = df_new[['Section']]

      dfdata = dfdata[['Section','position', 'S1', 'S2', 'S3']]
      dfdata['sum']=dfdata[['S1', 'S2', 'S3']].sum(axis=1)
      dfdata['Max_sum']=0
      block_size=3
      for i in range(0,len(dfdata),block_size):
           block_sum = dfdata.loc[i:i+block_size-1,'sum'].max()
           for j in range(i,i+block_size):
              dfdata.at[j,'Max_sum'] = block_sum
      dfr = pd.DataFrame()
      filename1 = "Design_principal_calc_ignore.xlsx"

      n1 = len(dfr.index)+2
      sheetname = 'Design_Principal_stress'
      title=" Design Cond. "
      Merge_Range='A1:E1'


      dfdata.set_index(['Section'],inplace= True)
      print(" principal stress condition calculation complete !!")

      #dfdata.style.format(precision=2)
      dfdata.style.format({
          "Max_Sum": "{:.2f}"})
      dfdata.to_excel('Design_Principal_stress_results.xlsx',sheet_name='Design_Principal_stress',engine='openpyxl')



#!!!!!!! End of principal stress calculation !!!

    else: 
      print("Other condition")

# Result dataframe to xlsx write module..!!!!
    writer = pd.ExcelWriter(filename1, engine='xlsxwriter')
    dfr.to_excel(writer,index=False, sheet_name=sheetname,startrow=1,startcol=0)

    workbook = writer.book
    worksheet = writer.sheets[sheetname]
    worksheet.set_zoom(90)

#!!!!!!!   Set header Format  !!!!!!!
    header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size':14,
            'border': 1}) 

#!!!!add title format
    format = workbook.add_format({'bold': True,'align': 'center','valign': 'vcenter','font_color':'black'})
    format.set_border(1)
    format.set_font_size(14)
   #format.set_font_color("black")

# number format
    num_format =workbook.add_format({'num_format': '0.00','border':1})
    #worksheet.set_column(Merge_Range,8,num_format)
    worksheet.merge_range(Merge_Range, title, format)
    worksheet.set_column('A:A',25)


# result number format
    rnum_format =workbook.add_format({'num_format': '0.00','align': 'center','valign': 'vcenter','font_size':11,'border':1})
    r2num_format =workbook.add_format({'align': 'center','valign': 'vcenter','font_size':11,'border':1})

    worksheet.set_row(0, 26) 
    worksheet.set_row(1, 43) # Set the header row height to 26,43
# puting it all together
# Write the column headers with the defined format.

    for col_num, value in enumerate(dfr.columns.values):
      print(col_num, value)
      worksheet.write(1, col_num, value, header_format)
      worksheet.set_column('A:A',25,r2num_format) # setting the column width
      worksheet.set_column('B:D',12,rnum_format) # setting the column width
      worksheet.set_column('E:E',22,rnum_format) # setting the column width
# Light green fill with bold text.
      lessthanone = workbook.add_format({'bold': True,'font_size':8,'bg_color':   '35FC03'})
      greaterthanone = workbook.add_format({'bold': True,'bg_color':   'red','font_color': '#FFC7CE'})


      #worksheet.conditional_format('D3:E50', {'type':     'cell','criteria':'between', 'minimum':  0.0001, 'maximum':  1,'format':   lessthanone})
      worksheet.conditional_format('D3:D'+str(n1), {'type':     'formula','criteria': '=$D3 >1','format':   greaterthanone})
      worksheet.conditional_format('D3:E'+str(n1), {'type':     'formula','criteria': '=$E3 >1','format':   greaterthanone})

      worksheet.conditional_format('E3:E'+str(n1), {'type':     'formula','criteria': '=AND($E3<1,$D3<$E3)','format':   lessthanone})
      worksheet.conditional_format('D3:D'+str(n1), {'type':     'formula','criteria': '=AND($D3<1,$E3<$D3)','format':   lessthanone})
     


# add borders
      full_border = workbook.add_format({'border' : 1, 'border_color': 'black'})
      worksheet.conditional_format('D3:E'+str(n1), {'type':  'formula','criteria': '=$D3<1','format':   full_border})
      worksheet.conditional_format('D3:E'+str(n1), {'type':  'formula','criteria': '=$D3>1','format':   full_border})
    writer.save()
  
global frame1
def clear():
    global frame1
    e1.delete(0 ,'end')
    e3.delete(0, 'end')
    for widgets in frame1.winfo_children():
        widgets.destroy()
    for widgets in frame.winfo_children():
        widgets.destroy()
    frame1.pack_forget()
    frame.pack_forget()

def close():
    root.destroy()

print("entering button controls")
#---------------------------------

button1 = Button(root,text = "  select input File (CSV)", height ="2", width = "25",\
                 font = ("Calibri",13),bg="light grey",fg ="black", command = openfile,state= DISABLED)
button1.place(x = 60, y = 335)

#---------------------------------
button2 = Button(root,text = "  Execute ",height ="1", width = "20",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white", command =execute)
button2.place(x = 60, y = 520)
#--------------------------------
#--------------------------------
#---------------------------------
readme = Button(root,text = "  READ ME ",height ="2", width = "25",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white", command =toplevel2)
readme.place(x = 60, y = 270)
#--------------------------------
#---------------------------------
button6 = Button(root,text = " Clear ",height ="1", width = "20",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white", command = clear)
button6.place(x = 60, y = 580)
#--------------------------------
#---------------------------------
button6 = Button(root,text = " Close ",height ="1", width = "20",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white", command = close)
button6.place(x = 60, y = 640)
#--------------------------------c
label = Label(root)
label.pack()
root.mainloop()

print("completed button controls")
root.state("zoomed")
root.mainloop()