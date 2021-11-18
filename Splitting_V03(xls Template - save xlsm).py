import pandas as pd
import numpy as np
from tkinter import *
import os.path
import xlwings as xw
####################################### Template Listing ######################################### 
def templateList(): #generate the list of files available     
    templates = os.listdir() 
    global template_list
    template_list = []
    for file in templates:
        if str.upper(file).startswith('TEMPLATE'):
            print('Template encontrado: ',file)
            template_list.append(file)
####################################### folder selection ######################################### 
def filesList(): #generate the list of files available     
    files = os.listdir() 
    global files_xls
    files_xls = []
    for file in files:
        if (file.endswith('.xls') | file.endswith('.xlsx') | file.endswith('.xlsm')) & (file not in template_list):
            print('Arquivo csv encontrado: ',file)
            files_xls.append(file)  
####################################### load and upload cols ######################################### 
def loadFile(file, skip_lines):
    global df,cols_of_df
    df = pd.read_excel(file, skiprows=skip_lines)
    cols_of_df = list(df.columns)
    
def to_excel_as_template(df_to_be_append,template_file, output_name):
    excel_app = xw.App(visible=False)
    global wb
    wb = xw.Book(template_file) ### KEY ITEM FOR TEMPLATE ###
    
    wb.sheets[0].range(2,1).options(index=False, header=True ,parse_dates=True, decimal='.').value = df_to_be_append
    	#wb.sheets[sheet index].range(row index, column number 'start in 1')

    wb.save(output_name)
    wb.close()
####################################### Exporting and repairing file name ######################################### 
def multipleExcel(col,sel_template,folder):
    global path, new_path, error_log,file_name
    error_log= ['',("Error found while splitting the data by: '" +(col)+"'"),'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - -']
    path = os.path.abspath('')
    new_path = os.path.join(path, folder)  
    try:
        os.mkdir(new_path)
    except OSError:
        print ("Creation of the directory %s failed" % new_path)
    else:
        print ("Successfully created the directory %s" % new_path)
    
    df[col] = df[col].astype(str)
    
    uniques = sorted(list(df[selected_col].unique()))
    print('==> Generating files:')
    
    for name in uniques:
        try:
            file_name = str(name[0:29]+'.xlsm')
            df_name = df.loc[df[col]==name]
            #df_name.to_excel(os.path.join(new_path,file_name), index=False)
            to_excel_as_template(df_name, sel_template, os.path.join(new_path,file_name))
            print('        ',file_name)
        except Exception:
            print('Error found while splitting! \n  Please check error log file to know those keys not created.')
            fileNameCheck_def(name)
            df_name = df.loc[df[col]==name]
            #df_name.to_excel(os.path.join(new_path,filename_checked), index=False)
            to_excel_as_template(df_name, sel_template, os.path.join(new_path,filename_checked))
            error_log.append(name+" - generated changing not acceptable character by ~")
            pass
    
    error_log_name = os.path.join(new_path,str('_Error_Log(-'+col+'-).txt'))
    f=open(error_log_name,'w')
    s1='\n'.join(error_log)
    f.write(s1)
    f.close()    

def fileNameCheck_def(filename):
    global filename_checked
    not_permited = []
    for ele in str("\/:*?<>"): not_permited.append(ele)
    
    filename_checked= ''
    for i in filename:
        if i in not_permited: filename_checked += '~'
        else: filename_checked +=i
    filename_checked +='.xlsm'
    print('File saved as: '+filename_checked)
####################################### GUI App #########################################    
###################################### Main code ######################################### 

templateList() ; filesList() # templateList() need to enter before fileList()

folder_init=""
path = os.path.abspath(folder_init) 

root = Tk()
root.title('Splitting DF by Column -DDS.V01-')
root.geometry("350x640+550+80") # window width x window height + X coordinate + Y coordinate
root.resizable(False, False)

def select_file():
    global selected_file
    selected_file = my_listbox1.get(ANCHOR)
    h_lines_output = int(h_lines.get())
    loadFile(selected_file, h_lines_output)
    my_listbox2.delete(0,END)
    for c in cols_of_df:
        my_listbox2.insert(END, c)

def splitting_df():
    global selected_col , input_name, headers_line, folder_name, selected_template
    
    folder_name = str(f_name.get())
    selected_col = my_listbox2.get(ANCHOR)
    selected_template = my_listbox3.get(ANCHOR)
    multipleExcel(selected_col,selected_template,folder_name)
    root.destroy()
    
def close_window():
    root.destroy()

def callback(url):
    webbrowser.open_new(url)   
# - - - - -Top text!
toptext = Label(root,text='Bem Vindo!! \n'+
              'Selecione o arquivo a ser Carregado:',fg= '#a30000' ,bg="#fcbbcd", width=100)
toptext.pack()
# - - - - - Listbox folders
my_listbox1 = Listbox(root,width=40,height=7, borderwidth=3,fg='#1500ff')
my_listbox1.pack(pady=5)
for item in files_xls:
    my_listbox1.insert(END, item)
# - - - - -text 1 + headers line
text1 = Label(root,text="Digite a quantidade de linhas a serem desconsideradas:")
text1.pack()
h_lines = Spinbox(root, width=7, borderwidth=3,fg='#1500ff',justify=CENTER, from_=0, to=25)
h_lines.pack()
# - - - - - load/update cols
loaddf_button=Button(root, text="Carregar Dados", command=select_file)
loaddf_button.pack(pady=5)  
# - - - - -text 2 + input file name
text2 = Label(root,text="Selecione a 'Coluna Filtro':")
text2.pack()
# - - - - - load/update cols
my_listbox2 = Listbox(root,width=40,height=7, borderwidth=3,fg='#1500ff')
my_listbox2.pack(pady=5)
# - - - - -text 3 + Template select Name
text3 = Label(root,text="Escolha o arquivo 'Template':")
text3.pack()
my_listbox3 = Listbox(root,width=40,height=5, borderwidth=3,bg='#dbdbdb')
my_listbox3.pack(pady=0)
for item in template_list :
    my_listbox3.insert(END, item)
# - - - - -text 4 + Folder Name
text4 = Label(root,text="Digite o nome da nova pasta:")
text4.pack()
f_name = Entry(root, width=40, borderwidth=3,fg='#1500ff')
f_name.insert(0,"Nova Pasta")
f_name.pack()
# - - - - - My Button + Close Button
my_button=Button(root, text="  Fragmentar Dados ✅  ", command=splitting_df)
my_button.pack(pady=10)
close_Button = Button(root, text=' Fechar ', command=close_window)
close_Button.pack(pady=5)
watermark = Label(root,text="Created by DDS", justify=RIGHT, fg='#00bbff' ) ;watermark.place(x=260, y=620)
root.mainloop()
    
######## END ########
#### Created By Danilo Uzelin da Silva ####

