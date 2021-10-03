#### Created By Danilo U da Silva ####
#### Automatização da rotina para gerar o relatório Daily Inventory ####
import pandas as pd
from tkinter import *
import os.path

####################################### folder selection ######################################### 
def folderSelect(folder): #generate the list of paths and files to be consolidated
    global path     
    path = os.path.abspath(folder) 
    files = os.listdir(path) 
    global files_xls
    files_xls = []
    for file in files:
        if file.endswith('.xls') | file.endswith('.xlsx') | file.endswith('.xlsm'):
            print('Arquivo csv encontrado: ',file)
            files_xls.append(file)
            
####################################### appending ######################################### 
def appendingFilesInFolder(listoffiles, skip_lines):
    global df1 , df2
    df1=pd.DataFrame()
    df2=pd.DataFrame()

    for file in listoffiles:
        file_in_folder = os.path.join(path,file)
        try:
        	df1 = df1.append(pd.read_excel(file_in_folder, sheet_name=0, skiprows=skip_lines))
        	df2 = df2.append(pd.read_excel(file_in_folder, sheet_name=1, skiprows=skip_lines))
        except:
        	print("Error found openning ",file)
        	pass
        
####################################### Exporting ######################################### 
def export_xls(dataframe1, dataframe2, file_name):
   
    output_name = str(file_name+'.xlsx')

    with pd.ExcelWriter(os.path.join(path,output_name)) as writer:
    	try:
    		dataframe1.to_excel(writer, sheet_name='CLT',index=False)
    	except:
    		print("Error found in sheet 0")
    		pass

    	try:
    		dataframe2.to_excel(writer, sheet_name='3s',index=False)
    	except:
    		print("Error found in sheet 1")
    		pass


#    dataframe.to_excel(os.path.join(path,output_name), index=False)

####################################### GUI App #########################################    
###################################### Main code ######################################### 
folder_init=""
path = os.path.abspath(folder_init) 
directory = os.listdir(path) 
list_dir= []

for i in directory:
    if os.path.isdir(i):
        list_dir.append(i)
    
root = Tk()
root.title('Join files of folde -DDS.V01-')
root.geometry("350x400+550+200") # window width x window height + X coordinate + Y coordinate
root.resizable(False, False)

def select_folder():
    global selected_folder , input_name, headers_line
    
    selected_folder = my_listbox.get(ANCHOR)
    input_name = name.get()
    h_lines_output = int(h_lines.get())

    folderSelect(selected_folder)
    print('Folder selected: ',selected_folder)
    
    appendingFilesInFolder(files_xls,h_lines_output)
    print('Files appended, main DF generated')
    print('Genetaring output file')
    export_xls(df1, df2, input_name)
    print('Ouput generated with success!!!')
    print('Script will be close!')
    root.destroy()


def close_window():
    root.destroy()

def callback(url):
    webbrowser.open_new(url)
    
# - - - - -Top text!
toptext = Label(root,text='Bem Vindo!! \n'+
              'Selecione a paste que contém os arquivos a serem consolidados:',fg= '#a30000' ,bg="#fcbbcd", width=100)
toptext.pack()

# - - - - - Listbox folders

scrollbar = Scrollbar(root, orient="vertical")
my_listbox = Listbox(root,width=40, borderwidth=3,fg='#1500ff')
my_listbox.pack(pady=5)


for item in list_dir:
    my_listbox.insert(END, item)

# - - - - -text 1 + headers line
text1 = Label(root,text="Digite a quantidade de linhas a serem desconsideradas:")
text1.pack()

h_lines = Entry(root, width=5, borderwidth=3,fg='#1500ff',justify=CENTER)
h_lines.insert(0,0)
h_lines.pack()
    
# - - - - -text 2 + input file name
text2 = Label(root,text="Digite o nome do arquivo de saída:")
text2.pack()

name = Entry(root, width=40, borderwidth=3,fg='#1500ff')
name.pack()

# - - - - - My Button + Close Button
    
my_button=Button(root, text="  Consolidar Dados  ", command=select_folder)
my_button.pack(pady=15)
close_Button = Button(root, text=' Fechar ', command=close_window)
close_Button.pack(pady=0)

watermark = Label(root,text="Created by DDS", justify=RIGHT, fg='#00bbff' )
watermark.place(x=260, y=380)

root.mainloop()
    
#### Created By Danilo U da Silva ####
