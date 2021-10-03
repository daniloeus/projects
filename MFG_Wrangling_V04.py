### Created by Daniloeus ###

from tkinter import *
import os.path
from sys import exit
import pandas as pd
#Created by Danilo Uzelin da Silva


def full_consolidation(folder):

    path = os.path.abspath(folder) 
    files = os.listdir(path) 

    files_mfg = []
    files_aduanas = []

    for file in files:
        if file.endswith('.xlsx') and file.startswith('5.5.1.10'):
            print('Arquivo MFG encontrado: ',file)
            files_mfg.append(file)
        elif file.startswith('Custos Log'):
            print('Arquivo de Custos Logísticos encontrado: ',file)
            files_aduanas.append(file)

    # > > > > > Outuput as request < < < < <
    output_cols = ['Plant','Part Number','Part Description','Transit Days','Prod Line','Status','Safety Days','Site']


    # > > > > > > > > > > Def tratamento dos dados base MFG < < < < < < < < < <
    def consolidation_mfg(files):

        df = pd.DataFrame() #Non global variable
        cols= ['Item',
           'Descrição do Item',
           'Local',
           'Dias Transporte',
           'Linha Produto',
           'Status',
           'Tempo Segur',
           'Compra/Fabric']

        for file in files:
            file_in_folder = os.path.join(path,file)
            df = df.append(pd.read_excel(file_in_folder, header=1, usecols=cols), ignore_index=True)


        # - - - - - remove rows if Item = NaN (inplace) - - - - -
        df.dropna(subset=['Item'], inplace=True)

        # - - - - - rename columns and create new datas from Local id - - - - -
        df['Plant']= df['Local'].replace([46911,18701,18711,14421,44111,4601,8000,10281,16201,29711,42311,86421],
                                         ['CB' ,'GV' , 'GV','PA' ,'PA' ,'RO','RO','RO' ,'SB' ,'SB' ,'SB' ,'SB' ])

        df['Site']= df['Local'].replace([46911,18701,18711,14421,44111,4601,8000,10281,16201,29711,42311,86421],
                                        [469  ,1870 ,1871 ,1442 ,441  ,460 ,800 ,1028 ,1620 ,297  ,423  ,8642])

        columns_change = {'Item':'Part Number',
                          'Descrição do Item':'Part Description',
                          'Dias Transporte':'Transit Days',
                          'Linha Produto': 'Prod Line',
                          'Tempo Segur':'Safety Days',}

        df.rename(columns=columns_change, inplace=True)

        global df_mfg_output
        df_mfg_output = df[output_cols]
        df_mfg_output.drop_duplicates(subset = output_cols, inplace= True)
    # > > > > > > > > > > FIM Def tratamento dos dados base MFG < < < < < < < < < < 
    # - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    # > > > > > > > > > > Def tratamento dos dados Planilha Aduanas < < < < < < < < < <    
    def consolidation_aduanas(files):
        df_a = pd.DataFrame() #Non global variable
        cols_aduanas = ["CHEGADA", "UNIDADE","PARTNUMBER"]

        for file in files:
            file_in_folder = os.path.join(path,file)
            df_a = df_a.append(pd.read_excel(file_in_folder, usecols=cols_aduanas, parse_dates=['CHEGADA']), ignore_index=True)

        df_a = df_a.loc[df_a['PARTNUMBER'].str.startswith('G')==True]
        df_a = df_a.loc[df_a['CHEGADA'] > pd.to_datetime('01/08/19')] #sem resultado efetivo, todos as chegadas são de datas pós 01/08/2019
        dict_local = {'PA':441,
                      'SB':297, 
                      'GV':1870, 
                      'CB':469}
        df_a['Site'] = df_a['UNIDADE'].map(dict_local)
        df_a['Part Description'],df_a['Transit Days'],df_a['Prod Line'],df_a['Status'],df_a['Safety Days'] = 'GCode', 2, 'GCode','GCode',2
        df_a.rename(columns={'UNIDADE':'Plant','PARTNUMBER':'Part Number'}, inplace=True)
        global df_a_output
        df_a_output = df_a[output_cols]
        df_a_output.drop_duplicates(subset = output_cols, inplace= True)
    # > > > > > > > > > > FIM Def tratamento dos dados Planilha Aduanas < < < < < < < < < <

    consolidation_mfg(files_mfg)
    consolidation_aduanas(files_aduanas)

    df_final = df_mfg_output.append(df_a_output)
    df_final.drop_duplicates(subset = output_cols, inplace= True)
    output_name = folder+'_(DSS_Py).xlsx'
    df_final.to_excel(output_name,index=False)


folder_init=""
path = os.path.abspath(folder_init) 
directory = os.listdir(path) 
list_dir= []

for i in directory:
    if os.path.isdir(i):
        list_dir.append(i)
    
root = Tk()
root.title('MFG Data Wrangling V04')
root.geometry("400x330")


def select_file():
    global my_file
    to_transform = my_listbox.get(ANCHOR)
    my_label.config(text="Arquivos da pasta '"+my_listbox.get(ANCHOR)+"' processados com sucesso!!")
    print(to_transform)
    full_consolidation(to_transform)
#    root.destroy()


def close_window():
    root.destroy()

# - - - - -Top text!
initial_text = Label(root,text="Selecione abaixo qual pasta contém os arquivos a serem consolidados \n e clique em 'Consolidar Dados'",fg= '#56DA43' ,bg="#303A6C")
initial_text.pack(pady=10)
# - - - - -Listbox!
scrollbar = Scrollbar(root, orient="vertical")
my_listbox = Listbox(root,width=60)

#my_listbox.geometry("120x50")
my_listbox.pack()

for item in list_dir:
    my_listbox.insert(END, item)
    
global my_label
my_label= Label(root,text='')
my_label.pack()

my_button=Button(root, text="Consolidar Dados", command=select_file)
my_button.pack(pady=10)


close_Button = Button(root, text='Fechar', command=close_window)
close_Button.pack(pady=10)


root.mainloop()