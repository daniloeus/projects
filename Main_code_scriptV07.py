#### Created By Danilo U da Silva ####
#### Automatização da rotina para gerar o relatório Daily Inventory ####

import pandas as pd
import numpy as np
from tkinter import *
from tkcalendar import *
from babel.dates import format_date, parse_date, get_day_names, get_month_names
from babel.numbers import *  # Additional Import
import webbrowser
from datetime import datetime
import os.path
import xlwings as xw

# ###################################### list of files CSV in selected folder ######################################################
def folderSelect(folder): #generate the list of paths and files to be consolidated
    global path     
    path = os.path.abspath(folder) 
    files = os.listdir(path) 
    global files_csv
    files_csv = []
    for file in files:
        if file.endswith('.csv'):
            print('Arquivo csv encontrado: ',file)
            files_csv.append(file)
            
# ###################################### appendig list of files csv ######################################################
def appendingFilesInFolder(listoffiles):
    #cols = ['Site','Part Number','Inventory Location','Qty On Hand','Total Inventory Value','Transit Days','Supplier Code','Supplier Name']
    global df_sar
    df_sar=pd.DataFrame()

    for file in listoffiles:
        file_in_folder = os.path.join(path,file)
        df_sar = df_sar.append(pd.read_csv(file_in_folder, sep=';',encoding='ISO-8859-1', header=1, decimal=',')) #usecols=cols,

    df_sar['Total Inventory Value'] = np.where((df_sar['Inventory Location'] == 'AUDI'), 0 ,df_sar['Total Inventory Value']) #Audi Inv = 0
    df_sar['Supplier Name'] = np.where((df_sar['Supplier Code'] == 'CONSIG V'),'CONSIGNADO VW' ,df_sar['Supplier Name']) #remove str on supplier code
    df_sar['Supplier Code']= df_sar['Supplier Code'].replace("CONSIG V", 0)
    df_sar['Supplier Code'].fillna(int(0), inplace=True) #remove NaN in supplier code
    df_sar['Supplier Code'] = df_sar['Supplier Code'].astype(int)

# ###################################### loading and merging SAP database ######################################################
def loadSAPInfos():
    global df_sar2
    SAP_suppliers = 'py base (do not delete)/ADNT_Suppliers_Base.xlsx'
    df_SAP = pd.read_excel(SAP_suppliers, usecols=['SupplierID','Company','CID'])
    df_SAP.rename(columns={'SupplierID':'Supplier Code','CID':'CID_SAP'}, inplace=True)

    df_sar2 = pd.merge(df_sar,df_SAP,how='left',on='Supplier Code')
    df_sar2['CID_SAP'].fillna(int(0), inplace=True) #remove NaN in supplier code
    
# ###################################### Classifying Local and Imported Material + applying it ###################################################### 
def local_importado (row):
    list_chemicals = [356712,366772,366724,381948,367657,366772,366724,381948,367657]
    ADNT_plants_BR = [29711,297,
                      42311,423,
                      16201,1620,
                      86421,864,
                      44111,441,
                      14421,1442,
                      18701,1870,
                      18711,1871,
                      46911,469,
                      459,461,43922]
    
    ADNT_plants_AR = [1028,8000,4601]

    val=''
    if row['Supplier Code'] in list_chemicals:
        val = "Importado"
    elif (row['Supplier Code'] in ADNT_plants_BR) and (row['Site'] in ADNT_plants_BR):
        val = "Local"    
    elif (row['Supplier Code'] in ADNT_plants_AR) and (row['Site'] in ADNT_plants_AR):
        val = "Local"    

    elif (row['Supplier Code'] in ADNT_plants_BR) and (row['Site'] in ADNT_plants_AR):
        val = "Importado"    
    elif (row['Supplier Code'] in ADNT_plants_AR) and (row['Site'] in ADNT_plants_BR):
        val = "Importado"    
         
    elif (row['CID_SAP'] == 'BR') and (row['Site'] in ADNT_plants_BR):
        val = "Local"
    elif (row['CID_SAP'] == 'AR') and (row['Site'] in ADNT_plants_AR):
        val = "Local"
    elif (row['CID_SAP'] == 'AR') and (row['Site'] in ADNT_plants_BR):
        val = "Importado"
    elif (row['CID_SAP'] == 'BR') and (row['Site'] in ADNT_plants_AR):
        val = "Importado"
    elif (row['Supplier Code'] == 0):
        val = "Local"
    elif (row['CID_SAP'] != 'BR') and (row['CID_SAP'] != 'AR') and (row['CID_SAP'] !=0):
        val = "Importado"
    elif (row['Transit Days'] > 0):
        val = "Importado"
    else:
        val = "Local"
    return val

def applying_local_importado():
    df_sar2['Origem'] = df_sar2.apply(local_importado,axis=1)
    
# ###################################### generating Output table ###################################################### 
def generatingOutput():
    global df_out1
    df_out1 = pd.pivot_table(df_sar2, index='Site',columns='Origem',values='Total Inventory Value', aggfunc='sum')
    df_out1['Total'] = df_out1['Importado']+df_out1['Local']
    df_out1=df_out1.reset_index()
    df_out1

    # ###################################### adding stock infos ###################################################### 

    negatives = {}
    devrec = {}
    loctran = {}

    for plant in df_sar2['Site'].unique():
        neg = (df_sar2.loc[(df_sar2['Site']==plant) & (df_sar2['Qty On Hand']<0)].shape[0])
        dev = (df_sar2.loc[(df_sar2['Site']==plant) & ((df_sar2['Inventory Location'].str.upper()=='DEVREC') | (df_sar2['Inventory Location'].str.upper()=='DEVOLPRO'))].shape[0] )
        loct = (df_sar2.loc[(df_sar2['Site']==plant) & (df_sar2['Inventory Location'].str.upper()=='LOCTRAN')].shape[0])
        negatives[plant] = neg
        devrec[plant] = dev
        loctran[plant] = loct    

    stock_info = pd.DataFrame.from_dict([negatives,devrec,loctran]).T
    stock_info.rename(columns={0:'Negative Items', 1:'DEVREC',2:'LOCTRAN'}, inplace=True)

    df_out1 = pd.merge(df_out1,stock_info,how='left', left_on='Site', right_index=True)

# ###################################### Storing Negative data in it's path ###################################################### 
def store_dfsar():
    df_sar3 = df_sar2.drop(columns=['Company','CID_SAP','Origem'], )
    df_sar3['Country'] = df_sar3['Country'].str.replace(' Genesis','', case=False)
    sar_file = str(selected_folder+'(consol).xlsx')
    df_sar3.to_excel(os.path.join(path,sar_file), index=False)    
    
# ###################################### applying rates to inventory ###################################################### 
def rate_import_inv(row):
    if row["Site"] in [4601,8000]:
        return (row['Importado'] /1000 ) / input_USDARG
    else:
        return (row['Importado'] /1000 ) / input_USDBRL

def rate_local_inv(row):
    if row["Site"] in [4601,8000]:
        return (row['Local'] /1000 ) / input_USDARG
    else:
        return (row['Local'] /1000 ) / input_USDBRL    
    
def applying_rates():
    df_out1['Importado_convertido'] = df_out1.apply(rate_import_inv, axis=1)
    df_out1['Local_convertido'] = df_out1.apply(rate_local_inv, axis=1)

    df_out1['rate_USDBRL'] = input_USDBRL
    df_out1['rate_USDARG'] = input_USDARG
    df_out1['report_date'] = input_date
#    df_out1['report_date']=df_out1['report_date'].astype('datetime64[ns]')
    df_out1['report_date']= pd.to_datetime(df_out1['report_date'], format="%d/%m/%Y")

# ###################################### adding output to excel file + closing file def ###################################################### 
def append_df_to_excel(df_to_be_append):
    excel_app = xw.App(visible=False)

    # Open a template file
    global wb
    wb = xw.Book('Inventory_Evolution_SAR_Template(in use).xlsm')

    # Assign data to last row +1 cell
    last_row = wb.sheets['Data'].range(1,1).end('down').row
    #print("The last row is {row}.".format(row=last_row))
    #print("The DataFrame df has {rows} rows.".format(rows=df.shape[0]))

    wb.sheets['Data'].range(last_row+1,1).options(index=False, header=False,parse_dates=True, decimal='.').value = df_to_be_append

    # Save under a new file name
    wb.save('Inventory_Evolution_SAR_Template(in use).xlsm')
    wb.close()
    excel_app.quit()

    txt_backup = str(selected_folder+'_append.txt')
    #f_txt = open(os.path.join(path,txt_backup), "w")
    df_to_be_append.to_csv(os.path.join(path,txt_backup), header=True, index=None, sep=';')

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
root.title('Daily Inventory Report -DDS.V04-')
root.geometry("480x730+450+30") # window width x window height + X coordinate + Y coordinate


def select_folder():
    global input_date,input_USDBRL ,input_USDARG, selected_folder
    
    input_date = report_date.get_date()
    input_USDBRL= float(rate_USDBRL.get().replace(',','.'))
    input_USDARG= float(rate_USDARG.get().replace(',','.'))
    
    selected_folder = my_listbox.get(ANCHOR)
    print(selected_folder)
    print(input_date,input_USDBRL,input_USDARG, sep='\n')
    
    folderSelect(selected_folder)
    print('Folder selected')
    appendingFilesInFolder(files_csv)
    print('Files appended')
    loadSAPInfos()
    print('SAP suppliers database loaded')
    applying_local_importado()
    print('Local and Imported suppliers classified')
    generatingOutput()
    print('Output template generated')
    store_dfsar()
    print('Consolidated Data stored')
    applying_rates()
    print('FX rates applied')
    print('Appending new data to excel file')
    append_df_to_excel(df_out1)
    print('Output added to excel spreadsheet')
    root.destroy()


def close_window():
    root.destroy()

def callback(url):
    webbrowser.open_new(url)
    
# - - - - -Top text!
toptext = Label(root,text='Bem Vindo!! \n'+
              'Preencha todos os campos e selecione a pasta que contém\n'+
              'os arquivos para a criação correta do relatório',fg= '#56DA43' ,bg="#303A6C", width=100)
toptext.pack(pady=5)
# - - - - -text 1 + data
text1 = Label(root,text="1) Selecione a data referente ao relatório")
text1.place(x=20 , y=60)
report_date = Calendar(root, selectmode='day',
                       year = datetime.today().year,
                       month = datetime.today().month,
                       day = datetime.today().day,date_pattern='dd/mm/yyyy')
report_date.pack(pady=25)

# - - - - -text 2 + Link + rate USDBRL
text2 = Label(root,text="2) Informe o rate 'Dolar x Real' (Taxa de Venda):")
text2.place(x=20 , y=300)

rate_USDBRL = Entry(root, width=20, borderwidth=3)
rate_USDBRL.place(x=300 , y=305)

link1 = Label(root, text="Site: Banco Central - Brasil", fg="blue", cursor="hand2")
link1.place(x=290 , y=330)
link1.bind("<Button-1>", lambda e: callback("https://www.bcb.gov.br/estabilidadefinanceira/fechamentodolar"))

# - - - - -text 3 + Link + rate USDARG
text3 = Label(root,text="3) Informe o rate 'Dolar x Peso ARG' (Mayorista):")
text3.place(x=20 , y=360)

rate_USDARG = Entry(root, width=20, borderwidth=3)
rate_USDARG.place(x=300 , y=365)

link2 = Label(root, text="Site: Banco Central - Argentina", fg="blue", cursor="hand2")
link2.place(x=290 , y=390)
link2.bind("<Button-1>", lambda e: callback("http://www.bcra.gov.ar"))



# - - - - -text 4 + Listbox folders
text4 = Label(root,text="4) Selecione a pasta com os arquivos MFG do relatório:")
text4.place(x=20 , y=420)

scrollbar = Scrollbar(root, orient="vertical")
my_listbox = Listbox(root,width=60, borderwidth=3)
my_listbox.place(x=58, y=440)

for item in list_dir:
    my_listbox.insert(END, item)
    
my_button=Button(root, text="Consolidar Dados", command=select_folder)
my_button.place(x=190, y=630)


close_Button = Button(root, text='Fechar', command=close_window)
close_Button.place(x=218, y=680)

watermark = Label(root,text="Created by adasild9", justify=RIGHT, fg='#00bbff' ) ;watermark.place(x=365, y=705)

root.mainloop()
###################################### END #########################################
print("  ______________________________________________________________________  ")
print(" |   ______  ______  __   ___                                   adasild9| ")
print(" |  ||||||| ||||||| |||\ /|||                                           | ")
print(" |  |||____ |||     |||\V/|||   __                                      | ")
print(" |  ||||||| |||     ||| V |||  |  \      ___       ___   __   _         | ")
print(" |   ___||| |||____ |||   |||  |   | /\ |_ _| /\    |   |__  /_\  |\/|  | ")
print(" |  ||||||| ||||||| |||   |||  |__/ /__\ |_| /__\   |   |__ /   \ |  |  | ")
print(" |______________________________________________________________________| ")
#### Created By Danilo U da Silva ####
