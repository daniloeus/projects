### Criado do Daniloeus ### 
### Script teste para indicação por similaridade de novos itens de engenharia ###


print('\n\n 		***** Indicação por similaridade - V02 [DDS] *****\n')
import pandas as pd
print('			Pandas, Imported\n')

file = input('	Digite o nome do arquivo csv contendo as descrições MFG (inserir .csv ao final): ')
df=pd.read_csv(file, parse_dates=['Data Inclusao'], nrows=12500)
#----- Dropping duplicates Descrição 1-----#
df.drop_duplicates(subset = ['Descrição 1'], keep = 'last', inplace=True)
print('	DataFrame Loaded\n')

from fuzzywuzzy import process
from fuzzywuzzy import fuzz #token_sort_ratio
print('	FuZzYwUzZy imported\n')

file_inputs = input('	Digite o nome do arquivo csv que contenha as estradas as serem avaliadas (inserir .csv ao final): ')
inputs= pd.read_csv(file_inputs, error_bad_lines=False,warn_bad_lines=True, header=None)

print (' Processando......')

list_of_strings = list(inputs[0]) #lista do arquivo inputs
searches=[]
df_out=pd.DataFrame()

for i1 in list_of_strings:
    string = i1
    choices = df['Descrição 1']
    results = process.extract(string, choices,scorer=fuzz.token_sort_ratio, limit=3) # scorer=fuzz.token_sort_ratio
    
    for i2 in results:
        found = i2[0];score = i2[1];indexed = i2[2];
        searches.append([string,found,df['Descrição 2'].loc[indexed],score,df['Número de Item'].loc[indexed]])
    
df_out = df_out.append(searches , ignore_index=True)
df_out.columns=['input','found_EN','found_BR','score%','Pn_found']
print('			FuzzyWuzzy List Generated\n')

file_out = input('	Digite o nome desejado para o arquivo (inserir .xls ao final): ')
df_out.to_excel(file_out, index=False)
print('			Outputs Saved\n\n		Bye!!')
