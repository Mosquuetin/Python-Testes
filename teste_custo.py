import pandas as pd
import numpy as np
import glob

def open_excel(filepath):
    filename = filepath.split('\\')[1].replace('.xlsb','')
    df = pd.ExcelFile(filepath)
    abas = df.sheet_names
    abas_filtradas = list(filter(lambda aba: 'BB_' in aba and len(aba)==4, abas))
    df1 = pd.DataFrame()
    for aba in abas_filtradas:
        print('lendo aba ' + aba)
        df2 = pd.read_excel(filepath, sheet_name=aba, header = 5, usecols = "A:AS")
        if df1.empty:
            df1 = df2
            print('criando primeiro DF')
        else:
            print('concatenando DFs posteriores')
            df1 = pd.concat([df1,df2])
    print('salvando arquivo')
    df1.to_csv("Outputs\\"+filename+".csv",index = False)


files = glob.glob('Files/*.xlsb')
for file in files:
    print(file)
    open_excel(file)
    

    