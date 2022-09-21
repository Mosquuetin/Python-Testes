import pandas as pd
import glob

def open_excel(filepath):
    df = pd.read_excel(filepath, sheet_name="Arquivos")
    print('Arquivo aberto')
    agentes = list(df['Agente'])
    arquivos = list(df['Arquivo'])
    caminhos = list(df['Caminho'])
    df = pd.read_excel(filepath, sheet_name="Abas")
    print(df)
    print('Executando Loop')
    i = 0
    for agente in agentes:
        print('Iniciando '+ agente)
        df1 = df[df.Agente.eq(agente)]
        abas = list(df1['Nome da Aba'])
        caminho = caminhos[i]
        arquivo = arquivos[i]
        print(abas)
        print('Compilando Base '+ agente)
        df2 = pd.concat(pd.read_excel(caminho, sheet_name=abas, header = 5, usecols = "A:AS"), ignore_index = True)
        print(df2)
        print('Salvando Arquivo')
        df2.to_excel("Outputs\\"+arquivo+".xlsx",index = False)
        i = i + 1

files = glob.glob('Auxiliar/*.xlsx')


for file in files:
    print(file)
    open_excel(file)