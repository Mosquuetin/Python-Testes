import pandas as pd
import os

#caminho_inicial = os.getcwd()+"\Auxiliar\Cronograma.xlsx"
#caminho_final = os.getcwd()+"\Análises\StarkBank\\"
caminho_inicial = os.getcwd()[ : -10]+r"\Directa24 Dropbox\Auditoria\Gerencial (temporaria)\2022\Projetos Python\Auxiliar\Cronograma.xlsx"
caminho_final = os.getcwd()[ : -10]+r"\Directa24 Dropbox\Auditoria\Gerencial (temporaria)\2022\Projetos Python\Análises\StarkBank\\"

df = pd.read_excel(caminho_inicial, sheet_name="Parâmetros Python")
aba_inicial = list(df['Aba Inicial'])
aba_inicial = aba_inicial[0]
df = pd.read_excel(caminho_inicial, sheet_name=aba_inicial)
print('Arquivo aberto')
agentes = list(df['Agente'])
arquivos = list(df['Arquivo'])
caminhos = list(df['Caminho'])
print(df)
print('Executando Loop')
i = 0
for agente in agentes:
    print('Iniciando '+ agente)
    caminho = caminhos[i]
    arquivo = arquivos[i]
    df1 = pd.DataFrame()
    i = i + 1
    df2 = pd.read_csv(caminho, low_memory=False, usecols=[1,3,7,8,11,22])
    df2 = df2[df2['Status / Aba'].str.contains('STA')]
    print('Salvando Arquivo')
    df2.to_csv(caminho_final+arquivo+".csv", index = False)
