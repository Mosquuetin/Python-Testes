import pandas as pd
import glob
import msoffcrypto
import io

def open_excel(filepath):
    df = pd.read_excel(filepath, sheet_name="Arquivos")
    print('Arquivo aberto')
    agentes = list(df['Agente'])
    arquivos = list(df['Arquivo'])
    caminhos = list(df['Caminho'])
    datas = list(df['Data'])
    df = pd.read_excel(filepath, sheet_name="Abas")
    print(df)
    print('Executando Loop')
    df2 = pd.DataFrame()
    i = 0
    for agente in agentes:
        print('Iniciando '+ agente)
        df1 = df[df.Agente.eq(agente)]
        abas = list(df1['Nome da Aba'])
        caminho = caminhos[i]
        arquivo = arquivos[i]
        data = datas[i]
        i = i + 1
        print(abas)
        x = 0
        for aba in abas:
            print('Lendo aba ' + aba + ' de ' + agente)
            responsaveis = list(df1['Responsável'])
            grupos = list(df1['País / Grupo'])
            bancos = list(df1['Banco'])
            agencias = list(df1['Agência'])
            contas = list(df1['Conta'])
            responsavel = responsaveis[x]
            grupo = grupos[x]
            banco = bancos[x]
            agencia = agencias[x]
            conta = contas [x]
            abertura_excel = io.BytesIO()
            with open(caminho, 'rb') as abertura:
                excel = msoffcrypto.OfficeFile(abertura)
                excel.load_key(('apc2016'))
                excel.decrypt(abertura_excel)
            df3 = pd.read_excel(abertura_excel, sheet_name = aba, header = 5, usecols = "A:AS", na_filter=False)
            del abertura_excel
            df3.insert(0,'Conta', value = conta)
            df3.insert(0,'Agência', value = agencia)
            df3.insert(0,'Banco', value = banco)
            df3.insert(0,'Agente', value = agente)
            df3.insert(0,'País / Grupo', value = grupo)
            df3.insert(0,'Status / Aba', value = aba)
            df3.insert(0,'Responsável', value = responsavel)
            df3 = df3[df3['DATA'] == data]
            x = x + 1
            if df2.empty:
                df2 = df3
            else:
                df2 = pd.concat([df2,df3])
            
        print('Compilando Base '+ agente)
        print('Salvando Arquivo')
        df2.to_csv("Outputs\\"+arquivo+".csv", index = False)
        df2 = pd.DataFrame()
        df3 = pd.DataFrame()

files = glob.glob('Auxiliar/*.xlsx')


for file in files:
    print(file)
    open_excel(file)