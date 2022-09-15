import numpy as np
import pandas as pd

see = pd.read_excel(r"C:\Users\danie\Downloads\nb.xlsx")
df = pd.DataFrame(see)
# df["chave"] = df['Artigo'].astype(str) +"-"+ df["Tamanho"].astype(str)
# df["Chave canal"] = df['Escritório Cód'].astype(str) +"-"+ df["Canal Cód"].astype(str)
df["chave"] = df['Artigo'].astype(str) + df["Tamanho"].astype(str)
df["Chave canal"] = df['Escritório Cód'].astype(str) + df["Canal Cód"].astype(str)

# Método para indice tipo de ordem
conditions = [
    (df['Documento Tipo Cód'] == 'ZBME') | (df['Documento Tipo Cód'] == 'ZBRI') | (df['Documento Tipo Cód'] == 'ZBRM') | (df['Documento Tipo Cód'] == 'ZBTM') | (df['Documento Tipo Cód'] == 'ZCO2') | (df['Documento Tipo Cód'] == 'ZLCM') | (df['Documento Tipo Cód'] == 'ZLWC') | (df['Documento Tipo Cód'] == 'ZMOT') | (df['Documento Tipo Cód'] == 'ZCO1'),
    (df['Documento Tipo Cód'] == 'ZEXF'),
    (df['Documento Tipo Cód'] == 'ZLOP') | (df['Documento Tipo Cód'] == 'ZREP'),
    (df['Documento Tipo Cód'] == 'ZLWE'),
    (df['Documento Tipo Cód'] == 'ZMAM') | (df['Documento Tipo Cód'] == 'ZMOS') | (df['Documento Tipo Cód'] == 'ZNOR') | (df['Documento Tipo Cód'] == 'ZKHD') | (df['Documento Tipo Cód'] == 'ZKIL'),
]
choices = [1, 2, 3, 4, 5]
df['Indice tipord'] = np.select(conditions, choices, default=1)

# Tratativa para canal
indcanal = [

    (df['Chave canal'] == '50FQ'),
    (df['Chave canal'] == '50VJ'),
    (df['Chave canal'] == '61FQ'),
    (df['Chave canal'] == '62FQ'),
    (df['Chave canal'] == '63DG'),
    (df['Chave canal'] == '63FQ'),
    (df['Chave canal'] == '64FQ'),
    (df['Chave canal'] == '65VJ'),
    (df['Chave canal'] == '66VJ'),
    (df['Chave canal'] == '67VJ'),
    (df['Chave canal'] == '68VJ'),
    (df['Chave canal'] == '68FQ'),
    (df['Chave canal'] == '69FQ'),
    (df['Chave canal'] == '70VJ'),
    (df['Chave canal'] == '70DG'),
    (df['Chave canal'] == '73VJ'),
    (df['Chave canal'] == '75VJ'),
    (df['Chave canal'] == '76FQ'),
    (df['Chave canal'] == '77FQ'),
    (df['Chave canal'] == '78VJ'),
    (df['Chave canal'] == '78FQ'),
    (df['Chave canal'] == '79FQ'),
    (df['Chave canal'] == '80VJ'),
    (df['Chave canal'] == '97FQ'),
    (df['Chave canal'] == '98FQ'),
    (df['Chave canal'] == '98VJ'),
    (df['Chave canal'] == '99VJ'),
    (df['Chave canal'] == '99FQ'),
    (df['Chave canal'] == '50MM'),
    (df['Chave canal'] == '65MM'),
    (df['Chave canal'] == '66MM'),
    (df['Chave canal'] == '67MM'),
    (df['Chave canal'] == '68MM'),
    (df['Chave canal'] == '70MM'),
    (df['Chave canal'] == '73MM'),
    (df['Chave canal'] == '75MM'),
    (df['Chave canal'] == '78MM'),
    (df['Chave canal'] == '80MM'),
    (df['Chave canal'] == '98MM'),
    (df['Chave canal'] == '99MM'),
    (df['Chave canal'] == '58VJ'),
    (df['Chave canal'] == '75FQ'),
    (df['Chave canal'] == '81PL'),
    (df['Chave canal'] == '99TX'),
    (df['Chave canal'] == '98EX'),
    (df['Chave canal'] == '70KA'),
    (df['Chave canal'] == '77LP'),
    (df['Chave canal'] == '58MM'),

]

canalfila = ['FQ', 'MM', 'FQ', 'LP', 'WEB', 'WEB', 'FQ', 'MM', 'MM', 'MM', 'MM', 'MM', 'FQ', 'KA', 'KA', 'MM', 'MM', 'LP', 'LP', 'MM', 'MM', 'FQ', 'MM', 'MI', 'MI', 'MI', 'MM', 'FQ', 'MM', 'MM', 'MM', 'MM', 'MM', 'KA', 'MM', 'MM', 'MM', 'MM', 'MI', 'MM', 'MM', 'MM', 'MM', 'MM', 'MI', 'KA', 'LP', 'MM']
df['Canal'] = np.select(indcanal, canalfila, default=df['Canal Cód'])

# Ordem de alocação dos canais
MI = 1
KA = 2
MM = 3
FQ = 4
LP = 5
WEB = 6

# Tratativa para canal
canais = [

    (df['Canal'] == 'MI'),
    (df['Canal'] == 'KA'),
    (df['Canal'] == 'MM'),
    (df['Canal'] == 'FQ'),
    (df['Canal'] == 'LP'),
    (df['Canal'] == 'WEB'),

]
canalind = [MI, KA, MM, FQ, LP, WEB]
df['Indice Canal'] = np.select(canais, canalind, default=1)

# var = df.sort_values(
#    by=["Indice tipord", "Indice Canal"]
# )[["Indice tipord", "Indice Canal"]]

var2 = df.sort_values(by=['chave','Data Embarque Original', 'Status Resumo','Indice tipord','Indice Canal' ])
#var2.to_excel("finalmente.xlsx")

#------------------------------Extração estoque disponível

dfest = pd.read_excel(r"C:\Users\danie\Downloads\base estoque.xlsx")
df4 = pd.DataFrame(dfest)
var = df4.loc[df4['Tipo estoque'] == "Físico"]
df2 = pd.DataFrame(var)
var5 = df2.groupby(['chave'])['QTD'].sum()
df3 = pd.DataFrame(var5)
#print(df3)

#Junção das tabelas-----------------------------------------------

vendas = pd.DataFrame(var2)
left_join = pd.merge(vendas, df3, on='chave', how='left')
#left_join.to_excel("desesperoelagrimas.xlsx")

basefinal = pd.DataFrame(left_join)

basefinal["Est liq"] = 0
basefinal["Cart acum"] = 0
basefinal["Unit"] = 0
basefinal["Pçs atende"] = 0
basefinal["Pçs falta"] = 0
basefinal["Vlr atende"] = 0
basefinal["Vlr falta"] = 0
#basefinal.to_excel("desespero.xlsx")
dados = basefinal.to_numpy()

lin,col = dados.shape
x = lin
i = 1

dados[0,17] = 0
dados[0,18] = dados[0,16] - dados[0,17]
dados[0,19] = dados[0,9] / dados[0,10]
#Começa em 1 depois de definir o conteudo de zero como sendo 0
while i<x:
    if dados[i,11] == dados[i-1,11]:
        dados[i,17] = dados[i-1,17] + dados[i-1,10]
    else:
        dados[i,17] = 0
    #Desconta do estoque original o consumo acumulado
    dados[i, 18] = dados[i, 16] - dados[i, 17]
    dados[i, 19] = dados[i, 9] / dados[i, 10]

    #Lógica peças atende
    i = i + 1

a = 0
while a<x:
    if dados[a,18] > 0 and dados[a,18] >= dados[a,10]:
        dados[a,20] = dados[a,10]
    elif dados[a,18] >0 and dados[a,18] < dados[a,10]:
        dados[a,20] = dados[a,18]
    else:
        dados[a,20] = 0

    #vlr atende, pçs falta, valor falta
    dados[a,21] = dados[a,20] * dados[a,19]
    dados[a,22] = dados[a,10] - dados[a, 20]
    dados[a,23] = dados[a,9] - dados[a, 21]
    a = a + 1

edicao = pd.DataFrame(dados)
edicao.to_excel("calculado.xlsx")





