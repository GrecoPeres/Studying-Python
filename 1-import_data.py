import pandas as pd

#AKI EU IMPORTO O DADO DA MINHA PLANILHA
data = pd.read_excel("data/VendaCarros.xlsx")
# print(data)

#LISTO OS PRIMEIROS 10 REGISTROS DOS MEUS DADOS
# print(data.head(10))

#LISTO OS ÚLTIMOS 10 REGISTROS
# print(data.tail(10))

# CONTAGEM DE VALORES POR FABRICANTE (COLUNA)
print(data["Fabricante"].value_counts())