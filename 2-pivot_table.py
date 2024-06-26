import pandas as pd

data = pd.read_excel("data/VendaCarros.xlsx")

# print (type(data))

# AKI EU SELECIONO COLUNAS ESPECÍFICAS DO DATAFRAME
df = data[["Fabricante", "ValorVenda", "Ano"]]
# print(df)

# AKI EU CRIO A TABELA PIVÔ
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)
print(pivot_table)

# EXPORTO A TABLE PIVO EM ARQUIVO EXCEL
pivot_table.to_excel("data/pivot_table.xlsx", "Relatorio")