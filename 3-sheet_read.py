# FAÇO ISSO PARA LER/CARREGAR A "PAGE" DE UMA PLANILHA EXCEL
from openpyxl import load_workbook

wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatorio"] #ESPECIFICO O NOME DA "PAGE"/WORKBOOK DA MINHA PLANILHA

# ACESSO UM VALOR ESPECIFICO
# print(sheet["B3"].value)
# print(sheet["A3"].value)

# PEGANDO OS VALORES MEIO DE LOOP
for i in range(2, 6):
    ano = sheet["A%s" %i].value
    cv = sheet["B%s" %i].value
    ft = sheet["C%s" %i].value
    print("{0} o Chevrolet vendeu {1} e o Fiat vendeu {2}".format(ano, cv, ft))