import numpy as np
import pandas as pd
import openpyxl
import xlsxwriter


planilha = r'raiz/VALOR REPOSIÇÃO.xlsx'

nome_da_pagina = pd.ExcelFile(planilha)
nome_da_pagina = nome_da_pagina.sheet_names

x = 0
for nome in nome_da_pagina:
    pagina = pd.read_excel(planilha, sheet_name=x)
    nome_do_arquivo = nome_da_pagina[x]

    PN = []
    TNF_FINAL = []

    for i in pagina["PN"]:
        i = i.replace("'", "")
        PN.append(i)

    for i in pagina["TNF FINAL"]:
        i = round(i, 2)
        i = str(i)
        i = i.replace(",", ".")
        TNF_FINAL.append(i)

    workbook = xlsxwriter.Workbook(f'raiz/{nome_do_arquivo}.xlsx')
    worksheet = workbook.add_worksheet("produtos")

    worksheet.write('A1', 'referencia_fabrica')
    worksheet.write('B1', 'descricao')
    worksheet.write('C1', 'codigo_barras')
    worksheet.write('D1', 'peso')
    worksheet.write('E1', 'preco_publico')
    worksheet.write('F1', 'preco_custo')
    worksheet.write('G1', 'preco_garantia')
    worksheet.write('H1', 'aliquota_ipi')
    worksheet.write('I1', 'ncm')
    worksheet.write('J1', 'classe_desconto')

    row = 1
    col = 0

    for pn in PN:
        worksheet.write_string(row, col, pn)
        row += 1

    row = 1
    col = 0

    for tnf in TNF_FINAL:
        worksheet.write_string(row, col + 5, tnf)
        worksheet.write_string(row, col + 6, tnf)
        row += 1

    workbook.close()

    print(f'terminei o {nome_do_arquivo}')
    x += 1
