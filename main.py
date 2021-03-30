import xlsxwriter

from datetime import datetime

import io

# Createaworkbookandaddaworksheet.

workbook = xlsxwriter.Workbook('test-project.xlsx')

worksheet = workbook.add_worksheet()

# Addaboldformattousetohighlightcells.

bold = workbook.add_format({'bold': 1})

# AddanExceldateformat.

date_now = datetime.now().isoformat(timespec='seconds')

# ARQUIVOCSVPRAPEGARDATOS

file = io.open("testcsv.csv", mode="r", encoding="utf-8")

data_join = []

for line in file:
    split_data = line.split(';')
    data_join.append(split_data)

# Cabeçalho

worksheet.write('A1', 'Data', bold)

worksheet.write('B1', 'Quantidade', bold)

worksheet.write('C1', 'NFE', bold)

worksheet.write('D1', 'nome_comprador', bold)

worksheet.write('E1', 'cpf_comprador', bold)

worksheet.write('F1', 'cod_produto', bold)

worksheet.write('G1', 'operacao_entrada', bold)

worksheet.write('H1', 'operacao_saida', bold)

worksheet.write('I1', 'valor_unit', bold)

worksheet.write('J1', 'valor', bold)

itens = []

for item in data_join:
    itens.append([item[0], item[1], item[2], item[3], item[4], item[5], item[6], float(item[7]), float(item[8])])

tuple_itens = tuple(itens)

# Startfromthefirstcellbelowtheheaders.

row = 1

col = 0

for quantidade, NFE, nome_comprador, cpf_comprador, cod_produto, operacao_entrada, operacao_saida, valor_unit, valor in tuple_itens:

    # Convertthedatestringintoadatetimeobject.

    worksheet.write_string(row, col, date_now)
    worksheet.write_string(row, col + 1, quantidade)
    worksheet.write_string(row, col + 2, NFE)
    worksheet.write_string(row, col + 3, nome_comprador)
    worksheet.write_string(row, col + 4, cpf_comprador)
    worksheet.write_string(row, col + 5, cod_produto)
    worksheet.write_string(row, col + 6, operacao_entrada)
    worksheet.write_string(row, col + 7, operacao_saida)
    worksheet.write_number(row, col + 8, valor_unit)
    worksheet.write_number(row, col + 9, valor)

    row += 1

workbook.close()
