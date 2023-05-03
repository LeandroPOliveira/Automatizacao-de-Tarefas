from tika import parser
import os
import pandas as pd

lista = [[], [], [], [], []]
mes = input('Digite o mês no formato "mm-aaaa": ').strip()

diretorio = rf'G:\GECOT\APURAÇÃO DA RECEITA\{mes[-4:]}\DANFE {mes}'
for nota in os.listdir(diretorio):
    conta = parser.from_file(os.path.join(diretorio, nota))
    linha_conta = conta['content'].splitlines()
    linha_conta = [linha.replace('.', '').replace(',', '.') for linha in linha_conta]
    pisco = 0
    for index, row in enumerate(linha_conta):
        if 'MMB' in row:
            lista[0].append(linha_conta[index+2])
        if 'V TOTAL PRODUTOS' in row:
            lista[1].append(linha_conta[index + 2])
        if 'VALOR DO ICMS' == row:
            lista[2].append(linha_conta[index+2])
        if 'VALOR DO PIS' in row:
            pisco = float(linha_conta[index+2])
        if 'VALOR DA COFINS' in row:
            pisco += float(linha_conta[index+2])
            lista[3].append(pisco)
        if '1 - SAÍDA 1' in row:
            lista[4].append(linha_conta[index + 2])

dados = pd.DataFrame(lista).T

# dados = dados.astype(float)

dados.columns = ['Qde', 'Valor Total', 'ICMS', 'PIS/COFINS', 'nota']

# dados['Valor Unit.'] = dados['Valor Total'] / dados['Qde']

dados.to_excel('gas.xlsx')
