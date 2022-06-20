from tika import parser
import pandas as pd
import os


razao = pd.read_excel('razao.xlsx')
razao['Nota'] = razao['Texto'].str.slice(10, 19)


dados = [[], [], [], []]
for nota in os.listdir('G:\GECOT\\NOTAS FISCAIS DIGITALIZADAS\\2022\\05 - MAIO\ENERGIA ELÉTRICA'):
    diretorio = 'G:\GECOT\\NOTAS FISCAIS DIGITALIZADAS\\2022\\05 - MAIO\ENERGIA ELÉTRICA'
    if nota.endswith('.pdf'):
        conta = parser.from_file(os.path.join(diretorio, nota))
        linha_conta = conta['content'].splitlines()
        outros_deb = 0
        for index, row in enumerate(linha_conta):
            if 'Série C' in row:
                dados[3].append(linha_conta[index].split(' ')[1]) if linha_conta[index].split(' ')[1] \
                                                                     not in dados[3] else None
            if 'CNPJ' in row:
                dados[0].append(linha_conta[index-4])
                dados[1].append(linha_conta[index - 2][10:].split('-')[0])
            if 'DÉBITOS' in row:
                outros_deb = float(linha_conta[index+2].split(' ')[6].replace(',', '.'))
            if 'Total a Pagar (R$)' in row:
                vr_total = float(linha_conta[index+1].strip().replace(',', '.'))
                imposto = (vr_total - outros_deb) * 0.0925
                vr_a_pagar = vr_total - imposto
                dados[2].append(round(vr_a_pagar, 2))

dados = pd.DataFrame(dados).T
dados.columns = ['Endereco', 'Cidade', 'Valor', 'Nota']

dados_a_completar = pd.merge(razao, dados[['Nota', 'Cidade']], on=['Nota'], how='left')


# dados.to_excel('energia.xlsx', index=False)
dados_a_completar.to_excel('energia.xlsx')