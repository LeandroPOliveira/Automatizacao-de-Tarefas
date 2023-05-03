import pandas as pd
from datetime import datetime

pd.options.mode.chained_assignment = None  # default='warn'
desired_width = 320
pd.set_option('display.width', desired_width)
pd.set_option('display.max_columns', 10)


planilha_base = pd.read_excel('Audit3.xlsx')
data = pd.DataFrame(planilha_base)
colunas = [17]
data.drop(data.columns[colunas], axis=1, inplace=True)
data.rename(columns={'Classe imobilizado': 'Classe', 'Incorporação em': 'Dt.incorp.'}, inplace=True)
print(data.columns)
data = data[data['Classe'].notnull()]
data = data[~data['Classe'].str.contains('Classe')]
data['Dt.incorp.'] = data['Dt.incorp.'].apply(lambda x: pd.to_datetime(x, format='%d.%m.%Y'))
data = data.sort_values(by=['Dt.incorp.'])
data.drop(data.tail(6).index, inplace=True)
data = data[data['Classe'] != 'IES-50']
numeros = ['    CAP InícEx', '    Dep.InícEx', 'ValContIníExer', '        Aquis.', '         Baixa',
           '     Transfer.', 'Deprec.do ano', '    Baixa dep.', '  Transf.depr.', '    CAP atuais',
           'Depr.acumulada', '  ValCon.atual']
data[numeros] = data[numeros].apply(pd.to_numeric, errors='coerce')


data_atual = input('Digite o mês (mm-aaaa): ').strip()
data_formatada = datetime.strptime(data_atual, '%m-%Y')
exercicio_anterior = str(int(data_formatada.strftime('%m')) + 12 - int(data_formatada.strftime('%m'))) + '-' + \
                     str(data_formatada.year - 1)
exercicio_atual = str(data_formatada.strftime('%m')) + '-' + str(data_formatada.year)
mes = data_atual[:2]
ano = data_atual[-4:]

path_rede = rf'G:\GECOT\ATIVO FINANCEIRO\ATIVO FINANCEIRO_{ano}\{mes}_{ano}\Ativo financeiro - ' \
            rf'{mes}_{ano}.xlsx'

estoque = pd.read_excel(path_rede, sheet_name='MAPA', usecols=[7, 8])
estoque.rename(columns={estoque.columns[0]: 'ANTERIOR', estoque.columns[1]: 'ATUAL'}, inplace=True)

estoque_anterior = estoque['ANTERIOR'].loc[180]
estoque_atual = estoque['ATUAL'].loc[180]


anterior = data[data['Dt.incorp.'].dt.year < data_formatada.year]
novo = list(anterior.iloc[:, 4:].sum(axis=0))
novo.insert(3, data_formatada)
while len(novo) < 17:
    novo.insert(0, '-')
novo = pd.DataFrame(novo).transpose()
anterior.loc['TOTAL'] = anterior.iloc[:, 4:].sum(axis=0)
anterior = anterior.append({'    Dep.InícEx': "Estoque", "Depr.acumulada": "Estoque",
                            'ValContIníExer': estoque_anterior,
                            '  ValCon.atual': estoque_atual}, ignore_index=True)
anterior = anterior.append({'ValContIníExer': estoque_anterior + anterior['ValContIníExer'].iloc[-2],
                            '    Dep.InícEx': exercicio_anterior,
                            '  ValCon.atual': estoque_atual + anterior['  ValCon.atual'].iloc[-2],
                            'Depr.acumulada': exercicio_atual},
                           ignore_index=True)
anterior['Dt.incorp.'] = anterior['Dt.incorp.'].apply(lambda x: pd.to_datetime(x, format='%d.%m.%Y'))


dados_atuais = data[data['Dt.incorp.'].dt.year >= data_formatada.year]
novo.columns = dados_atuais.columns
dados_atuais = pd.concat([novo, dados_atuais])
dados_atuais.loc['TOTAL'] = dados_atuais.iloc[:, 4:].sum(axis=0)
dados_atuais = dados_atuais.append({'    Dep.InícEx': "Estoque", "Depr.acumulada": "Estoque",
                                    'ValContIníExer': estoque_anterior,
                                    '  ValCon.atual': estoque_atual}, ignore_index=True)
dados_atuais = dados_atuais.append({'ValContIníExer': estoque_anterior + dados_atuais['ValContIníExer'].iloc[-2],
                                    '    Dep.InícEx': exercicio_anterior,
                                    '  ValCon.atual': estoque_atual + dados_atuais['  ValCon.atual'].iloc[-2],
                                    'Depr.acumulada': exercicio_atual},
                                   ignore_index=True)
dados_atuais['Dt.incorp.'] = dados_atuais['Dt.incorp.'].apply(lambda x: pd.to_datetime(x, format='%d.%m.%Y'))


depreciados = data[data['Classe'].str.contains('IES')]
depreciados = depreciados[depreciados['  ValCon.atual'] == 0]


writer = pd.ExcelWriter(
    fr'G:\GECOT\CONTROLE PATRIMONIAL\\{ano}\Imobilizado-' + str(datetime.strftime(data_formatada, '%m-%Y') + 'rev')
    + '.xlsx', engine='xlsxwriter',
    date_format='DD/MM/YYYY')

data['Dt.incorp.'] = data['Dt.incorp.'].dt.date
anterior['Dt.incorp.'] = anterior['Dt.incorp.'].dt.date
dados_atuais['Dt.incorp.'] = dados_atuais['Dt.incorp.'].dt.date
depreciados['Dt.incorp.'] = depreciados['Dt.incorp.'].dt.date

data.to_excel(writer, sheet_name='Ativo Geral', index=False)
anterior.to_excel(writer, sheet_name='Ativo ' + str(exercicio_anterior), index=False)
dados_atuais.to_excel(writer, sheet_name='Ativo ' + str(exercicio_atual), index=False)
depreciados.to_excel(writer, sheet_name='Depreciados', index=False)
workbook = writer.book
worksheet_total = writer.sheets['Ativo Geral']
worksheet_anterior = writer.sheets['Ativo ' + str(exercicio_anterior)]
worksheet_atual = writer.sheets['Ativo ' + str(exercicio_atual)]
worksheet_depr = writer.sheets['Depreciados']

format_numero = workbook.add_format({'num_format': '#,##0.00'})
format_texto = workbook.add_format({'num_format': '0', 'align': 'center'})
format_texto2 = workbook.add_format({'num_format': '0', 'align': 'left'})
# format_texto3 = workbook.add_format({'bold': True, 'fg_color': '#B7DEE8'})

format_list = [worksheet_total, worksheet_anterior, worksheet_atual, worksheet_depr]
for tab in format_list:
    tab.freeze_panes(1, 0)
    tab.set_column('A:A', 10, format_texto)
    tab.set_column('B:B', 15, format_texto)
    tab.set_column('C:C', 8, format_texto)
    tab.set_column('D:D', 40, format_texto2)
    colunas = ['E:E', 'F:F', 'G:G', 'H:H', 'I:I', 'J:J', 'K:K', 'L:L', 'M:M', 'N:N', 'O:O', 'P:P', 'Q:Q']
    for c in colunas:
        tab.set_column(c, 14, format_numero)

writer.close()
