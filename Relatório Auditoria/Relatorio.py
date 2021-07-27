import pandas as pd
from datetime import datetime

pd.options.mode.chained_assignment = None  # default='warn'
desired_width=320
pd.set_option('display.width', desired_width)
pd.set_option('display.max_columns', 10)


def ativo_geral(data):
    data = pd.DataFrame(data)
    # data.columns = [i for i in range(data.shape[1])]
    colunas = [0, 1, 4, 5, 8, 9, 10, 12, 25, 26]
    data.drop(data.columns[colunas], axis=1, inplace=True)
    data = data[data['Classe'].notnull()]
    data = data[~data['Classe'].str.contains('Classe')]
    data['Dt.incorp.'] = data['Dt.incorp.'].apply(lambda x: pd.to_datetime(x, format='%d.%m.%Y'))
    data = data.sort_values(by=['Dt.incorp.'])
    data.drop(data.tail(6).index, inplace=True)
    data = data[data['Classe'] != 'IES-50']
    numeros = ['     CAP InícEx', '      Dep.InícEx', ' ValContIníExer', '        Aquis.', '         Baixa',
               '     Transfer.', ' Deprec.do ano', '    Baixa dep.', '  Transf.depr.', '     CAP atuais',
               '  Depr.acumulada', '   ValCon.atual']
    data[numeros] = data[numeros].apply(pd.to_numeric, errors='coerce')
    return data

dados_inicio = ativo_geral(pd.read_excel('Audit.xlsx', skiprows=8))

def define_datas(mes):
    mes = datetime.strptime(mes, '%m-%Y')
    global exercicio_anterior
    global exercicio_atual
    exercicio_anterior = str(int(mes.strftime('%m')) + 12-int(mes.strftime('%m'))) + '-' + str((mes.year)-1)
    exercicio_atual = str(mes.strftime('%m')) + '-' + str(mes.year)
    return mes

mes_relatorio = define_datas(input('Digite o mês (mm-aaaa): '))

def busca_estoque(estoque):
    estoque.rename(columns={'SALDO EM 31.12.2020': 'ANTERIOR', 'SALDO EM 30.06.2021': 'ATUAL'}, inplace=True)
    global estoque_anterior
    global estoque_atual
    estoque_anterior = estoque['ANTERIOR'].loc[180]
    estoque_atual = estoque['ATUAL'].loc[180]
    return estoque

# saldo_estoque = busca_estoque(pd.read_excel('Ativo financeiro - 05_2021.xlsx'))
saldo_estoque = busca_estoque(pd.read_excel(r'G:\GECOT\ATIVO FINANCEIRO\ATIVO FINANCEIRO_' + str(mes_relatorio.year) + '\\'
                                            + str(mes_relatorio.strftime('%m')) + '_' + str(mes_relatorio.year) +
                                            '\Ativo financeiro - 06_2021.xlsx', sheet_name='MAPA', usecols=[7, 8]))


def ativo_ano_anterior(anterior):
    anterior = anterior[anterior['Dt.incorp.'].dt.year < mes_relatorio.year]
    global novo
    # totais = anterior.loc['TOTAL'] = anterior.iloc[:, 4:].sum(axis=0)
    novo = list(anterior.iloc[:, 4:].sum(axis=0))
    novo.insert(3, mes_relatorio)
    while len(novo) < 17:
        novo.insert(0, '-')
    novo = pd.DataFrame(novo).transpose()
    anterior.loc['TOTAL'] = anterior.iloc[:, 4:].sum(axis=0)
    anterior = anterior.append({'      Dep.InícEx': "Estoque", "  Depr.acumulada": "Estoque",
                                ' ValContIníExer': estoque_anterior,
                                '   ValCon.atual': estoque_atual}, ignore_index=True)
    anterior = anterior.append({' ValContIníExer': estoque_anterior + anterior[' ValContIníExer'].iloc[-2],
                                '      Dep.InícEx': exercicio_anterior,
                                '   ValCon.atual': estoque_atual + anterior['   ValCon.atual'].iloc[-2],
                                '  Depr.acumulada': exercicio_atual},
                               ignore_index=True)
    anterior['Dt.incorp.'] = anterior['Dt.incorp.'].apply(lambda x: pd.to_datetime(x, format='%d.%m.%Y'))
    return anterior

dados_anterior = ativo_ano_anterior(dados_inicio)

def ativo_atual(dados_atuais):
    dados_atuais = dados_atuais[dados_atuais['Dt.incorp.'].dt.year >= mes_relatorio.year]
    novo.columns = dados_atuais.columns
    dados_atuais = pd.concat([novo, dados_atuais])
    dados_atuais.loc['TOTAL'] = dados_atuais.iloc[:, 4:].sum(axis=0)
    dados_atuais = dados_atuais.append({'      Dep.InícEx': "Estoque", "  Depr.acumulada": "Estoque",
                                ' ValContIníExer': estoque_anterior,
                                '   ValCon.atual': estoque_atual}, ignore_index=True)
    dados_atuais = dados_atuais.append({' ValContIníExer': estoque_anterior + dados_atuais[' ValContIníExer'].iloc[-2],
                                '      Dep.InícEx': exercicio_anterior,
                                '   ValCon.atual': estoque_atual + dados_atuais['   ValCon.atual'].iloc[-2],
                                '  Depr.acumulada': exercicio_atual},
                               ignore_index=True)
    dados_atuais['Dt.incorp.'] = dados_atuais['Dt.incorp.'].apply(lambda x: pd.to_datetime(x, format='%d.%m.%Y'))
    return dados_atuais

base_atual = ativo_atual(dados_inicio)

def ativos_depreciados(depreciados):
    depreciados = depreciados[depreciados['Classe'].str.contains('IES')]
    depreciados = depreciados[depreciados['   ValCon.atual'] == 0]
    return depreciados

deprec = ativos_depreciados(dados_inicio)

def formatar_relatorio(total, antes, atual, depr):
    writer = pd.ExcelWriter('G:\GECOT\CONTROLE PATRIMONIAL\\2021\Imobilizado-' + str(datetime.strftime(mes_relatorio, '%m-%Y'))
                                                                                      + '.xlsx', engine='xlsxwriter',
                            date_format='DD/MM/YYYY')

    total['Dt.incorp.'] = total['Dt.incorp.'].dt.date
    antes['Dt.incorp.'] = antes['Dt.incorp.'].dt.date
    atual['Dt.incorp.'] = atual['Dt.incorp.'].dt.date
    depr['Dt.incorp.'] = depr['Dt.incorp.'].dt.date

    total.to_excel(writer, sheet_name='Ativo Geral', index=False)
    antes.to_excel(writer, sheet_name='Ativo ' + str(exercicio_anterior), index=False)
    atual.to_excel(writer, sheet_name='Ativo ' + str(exercicio_atual), index=False)
    depr.to_excel(writer, sheet_name='Depreciados', index=False)
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


    writer.save()

formatar_relatorio(dados_inicio, dados_anterior, base_atual, deprec)