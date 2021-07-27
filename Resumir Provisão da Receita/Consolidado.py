import pandas as pd
from datetime import date

desired_width=320

pd.set_option('display.width', desired_width)

pd.set_option('display.max_columns', 10)

dataset = pd.read_csv('VALOR_CONSOLIDADO.CSV', sep=';', decimal=',', header=None)
pd.options.display.float_format = "{:,.2f}".format
dados = pd.DataFrame(dataset)


dados.columns = ['cliente', 'nome', 'cidade', 'segmento', 'vol prov ini', 'vol prov mes', 'vol prov fin', 'rec prov ini', 'rec prov mes',
                 'rec prov fin', 'vol fat ini', 'vol fat mes', 'vol fat fin', 'rec fat ini', 'rec fat mes', 'rec fat fin', 'vol ñ fat',
                 'rec ñ fat', 'pis', 'cofins', 'desconto', 'icms', 'icms/st']

dados.index = ['linha' + str(i) for i in range(len(dados))]

novo = pd.DataFrame(dados)

novo['for de gn'] = novo.apply(lambda x: x['rec ñ fat'] - x['desconto'] + x['icms/st'], axis=1)


sistema_2 = ['PORTO FERREIRA', 'SAO CARLOS', 'DESCALVADO']

selecao = novo['cidade'].isin(sistema_2)

dados_sistema_2 = novo[selecao]

dados_sistema_2.index = ['linha' + str(i) for i in range(len(dados_sistema_2))]

segmentos = list(novo['segmento'].drop_duplicates())

for l in segmentos:
    if l == 2:
        novo['segmento'] = novo['segmento'].replace([l], 'Res. Coletivo')
    elif l == 1:
        novo['segmento'] = novo['segmento'].replace([l], 'Residencial')
    elif l == 3:
        novo['segmento'] = novo['segmento'].replace([l], 'Comercial')
    elif l == 4:
        novo['segmento'] = novo['segmento'].replace([l], 'Industrial')
    elif l == 5:
        novo['segmento'] = novo['segmento'].replace([l], 'Industrial')
    elif l == 6:
        novo['segmento'] = novo['segmento'].replace([l], 'GNV')
    elif l == 8:
        novo['segmento'] = novo['segmento'].replace([l], 'GNV - Frotas')
    elif l == 14:
        novo['segmento'] = novo['segmento'].replace([l], 'GNC')
    elif l == 17:
        novo['segmento'] = novo['segmento'].replace([l], 'Residencial')
    elif l == 13:
        novo['segmento'] = novo['segmento'].replace([l], 'Industrial')
    elif l == 12:
        novo['segmento'] = novo['segmento'].replace([l], 'Industrial')

segmentos = list(novo['segmento'].drop_duplicates())

tabela = [[], [], [], [], [], [], [], [], [], []]
tabela_ca = [[], []]


for s in segmentos:
    tabela[0].append(novo[novo['segmento'] == s]['pis'].sum())
    tabela[1].append(novo[novo['segmento'] == s]['cofins'].sum())
    tabela[2].append(novo[novo['segmento'] == s]['desconto'].sum())
    tabela[3].append(novo[novo['segmento'] == s]['icms'].sum())
    tabela[4].append(novo[novo['segmento'] == s]['icms/st'].sum())
    tabela[5].append(s)
    tabela[6].append(novo[novo['segmento'] == s]['vol ñ fat'].sum())
    tabela[7].append(novo[novo['segmento'] == s]['rec ñ fat'].sum())
    tabela[8].append(novo[novo['segmento'] == s]['for de gn'].sum())
    tabela[9].append(novo[novo['segmento'] == s]['cliente'].count())
    tabela_ca[0].append(novo[novo['segmento'] == s]['rec prov mes'].sum())

#soma3 = novo['icms'].sum()
#print(soma3.round(2))

tabela_nova = pd.DataFrame({'Segmento': tabela[5], 'Nº de Clientes': tabela[9], 'Volume não Faturado': [float(i) for i in tabela[6]],
                            'Forcecimento de GN': [float(i) for i in tabela[8]],
                            'Receita não Faturada': [float(i) for i in tabela[7]],
                            'DESCONTO': [float(i) for i in tabela[2]], 'ICMS': [float(i) for i in tabela[3]],
                            'ICMS/ST': [float(i) for i in tabela[4]], 'PIS': [float(i) for i in tabela[0]],
                            'COFINS': [float(i) for i in tabela[1]]})

tabela_nova.loc['TOTAL GERAL'] = tabela_nova.iloc[:, 1:].sum(axis=0)

tabela_ca = pd.DataFrame({'Segmento': tabela[5], 'Rec Bruta Prov': tabela_ca[0], 'Desconto': tabela[2]})

tabela_ca.loc['TOTAL GERAL'] = tabela_ca.iloc[:, 1:].sum(axis=0)

print(tabela_nova)

tabela_nova.to_excel('Receita_Segmento.xlsx', engine='xlsxwriter')
tabela_ca.to_excel('Relatório GEFIN.xlsx', engine='xlsxwriter')

def formatar_consolidado(item):
    writer = pd.ExcelWriter('Consolidado - 0' + str(date.today().month) + '-' + str(date.today().year) + '.xlsx', engine='xlsxwriter')
    item.to_excel(writer, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'num_format': '0'})


    worksheet.set_column('B:B', 15, format2)
    worksheet.set_column('C:C', 15, format2)
    worksheet.set_column('D:D', 20, format1)
    worksheet.set_column('E:E', 20, format1)
    worksheet.set_column('F:F', 20, format1)
    worksheet.set_column('G:G', 15, format1)
    worksheet.set_column('H:H', 15, format1)
    worksheet.set_column('I:I', 15, format1)
    worksheet.set_column('J:J', 15, format1)
    worksheet.set_column('K:K', 15, format1)

    writer.save()


formatar_consolidado(tabela_nova)
#formatar_consolidado(tabela_ca)
