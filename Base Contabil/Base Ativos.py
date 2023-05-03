import numpy as np
import pandas as pd
from datetime import datetime

base = pd.read_excel('base.xlsx', dtype={'Subnº': str})
tabela_anla = pd.read_excel('ANLA.xlsx', dtype={'ANLN2': str, 'ANLN1': str, 'AIBN1': str, 'AIBN2': str})
tabela_anlh = pd.read_excel('ANLH.xlsx', dtype={'ANLN1': str})
base = base.rename(columns=lambda x: x.strip())
base = base[base['Subnº'].notnull()]
base2 = base['Imobilizado'] + ' ' + base['Subnº']
base.insert(loc=0, column='Num2', value=base2)

tabela_anla = pd.merge(tabela_anla, tabela_anlh[['ANLN1', 'ANLHTXT']], on=['ANLN1'], how='left')

num_2 = tabela_anla['ANLN1'] + ' ' + tabela_anla['ANLN2']
tabela_anla.insert(loc=0, column='Num2', value=num_2)

base = pd.merge(base,
                tabela_anla[['Num2', 'ANLKL', 'TXT50', 'TXA50', 'ANLHTXT', 'MENGE', 'MEINS', 'SERNR', 'INVNR', 'ORD41',
                             'ORD42', 'ORD43', 'GDLGRP', 'AIBN1', 'AIBN2']], on=['Num2'], how='left')

base['Num2'] = base['AIBN1'] + ' ' + base['AIBN2']

base['Incorporação em'] = pd.to_datetime(base['Incorporação em']).dt.date

base = pd.merge(base, tabela_anla[['Num2', 'POSID', 'EAUFN']], on=['Num2'], how='left')

base['POSID'] = np.where(base['POSID'].isnull(), base['EAUFN'], base['POSID'])

base = base.drop(columns=['Num2', 'Denominação do imobilizado', 'Moeda', 'AIBN1', 'AIBN2', 'EAUFN'], axis=1)

base = base.rename(columns={'ANLKL': 'Classe Imobilizado', 'Incorporação em': 'Data Inic. Deprec.',
                            'ValAquis.': 'Vlr. Aquisi.', 'TXT50': 'Descrição normalizada',
                            'TXA50': 'Descrição (Adicional)',
                            'ANLHTXT': 'Descrição Livre', 'MENGE': 'Quant.', 'MEINS': 'U.M.', 'SERNR': 'Nº Série',
                            'INVNR': 'Nº Inventário', 'ORD41': 'Localidade', 'ORD42': 'SDGN', 'ORD43': 'Tipo Administ.',
                            'GDLGRP': 'Diâmetro', 'POSID': 'OSI'})

tipo_material = ['IES-04', 'IES-06']

base['Tipo de Material'] = np.where(base['Classe Imobilizado'] == 'IES-03', 'AÇO',
                                    (np.where(base['Classe Imobilizado'].isin(tipo_material), 'PEAD', '')))

centro_custo_adm = ['IES-20', 'IES-19', 'IES-32']
centro_custo_ti = ['IES-22', 'IES-34']


base['Centro Custo'] = np.where(base['Classe Imobilizado'].isin(centro_custo_ti), '11330',
                                (np.where(base['Classe Imobilizado'].isin(centro_custo_adm), '11310', '11440')))

base = base[['Imobilizado', 'Subnº', 'Classe Imobilizado', 'Descrição normalizada', 'Descrição (Adicional)',
             'Descrição Livre', 'Quant.', 'U.M.', 'Nº Série', 'Nº Inventário', 'OSI', 'Centro Custo', 'Localidade', 'SDGN',
             'Tipo Administ.',  'Tipo de Material', 'Diâmetro', 'Data Inic. Deprec.', 'Vlr. Aquisi.',
             'Depreciação ac.', 'Valor contábil']]

base['Imobilizado'] = base['Imobilizado'].astype(dtype=int, errors='ignore')
base['Subnº'] = base['Subnº'].astype(int)

writer = pd.ExcelWriter('Base Contábil.xlsx', engine='xlsxwriter', date_format='DD/MM/YYYY')
base.to_excel(writer, sheet_name='Sheet1', index=False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

format1 = workbook.add_format({'num_format': '#,##0.00', 'align': 'left'})
format2 = workbook.add_format({'num_format': '0', 'align': 'left'})
format3 = workbook.add_format({'num_format': '0', 'align': 'left'})


worksheet.set_column('A:A', 11, format2)
worksheet.set_column('B:B', 5, format2)
worksheet.set_column('C:C', 14, format2)
worksheet.set_column('D:D', 45)
worksheet.set_column('E:E', 45)
worksheet.set_column('F:F', 45)
worksheet.set_column('G:G', 9, format1)
worksheet.set_column('H:H', 5)
worksheet.set_column('I:I', 17)
worksheet.set_column('J:J', 12)
worksheet.set_column('K:K', 18)
worksheet.set_column('L:L', 10)
worksheet.set_column('M:M', 8)
worksheet.set_column('N:N', 5)
worksheet.set_column('O:O', 11)
worksheet.set_column('P:P', 13)
worksheet.set_column('Q:Q', 17)
worksheet.set_column('R:R', 15, format3)
worksheet.set_column('S:S', 14, format1)
worksheet.set_column('T:T', 14, format1)
worksheet.set_column('U:U', 14, format1)

writer.close()



