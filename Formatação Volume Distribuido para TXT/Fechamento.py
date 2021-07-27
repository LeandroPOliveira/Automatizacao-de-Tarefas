import pandas as pd

desired_width=320

pd.set_option('display.width', desired_width)

pd.set_option('display.max_columns', 10)


data = pd.read_excel('Fechamento_Mês.xlsx', sheet_name='Fechamento', dtype=str)

data = pd.DataFrame(data)

data['Volume de Gás Distribuido'] = data['Volume de Gás Distribuido'].str.replace(".",",", regex=True)

data['Cliente'] = data['Cliente'].apply(lambda x: '{0:0>10}'.format(x))

data['numero'] = ['01' for i in range(data.shape[0])]

lista = data.columns.tolist()

data = data[['Período', 'Organização de Vendas', 'Canal de Distribuição', 'Segmento', 'numero', 'Cliente',
                'Nome do Cliente', 'Município', 'Volume de Gás Distribuido']]

data.replace(to_replace='GUAIÇARA', value ='GUAICARA', inplace=True)
data.replace(to_replace='AMÉRICO BRASILIENSE', value='AMERICO BRASILIENSE', inplace=True)

data.dropna(inplace=True)

print(data)

data.to_csv('Fechamento_Volumes.txt', sep=';', index=False, header=None, encoding='utf8')

