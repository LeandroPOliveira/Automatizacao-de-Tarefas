import pandas as pd
from tkinter import *
from tkinter import messagebox, filedialog as fd
from tkinter.filedialog import asksaveasfilename
import tkinter.messagebox


class ProvisaoReceita:

    def __init__(self, janela):
        self.local2 = None
        self.local = None
        self.dados = None
        self.janela = janela
        self.janela.geometry('600x600')
        self.frame1 = Frame(self.janela, height=600, width=600, bg='white').place(x=0, y=0)
        self.borda = Canvas(self.janela, height=400, width=400).place(x=100, y=50)
        self.label = Label(self.frame1, text='Consolidado Provisão Automação-de-Tarefas', font=('arial', 16, 'bold'),
                     bg='white').place(x=130, y=100)

        # botão procurar
        Label(self.frame1, text='Abrir Arquivo Origem ".CSV"').place(x=120, y=150)
        self.btn_origem = Button(self.frame1, text='Procurar', width=10, command=self.abrir).place(x=350, y=180)
        self.entrada_origem = Entry(self.frame1, bd=1)
        self.entrada_origem.place(x=120, y=180, width=200)

        # botão salvar arquivo convertido
        self.label_resumo = Label(self.frame1, text='Resumo Provisão da Automação-de-Tarefas').place(x=120, y=230)
        self.btn_resumo = Button(self.frame1, text='Salvar como', width=10, command=self.salvar).place(x=350, y=260)
        self.entrada_resumo = Entry(self.frame1, bd=1)
        self.entrada_resumo.place(x=120, y=260, width=200)

        # botão salvar arquivo convertido GEFIN
        Label(self.frame1, text='Resumo GEFIN').place(x=120, y=290)
        self.btn_resumo2 = Button(self.frame1, text='Salvar como', width=10, command=self.salvar_2).place(x=350, y=320)
        self.entrada_resumo2 = Entry(self.frame1, bd=1)
        self.entrada_resumo2.place(x=120, y=320, width=200)

        # botão para executar script
        self.btn_converter = Button(self.frame1, text='Converter Relatório', command=self.calcular)
        self.btn_converter.place(x=235, y=380)


    def abrir(self):
        self.dados = fd.askopenfilename(title='Abrir arquivo', initialdir='\C:')
        self.entrada_origem.delete(0, END)
        self.entrada_origem.insert(0, self.dados)

    # selecionar local do arquivo convertido
    def salvar(self):
        files = [("Excel files", "*.xlsx")]
        self.local = asksaveasfilename(filetypes=files, defaultextension=files[0][1])
        self.alterado = self.entrada_resumo.insert(0, self.local)

    def salvar_2(self):
        files = [("Excel files", "*.xlsx")]
        self.local2 = asksaveasfilename(filetypes=files, defaultextension=files[0][1])
        self.alterado2 = self.entrada_resumo2.insert(0, self.local2)

    def calcular(self):
        dataset = pd.read_csv(self.dados, sep=';', decimal=',', header=None)
        pd.options.display.float_format = "{:,.2f}".format
        dados = pd.DataFrame(dataset)

        dados.columns = ['cliente', 'nome', 'cidade', 'segmento', 'vol prov ini', 'vol prov mes', 'vol prov fin',
                         'rec prov ini', 'rec prov mes', 'rec prov fin', 'vol fat ini', 'vol fat mes', 'vol fat fin',
                         'rec fat ini', 'rec fat mes', 'rec fat fin', 'vol ñ fat', 'rec ñ fat', 'pis', 'cofins',
                         'desconto', 'icms', 'icms/st']

        dados.index = ['linha' + str(i) for i in range(len(dados))]

        novo = pd.DataFrame(dados)

        novo['for de gn'] = novo.apply(lambda x: x['rec ñ fat'] - x['desconto'] + x['icms/st'], axis=1)

        sistema_2 = ['PORTO FERREIRA', 'SAO CARLOS', 'DESCALVADO']

        selecao = novo['cidade'].isin(sistema_2)

        dados_sistema_2 = novo[selecao]

        dados_sistema_2.index = ['linha' + str(i) for i in range(len(dados_sistema_2))]

        segmentos = list(novo['segmento'].drop_duplicates())

        for segmento in segmentos:
            if segmento == 2:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Res. Coletivo')
            elif segmento == 1:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Residencial')
            elif segmento == 3:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Comercial')
            elif segmento == 4:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Industrial')
            elif segmento == 5:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Industrial')
            elif segmento == 6:
                novo['segmento'] = novo['segmento'].replace([segmento], 'GNV')
            elif segmento == 8:
                novo['segmento'] = novo['segmento'].replace([segmento], 'GNV - Frotas')
            elif segmento == 14:
                novo['segmento'] = novo['segmento'].replace([segmento], 'GNC')
            elif segmento == 17:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Residencial')
            elif segmento == 13:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Industrial')
            elif segmento == 12:
                novo['segmento'] = novo['segmento'].replace([segmento], 'Industrial')

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

        tabela_nova = pd.DataFrame({'Segmento': tabela[5], 'Nº de Clientes': tabela[9], 'Volume não Faturado':
                                    [float(i) for i in tabela[6]],
                                    'Forcecimento de GN': [float(i) for i in tabela[8]],
                                    'Receita não Faturada': [float(i) for i in tabela[7]],
                                    'DESCONTO': [float(i) for i in tabela[2]], 'ICMS': [float(i) for i in tabela[3]],
                                    'ICMS/ST': [float(i) for i in tabela[4]], 'PIS': [float(i) for i in tabela[0]],
                                    'COFINS': [float(i) for i in tabela[1]]})

        tabela_nova.loc['TOTAL GERAL'] = tabela_nova.iloc[:, 1:].sum(axis=0)

        tabela_ca = pd.DataFrame({'Segmento': tabela[5], 'Rec Bruta Prov': tabela_ca[0], 'Desconto': tabela[2]})

        tabela_ca.loc['TOTAL GERAL'] = tabela_ca.iloc[:, 1:].sum(axis=0)

        tabela_nova.to_excel(self.local, engine='xlsxwriter')
        tabela_ca.to_excel(self.local2, engine='xlsxwriter')

        def formatar_consolidado(item):
            writer = pd.ExcelWriter(self.local, engine='xlsxwriter')
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
        tkinter.messagebox.showinfo('', 'Arquivo Gerado com Sucesso!')


if __name__=='__main__':
    janela = Tk()
    aplicacao = ProvisaoReceita(janela)
    janela.mainloop()