import pandas as pd
from tkinter import *
from tkinter import messagebox, filedialog as fd
from tkinter.filedialog import asksaveasfilename


class FechamentoVolumes:

    def __init__(self, fechamento):
        self.fechamento = fechamento
        self.fechamento.geometry('500x500+500+100')
        self.fechamento.title('Fechamento de Volumes')
        self.fechamento.resizable(0,0)

        # estilo do cores do frame
        x1 = 0
        c1 = 0
        for i in range(100):
            c = str(222222+c1)
            Frame(fechamento, width=10, height=500, bg='#'+c).place(x=x1, y=0)
            x1 += 10
            c1 += 1

        # Frame e título
        Frame(fechamento, width=400, height=400, bg='white').place(x=50, y=50)
        self.fonte = ('consolas', 11)
        self.l1 = Label(fechamento, text='Relatório de Fechamento de Volumes', bg='white', font=self.fonte).place(x=70, y=100)

        # botão procurar
        self.b1 = Button(fechamento, text='Procurar', width=10, command=self.abrir).place(x=330, y=180)
        self.e1 = Entry(fechamento, bd=1)
        self.e1.place(x=70, y=180, width=250)

        # botão salvar arquivo convertido
        self.b2 = Button(fechamento, text='Salvar como', width=10, command=self.salvar).place(x=330, y=280)
        self.e2 = Entry(fechamento, bd=1)
        self.e2.place(x=70, y=280, width=250)

        # botão para executar script
        self.b3 = Button(fechamento, text='Converter Relatório', font=self.fonte, command=self.gerar)
        self.b3.place(x=185, y=360)


    # selecionar arquivo excel dos volumes
    def abrir(self):
        self.dados = fd.askopenfilename(title='Abrir arquivo', initialdir='\C:')
        self.e1.delete(0, END)
        self.e1.insert(0, self.dados)


    # selecionar local do arquivo convertido
    def salvar(self):
        files = [('Text Document', '*.txt')]
        self.local = asksaveasfilename(filetypes=files, defaultextension=files)
        self.alterado = self.e2.insert(0, self.local)


    # executar conversão e formatações
    def gerar(self):
        self.data = pd.read_excel(self.dados, sheet_name='Fechamento', dtype=str)
        self.data = pd.DataFrame(self.data)
        self.data['Volume de Gás Distribuido'] = self.data['Volume de Gás Distribuido'].str.replace(".",",", regex=True)
        self.data['Cliente'] = self.data['Cliente'].apply(lambda x: '{0:0>10}'.format(x))
        self.data['numero'] = ['01' for i in range(self.data.shape[0])]
        lista = self.data.columns.tolist()
        self.data = self.data[['Período', 'Organização de Vendas', 'Canal de Distribuição', 'Segmento', 'numero', 'Cliente',
                        'Nome do Cliente', 'Município', 'Volume de Gás Distribuido']]
        self.data.replace(to_replace='GUAIÇARA', value ='GUAICARA', inplace=True)
        self.data.replace(to_replace='AMÉRICO BRASILIENSE', value='AMERICO BRASILIENSE', inplace=True)
        self.data.replace(to_replace='ITÁPOLIS', value='ITAPOLIS', inplace=True)
        self.data.dropna(inplace=True)
        self.data.to_csv(self.local, sep=';', index=False, header=None, encoding='utf8')
        messagebox.showinfo('', 'Relatório gerado com sucesso!')


if __name__=='__main__':
    fechamento = Tk()
    aplicacao = FechamentoVolumes(fechamento)
    fechamento.mainloop()