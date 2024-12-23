import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import requests
from tkinter.filedialog import askopenfilename
import pandas as pd
from datetime import datetime
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegar_cotacao():    
    moeda = combobox_selecionar_moeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'    
    requisicao_moeda =  requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_texto_cotacao['text'] = f'A cotação do {moeda} no dia {data_cotacao} foi de : R${valor_moeda}' 

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de moeda')
    var_caminho_arquivo.set(caminho_arquivo)
    
    if caminho_arquivo:
        label_arquivo_selecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"

def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminho_arquivo.get())
        moedas = df.iloc[:,0]
        data_inicial = calendario_data_inicial.get()
        data_final = calendario_data_final.get()
        
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]
        
        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}'    
            print(link)
            requisicao_moeda =  requests.get(link)
            print(requisicao_moeda)
            cotacoes = requisicao_moeda.json()
            print(cotacoes)
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime("%d/%m/%Y")
                if data not in df:
                    print(data)
                    df[data] = np.nan
                    
                df.loc[df.iloc[:, 0]==moeda,data] = bid
        df.to_excel('Teste.xlsx')
        label_atualizar_cotacoes['text'] = "Arquivo Atualizado Com Sucesso"
    except:
        label_atualizar_cotacoes['text'] = "Selecione um Arquivo Excel no Formato correto"
        


janela = tk.Tk()
janela.title('Ferramenta de Cotação de Moedas')

label_cotacao_moeda = ttk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
label_cotacao_moeda.grid(row=0,column=0, padx=10,pady=10, sticky='nsew', columnspan=3)

label_selecionar_moeda = ttk.Label(text='Selecionar Moeda', anchor='e')
label_selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

combobox_selecionar_moeda = ttk.Combobox(values=lista_moedas)
combobox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

label_selecionar_dia = ttk.Label(text='Selecione o dia que deseje pegar a cotação', anchor='e')
label_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)



calendario_moeda = DateEntry(year=2024, locale='pt_br')
calendario_moeda.grid(row= 2, column= 2, padx= 10, pady= 10, sticky= 'nsew')

label_texto_cotacao = ttk.Label(text="")
label_texto_cotacao.grid(row=3, column=0, columnspan=2, padx=10,pady=10, sticky='nsew')

botao_pegar_cotacao = ttk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

#cotação de várias moedas

label_cotacao_varias_moedas = ttk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
label_cotacao_varias_moedas.grid(row=4,column=0, padx=10,pady=10, sticky='nsew', columnspan=3)

label_selecionar_arquivo = ttk.Label(text='Selecione um arquivo em Excel com as Moedas na Coluna A')
label_selecionar_arquivo.grid(row=5, column=0,columnspan=2, padx=10,pady=10, sticky='nswe')

var_caminho_arquivo = tk.StringVar()



botao_selecionar_arquivo = ttk.Button(text='Clique para Selecionar', command=selecionar_arquivo)
botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivo_selecionado = ttk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivo_selecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_data_inicial = ttk.Label(text='Data Inicial',anchor = 'e')
label_data_final = ttk.Label(text='Data Final',anchor = 'e')

label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')


calendario_data_inicial = DateEntry(year = 2024, locale= "pt_br")
calendario_data_final = DateEntry(year = 2024, locale= "pt_br")

calendario_data_inicial.grid(row=7, column = 1, padx= 10, pady = 10, sticky = 'nswe')
calendario_data_final.grid(row=8, column = 1, padx= 10, pady = 10, sticky = 'nswe')

botao_atualizar_cotacoes = ttk.Button(text='Atualizar Cotacoes', command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizar_cotacoes = ttk.Label(text='')
label_atualizar_cotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = ttk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10 , column=2, padx=10, pady=10, sticky='nswe')



janela.mainloop()