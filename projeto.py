import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename, asksaveasfilename
import tkinter.messagebox as messagebox
import pandas as pd
import numpy as np
import requests
from datetime import datetime

# Requisição inicial para obter a lista de moedas
requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()
print(dicionario_moedas)

lista_moedas = list(dicionario_moedas.keys())

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminho_arquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivo_selecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"
    
def pegar_cotacao():
    moeda = combobox_selecionar_moeda.get()
    data_cotacao = calendario_moeda.get()
    
    if moeda not in lista_moedas:
        label_texto_cotacao['text'] = "Moeda não válida. Selecione outra."
        return
    
    ano, mes, dia = data_cotacao[-4:], data_cotacao[3:5], data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    
    try:
        requisicao_moeda = requests.get(link)
        requisicao_moeda.raise_for_status()
        cotacao = requisicao_moeda.json()
        
        if not cotacao:
            label_texto_cotacao['text'] = "Nenhuma cotação encontrada para esta data."
            return
        
        valor_moeda = cotacao[0]['bid']
        label_texto_cotacao['text'] = f"A cotação da moeda '{moeda}' no dia {data_cotacao} foi de: R${valor_moeda}"
    
    except requests.exceptions.RequestException as e:
        label_texto_cotacao['text'] = f"Erro ao buscar cotação: {e}"
    

def atualizar_cotacoes():
    try:
        # Ler o dataframe de moedas
        df = pd.read_excel(var_caminho_arquivo.get())
        moedas = df.iloc[:, 0]
        
        # Pegar a data inicial e data final das cotações
        data_inicial = calendario_data_inicial.get()
        data_final = calendario_data_final.get()

        # Extrair ano, mês e dia da data inicial
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        # Extrair ano, mês e dia da data final
        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        # Gerar lista de datas entre a data inicial e a data final
        datas = pd.date_range(start=f"{ano_inicial}-{mes_inicial}-{dia_inicial}", 
                              end=f"{ano_final}-{mes_final}-{dia_final}")

        for moeda in moedas:
            for data in datas:
                data_formatada = data.strftime('%d/%m/%Y')
                
                # Formatar a URL para cada data
                link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?" \
                       f"start_date={data.strftime('%Y%m%d')}&end_date={data.strftime('%Y%m%d')}"
                
                requisicao_moeda = requests.get(link)
                cotacoes = requisicao_moeda.json()
                
                if cotacoes:
                    # Assume que sempre haverá uma cotação para a data solicitada
                    bid = float(cotacoes[0]['bid'])
                    if data_formatada not in df.columns:
                        df[data_formatada] = np.nan

                    df.loc[df.iloc[:, 0] == moeda, data_formatada] = bid

        # Avisar o usuário que ele deve nomear e escolher onde salvar o arquivo
        messagebox.showinfo("Salvar Arquivo", "Escolha um local e nome para salvar o arquivo gerado.")

        # Perguntar ao usuário onde salvar o arquivo
        caminho_arquivo = asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if caminho_arquivo:
            df.to_excel(caminho_arquivo, index=False)
            label_atualizar_cotacoes['text'] = "Arquivo Atualizado com Sucesso"
        else:
            label_atualizar_cotacoes['text'] = "Salvar arquivo cancelado."
    except Exception as e:
        label_atualizar_cotacoes['text'] = f"Erro: {e}"



# Configuração da interface gráfica
janela = tk.Tk()
janela.title('Ferramenta de Cotação De Moedas')


# Interface para pegar cotação específica
label_cotacao_moeda = tk.Label(text="Cotação de moeda específica", borderwidth=2, relief='solid')
label_cotacao_moeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionar_moeda = tk.Label(text="Selecionar Moeda", anchor='e')
label_selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionar_moeda = ttk.Combobox(values=lista_moedas)
combobox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

label_selecionar_dia = tk.Label(text="Selecione o dia que deseja pegar a cotação", anchor='e')
label_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moeda = DateEntry(year=2024, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

label_texto_cotacao = tk.Label(text="")
label_texto_cotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

botao_pegar_cotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')

# Interface para cotação de várias moedas
label_cotacao_varias_moedas = tk.Label(text="Cotação de Múltiplas Moedas", borderwidth=2, relief='solid')
label_cotacao_varias_moedas.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionar_arquivo = tk.Label(text="Selecione um arquivo em Excel com as Moedas na Coluna A:")
label_selecionar_arquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

var_caminho_arquivo = tk.StringVar()

botao_selecionar_arquivo = tk.Button(text='Clique para Selecionar', command=selecionar_arquivo)
botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivo_selecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
label_arquivo_selecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_data_inicial = tk.Label(text="Data Inicial", anchor='e')
label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky="nsew")

label_data_final = tk.Label(text="Data Final", anchor='e')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky="nsew")

calendario_data_inicial = DateEntry(year=2024, locale='pt_br')
calendario_data_inicial.grid(row=7, column=1, padx=10, pady=10, sticky="nsew")

calendario_data_final = DateEntry(year=2024, locale='pt_br')
calendario_data_final.grid(row=8, column=1, padx=10, pady=10, sticky="nsew")

botao_atualizar_cotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0, padx=10, pady=10, sticky="nsew")

label_atualizar_cotacoes = tk.Label(text="")
label_atualizar_cotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()