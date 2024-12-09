# Portifolio 
import requests
from datetime import datetime
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import messagebox


api_key = 'c92e8302f64af150710ecfc4747078a2'


def obter_dados_tempo(cidade):
    url = f'http://api.openweathermap.org/data/2.5/weather?q={cidade}&appid={api_key}&units=metric'
    response = requests.get(url)
    dados = response.json()
    
    if response.status_code == 200:
        temperatura = dados['main']['temp']
        umidade = dados['main']['humidity']
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return data_hora, temperatura, umidade
    else:
        messagebox.showerror("Erro", f"Não foi possível obter os dados para '{cidade}': {dados.get('message', 'Erro desconhecido')}")
        return None, None, None


def salvar_dados_em_planilha(data_hora, cidade, temperatura, umidade):
    nome_arquivo = 'historico_tempo.xlsx'

    try:
        workbook = load_workbook(nome_arquivo)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['Data/Hora', 'Cidade', 'Temperatura (°C)', 'Umidade (%)'])  # Cabeçalho

    sheet.append([data_hora, cidade, temperatura, umidade])
    workbook.save(nome_arquivo)


def capturar_e_exibir():
    cidade = entry_cidade.get().strip()
    if not cidade:
        messagebox.showwarning("Atenção", "Insira o nome da cidade de São Paulo.")
        return
    
    data_hora, temperatura, umidade = obter_dados_tempo(cidade)
    
    if data_hora:
        salvar_dados_em_planilha(data_hora, cidade, temperatura, umidade)
        label_resultado.config(text=f"Data/Hora: {data_hora}\nCidade: {cidade}\nTemperatura: {temperatura}°C\nUmidade: {umidade}%")
        messagebox.showinfo("Sucesso", f"Dados capturados e salvos com sucesso para {cidade}!")


app = tk.Tk()
app.title("Captador de Temperatura")
app.geometry("300x250")

label_instrucoes = tk.Label(app, text="Digite o nome da cidade de São Paulo e clique em buscar")
label_instrucoes.pack(pady=10)

entry_cidade = tk.Entry(app)
entry_cidade.pack(pady=5)

btn_capturar = tk.Button(app, text="Buscaar Cidade", command=capturar_e_exibir)
btn_capturar.pack(pady=10)

label_resultado = tk.Label(app, text="", font=("Arial", 10), justify="left")
label_resultado.pack(pady=10)

app.mainloop()
