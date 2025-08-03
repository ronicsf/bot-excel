import pandas as pd
from datetime import datetime
import pywhatkit as kit
import time
import tkinter as tk
from tkinter import messagebox

ARQUIVO_EXCEL = "clientes.xlsx"

# Garante que o arquivo exista
def criar_arquivo_se_nao_existir():
    try:
        pd.read_excel(ARQUIVO_EXCEL)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Nome", "Telefone", "Data Criacao"])
        df.to_excel(ARQUIVO_EXCEL, index=False)

criar_arquivo_se_nao_existir()

# ------------------- ENVIO WHATSAPP -------------------
def enviar_mensagem_whatsapp(telefone, mensagem):
    try:
        kit.sendwhatmsg_instantly(telefone, mensagem, wait_time=10, tab_close=True)
        print(f"✅ Mensagem enviada para {telefone}")
        time.sleep(10)
    except Exception as e:
        print(f"❌ Erro ao enviar mensagem para {telefone}: {e}")

def verificaData():
    tabela = pd.read_excel(ARQUIVO_EXCEL)
    tabela["Data Criacao"] = pd.to_datetime(tabela["Data Criacao"]).dt.date
    data_hoje = datetime.today().date()
    encontrou = False

    for index, row in tabela.iterrows():
        if row["Data Criacao"] == data_hoje:
            nome = row['Nome']
            telefone = str(row['Telefone'])

            if not telefone.startswith('+'):
                telefone = '+55' + telefone

            mensagem = f"Olá {nome}, parabéns pelo seu aniversário de cadastro! 🎉"
            enviar_mensagem_whatsapp(telefone, mensagem)
            encontrou = True

    if not encontrou:
        print("ℹ️ Nenhuma data bate com a data de hoje.")

# ------------------- GUI CADASTRO -------------------
def salvar_cliente():
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    data_criacao = datetime.today().date()

    if not nome or not telefone:
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    # Lê o arquivo e adiciona o novo cliente
    try:
        df = pd.read_excel(ARQUIVO_EXCEL)
    except:
        df = pd.DataFrame(columns=["Nome", "Telefone", "Data Criacao"])

    novo_cliente = pd.DataFrame([[nome, telefone, data_criacao]], columns=["Nome", "Telefone", "Data Criacao"])
    df = pd.concat([df, novo_cliente], ignore_index=True)
    df.to_excel(ARQUIVO_EXCEL, index=False)

    messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")
    entry_nome.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)

# Interface Tkinter
janela = tk.Tk()
janela.title("Cadastro de Clientes")
janela.geometry("300x200")

tk.Label(janela, text="Nome:").pack(pady=5)
entry_nome = tk.Entry(janela, width=30)
entry_nome.pack()

tk.Label(janela, text="Telefone:").pack(pady=5)
entry_telefone = tk.Entry(janela, width=30)
entry_telefone.pack()

btn_salvar = tk.Button(janela, text="Cadastrar Cliente", command=salvar_cliente)
btn_salvar.pack(pady=15)

btn_verificar = tk.Button(janela, text="Verificar Datas e Enviar WhatsApp", command=verificaData)
btn_verificar.pack(pady=5)


janela.mainloop()
