import pandas as pd
from datetime import datetime
import time
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from dateutil.relativedelta import relativedelta
import pyautogui
import traceback
import os
import pyperclip

ARQUIVO_EXCEL = "clientes.xlsx"
IMAGEM_BOTAO = "botao_enviar.png"

# ------------------- CRIA ARQUIVO SE NÃO EXISTIR -------------------
def criar_arquivo_se_nao_existir():
    if not os.path.exists(ARQUIVO_EXCEL):
        df = pd.DataFrame(columns=["Nome", "Telefone", "Data Criacao", "Vencimento", "Pagou"])
        df.to_excel(ARQUIVO_EXCEL, index=False)

criar_arquivo_se_nao_existir()

# ------------------- ENVIO WHATSAPP -------------------
def enviar_mensagem_whatsapp(telefone, mensagem):
    try:
        print(f"▶️ Enviando para {telefone}...")

        # Abre conversa no WhatsApp Web
        url = f"https://web.whatsapp.com/send?phone={telefone}"
        pyautogui.hotkey('ctrl', 'l')  # seleciona a barra de endereços
        pyperclip.copy(url)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        time.sleep(12)  # Espera carregar a conversa

        # Cola e envia a mensagem
        pyperclip.copy(mensagem)
        pyautogui.hotkey("ctrl", "v")

        # Tenta clicar no botão "enviar"
        for tentativa in range(10):
            botao = pyautogui.locateCenterOnScreen(IMAGEM_BOTAO, confidence=0.8)
            if botao:
                pyautogui.click(botao)
                print(f"✅ Mensagem enviada para {telefone}")
                break
            time.sleep(1)
        else:
            pyautogui.press("enter")
            print("⚠️ Botão não encontrado, enviando com ENTER.")

        time.sleep(2)

    except Exception as e:
        traceback.print_exc()
        print(f"❌ Erro ao enviar mensagem para {telefone}: {e}")

# ------------------- VERIFICAR VENCIMENTO E ENVIAR -------------------
def verificaData_com_progresso():
    df = pd.read_excel(ARQUIVO_EXCEL)
    df["Vencimento"] = pd.to_datetime(df["Vencimento"]).dt.date
    df["Pagou"] = df["Pagou"].fillna(False)

    hoje = datetime.today().date()
    clientes_pendentes = df[(df["Vencimento"] == hoje) & (df["Pagou"] == False)]

    total = len(clientes_pendentes)
    if total == 0:
        messagebox.showinfo("Nenhum vencimento hoje", "Nenhum cliente com vencimento hoje e pendente.")
        return

    progresso["maximum"] = total
    progresso["value"] = 0

    for index, row in clientes_pendentes.iterrows():
        nome = row['Nome']
        telefone = str(row['Telefone'])

        if not telefone.startswith('+'):
            telefone = '+55' + telefone

        mensagem = f"Olá {nome}, seu plano vence hoje. Entre em contato para renovar! 📆"
        status_label.config(text=f"Enviando para: {nome}")
        janela.update_idletasks()

        enviar_mensagem_whatsapp(telefone, mensagem)

        progresso["value"] += 1
        janela.update_idletasks()

    status_label.config(text="Envio concluído!")
    messagebox.showinfo("Mensagens Enviadas", "Mensagens enviadas. Agora confirme os pagamentos.")
    exibir_confirmacoes(clientes_pendentes)

# ------------------- CONFIRMAÇÃO DE PAGAMENTO -------------------
def exibir_confirmacoes(clientes):
    def confirmar_pagamento(i):
        nome = clientes.at[i, "Nome"]
        df = pd.read_excel(ARQUIVO_EXCEL)
        idx = df[df["Nome"] == nome].index[0]

        df.at[idx, "Pagou"] = True
        novo_vencimento = datetime.today().date() + relativedelta(months=1)
        df.at[idx, "Vencimento"] = novo_vencimento
        df.to_excel(ARQUIVO_EXCEL, index=False)

        messagebox.showinfo("Pagamento Confirmado", f"{nome} renovado até {novo_vencimento}")
        janela_confirma.destroy()
        atualizar_vencimento_restante()

    def atualizar_vencimento_restante():
        df_atual = pd.read_excel(ARQUIVO_EXCEL)
        df_atual["Vencimento"] = pd.to_datetime(df_atual["Vencimento"]).dt.date
        df_atual["Pagou"] = df_atual["Pagou"].fillna(False)
        pendentes_restantes = df_atual[(df_atual["Vencimento"] == datetime.today().date()) & (df_atual["Pagou"] == False)]
        if not pendentes_restantes.empty:
            exibir_confirmacoes(pendentes_restantes)

    janela_confirma = tk.Toplevel(janela)
    janela_confirma.title("Confirmar Pagamentos")
    janela_confirma.geometry("400x300")

    tk.Label(janela_confirma, text="Clientes com vencimento hoje e pendentes:").pack(pady=10)

    for i, row in clientes.iterrows():
        frame = tk.Frame(janela_confirma)
        frame.pack(fill="x", pady=2, padx=10)

        nome = row["Nome"]
        telefone = row["Telefone"]
        tk.Label(frame, text=f"{nome} - {telefone}", width=30, anchor="w").pack(side="left")
        btn = tk.Button(frame, text="Confirmar Pagamento", command=lambda i=i: confirmar_pagamento(i))
        btn.pack(side="right")

    if len(clientes) == 0:
        tk.Label(janela_confirma, text="Nenhum cliente pendente.").pack(pady=20)

# ------------------- CADASTRO -------------------
def salvar_cliente():
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    data_criacao = datetime.today().date()
    vencimento = data_criacao + relativedelta(months=1)
    pagou = var_pagamento.get()

    if not nome or not telefone:
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    try:
        df = pd.read_excel(ARQUIVO_EXCEL)
    except:
        df = pd.DataFrame(columns=["Nome", "Telefone", "Data Criacao", "Vencimento", "Pagou"])

    novo_cliente = pd.DataFrame([[nome, telefone, data_criacao, vencimento, pagou]],
                                columns=["Nome", "Telefone", "Data Criacao", "Vencimento", "Pagou"])
    df = pd.concat([df, novo_cliente], ignore_index=True)
    df.to_excel(ARQUIVO_EXCEL, index=False)

    messagebox.showinfo("Sucesso", f"Cliente cadastrado com vencimento em {vencimento}")
    entry_nome.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)
    var_pagamento.set(False)
    atualizar_vencimento_automatico()

def atualizar_vencimento_automatico():
    vencimento = (datetime.today().date() + relativedelta(months=1)).strftime("%d/%m/%Y")
    entry_vencimento.config(state="normal")
    entry_vencimento.delete(0, tk.END)
    entry_vencimento.insert(0, vencimento)
    entry_vencimento.config(state="readonly")

def iniciar_envio_em_thread():
    thread = threading.Thread(target=verificaData_com_progresso)
    thread.start()

# ------------------- INTERFACE PRINCIPAL -------------------
janela = tk.Tk()
janela.title("Cadastro e Controle de Clientes")
janela.geometry("420x500")

tk.Label(janela, text="Nome:").pack(pady=5)
entry_nome = tk.Entry(janela, width=35)
entry_nome.pack()

tk.Label(janela, text="Telefone (somente números):").pack(pady=5)
entry_telefone = tk.Entry(janela, width=35)
entry_telefone.pack()

tk.Label(janela, text="Vencimento (automático):").pack(pady=5)
entry_vencimento = tk.Entry(janela, width=35, state="readonly")
entry_vencimento.pack()
atualizar_vencimento_automatico()

var_pagamento = tk.BooleanVar()
check_pagamento = tk.Checkbutton(janela, text="Pagou?", variable=var_pagamento)
check_pagamento.pack(pady=5)

btn_salvar = tk.Button(janela, text="Cadastrar Cliente", command=salvar_cliente)
btn_salvar.pack(pady=10)

btn_verificar = tk.Button(janela, text="Verificar Vencimentos e Enviar WhatsApp", command=iniciar_envio_em_thread)
btn_verificar.pack(pady=10)

progresso = ttk.Progressbar(janela, orient="horizontal", length=300, mode="determinate")
progresso.pack(pady=10)

status_label = tk.Label(janela, text="")
status_label.pack()

janela.mainloop()
