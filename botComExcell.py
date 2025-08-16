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
import webbrowser
from tkcalendar import DateEntry  # <-- IMPORTANTE

ARQUIVO_EXCEL = "clientes.xlsx"

def criar_arquivo_se_nao_existir():
    if not os.path.exists(ARQUIVO_EXCEL):
        df = pd.DataFrame(columns=["Nome", "Telefone", "Data Criacao", "Vencimento", "Pagou"])
        df.to_excel(ARQUIVO_EXCEL, index=False)

criar_arquivo_se_nao_existir()

def abrir_whatsapp_web():
    if not hasattr(abrir_whatsapp_web, "ja_abriu"):
        webbrowser.open("https://web.whatsapp.com")
        print("▶️ Abrindo WhatsApp Web...")
        time.sleep(10)  
        abrir_whatsapp_web.ja_abriu = True

def enviar_mensagem_whatsapp(telefone, mensagem):
    try:
        print(f"▶️ Enviando para {telefone}...")
        abrir_whatsapp_web()

        busca = None
        for _ in range(10):
            busca = pyautogui.locateCenterOnScreen("barra_busca.png", confidence=0.8)
            if busca:
                break
            time.sleep(1)

        if not busca:
            raise Exception("❌ Campo de busca do WhatsApp Web não encontrado. Verifique a imagem 'barra_busca.png'.")

        pyautogui.click(busca)
        time.sleep(1)

        pyautogui.write(telefone)
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(2)

        pyautogui.write(mensagem)
        time.sleep(1)
        pyautogui.press("enter")

        print(f"✅ Mensagem enviada para {telefone}")
        time.sleep(2)

    except Exception as e:
        traceback.print_exc()
        print(f"❌ Erro ao enviar mensagem para {telefone}: {e}")

def verificaData_com_progresso():
    abrir_whatsapp_web()

    df = pd.read_excel(ARQUIVO_EXCEL)
    df["Vencimento"] = pd.to_datetime(df["Vencimento"]).dt.date
    df["Pagou"] = df["Pagou"].fillna(False)

    hoje = datetime.today().date()
    clientes_pendentes = df[(df["Vencimento"] <= hoje) & (df["Pagou"] == False)]

    total = len(clientes_pendentes)
    if total == 0:
        messagebox.showinfo("Nenhum vencimento", "Nenhum cliente pendente ou vencido encontrado.")
        return

    progresso["maximum"] = total
    progresso["value"] = 0

    for index, row in clientes_pendentes.iterrows():
        nome = row['Nome']
        telefone = str(row['Telefone'])

        if not telefone.startswith('+'):
            telefone = '+55' + telefone

        mensagem = f"Olá {nome}, seu plano venceu em {row['Vencimento']}. Entre em contato para renovar! 📆"
        status_label.config(text=f"Enviando para: {nome}")
        janela.update_idletasks()

        enviar_mensagem_whatsapp(telefone, mensagem)

        progresso["value"] += 1
        janela.update_idletasks()

    status_label.config(text="Envio concluído!")
    messagebox.showinfo("Mensagens Enviadas", "Mensagens enviadas. Agora confirme os pagamentos.")
    exibir_confirmacoes(clientes_pendentes)

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

    def deletar_cliente(i):
        nome = clientes.at[i, "Nome"]
        df = pd.read_excel(ARQUIVO_EXCEL)
        df = df[df["Nome"] != nome]
        df.to_excel(ARQUIVO_EXCEL, index=False)
        messagebox.showinfo("Cliente Deletado", f"{nome} foi removido do sistema.")
        janela_confirma.destroy()
        atualizar_vencimento_restante()

    def atualizar_vencimento_restante():
        df_atual = pd.read_excel(ARQUIVO_EXCEL)
        df_atual["Vencimento"] = pd.to_datetime(df_atual["Vencimento"]).dt.date
        df_atual["Pagou"] = df_atual["Pagou"].fillna(False)
        pendentes_restantes = df_atual[(df_atual["Vencimento"] <= datetime.today().date()) & (df_atual["Pagou"] == False)]
        if not pendentes_restantes.empty:
            exibir_confirmacoes(pendentes_restantes)

    janela_confirma = tk.Toplevel(janela)
    janela_confirma.title("Confirmar Pagamentos")
    janela_confirma.geometry("500x400")

    tk.Label(janela_confirma, text="Clientes vencidos ou vencendo hoje:").pack(pady=10)

    for i, row in clientes.iterrows():
        frame = tk.Frame(janela_confirma)
        frame.pack(fill="x", pady=2, padx=10)

        nome = row["Nome"]
        telefone = row["Telefone"]
        vencimento = row["Vencimento"]
        tk.Label(frame, text=f"{nome} - {telefone} | Vencido desde: {vencimento}", width=50, anchor="w").pack(side="left")

        btn_confirma = tk.Button(frame, text="Confirmar Pagamento", command=lambda i=i: confirmar_pagamento(i))
        btn_confirma.pack(side="right", padx=5)

        btn_delete = tk.Button(frame, text="Deletar Cliente", command=lambda i=i: deletar_cliente(i))
        btn_delete.pack(side="right")

    if len(clientes) == 0:
        tk.Label(janela_confirma, text="Nenhum cliente pendente.").pack(pady=20)

def salvar_cliente():
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    data_criacao = datetime.today().date()
    vencimento = entry_vencimento.get_date()  # <-- pega a data do calendário

    if not nome or not telefone or not vencimento:
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    pagou = var_pagamento.get()

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

def iniciar_envio_em_thread():
    thread = threading.Thread(target=verificaData_com_progresso)
    thread.start()

# ------------------- INTERFACE -------------------
janela = tk.Tk()
janela.title("Cadastro e Controle de Clientes")
janela.geometry("480x550")

tk.Label(janela, text="Nome:").pack(pady=5)
entry_nome = tk.Entry(janela, width=35)
entry_nome.pack()

tk.Label(janela, text="Telefone (somente números):").pack(pady=5)
entry_telefone = tk.Entry(janela, width=35)
entry_telefone.pack()

tk.Label(janela, text="Data de Vencimento:").pack(pady=5)
entry_vencimento = DateEntry(janela, width=32, background='darkblue',
                             foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
entry_vencimento.pack(pady=5)

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
