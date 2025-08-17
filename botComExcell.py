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
import base64
from tkcalendar import DateEntry   # 📅 calendário

ARQUIVO_EXCEL = "clientes.xlsx"
ARQUIVO_LICENCA = "licenca.txt"

# =================== SISTEMA DE LICENÇA ===================
def verificar_licenca():
    if not os.path.exists(ARQUIVO_LICENCA):
        return False
    try:
        with open(ARQUIVO_LICENCA, "r") as f:
            chave = f.read().strip()
        conteudo = base64.b64decode(chave).decode()
        if not conteudo.startswith("EXPIRA:"):
            return False
        data_expira = datetime.strptime(conteudo.split(":")[1], "%Y-%m-%d").date()
        return datetime.today().date() <= data_expira
    except:
        return False

def salvar_licenca(chave):
    with open(ARQUIVO_LICENCA, "w") as f:
        f.write(chave)

def pedir_licenca():
    def validar_chave():
        chave = entry_chave.get().strip()
        try:
            conteudo = base64.b64decode(chave).decode()
            if not conteudo.startswith("EXPIRA:"):
                raise Exception("Formato inválido")
            data_expira = datetime.strptime(conteudo.split(":")[1], "%Y-%m-%d").date()
            if datetime.today().date() > data_expira:
                messagebox.showerror("Licença Expirada", "Sua licença expirou. Solicite uma nova.")
                return
            salvar_licenca(chave)
            messagebox.showinfo("Sucesso", f"Licença válida até {data_expira}")
            licenca_janela.destroy()
            iniciar_programa()
        except:
            messagebox.showerror("Erro", "Chave inválida.")

    licenca_janela = tk.Tk()
    licenca_janela.title("Ativação Necessária")
    licenca_janela.geometry("400x200")
    tk.Label(licenca_janela, text="Insira sua chave de licença:").pack(pady=10)
    entry_chave = tk.Entry(licenca_janela, width=45)
    entry_chave.pack(pady=5)
    tk.Button(licenca_janela, text="Ativar", command=validar_chave).pack(pady=10)
    licenca_janela.mainloop()

# =================== FUNÇÕES CLIENTES ===================
def criar_arquivo_se_nao_existir():
    if not os.path.exists(ARQUIVO_EXCEL):
        df = pd.DataFrame(columns=["Nome", "Telefone", "Data Criacao", "Vencimento", "Pagou"])
        df.to_excel(ARQUIVO_EXCEL, index=False)

def abrir_whatsapp_web():
    if not hasattr(abrir_whatsapp_web, "ja_abriu"):
        webbrowser.open("https://web.whatsapp.com")
        time.sleep(10)
        abrir_whatsapp_web.ja_abriu = True

def enviar_mensagem_whatsapp(telefone, mensagem):
    try:
        abrir_whatsapp_web()
        busca = None
        for _ in range(10):
            busca = pyautogui.locateCenterOnScreen("barra_busca.png", confidence=0.8)
            if busca: break
            time.sleep(1)
        if not busca:
            raise Exception("❌ Campo de busca não encontrado.")
        pyautogui.click(busca)
        time.sleep(1)
        pyautogui.write(telefone)
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.write(mensagem)
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(2)
    except Exception as e:
        traceback.print_exc()
        print(f"❌ Erro ao enviar mensagem: {e}")

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
        # 🔄 Atualiza vencimento corretamente
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
        messagebox.showinfo("Cliente Removido", f"{nome} foi deletado do sistema.")
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
    janela_confirma.geometry("700x400")
    tk.Label(janela_confirma, text="Clientes vencidos ou vencendo hoje:").pack(pady=10)

    for i, row in clientes.iterrows():
        frame = tk.Frame(janela_confirma)
        frame.pack(fill="x", pady=2, padx=10)
        nome = row["Nome"]
        telefone = row["Telefone"]
        vencimento = row["Vencimento"]
        tk.Label(frame, text=f"{nome} - {telefone} | Vencimento: {vencimento}", width=50, anchor="w").pack(side="left")
        tk.Button(frame, text="Confirmar Pagamento", command=lambda i=i: confirmar_pagamento(i)).pack(side="right", padx=5)
        tk.Button(frame, text="Deletar", command=lambda i=i: deletar_cliente(i)).pack(side="right")

def salvar_cliente():
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    data_criacao = datetime.today().date()
    try:
        vencimento = entry_vencimento.get_date()  # 📅 agora usa DateEntry
    except:
        messagebox.showerror("Erro", "Data de vencimento inválida.")
        return
    pagou = False
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
    messagebox.showinfo("Sucesso", f"Cliente cadastrado com vencimento em {vencimento.strftime('%d/%m/%Y')}")
    entry_nome.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)

def iniciar_envio_em_thread():
    thread = threading.Thread(target=verificaData_com_progresso)
    thread.start()

# =================== INTERFACE ===================
def iniciar_programa():
    global janela, entry_nome, entry_telefone, entry_vencimento, progresso, status_label
    criar_arquivo_se_nao_existir()
    janela = tk.Tk()
    janela.title("Cadastro e Controle de Clientes")
    janela.geometry("700x550")

    tk.Label(janela, text="Nome:").pack(pady=5)
    entry_nome = tk.Entry(janela, width=40)
    entry_nome.pack()

    tk.Label(janela, text="Telefone (somente números):").pack(pady=5)
    entry_telefone = tk.Entry(janela, width=40)
    entry_telefone.pack()

    tk.Label(janela, text="Vencimento:").pack(pady=5)
    entry_vencimento = DateEntry(janela, width=37, date_pattern="dd/mm/yyyy")  # 📅 seletor de data
    entry_vencimento.pack()

    tk.Button(janela, text="Cadastrar Cliente", command=salvar_cliente).pack(pady=10)
    tk.Button(janela, text="Verificar Vencimentos e Enviar WhatsApp", command=iniciar_envio_em_thread).pack(pady=10)

    progresso = ttk.Progressbar(janela, orient="horizontal", length=350, mode="determinate")
    progresso.pack(pady=10)
    status_label = tk.Label(janela, text="")
    status_label.pack()
    janela.mainloop()

# =================== MAIN ===================
if __name__ == "__main__":
    if verificar_licenca():
        iniciar_programa()
    else:
        pedir_licenca()
