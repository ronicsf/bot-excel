import pandas as pd
from datetime import datetime, timedelta
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

COLUNAS = ["Nome", "Telefone", "Data Criacao", "Vencimento", "Pagou"]

# =================== UTIL: DADOS ===================
def criar_arquivo_se_nao_existir():
    if not os.path.exists(ARQUIVO_EXCEL):
        df = pd.DataFrame(columns=COLUNAS)
        df.to_excel(ARQUIVO_EXCEL, index=False)

def carregar_df():
    """Carrega Excel garantindo colunas e tipos."""
    criar_arquivo_se_nao_existir()
    df = pd.read_excel(ARQUIVO_EXCEL)
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = None
    # Tipos
    df = df[COLUNAS]
    if not df.empty:
        df["Vencimento"] = pd.to_datetime(df["Vencimento"]).dt.date
        df["Pagou"] = df["Pagou"].fillna(False).astype(bool)
    return df

def salvar_df(df):
    df = df[COLUNAS]
    df.to_excel(ARQUIVO_EXCEL, index=False)

def encontrar_indice(df, nome, telefone):
    """Tenta encontrar a linha usando Nome + Telefone; se não achar, cai no primeiro Nome."""
    match = df[(df["Nome"] == nome) & (df["Telefone"].astype(str) == str(telefone))]
    if not match.empty:
        return match.index[0]
    match2 = df[df["Nome"] == nome]
    return match2.index[0] if not match2.empty else None

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

# =================== WHATSAPP ===================
def abrir_whatsapp_web():
    """Abre WhatsApp Web apenas uma vez."""
    if not hasattr(abrir_whatsapp_web, "ja_abriu"):
        webbrowser.open("https://web.whatsapp.com")
        time.sleep(10)
        abrir_whatsapp_web.ja_abriu = True

def enviar_mensagem_whatsapp(telefone, mensagem):
    """
    Requer a imagem 'barra_busca.png' (print do campo de busca do WhatsApp Web) na mesma pasta.
    """
    try:
        abrir_whatsapp_web()
        busca = None
        for _ in range(15):
            busca = pyautogui.locateCenterOnScreen("barra_busca.png", confidence=0.8)
            if busca:
                break
            time.sleep(1)
        if not busca:
            raise Exception("❌ Campo de busca não encontrado (barra_busca.png).")
        pyautogui.click(busca)
        time.sleep(0.8)
        pyautogui.write(telefone)
        time.sleep(1.8)
        pyautogui.press("enter")
        time.sleep(1.5)
        pyautogui.write(mensagem)
        time.sleep(0.6)
        pyautogui.press("enter")
        time.sleep(1.0)
    except Exception as e:
        traceback.print_exc()
        print(f"❌ Erro ao enviar mensagem: {e}")

# =================== CLIENTES: CADASTRO ===================
def salvar_cliente():
    nome = entry_nome.get().strip()
    telefone = entry_telefone.get().strip()
    data_criacao = datetime.today().date()
    try:
        vencimento = entry_vencimento.get_date()
    except:
        messagebox.showerror("Erro", "Data de vencimento inválida.")
        return

    if not nome or not telefone:
        messagebox.showerror("Erro", "Preencha nome e telefone.")
        return

    df = carregar_df()
    novo = pd.DataFrame([[nome, telefone, data_criacao, vencimento, False]], columns=COLUNAS)
    df = pd.concat([df, novo], ignore_index=True)
    salvar_df(df)

    messagebox.showinfo("Sucesso", f"Cliente cadastrado com vencimento em {vencimento.strftime('%d/%m/%Y')}")
    entry_nome.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)

# =================== CLIENTES: VENCIMENTOS + ENVIO ===================
def verificaData_com_progresso():
    abrir_whatsapp_web()
    df = carregar_df()

    hoje = datetime.today().date()
    clientes_pendentes = df[(df["Vencimento"] <= hoje) & (df["Pagou"] == False)]

    total = len(clientes_pendentes)
    if total == 0:
        messagebox.showinfo("Nenhum vencimento", "Nenhum cliente pendente ou vencido encontrado.")
        return

    progresso["maximum"] = total
    progresso["value"] = 0

    for _, row in clientes_pendentes.iterrows():
        nome = row['Nome']
        telefone = str(row['Telefone'])
        if not telefone.startswith('+'):
            telefone = '+55' + telefone  # ajuste para BR se vier só número
        mensagem = f"Olá {nome}, seu plano venceu em {row['Vencimento']}. Entre em contato para renovar! 📆"

        status_label.config(text=f"Enviando para: {nome}")
        janela.update_idletasks()
        enviar_mensagem_whatsapp(telefone, mensagem)

        progresso["value"] += 1
        janela.update_idletasks()

    status_label.config(text="Envio concluído!")
    messagebox.showinfo("Mensagens Enviadas", "Mensagens enviadas. Agora confirme os pagamentos.")
    exibir_confirmacoes()  # abre lista para confirmar e remover da tela conforme confirma

def iniciar_envio_em_thread():
    thread = threading.Thread(target=verificaData_com_progresso)
    thread.daemon = True
    thread.start()

def exibir_confirmacoes():
    """
    Lista todos os clientes vencidos/pendentes AGORA e permite:
    - Confirmar Pagamento (renova +1 mês e marca Pagou=True) -> some da lista
    - Deletar -> some da lista
    A janela fecha sozinha quando não restar ninguém.
    """
    df = carregar_df()
    hoje = datetime.today().date()
    pendentes = df[(df["Vencimento"] <= hoje) & (df["Pagou"] == False)]

    if pendentes.empty:
        messagebox.showinfo("Info", "Nenhum pagamento pendente.")
        return

    win = tk.Toplevel(janela)
    win.title("Confirmar Pagamentos")
    win.geometry("760x420")

    tk.Label(win, text="Clientes vencidos ou vencendo hoje:").pack(pady=10)

    lista_frame = tk.Frame(win)
    lista_frame.pack(fill="both", expand=True, padx=10, pady=5)

    rows_widgets = []

    def remover_linha(frame):
        frame.destroy()
        # se esvaziou a lista, fecha a janela
        if all(not w.winfo_exists() for w in rows_widgets):
            win.destroy()

    def confirmar_pagamento(nome, telefone, frame):
        df2 = carregar_df()
        idx = encontrar_indice(df2, nome, telefone)
        if idx is None:
            messagebox.showerror("Erro", "Registro não encontrado.")
            return
        df2.at[idx, "Pagou"] = True
        df2.at[idx, "Vencimento"] = datetime.today().date() + relativedelta(months=1)
        salvar_df(df2)
        messagebox.showinfo("Pagamento", f"{nome} renovado por +1 mês.")
        remover_linha(frame)

    def deletar_cliente(nome, telefone, frame):
        if not messagebox.askyesno("Confirmação", f"Deletar cliente {nome}?"):
            return
        df2 = carregar_df()
        idx = encontrar_indice(df2, nome, telefone)
        if idx is None:
            messagebox.showerror("Erro", "Registro não encontrado.")
            return
        df2 = df2.drop(idx)
        salvar_df(df2)
        messagebox.showinfo("Removido", f"{nome} deletado.")
        remover_linha(frame)

    for _, row in pendentes.iterrows():
        f = tk.Frame(lista_frame)
        f.pack(fill="x", pady=2)
        rows_widgets.append(f)

        info = f"{row['Nome']} - {row['Telefone']} | Vencimento: {row['Vencimento']}"
        tk.Label(f, text=info, anchor="w", width=60).pack(side="left")
        tk.Button(f, text="Confirmar Pagamento",
                  command=lambda n=row['Nome'], t=row['Telefone'], fr=f: confirmar_pagamento(n, t, fr)
                 ).pack(side="right", padx=4)
        tk.Button(f, text="Deletar",
                  command=lambda n=row['Nome'], t=row['Telefone'], fr=f: deletar_cliente(n, t, fr)
                 ).pack(side="right", padx=4)

# =================== PESQUISAR / EDITAR / DELETAR / CONFIRMAR ===================
def pesquisar_cliente():
    def buscar():
        termo = entry_busca.get().strip().lower()
        for w in frame_resultados.winfo_children():
            w.destroy()

        df = carregar_df()
        if not termo:
            tk.Label(frame_resultados, text="Digite um nome ou telefone para pesquisar.").pack()
            return

        resultados = df[df.apply(lambda r: termo in str(r["Nome"]).lower()
                                           or termo in str(r["Telefone"]), axis=1)]
        if resultados.empty:
            tk.Label(frame_resultados, text="Nenhum cliente encontrado.").pack()
            return

        def editar_cliente(row):
            ew = tk.Toplevel(janela_pesquisa)
            ew.title(f"Editar - {row['Nome']}")
            ew.geometry("400x260")

            tk.Label(ew, text="Nome:").pack()
            e_nome = tk.Entry(ew, width=40)
            e_nome.insert(0, row["Nome"])
            e_nome.pack()

            tk.Label(ew, text="Telefone:").pack()
            e_tel = tk.Entry(ew, width=40)
            e_tel.insert(0, row["Telefone"])
            e_tel.pack()

            tk.Label(ew, text="Vencimento:").pack()
            e_venc = DateEntry(ew, date_pattern="dd/mm/yyyy")
            e_venc.set_date(row["Vencimento"])
            e_venc.pack()

            def salvar():
                df2 = carregar_df()
                idx = encontrar_indice(df2, row["Nome"], row["Telefone"])
                if idx is None:
                    messagebox.showerror("Erro", "Registro não encontrado.")
                    return
                df2.at[idx, "Nome"] = e_nome.get().strip()
                df2.at[idx, "Telefone"] = e_tel.get().strip()
                df2.at[idx, "Vencimento"] = e_venc.get_date()
                salvar_df(df2)
                messagebox.showinfo("Sucesso", "Cliente atualizado.")
                ew.destroy()
                buscar()  # atualiza lista

            tk.Button(ew, text="Salvar", command=salvar).pack(pady=10)

        def deletar(nome, telefone):
            if not messagebox.askyesno("Confirmação", f"Deletar {nome}?"):
                return
            df2 = carregar_df()
            idx = encontrar_indice(df2, nome, telefone)
            if idx is None:
                messagebox.showerror("Erro", "Registro não encontrado.")
                return
            df2 = df2.drop(idx)
            salvar_df(df2)
            messagebox.showinfo("Removido", f"{nome} deletado.")
            buscar()

        def confirmar_pagamento(nome, telefone):
            df2 = carregar_df()
            idx = encontrar_indice(df2, nome, telefone)
            if idx is None:
                messagebox.showerror("Erro", "Registro não encontrado.")
                return
            df2.at[idx, "Pagou"] = True
            df2.at[idx, "Vencimento"] = datetime.today().date() + relativedelta(months=1)
            salvar_df(df2)
            messagebox.showinfo("Pagamento", f"{nome} renovado por +1 mês.")
            buscar()

        # Render resultados
        for _, r in resultados.iterrows():
            fr = tk.Frame(frame_resultados)
            fr.pack(fill="x", pady=2, padx=6)

            txt = f"{r['Nome']} | {r['Telefone']} | Vence: {r['Vencimento']} | Pagou: {'Sim' if r['Pagou'] else 'Não'}"
            tk.Label(fr, text=txt, anchor="w", width=75).pack(side="left")
            tk.Button(fr, text="Editar", command=lambda rr=r: editar_cliente(rr)).pack(side="right", padx=3)
            tk.Button(fr, text="Deletar", command=lambda n=r['Nome'], t=r['Telefone']: deletar(n, t)).pack(side="right", padx=3)
            tk.Button(fr, text="Confirmar Pagamento",
                      command=lambda n=r['Nome'], t=r['Telefone']: confirmar_pagamento(n, t)
                     ).pack(side="right", padx=3)

    janela_pesquisa = tk.Toplevel(janela)
    janela_pesquisa.title("Pesquisar Cliente")
    janela_pesquisa.geometry("820x460")

    tk.Label(janela_pesquisa, text="Nome ou telefone:").pack(pady=6)
    entry_busca = tk.Entry(janela_pesquisa, width=40)
    entry_busca.pack()
    tk.Button(janela_pesquisa, text="Buscar", command=buscar).pack(pady=6)

    frame_resultados = tk.Frame(janela_pesquisa)
    frame_resultados.pack(fill="both", expand=True, pady=8)

# =================== INTERFACE ===================
def iniciar_programa():
    global janela, entry_nome, entry_telefone, entry_vencimento, progresso, status_label
    criar_arquivo_se_nao_existir()

    janela = tk.Tk()
    janela.title("Cadastro e Controle de Clientes")
    janela.geometry("740x580")
    

    # Cadastro
    tk.Label(janela, text="Nome:").pack(pady=5)
    entry_nome = tk.Entry(janela, width=40)
    entry_nome.pack()

    tk.Label(janela, text="Telefone (somente números):").pack(pady=5)
    entry_telefone = tk.Entry(janela, width=40)
    entry_telefone.pack()

    tk.Label(janela, text="Vencimento:").pack(pady=5)
    entry_vencimento = DateEntry(janela, width=37, date_pattern="dd/mm/yyyy")
    entry_vencimento.pack()

    tk.Button(janela, text="Cadastrar Cliente", command=salvar_cliente).pack(pady=10)

    # Ações
    tk.Button(janela, text="Pesquisar Cliente", command=pesquisar_cliente).pack(pady=5)
    tk.Button(janela, text="Verificar Vencimentos e Enviar WhatsApp", command=iniciar_envio_em_thread).pack(pady=5)

    # Progresso envio
    progresso = ttk.Progressbar(janela, orient="horizontal", length=420, mode="determinate")
    progresso.pack(pady=12)
    status_label = tk.Label(janela, text="")
    status_label.pack()

    janela.mainloop()

# =================== MAIN ===================
if __name__ == "__main__":
    if verificar_licenca():
        iniciar_programa()
    else:
        pedir_licenca()
