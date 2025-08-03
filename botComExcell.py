import pandas as pd
from datetime import datetime
import pywhatkit as kit
import time

# Carrega os dados da planilha
tabela = pd.read_excel("clientes.xlsx", usecols=["Nome", "Telefone", "Data Criacao"])
tabela["Data Criacao"] = pd.to_datetime(tabela["Data Criacao"]).dt.date

def enviar_mensagem_whatsapp(telefone, mensagem):
    try:
        # Envio instantâneo com tempo de espera fixo
        kit.sendwhatmsg_instantly(telefone, mensagem, wait_time=10, tab_close=True)
        print(f" Mensagem enviada para {telefone}")
        time.sleep(10)  # Tempo de espera entre mensagens
    except Exception as e:
        print(f" Erro ao enviar mensagem para {telefone}: {e}")

def verificaData(tabela):
    data_hoje = datetime.today().date()
    encontrou = False

    for index, row in tabela.iterrows():
        if row["Data Criacao"] == data_hoje:
            nome = row['Nome']
            telefone = str(row['Telefone'])

            # Ajusta número para formato internacional (Brasil como exemplo)
            if not telefone.startswith('+'):
                telefone = '+55' + telefone

            mensagem = f"Olá {nome}, sou um bot enviado pelo Roni para teste "
            enviar_mensagem_whatsapp(telefone, mensagem)
            encontrou = True

    if not encontrou:
        print(" Nenhuma data bate com a data de hoje.")

# Executa a verificação
verificaData(tabela)
