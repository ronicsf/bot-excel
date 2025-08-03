import pandas as pd
from datetime import datetime

# Carrega os dados
tabela = pd.read_excel("clientes.xlsx", usecols=["Nome", "Telefone", "Data Criacao"])

# Garante que a coluna de data está no formato datetime
tabela["Data Criacao"] = pd.to_datetime(tabela["Data Criacao"]).dt.date

def verificaData(tabela):
    data_hoje = datetime.today().date()  # Data atual
    encontrou = False

    for index, row in tabela.iterrows():
        if row["Data Criacao"] == data_hoje:
            print(f"Achei: {row['Nome']} - {row['Telefone']}")
            encontrou = True

    if not encontrou:
        print("Nenhuma data bate com a data de hoje.")

verificaData(tabela)