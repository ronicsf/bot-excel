import base64
from datetime import datetime, timedelta

def gerar_chave(dias_validade: int) -> str:
    """Gera chave de licença válida por X dias"""
    data_expira = datetime.today().date() + timedelta(days=dias_validade)
    conteudo = f"EXPIRA:{data_expira}"
    chave = base64.b64encode(conteudo.encode()).decode()
    return chave, data_expira

if __name__ == "__main__":
    print("=== GERADOR DE LICENÇAS ===")
    dias = int(input("Digite quantos dias a licença deve durar: "))
    chave, data_expira = gerar_chave(dias)
    print("\n✅ Licença gerada com sucesso!")
    print(f"   - Expira em: {data_expira}")
    print(f"   - Chave: {chave}\n")
    print("👉 Copie a chave acima e envie para o cliente.")