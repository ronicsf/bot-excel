import base64
import secrets
from datetime import datetime, timedelta

def gerar_chave(dias_validade: int) -> tuple[str, str]:
    """Gera chave de licença válida por X dias"""
    data_expira = (datetime.today().date() + timedelta(days=dias_validade)).isoformat()

    # Gera um token aleatório seguro
    token_unico = secrets.token_hex(8)  # 16 caracteres aleatórios

    conteudo = f"EXPIRA:{data_expira}|TOKEN:{token_unico}"
    chave = base64.b64encode(conteudo.encode()).decode()

    return chave, data_expira


def validar_chave(chave: str) -> bool:
    """Valida se a chave ainda é válida"""
    try:
        # Decodifica a chave
        conteudo = base64.b64decode(chave.encode()).decode()

        # Extrai os dados
        partes = dict(p.split(":") for p in conteudo.split("|"))
        data_expira = datetime.fromisoformat(partes["EXPIRA"]).date()

        # Verifica se está no prazo
        return datetime.today().date() <= data_expira

    except Exception:
        return False


if __name__ == "__main__":
    print("=== GERADOR DE LICENÇAS ===")
    dias = int(input("Digite quantos dias a licença deve durar: "))
    chave, data_expira = gerar_chave(dias)

    print("\n✅ Licença gerada com sucesso!")
    print(f"   - Expira em: {data_expira}")
    print(f"   - Chave: {chave}\n")

    print("=== TESTE DE VALIDAÇÃO ===")
    print("Chave válida?" , "SIM ✅" if validar_chave(chave) else "NÃO ❌")
