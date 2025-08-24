import base64
import secrets
import uuid
from datetime import datetime, timedelta
import os

# Arquivo para armazenar o UUID do executável
UUID_FILE = "app_uuid.txt"

def get_or_create_uuid() -> str:
    """Obtém o UUID fixo para este executável, cria se não existir"""
    if os.path.exists(UUID_FILE):
        with open(UUID_FILE, "r") as f:
            return f.read().strip()
    else:
        novo_uuid = str(uuid.uuid4())
        with open(UUID_FILE, "w") as f:
            f.write(novo_uuid)
        return novo_uuid


def gerar_chave(dias_validade: int) -> tuple[str, str]:
    """Gera chave de licença única para este executável"""
    data_expira = (datetime.today().date() + timedelta(days=dias_validade)).isoformat()
    token_unico = secrets.token_hex(8)  # parte aleatória
    app_uuid = get_or_create_uuid()      # identificador único

    conteudo = f"UUID:{app_uuid}|EXPIRA:{data_expira}|TOKEN:{token_unico}"
    chave = base64.b64encode(conteudo.encode()).decode()

    return chave, data_expira


def validar_chave(chave: str) -> bool:
    """Valida se a chave pertence a este executável e se ainda está no prazo"""
    try:
        conteudo = base64.b64decode(chave.encode()).decode()
        partes = dict(p.split(":") for p in conteudo.split("|"))

        app_uuid = get_or_create_uuid()
        data_expira = datetime.fromisoformat(partes["EXPIRA"]).date()

        # Valida UUID + prazo
        if partes["UUID"] != app_uuid:
            return False
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
