import base64
from datetime import datetime, timedelta

def gerar_licenca(mac: str, dias_validade: int = 30) -> str:
    """
    Gera uma chave de licença vinculada ao MAC informado.
    Remove ':' do MAC para evitar problemas no parse do programa principal.
    """
    mac_normalizado = mac.replace(":", "").upper()  # remove ':' e coloca maiúsculas

    data_expira = (datetime.today() + timedelta(days=dias_validade)).date().isoformat()
    conteudo = f"MAC:{mac_normalizado}|EXPIRA:{data_expira}"

    chave = base64.b64encode(conteudo.encode()).decode()
    return chave

if __name__ == "__main__":
    mac_usuario = input("Digite o MAC da máquina (ex: 00:1A:2B:3C:4D:5E): ").strip()
    dias = input("Dias de validade da licença [padrão 30]: ").strip()
    dias = int(dias) if dias.isdigit() else 30

    try:
        chave_gerada = gerar_licenca(mac_usuario, dias)
        print("\n✅ Chave de licença gerada (copie exatamente):")
        print(chave_gerada)
    except Exception as e:
        print("❌ Erro ao gerar licença:", e)
