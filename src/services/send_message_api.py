import json
import requests as api
import urllib3
from tqdm import tqdm
from src.variables import VARIABLES

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def send_message_api(phone_number: str, message: str) -> api.Response:
    """
    Envia uma mensagem SMS para o número informado.
    """
    url = VARIABLES["URL"]
    headers = {
        "content-type": "application/json",
        "auth-key": VARIABLES["AUTH_KEY"],
    }
    payload = {
        "Receivers": phone_number,
        "Content": message
    }

    try:
        response = api.post(url, data=json.dumps(payload), headers=headers, verify=False)
        return response
    except Exception as e:
        print(f"Erro ao enviar mensagem: {e}")
        return None

def validar_parametros_envio(telefone: str, mensagem: str) -> bool:
    """
    Valida o telefone e a mensagem antes do envio.
    """
    if not telefone or telefone.strip() == "":
        print(f"Erro: Telefone inválido. Telefone: {telefone}")
        return False
    if not mensagem or mensagem.strip() == "":
        print(f"Erro: Mensagem inválida. Mensagem: {mensagem}")
        return False
    return True

def enviar_mensagens(result_df):
    """
    Envia mensagens SMS para os números no DataFrame.
    """
    if "Produto" not in result_df.columns:
        raise KeyError("A coluna 'Produto' não foi encontrada no DataFrame.")

    total_mensagens = len(result_df)
    print(f"Iniciando o envio de {total_mensagens} mensagens SMS...")

    for _, row in tqdm(result_df.iterrows(), total=total_mensagens, desc="Enviando SMS"):
        telefone = row['Telefone Celular']
        mensagem = row['Mensagem']
        produto = row['Produto']

        if isinstance(mensagem, str) and "!prd!" in mensagem:
            mensagem = mensagem.replace("!prd!", str(produto))

        if not validar_parametros_envio(telefone, mensagem):
            continue

        response = send_message_api(telefone, mensagem)
        if response is None:
            print(f"Erro ao enviar mensagem para {telefone}: Resposta da API é None.")
        elif response.status_code != 200:
            print(f"Erro ao enviar mensagem para {telefone}: {response.status_code} - {response.text}")

    print("Envio de mensagens concluído com sucesso!")