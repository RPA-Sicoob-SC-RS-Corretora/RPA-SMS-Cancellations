import json
import requests as api
import os
import urllib3
import re
from src.variables import VARIABLES

# Suprimir avisos de InsecureRequestWarning
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def send_message_api(phone_number: str, message: str) -> api.Response:
    """
    Sends a message to a phone number using the API.

    :param phone_number: The phone number to send the message to.
    :type phone_number: str
    :param message: The message to be sent.
    :type message: str
    :return: The response from the API.
    :rtype: api.Response
    """
    url = VARIABLES["URL"]  # URL da API
    headers = {
        "content-type": "application/json",
        "auth-key": VARIABLES["AUTH_KEY"],
    }
    payload = {
        "Receivers": phone_number,
        "Content": message
    }

    # Log do payload
    print(f"Enviando payload: {json.dumps(payload)}")

    try:
        # Enviar a requisição POST
        response = api.post(url, data=json.dumps(payload), headers=headers, verify=False)
        # Log detalhado da resposta
        print(f"Resposta da API: {response.status_code} - {response.text}")
        return response
    except Exception as e:
        print(f"Erro ao enviar mensagem: {e}")
        return None

def validar_parametros_envio(telefone: str, mensagem: str) -> bool:
    """
    Valida os parâmetros de envio para garantir que não estejam vazios ou inválidos.

    :param telefone: Número de telefone a ser validado.
    :param mensagem: Mensagem a ser validada.
    :return: True se os parâmetros forem válidos, False caso contrário.
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
    Envia mensagens para os números de telefone presentes no DataFrame.

    :param result_df: DataFrame contendo as colunas 'Telefone Celular' e 'Mensagem'.
    """
    for _, row in result_df.iterrows():
        telefone = row['Telefone Celular']
        mensagem = row['Mensagem']

        # Validar os parâmetros antes de enviar
        if not validar_parametros_envio(telefone, mensagem):
            continue

        response = send_message_api(telefone, mensagem)
        if response is None:
            print(f"Erro ao enviar mensagem para {telefone}: Resposta da API é None.")
        elif response.status_code == 200:
            print(f"Mensagem enviada com sucesso para {telefone}.")
        else:
            print(f"Erro ao enviar mensagem para {telefone}: {response.status_code} - {response.text}")
