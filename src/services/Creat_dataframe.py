import pandas as pd
from tabulate import tabulate
from src.variables import VARIABLES
import os
import re
import pyodbc

def organizar_dados(cabecalhos, resultados):
    """
    Organiza os dados retornados de uma consulta SQL em um DataFrame.

    :param cabecalhos: Lista de cabeçalhos das colunas.
    :param resultados: Lista de resultados retornados pela consulta.
    :return: DataFrame organizado ou vazio em caso de erro.
    """
    if not resultados:
        print("Erro: Nenhum resultado encontrado para a consulta.")
        return pd.DataFrame()

    # Converter resultados de pyodbc.Row para lista de listas, se necessário
    if isinstance(resultados[0], tuple) or isinstance(resultados[0], pyodbc.Row):
        print("Convertendo resultados de pyodbc.Row para lista de listas...")
        resultados = [list(row) for row in resultados]

    # Validar se o número de colunas nos resultados corresponde ao número de cabeçalhos
    if any(len(linha) != len(cabecalhos) for linha in resultados):
        print(f"Erro: Número de colunas nos resultados não corresponde ao número de cabeçalhos.")
        print(f"Esperado: {len(cabecalhos)} colunas, Encontrado: {set(len(linha) for linha in resultados)} colunas.")
        return pd.DataFrame()

    try:
        # Criar o DataFrame
        df = pd.DataFrame(resultados, columns=cabecalhos)
        if df.empty:
            print("Erro: DataFrame criado está vazio.")
        return df
    except Exception as e:
        print(f"Erro ao criar DataFrame: {e}")
        return pd.DataFrame()

def exibir_dados_em_tabela(df):
    return tabulate(df, headers="keys", tablefmt="pipe")

def salvar_dados_em_csv(df, filepath):
    df.to_csv(filepath, index=False)

def formatar_cpf_cnpj(valor):
    valor = re.sub(r"\D", "", valor)
    if len(valor) == 11:
        return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
    elif len(valor) == 14:
        return f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
    return valor

def formatar_telefone(telefone):
    return re.sub(r"[^\d]", "", telefone)

def limpar_coluna_telefone(df):
    if "Telefone Celular" in df.columns:
        df["Telefone Celular"] = df["Telefone Celular"].astype(str).apply(formatar_telefone)
    else:
        raise KeyError("A coluna 'Telefone Celular' não foi encontrada no DataFrame.")
    return df

def merge_dataframes(df1, df2):
    df1["Cpf_cnpj"] = df1["Cpf_cnpj"].astype(str).apply(formatar_cpf_cnpj)
    df2["CPF/CNPJ"] = df2["CPF/CNPJ"].astype(str).apply(formatar_cpf_cnpj)
    merged_df = pd.merge(df1, df2[["CPF/CNPJ", "Telefone Celular"]], left_on="Cpf_cnpj", right_on="CPF/CNPJ", how="left")
    df1["Telefone Celular"] = merged_df["Telefone Celular"]
    return df1

def adicionar_mensagem_por_en2(df):
    mensagens = {
        "3067 SICOOB CREDIAUC": VARIABLES["MESSAGE_CREDIAUC"],
        "3258 SICOOB CREDISC": VARIABLES["MESSAGE_CREDISC"]
    }
    df["Mensagem"] = df["EN2"].map(mensagens).fillna("Mensagem não definida")
    return df

def salvar_dataframe(df, file_path, filename):
    filepath = os.path.join(file_path, filename)
    df.to_csv(filepath, index=False)

def separar_por_en2(df, file_path):
    """
    Separa o DataFrame em arquivos CSV com base nos valores únicos da coluna EN2.

    :param df: DataFrame contendo a coluna EN2.
    :param file_path: Diretório base onde os arquivos separados serão salvos.
    """
    if "EN2" not in df.columns:
        raise KeyError("A coluna 'EN2' não foi encontrada no DataFrame.")

    os.makedirs(file_path, exist_ok=True)  # Garantir que o diretório de saída exista

    valores_unicos = df["EN2"].unique()
    for valor in valores_unicos:
        df_filtrado = df[df["EN2"] == valor]
        filename = f"{valor.replace(' ', '_').replace('/', '_')}.csv"  # Alterado para salvar como CSV
        filepath = os.path.join(file_path, filename)
        df_filtrado.to_csv(filepath, index=False)  # Salvar como CSV
        print(f"Arquivo salvo: {filepath}")

