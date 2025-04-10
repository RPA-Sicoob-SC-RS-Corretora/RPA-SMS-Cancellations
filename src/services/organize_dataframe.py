import pandas as pd
from tabulate import tabulate
import re
import pyodbc

def organizar_dados(cabecalhos, resultados):
    """
    Organiza os resultados de uma query em um DataFrame.
    """
    if not resultados:
        return pd.DataFrame()

    if isinstance(resultados[0], tuple) or isinstance(resultados[0], pyodbc.Row):
        resultados = [list(row) for row in resultados]

    if any(len(linha) != len(cabecalhos) for linha in resultados):
        return pd.DataFrame()

    try:
        df = pd.DataFrame(resultados, columns=cabecalhos)
        return df
    except Exception as e:
        return pd.DataFrame()

def exibir_dados_em_tabela(df):
    """
    Exibe o DataFrame em formato de tabela.
    """
    return tabulate(df, headers="keys", tablefmt="pipe")

def salvar_dados_em_csv(df, filepath):
    """
    Salva o DataFrame em um arquivo CSV.
    """
    df.to_csv(filepath, index=False)

def formatar_cpf_cnpj(valor):
    """
    Formata um CPF ou CNPJ para o padrão correto.
    """
    valor = re.sub(r"\D", "", valor)
    if len(valor) == 11:
        return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
    elif len(valor) == 14:
        return f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
    return valor

def formatar_telefone(telefone):
    """
    Remove caracteres não numéricos de um telefone.
    """
    return re.sub(r"[^\d]", "", telefone)

def limpar_coluna_telefone(df):
    """
    Limpa e formata a coluna 'Telefone Celular' no DataFrame.
    """
    if "Telefone Celular" in df.columns:
        df["Telefone Celular"] = df["Telefone Celular"].astype(str).apply(formatar_telefone)
    else:
        raise KeyError("A coluna 'Telefone Celular' não foi encontrada no DataFrame.")
    return df