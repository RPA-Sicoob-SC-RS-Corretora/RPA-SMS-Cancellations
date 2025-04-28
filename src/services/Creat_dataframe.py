import pandas as pd
from tabulate import tabulate
from src.variables import VARIABLES
import os
import re
import pyodbc
import datetime
import unicodedata

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

def remover_acentos(texto):
    """
    Remove acentos de um texto.
    """
    if isinstance(texto, str):
        return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

def salvar_dados_em_csv(df, filepath):
    """
    Salva o DataFrame em um arquivo CSV, removendo acentos das palavras.
    """
    df = df.applymap(remover_acentos)  # Remove acentos de todas as células do DataFrame
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

def merge_dataframes(df1, df2):
    """
    Mescla dois DataFrames com base no CPF/CNPJ.
    """
    df1 = df1.drop_duplicates()
    print("Pré-visualização após remover duplicatas de df1:")
    print(df1.head())

    df1["Cpf_cnpj"] = df1["Cpf_cnpj"].astype(str).apply(formatar_cpf_cnpj)
    df2["CPF/CNPJ"] = df2["CPF/CNPJ"].astype(str).apply(formatar_cpf_cnpj)
    merged_df = pd.merge(df1, df2[["CPF/CNPJ", "Telefone Celular"]], left_on="Cpf_cnpj", right_on="CPF/CNPJ", how="left")
    print("Pré-visualização após mesclar os DataFrames:")
    print(merged_df.head())

    df1["Telefone Celular"] = merged_df["Telefone Celular"]
    return df1

def adicionar_mensagem_por_en2(df):
    """
    Adiciona mensagens ao DataFrame com base na coluna 'EN2' e inicializa 'Status_Envio'.
    """
    mensagens = {
        "3067 SICOOB CREDIAUC": VARIABLES["MESSAGE_CREDIAUC"],
        "3258 SICOOB CREDISC": VARIABLES["MESSAGE_CREDISC"]
    }
    df["Mensagem"] = df["EN2"].map(mensagens).fillna("Mensagem não definida")
    df["Status_Envio"] = "Pendente"  # Inicializa a coluna Status_Envio
    return df
    

def salvar_dataframe(df, file_path, filename):
    """
    Salva o DataFrame em um arquivo CSV no caminho especificado.
    """
    filepath = os.path.join(file_path, filename)
    df.to_csv(filepath, index=False)

def separar_por_en2(df, file_path):
    """
    Separa o DataFrame em arquivos CSV por valores únicos da coluna 'EN2' usando codificação UTF-8.
    """
    if "EN2" not in df.columns:
        raise KeyError("A coluna 'EN2' não foi encontrada no DataFrame.")
    if "Mensagem" not in df.columns:
        raise KeyError("A coluna 'Mensagem' não foi encontrada no DataFrame.")

    os.makedirs(file_path, exist_ok=True)

    colunas_relevantes = [
        "Familia", "EN3", "TipoProposta", "dt_emissao", "Cpf_cnpj", 
        "Seguradora", "Numero_apolice_certificado", "Telefone Celular", 
        "Mensagem", "Status_Envio"
    ]

    valores_unicos = df["EN2"].unique()
    for valor in valores_unicos:
        df_filtrado = df[df["EN2"] == valor]
        df_filtrado = df_filtrado[colunas_relevantes]
        df_filtrado = df_filtrado.applymap(remover_acentos)  # Remove acentos antes de salvar
        filename = f"{valor.replace(' ', '_').replace('/', '_')}.csv"
        filepath = os.path.join(file_path, filename)
        df_filtrado.to_csv(filepath, index=False, sep=';', encoding='utf-8')

def filter_date_by_en2(df):
    """
    Filtra o DataFrame pela coluna 'dt_emissao' com base no valor de 'EN2'.
    3067 SICOOB CREDIAUC: 5 dias para trás
    3258 SICOOB CREDISC: 8 dias para trás
    """
    current_date = datetime.date.today()
    
    if 'dt_emissao' not in df.columns:
        print("Coluna 'dt_emissao' não encontrada no DataFrame")
        return df
    
    df['dt_emissao'] = pd.to_datetime(df['dt_emissao'], errors='coerce')
    
    filtered_df = pd.DataFrame()

    for en2_value in df['EN2'].unique():
        df_subset = df[df['EN2'] == en2_value].copy()
        
        if en2_value == "3067 SICOOB CREDIAUC":
            date_threshold = current_date - datetime.timedelta(days=5)
            df_subset = df_subset[df_subset['dt_emissao'].dt.date >= date_threshold]
            
        elif en2_value == "3258 SICOOB CREDISC":
            date_threshold = current_date - datetime.timedelta(days=8)
            df_subset = df_subset[df_subset['dt_emissao'].dt.date >= date_threshold]
        
        print(f"Pré-visualização após filtrar por data para EN2 = {en2_value}:")
        print(df_subset.head())
        
        filtered_df = pd.concat([filtered_df, df_subset])
    
    print(f"DataFrame antes de filtrar por data: {len(df)} linhas")
    print(f"DataFrame após filtrar por data: {len(filtered_df)} linhas")
    print("Pré-visualização do DataFrame final após o filtro por data:")
    print(filtered_df.head())
    return filtered_df