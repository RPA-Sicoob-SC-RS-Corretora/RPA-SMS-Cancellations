import sys
import os

# Adicionar o diretório raiz ao sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import pandas as pd  # Adicionado para corrigir o erro 'name 'pd' is not defined'
from src.services.acess_banck import conectar_ao_banco, executar_query
from src.services.fetch_query import query1, query2
from src.services.Creat_dataframe import (
    organizar_dados, merge_dataframes, salvar_dataframe,
    exibir_dados_em_tabela, salvar_dados_em_csv, adicionar_mensagem_por_en2, limpar_coluna_telefone,
    separar_por_en2
)
from src.variables import VARIABLES
from src.services.send_message_api import enviar_mensagens
from src.services.ConfigEmail import MailConfig  # Importar a classe de configuração de email

def main():
    try:
        conn = conectar_ao_banco()
        if conn:
            dados_organizados1 = None
            dados_organizados2 = None
            
            # Executar a primeira query
            cabecalhos1, resultados1 = executar_query(conn, query1)
            if resultados1:
                print(f"Formato dos resultados da primeira query: {type(resultados1[0])}")
                print(f"Número de colunas retornadas pela primeira query: {len(resultados1[0]) if isinstance(resultados1[0], (list, tuple)) else 'Formato inválido'}")
                print(f"Cabeçalhos esperados: {len(cabecalhos1)}")
                try:
                    print(f"Cabeçalhos da primeira query: {cabecalhos1}")
                    print(f"Número de resultados da primeira query: {len(resultados1)}")
                    dados_organizados1 = organizar_dados(cabecalhos1, resultados1)
                    if dados_organizados1.empty:
                        print("Erro ao organizar dados da primeira query: DataFrame vazio ou erro de estrutura.")
                    else:
                        print("Pré-visualização da primeira query:")
                        print(dados_organizados1.head())  # Pré-visualizar as primeiras linhas
                except Exception as e:
                    print(f"Erro ao organizar dados da primeira query: {e}")
                    print(f"Resultados brutos da primeira query: {resultados1}")
            else:
                print("Nenhum resultado encontrado para a primeira query.")
            
            # Executar a segunda query
            cabecalhos2, resultados2 = executar_query(conn, query2)
            if resultados2:
                print(f"Formato dos resultados da segunda query: {type(resultados2[0])}")
                print(f"Número de colunas retornadas pela segunda query: {len(resultados2[0]) if isinstance(resultados2[0], (list, tuple)) else 'Formato inválido'}")
                print(f"Cabeçalhos esperados: {len(cabecalhos2)}")
                try:
                    print(f"Cabeçalhos da segunda query: {cabecalhos2}")
                    print(f"Número de resultados da segunda query: {len(resultados2)}")
                    dados_organizados2 = organizar_dados(cabecalhos2, resultados2)
                    if dados_organizados2.empty:
                        print("Erro ao organizar dados da segunda query: DataFrame vazio ou erro de estrutura.")
                    else:
                        print("Pré-visualização da segunda query:")
                        print(dados_organizados2.head())  # Pré-visualizar as primeiras linhas
                except Exception as e:
                    print(f"Erro ao organizar dados da segunda query: {e}")
                    print(f"Resultados brutos da segunda query: {resultados2}")
            else:
                print("Nenhum resultado encontrado para a segunda query.")
            
            # Mesclar os dataframes e salvar o resultado
            if dados_organizados1 is not None and not dados_organizados1.empty and \
               dados_organizados2 is not None and not dados_organizados2.empty:
                try:
                    result_df = merge_dataframes(dados_organizados1, dados_organizados2)
                    if 'Telefone Celular' in result_df.columns:
                        # Limpar e formatar a coluna 'Telefone Celular'
                        result_df = limpar_coluna_telefone(result_df)
                        
                        # Adicionar coluna de mensagem com base na coluna EN2
                        result_df = adicionar_mensagem_por_en2(result_df)
                        print("Mensagens associadas aos valores da coluna EN2:")
                        print(result_df[['EN2', 'Mensagem']].drop_duplicates())  # Visualizar valores únicos de EN2 e suas mensagens
                        print("Pré-visualização do relatório mesclado com mensagens:")
                        print(result_df.head())  # Pré-visualizar as primeiras linhas do DataFrame mesclado
                        
                        # Enviar mensagens usando a função centralizada
                        enviar_mensagens(result_df)
                        
                        # Separar o DataFrame final pelos valores da coluna EN2
                        separar_por_en2(result_df, VARIABLES["FILE_PATH"])
                        
                        # Enviar os arquivos separados por email
                        mail_config = MailConfig()
                        mail_config.send_email_with_attachment(VARIABLES["FILE_PATH"])
                        print("Emails enviados com sucesso para os arquivos separados.")
                    else:
                        print("Erro: A coluna 'Telefone Celular' não foi adicionada ao DataFrame final.")
                except Exception as e:
                    print(f"Erro ao mesclar os DataFrames: {e}")
            else:
                print("Os DataFrames não estão disponíveis ou estão vazios para mesclagem.")
        else:
            print("Erro ao conectar ao banco de dados.")
    except Exception as e:
        print(f"Erro durante a execução: {e}")

if __name__ == "__main__":
    main()