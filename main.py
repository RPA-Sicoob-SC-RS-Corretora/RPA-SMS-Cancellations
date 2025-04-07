import sys
import os
import pandas as pd
from src.services.acess_banck import conectar_ao_banco, executar_query
from src.services.fetch_query import query1, query2
from src.services.Creat_dataframe import *
from src.variables import VARIABLES
from src.services.send_message_api import enviar_mensagens
from src.services.ConfigEmail import MailConfig

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def main():
    """
    Executa o fluxo principal do programa: conecta ao banco, processa dados e envia mensagens.
    """
    try:
        conn = conectar_ao_banco()
        if conn:
            dados_organizados1 = None
            dados_organizados2 = None
            
            cabecalhos1, resultados1 = executar_query(conn, query1)
            if resultados1:
                print(f"Número de resultados da primeira query: {len(resultados1)}")
                try:
                    dados_organizados1 = organizar_dados(cabecalhos1, resultados1)
                    if not dados_organizados1.empty:
                        dados_organizados1 = dados_organizados1.drop_duplicates()
                except Exception as e:
                    pass
            
            cabecalhos2, resultados2 = executar_query(conn, query2)
            if resultados2:
                print(f"Número de resultados da segunda query: {len(resultados2)}")
                try:
                    dados_organizados2 = organizar_dados(cabecalhos2, resultados2)
                except Exception as e:
                    pass
            
            if dados_organizados1 is not None and not dados_organizados1.empty and \
               dados_organizados2 is not None and not dados_organizados2.empty:
                try:
                    result_df = merge_dataframes(dados_organizados1, dados_organizados2)
                    if 'Telefone Celular' in result_df.columns:
                        result_df = limpar_coluna_telefone(result_df)
                        result_df = adicionar_mensagem_por_en2(result_df)
                        
                        # Filtrar o DataFrame pela data antes de enviar mensagens
                        result_df = filter_date_by_en2(result_df)
                        
                        enviar_mensagens(result_df)
                        separar_por_en2(result_df, VARIABLES["FILE_PATH"])
                        mail_config = MailConfig()
                        mail_config.send_email_with_attachment(VARIABLES["FILE_PATH"])
                except Exception as e:
                    pass
        else:
            pass
    except Exception as e:
        pass

if __name__ == "__main__":
    main()