import sys
import os
import pandas as pd
from src.services.acess_banck import conectar_ao_banco
from src.services.fetch_query import get_producao_e_carteira
from src.services.Creat_dataframe import *
from src.variables import VARIABLES
from src.services.send_message_api import enviar_mensagens
from src.services.ConfigEmail import MailConfig
import datetime

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def main():
    """
    Executa o fluxo principal do programa: conecta ao banco, processa dados e envia mensagens.
    """
    try:
        conn = conectar_ao_banco()
        if conn:
            df_producao, df_carteira = get_producao_e_carteira(conn)
            
            if df_producao is not None and not df_producao.empty:
                print(f"Número de linhas em df_producao: {len(df_producao)}")
                df_producao = df_producao.drop_duplicates()
            else:
                print("Nenhum dado retornado para df_producao.")
                return
            
            if df_carteira is not None and not df_carteira.empty:
                print(f"Número de linhas em df_carteira: {len(df_carteira)}")
            else:
                print("Nenhum dado retornado para df_carteira.")
                return
            
            try:
                result_df = merge_dataframes(df_producao, df_carteira)
                if 'Telefone Celular' in result_df.columns:
                    result_df = limpar_coluna_telefone(result_df)
                    result_df = adicionar_mensagem_por_en2(result_df)
                
                    result_df = filter_date_by_en2(result_df)
                   
                    today = datetime.datetime.now().weekday()
                    is_monday = today == 0
                    
                    df_credisc = result_df[result_df['EN2'] == "3258 SICOOB CREDISC"]
                    
                    if is_monday:
                        if not result_df.empty:
                            result_df = enviar_mensagens(result_df)  # Envia todos os dados
                            separar_por_en2(result_df, VARIABLES["FILE_PATH"])
                            mail_config = MailConfig()
                            mail_config.send_email_with_attachment(VARIABLES["FILE_PATH"])
                        else:
                            print("Nenhum dado a ser enviado na segunda-feira.")
                    else:
                        if not df_credisc.empty:
                            df_credisc = enviar_mensagens(df_credisc)  # Envia apenas 3258 SICOOB CREDISC
                            separar_por_en2(df_credisc, VARIABLES["FILE_PATH"])
                            mail_config = MailConfig()
                            mail_config.send_email_with_attachment(VARIABLES["FILE_PATH"])
                        else:
                            print("Nenhum dado de 3258 SICOOB CREDISC a ser enviado hoje.")
                    
            except Exception as e:
                print(f"Erro ao processar os DataFrames: {e}")
        else:
            print("Falha ao conectar ao banco de dados.")
    except Exception as e:
        print(f"Erro no fluxo principal: {e}")

if __name__ == "__main__":
    main()