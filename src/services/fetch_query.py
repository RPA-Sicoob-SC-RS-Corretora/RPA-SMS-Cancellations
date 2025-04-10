import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def get_producao_e_carteira(conexao):
    """
    Retorna dois DataFrames com dados de produção e carteira para cooperativas específicas.
    O filtro de data é ajustado para o início do mês passado até o final do mês corrente.
    """
    # Obtém a data atual
    hoje = datetime.now()
    
    inicio_mes_passado = (hoje - relativedelta(months=1)).replace(day=1, hour=0, minute=0, second=0, microsecond=0)
 
    fim_mes_corrente = (hoje.replace(day=1, hour=23, minute=59, second=59, microsecond=999) + relativedelta(months=1) - relativedelta(days=1))

    inicio_formatado = inicio_mes_passado.strftime('%Y-%m-%dT%H:%M:%S.000Z')
    fim_formatado = fim_mes_corrente.strftime('%Y-%m-%dT%H:%M:%S.999Z')

  
    print(f"Filtro de data para as queries: Início = {inicio_formatado}, Fim = {fim_formatado}")

    query1 = f"""
    SELECT
        Produtinho, Produto, Familia, EN1, EN2, EN3, TipoProposta, dt_emissao, 
        Segurado, Cpf_cnpj, Seguradora, Numero_proposta, Numero_apolice_certificado
    FROM
        Sigas.vwProduca_RPA_Cancelamentos
    WHERE
        EN2 IN ('3067 SICOOB CREDIAUC', '3258 SICOOB CREDISC')
        AND dt_emissao BETWEEN '{inicio_formatado}' AND '{fim_formatado}';
    """

    query2 = """
    SELECT
        [CPF/CNPJ], [Telefone Celular]
    FROM
        Sisbr.Carteira
    WHERE
        [Número Cooperativa] IN (3067, 3258);
    """

    df_producao = pd.read_sql(query1, conexao)
    df_carteira = pd.read_sql(query2, conexao)
    
    return df_producao, df_carteira