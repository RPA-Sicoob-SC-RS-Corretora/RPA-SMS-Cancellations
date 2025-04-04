import pandas as pd

query1 = """
SELECT
    Produtinho, Produto, Familia, EN1, EN2, EN3, TipoProposta,dt_emissao, Segurado, Cpf_cnpj, Seguradora, Numero_proposta, Numero_apolice_certificado
FROM
    Sigas.vwProducao
WHERE
    EN2 IN ('3067 SICOOB CREDIAUC', '3258 SICOOB CREDISC')
    AND dt_emissao BETWEEN '2025-03-31T00:00:00.000Z' AND '2025-12-31T23:59:59.999Z'
    AND TipoProposta = 'CANCELAMENTO';
"""

query2 = """
SELECT
    [CPF/CNPJ], [Telefone Celular]
FROM
    Sisbr.Carteira
WHERE
    [Número Cooperativa] IN (3067, 3258);
"""
