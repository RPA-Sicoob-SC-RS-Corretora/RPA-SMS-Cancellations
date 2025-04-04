import os
import pyodbc
from dotenv import load_dotenv

load_dotenv()

DB_CONFIG = {
    "host": os.getenv("HOST"),
    "port": os.getenv("PORT"),
    "user": os.getenv("USER"),
    "password": os.getenv("PASSWORD"),
    "database": os.getenv("DATABASE")
}

def conectar_ao_banco():
    """
    Conecta ao banco de dados e retorna a conexão.
    """
    try:
        if not all(DB_CONFIG.values()):
            raise ValueError("Uma ou mais variáveis de ambiente do banco de dados não estão definidas.")
        
        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={DB_CONFIG['host']},{DB_CONFIG['port']};"
            f"DATABASE={DB_CONFIG['database']};"
            f"UID={DB_CONFIG['user']};"
            f"PWD={DB_CONFIG['password']}"
        )
        return pyodbc.connect(conn_str)
    except pyodbc.Error as e:
        print(f"Erro de conexão com o banco de dados: {e}")
        return None
    except Exception as e:
        print(f"Erro inesperado: {e}")
        return None

def criar_query_listar_tabelas():
    """
    Retorna uma query para listar tabelas do banco.
    """
    return "SELECT name FROM sys.tables;"

def executar_query(conn, query):
    """
    Executa uma query SQL e retorna cabeçalhos e resultados.
    """
    cursor = conn.cursor()
    cursor.execute(query)
    colunas = [desc[0] for desc in cursor.description]
    resultados = cursor.fetchall()
    return colunas, resultados