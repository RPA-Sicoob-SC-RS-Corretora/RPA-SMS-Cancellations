import os
import pyodbc
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

# Obter informações do banco de dados
DB_CONFIG = {
    "host": os.getenv("HOST"),
    "port": os.getenv("PORT"),
    "user": os.getenv("USER"),
    "password": os.getenv("PASSWORD"),
    "database": os.getenv("DATABASE")
}

def conectar_ao_banco():
    try:
        # Verificar se todas as variáveis de ambiente estão definidas
        if not all(DB_CONFIG.values()):
            raise ValueError("Uma ou mais variáveis de ambiente do banco de dados não estão definidas.")
        
        # Configurar a string de conexão
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
    return "SELECT name FROM sys.tables;"

def executar_query(conn, query):
    cursor = conn.cursor()
    cursor.execute(query)
    colunas = [desc[0] for desc in cursor.description]
    resultados = cursor.fetchall()
    return colunas, resultados