from src.services.acess_banck import conectar_ao_banco
from src.services.tables import listar_tabelas

def main():
    conn = conectar_ao_banco()
    if conn:
        # Mostrar as tabelas disponíveis
        listar_tabelas(conn)
        # Fechar a conexão
        conn.close()

if __name__ == "__main__":
    main()
