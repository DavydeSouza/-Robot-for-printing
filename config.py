import pyodbc

# Passo 1: Estabeleça a conexão com o banco de dados
conn_str = (
    'DRIVER={SQL Server};'
    'SERVER=#;'  # Remova as barras e deixe apenas o endereço IP
    'DATABASE=Impressora;'
    'UID=#;'
    'PWD=#'
)

try:
    # Conectando ao banco de dados
    with pyodbc.connect(conn_str) as conn:
        with conn.cursor() as cursor:
            # Executando a consulta
            cursor.execute('SELECT * FROM impressora')  # Substitua 'nome_tabela' pela sua tabela

            # Puxando os dados
            rows = cursor.fetchall()

            # Processando os dados
            for row in rows:
                print(row)

except pyodbc.Error as e:
    print(f"Erro na conexão ou na execução da consulta: {e}")
