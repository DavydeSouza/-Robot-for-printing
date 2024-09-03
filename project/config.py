import pyodbc

def obter_ips_do_banco():
    conn_str = (
        'DRIVER={SQL Server};'
        'SERVER=10.10.13.250;'
        'DATABASE=Impressora;'
        'UID=netazzurra;'
        'PWD=Azzurra@@2023'
    )
    ips = []
    
    try:
        # Conectando ao banco de dados
        with pyodbc.connect(conn_str) as conn:
            with conn.cursor() as cursor:
                # Executando a consulta para pegar os IPs
                cursor.execute('SELECT ipImp FROM impressora')  # Ajuste a consulta conforme necessário
                rows = cursor.fetchall()

                # Armazenando os IPs na lista
                ips = [row[0] for row in rows]
    except pyodbc.Error as e:
        print(f"Erro ao consultar o banco de dados: {e}")
    return ips
