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

def obter_setor_por_ip(ip):
    conn_str = (
        'DRIVER={SQL Server};'
        'SERVER=10.10.13.250;'
        'DATABASE=Impressora;'
        'UID=netazzurra;'
        'PWD=Azzurra@@2023'
    )
    
    setor = None
    try:
        # Conectar ao banco de dados
        with pyodbc.connect(conn_str) as conn:
            with conn.cursor() as cursor:
                # Executar a stored procedure e capturar o resultado
                cursor.execute("{CALL pcons_setor (?)}", (ip,))
                setor = cursor.fetchone()  # Obter o resultado
                
                if setor:
                    setor = setor[0]  # Ajuste conforme o formato do resultado
                print(f"Setor obtido para IP {ip}: {setor}")
                
    except pyodbc.Error as e:
        print(f"Erro ao obter o setor: {e}")
    
    return setor

def executar_stored_procedure(id_imp):
    conn_str = (
        'DRIVER={SQL Server};'
        'SERVER=10.10.13.250;'
        'DATABASE=Impressora;'
        'UID=netazzurra;'
        'PWD=Azzurra@@2023'
    )
    
    try:
        with pyodbc.connect(conn_str) as conn:
            with conn.cursor() as cursor:
                # Corrigir a chamada da stored procedure
                cursor.execute("{CALL pupd_log(?)}", (id_imp,))
                conn.commit()
                print(f"Stored procedure executada com sucesso para ID = {id_imp}")
    except pyodbc.Error as e:
        print(f"Erro ao executar a stored procedure: {e}")
