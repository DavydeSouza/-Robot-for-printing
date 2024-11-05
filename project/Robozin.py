import tkinter as tk
from tkinter import messagebox, Canvas
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import pyodbc
from PIL import Image, ImageTk
from config import executar_stored_procedure, obter_ips_do_banco, obter_setor_por_ip ,obter_tipo_impressora_por_ip,obter_loja_por_ip # Importando a função do arquivo de banco de dados
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk
from tkinter import filedialog
from datetime import datetime
import threading

dados_acumulados = []


def abrir_navegador(url):
    try:
        driver = webdriver.Chrome()  # Certifique-se de que o chromedriver esteja instalado e no PATH
        driver.get(url)
        time.sleep(5)  # Aguarde o carregamento da página
        return driver
    except Exception as e:
        # Exibir alerta ao usuário caso não seja possível acessar a impressora
        messagebox.showerror("Erro de Conexão", f"Não foi possível acessar a impressora no endereço {url}.\nErro: {e}")
        return None

def fazer_login(driver, senha):
    senha_campo = driver.find_element(By.ID, "LogBox")
    senha_campo.send_keys(senha)
    driver.find_element(By.ID, "login").click()
    time.sleep(5)  # Aguarde a autenticação

def acessar_pagina_admin(driver):
    admin_link = driver.find_element(By.LINK_TEXT, "Administrator")
    admin_link.click()
    time.sleep(5)

def acessar_pagina_controle(driver):
    controle_link = driver.find_element(By.LINK_TEXT, "Restricted Functions 1-25")
    controle_link.click()
    time.sleep(5)

def reseta_contador(driver):
    driver.find_element(By.XPATH, "//input[@value='All Counter Reset']").click()
    driver.find_element(By.XPATH, "//input[@value='Submit']").click()

def coletar_dados(driver, qt, setor, tipo_impressora, loja):
    try:
        wait = WebDriverWait(driver, 10)  # Espera de até 10 segundos
        
        # Coletando dados para impressora Multifuncional
        if tipo_impressora == 'Multifuncional':
            tabela = wait.until(EC.presence_of_element_located((By.ID, "lock")))  # Aguarda a tabela aparecer
            linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")

            for linha in linhas:
                try:
                    colunas = linha.find_elements(By.TAG_NAME, "td")

                    lock_num = colunas[0].text.strip()  # ID do usuário
                    lock_name = colunas[1].find_element(By.CSS_SELECTOR, "input").get_attribute("value").strip()  # Nome do usuário
                    page_limit_max = colunas[12].find_element(By.CSS_SELECTOR, "input").get_attribute("value").strip()  # Limite máximo de páginas
                    last_td_value = colunas[13].text.strip()  # Contador de páginas

                    if lock_name == "":
                        continue

                    # Adiciona os dados coletados à lista global
                    dados_acumulados.append([setor, lock_num, lock_name, page_limit_max, last_td_value, loja])
                except Exception as e:
                    print(f"Erro ao processar linha (Multifuncional): {e}")

        # Coletando dados para impressora Multifuncional/Colorida
        elif tipo_impressora == 'Multifuncional/Colorida':
            tabela = wait.until(EC.presence_of_element_located((By.ID, "lock")))
            linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")

            for linha in linhas:
                try:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    lock_num = colunas[0].find_element(By.TAG_NAME, "label").text.strip()  # ID do usuário
                    lock_name = colunas[1].find_element(By.CSS_SELECTOR, "input.lockName").get_attribute("value").strip()  # Nome do usuário
                    page_limit_max = colunas[14].find_element(By.CSS_SELECTOR, "input.lockPageLimitMax").get_attribute("value").strip()  # Limite máximo de páginas
                    last_td_value = colunas[15].text.strip()  # Contador de páginas

                    if lock_name == "":
                        continue

                    # Adiciona os dados coletados à lista global
                    dados_acumulados.append([setor, lock_num, lock_name, page_limit_max, last_td_value, loja])

                except Exception as e:
                    print(f"Erro ao processar linha (Multifuncional/Colorida): {e}")

        # Coletando dados para outros tipos de impressora
        else:
            tabela = wait.until(EC.presence_of_element_located((By.ID, "lock")))
            linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")

            for linha in linhas:
                try:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    lock_num = colunas[0].find_element(By.TAG_NAME, "label").text.strip()  # ID do usuário
                    lock_name = colunas[1].find_element(By.CSS_SELECTOR, "input.lockName").get_attribute("value").strip()  # Nome do usuário
                    page_limit_max = colunas[5].find_element(By.CSS_SELECTOR, "input.lockPageLimitMax").get_attribute("value").strip()  # Limite máximo de páginas
                    last_td_value = colunas[6].text.strip()  # Contador de páginas

                    if lock_name == "":
                        continue

                    # Adiciona os dados coletados à lista global
                    dados_acumulados.append([setor, lock_num, lock_name, page_limit_max, last_td_value, loja])

                except Exception as e:
                    print(f"Erro ao processar linha (Outro tipo): {e}")

    except Exception as e:
        print(f"Erro ao coletar os dados: {e}")

# Função para salvar todos os dados em um único arquivo Excel
def salvar_dados_em_excel(nome_arquivo="dados_completos.xlsx"):
    try:
        df = pd.DataFrame(dados_acumulados, columns=["Setor", "ID", "Nome", "Limite de Folhas", "Qtd Impressa","Loja"])
        df.to_excel(nome_arquivo, index=False)
        print(f"Dados salvos em '{nome_arquivo}'.")
    except Exception as e:
        print(f"Erro ao salvar os dados: {e}")


#def contar_linhas_na_tabela(driver):
    # Implementar a lógica para contar as linhas na tabela
    # Por exemplo:
    # tabela = driver.find_element_by_id("id_da_tabela")  # Substitua pelo ID ou outro seletor da tabela
    # linhas = tabela.find_elements_by_tag_name("tr")
    # return len(linhas) - 1  # Subtrair 1 se a primeira linha for o cabeçalho


# Função para abrir a janela de seleção de arquivo
def selecionar_arquivo():
    Tk().withdraw()  # Para ocultar a janela principal do Tkinter
    arquivo = filedialog.askopenfilename(title="Selecione o arquivo", 
                                         filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    return arquivo

# Função para importar os dados e inserir na tabela através da Stored Procedure
def importar_dados():
    # Selecionar o arquivo usando a janela
    arquivo = selecionar_arquivo()
    
    # Verificar o tipo do arquivo
    if arquivo.endswith('.xlsx'):
        df = pd.read_excel(arquivo)
    elif arquivo.endswith('.csv'):
        df = pd.read_csv(arquivo)
    else:
        raise ValueError("Tipo de arquivo não suportado. Use Excel (.xlsx) ou CSV (.csv)")
    
    # Conectar ao SQL Server
    conn = pyodbc.connect(
        'DRIVER={SQL Server};'
        'SERVER=10.10.13.250;'
        'DATABASE=Impressora;'
        'UID=netazzurra;'
        'PWD=Azzurra@@2023'
    )
    cursor = conn.cursor()

    # Para cada linha no DataFrame, chamamos a Stored Procedure para inserir os dados
    for index, row in df.iterrows():
        cursor.execute("""
            EXEC pins_dados ?, ?, ?, ?, ?, ?
        """, 
        row['Setor'], row['Nome'], row['Limite de Folhas'], row['Qtd Impressa'], datetime.now(),row['Loja'])

    conn.commit()  # Salva as mudanças no banco de dados
    conn.close()
    messagebox.showinfo("Concluído", "Dados importados com sucesso")

    print("Dados importados com sucesso!")

# Chamar a função de importação

def iniciar_automacao(botao, progress_var):
    botao.config(text="Processando...", state="disabled")
    progress_var.set(10)
    dados_acumulados =[]
    senha = 'initpass'
    ips = obter_ips_do_banco()
    
    qt = 1
    index = 0
    while index < len(ips):
        ip = ips[index]
        print(f"Iniciando automação para o IP: {ip}")
        login_url = f"http://{ip}"
        
        driver = abrir_navegador(login_url)
        if driver is None:
            index += 1
            continue
        
        fazer_login(driver, senha)
        acessar_pagina_admin(driver)
        acessar_pagina_controle(driver)
        
        setor = obter_setor_por_ip(ip)
        tipo_impressora = obter_tipo_impressora_por_ip(ip)
        loja = obter_loja_por_ip(ip)  # Obter o tipo da impressora
        
        if tipo_impressora:
            coletar_dados(driver, qt, setor, tipo_impressora,loja)
            reseta_contador(driver)
            if setor =="ShowRoom":
                controle_link = driver.find_element(By.LINK_TEXT, "Restricted Functions 26-50")
                controle_link.click()
                coletar_dados(driver, qt, setor, tipo_impressora,loja)
                reseta_contador(driver)

        else:
            print(f"Tipo de impressora não encontrado para o IP {ip}")
        
        executar_stored_procedure(ip)
        driver.quit()
        
        index += 1
        qt += 1
        progress_var.set(100 * (index / len(ips)))
    
    messagebox.showinfo("Concluído", "Automação finalizada com sucesso!")
    salvar_dados_em_excel()
    botao.config(text="Iniciar Automação", state="normal")
    progress_var.set(0)

def iniciar_botao(botao, progress_var):
    
    try:
        thread = threading.Thread(target=iniciar_automacao, args=(botao, progress_var))
        thread.start()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        botao.config(text="Iniciar Automação", state="normal")

def iniciar_dados(botao):
    try:
        importar_dados()
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        botao.config(text="Importar Dados", state="normal")


def on_closing():
    root.destroy()
    
root = tk.Tk()
root.title("Automoção de dados")
root.resizable(False, False)

# Definir um tema moderno
style = ttk.Style(root)
style.theme_use("clam")

window_width = 400
window_height = 350

# Obter a largura e altura da tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calcular a posição do canto superior esquerdo para centralizar a janela
x_cordinate = int((screen_width / 2) - (window_width / 2))
y_cordinate = int((screen_height / 2) - (window_height / 2))

root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")


# Configurar estilos personalizados para botões
style.configure("TButton", font=("Helvetica", 10), padding=6, relief="flat", background="#4A90E2", foreground="white")
style.map("TButton", background=[("active", "#357ABD")])

# Carregar e redimensionar a imagem de fundo
imagem_fundo = Image.open(r"assets/Impressora.png")
imagem_fundo = imagem_fundo.resize((300, 200), Image.LANCZOS)
imagem_fundo = ImageTk.PhotoImage(imagem_fundo)

# Criar um canvas para a imagem de fundo
canvas = tk.Canvas(root, width=400, height=300, highlightthickness=0, bg="#f0f0f0")
canvas.pack(fill="both", expand=True)

# Adicionar a imagem de fundo ao canvas
canvas.create_image(200, 125, anchor="center", image=imagem_fundo)

# Variável e barra de progresso estilizada
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, style="TProgressbar")
style.configure("TProgressbar", thickness=5)

# Criar botões modernos
botao_iniciar = ttk.Button(root, text="Iniciar Automação", command=lambda: iniciar_botao(botao_iniciar, progress_var), width=20)
botao_importar = ttk.Button(root, text="Importar Dados", command=lambda: iniciar_dados(botao_importar), width=20)

# Adicionar elementos ao canvas com melhor posicionamento
canvas.create_window(200, 250, anchor="center", window=progress_bar)
canvas.create_window(115, 300, anchor="center", window=botao_iniciar)
canvas.create_window(285, 300, anchor="center", window=botao_importar)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()