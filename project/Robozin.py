import tkinter as tk
from tkinter import messagebox, Canvas
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
from PIL import Image, ImageTk
from config import executar_stored_procedure, obter_ips_do_banco, obter_setor_por_ip ,obter_tipo_impressora_por_ip # Importando a função do arquivo de banco de dados
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

def coletar_dados(driver, qt, setor, tipo_impressora):
    try:
        wait = WebDriverWait(driver, 10)  # Espera de até 10 segundos

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
                    dados_acumulados.append([setor, lock_num, lock_name, page_limit_max, last_td_value])
                except Exception as e:
                    print(f"Erro ao processar linha: {e}")
        
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
                    dados_acumulados.append([setor, lock_num, lock_name, page_limit_max, last_td_value])
                except Exception as e:
                    print(f"Erro ao processar linha: {e}")

    except Exception as e:
        print(f"Erro ao coletar os dados: {e}")

# Função para salvar todos os dados em um único arquivo Excel
def salvar_dados_em_excel(nome_arquivo="dados_completos.xlsx"):
    try:
        df = pd.DataFrame(dados_acumulados, columns=["Setor", "ID", "Nome", "Limite de Folhas", "Qtd Impressa"])
        df.to_excel(nome_arquivo, index=False)
        print(f"Dados salvos em '{nome_arquivo}'.")
    except Exception as e:
        print(f"Erro ao salvar os dados: {e}")


def iniciar_automacao(botao, progress_var):
    botao.config(text="Processando...", state="disabled")
    progress_var.set(10)
    
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
        tipo_impressora = obter_tipo_impressora_por_ip(ip)  # Obter o tipo da impressora
        
        if tipo_impressora:
            coletar_dados(driver, qt, setor, tipo_impressora)
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
        iniciar_automacao(botao, progress_var)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        botao.config(text="Iniciar Automação", state="normal")

# Criar a interface gráfica com tkinter
root = tk.Tk()
root.title("Automação Impressora")

# Configurar o tamanho da janela e impedir o redimensionamento
root.geometry("400x300")
root.resizable(False, False)

# Carregar a imagem de fundo
imagem_fundo = Image.open(r"C:/Users/Davy/Downloads/ImpZin.png")  # Certifique-se que o caminho da imagem está correto
imagem_fundo = imagem_fundo.resize((400, 300), Image.LANCZOS) # Redimensionar a imagem para 400x300
imagem_fundo = ImageTk.PhotoImage(imagem_fundo)

# Criar um canvas onde a imagem de fundo será desenhada
canvas = Canvas(root, width=400, height=300)
canvas.pack(fill="both", expand=True)

# Adicionar a imagem de fundo ao canvas
canvas.create_image(0, 0, anchor="nw", image=imagem_fundo)

# Variável para a barra de progresso
progress_var = tk.DoubleVar()


# Criar o botão "Iniciar Automação" sobre o canvas
botao_iniciar = ttk.Button(root, text="Iniciar Automação", command=lambda: iniciar_botao(botao_iniciar, progress_var), width=20)

# Adicionar o botão ao canvas em uma posição visível
canvas.create_window(200, 250, anchor="center", window=botao_iniciar)

# Executar o loop principal do tkinter
root.mainloop()