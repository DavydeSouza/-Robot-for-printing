from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
from config import executar_stored_procedure, obter_ips_do_banco  # Importando a função do arquivo de banco de dados
import pyodbc  # Para conexão com o banco de dados SQL Server

def abrir_navegador(url):
    driver = webdriver.Chrome()  # Certifique-se de que o chromedriver esteja instalado e no PATH
    driver.get(url)
    time.sleep(5)  # Aguarde o carregamento da página
    return driver

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

def coletar_dados(driver):
    try:
        tabela = driver.find_element(By.ID, "lock")
        linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")

        dados = []
        for linha in linhas:
            try:
                colunas = linha.find_elements(By.TAG_NAME, "td")
                lock_num = colunas[0].find_element(By.TAG_NAME, "label").text
                lock_name = colunas[1].find_element(By.CSS_SELECTOR, "input.lockName").get_attribute("value")
                page_limit_max = colunas[5].find_element(By.CSS_SELECTOR, "input.lockPageLimitMax").get_attribute("value")
                last_td_value = colunas[6].text

                dados.append([lock_num, lock_name, page_limit_max, last_td_value])

            except Exception as e:
                print(f"Erro ao processar linha: {e}")
        
        df = pd.DataFrame(dados, columns=["ID", "Nome", "Paginas Disponiveis", "Qtd Impressa"])
        df.to_excel("dados_completos.xlsx", index=False)
        print("Dados salvos em 'dados_completos.xlsx'.")
        
    except Exception as e:
        print(f"Erro ao coletar os dados: {e}")



def iniciar_automacao():
    senha = 'initpass'
    
    # Obter os IPs do banco de dados
    ips = obter_ips_do_banco()
    
    for ip in ips:
        print(f"Iniciando automação para o IP: {ip}")
        login_url = f"http://{ip}"  # Construir a URL com base no IP
        
        # Executar a stored procedure antes de iniciar o processo
        

        # Abrir o navegador e fazer login
        driver = abrir_navegador(login_url)
        fazer_login(driver, senha)
        
        # Acessar as páginas necessárias e coletar os dados
        acessar_pagina_admin(driver)
        acessar_pagina_controle(driver)
        coletar_dados(driver)
        executar_stored_procedure(id_imp=1)
        
        # Fechar o navegador
        driver.quit()

# Iniciar o processo de automação
iniciar_automacao()
