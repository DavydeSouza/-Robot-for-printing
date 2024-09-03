from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

def abrir_navegador(url):
    driver = webdriver.Chrome()  # Certifique-se de que o chromedriver esteja instalado e no PATH
    driver.get(url)
    time.sleep(5)  # Aguarde o carregamento da página
    return driver

def fazer_login(driver, senha):
    senha_campo = driver.find_element(By.ID, "LogBox")  # Ajuste o seletor conforme necessário
    senha_campo.send_keys(senha)
    
    # Pressiona o botão de login
    driver.find_element(By.ID, "login").click()
    time.sleep(5)  # Aguarde a autenticação

def acessar_pagina_admin(driver):
    admin_link = driver.find_element(By.LINK_TEXT, "Administrator")
    admin_link.click()
    time.sleep(5)  # Aguarde o carregamento da página

def acessar_pagina_controle(driver):
    controle_link = driver.find_element(By.LINK_TEXT, "Restricted Functions 1-25")
    controle_link.click()
    time.sleep(5)  # Aguarde o carregamento da página

def coletar_dados(driver):
    try:
        # Coletar dados da tabela
        tabela = driver.find_element(By.ID, "lock")  # Ajuste o ID conforme necessário

        # Extrair linhas da tabela (even/odd)
        linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")

        dados = []
        for linha in linhas:
            try:
                # Pegar as colunas da linha
                colunas = linha.find_elements(By.TAG_NAME, "td")

                # Coletar os dados relevantes
                lock_num = colunas[0].find_element(By.TAG_NAME, "label").text  # Número do bloqueio
                lock_name = colunas[1].find_element(By.CSS_SELECTOR, "input.lockName").get_attribute("value")  # Nome do bloqueio
                page_limit_max = colunas[5].find_element(By.CSS_SELECTOR, "input.lockPageLimitMax").get_attribute("value")  # Limite máximo de página
                last_td_value = colunas[6].text  # Valor da última célula (ex: 130)

                # Adicionar dados à lista
                dados.append([lock_num, lock_name, page_limit_max, last_td_value])

            except Exception as e:
                print(f"Erro ao processar linha: {e}")
        
        # Criar DataFrame com os dados extraídos
        df = pd.DataFrame(dados, columns=["ID", "Nome", "Paginas Disponiveis", "Qtd Impressa"])
        
        # Salvar os dados em um arquivo Excel
        df.to_excel("dados_completos.xlsx", index=False)
        print("Dados salvos em 'dados_completos.xlsx'.")
        
    except Exception as e:
        print(f"Erro ao coletar os dados: {e}")

def iniciar_automacao():
    login_url = '#'
    senha = '#'
    
    # Abrir o navegador e fazer login
    driver = abrir_navegador(login_url)
    fazer_login(driver, senha)
    
    # Acessar as páginas necessárias e coletar os dados
    acessar_pagina_admin(driver)
    acessar_pagina_controle(driver)
    coletar_dados(driver)
    
    # Fechar o navegador
    driver.quit()

# Iniciar o processo de automação
iniciar_automacao()
 