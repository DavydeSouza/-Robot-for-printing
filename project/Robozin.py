import tkinter as tk
from tkinter import messagebox, Canvas
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
from PIL import Image, ImageTk
from config import executar_stored_procedure, obter_ips_do_banco, obter_setor_por_ip  # Importando a função do arquivo de banco de dados
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

def coletar_dados(driver, qt, setor):
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

                # Verificar se o nome não está vazio antes de adicionar
                if lock_name.strip() == "":
                    continue  # Pular a linha se o nome estiver vazio

                # Adicionar o setor aos dados coletados
                dados.append([setor, lock_num, lock_name, page_limit_max, last_td_value])

            except Exception as e:
                print(f"Erro ao processar linha: {e}")
        
        # Gerar nome de arquivo com base no contador 'qt'
        nome_arquivo = f"dados_completos_{qt}.xlsx"
        
        # Adicionar o setor como uma nova coluna no DataFrame
        df = pd.DataFrame(dados, columns=["Setor", "ID", "Nome", "Limite de Folhas", "Qtd Impressa"])
        
        # Salvar os dados no arquivo Excel com o nome gerado
        df.to_excel(nome_arquivo, index=False)
        print(f"Dados salvos em '{nome_arquivo}'.")
        
    except Exception as e:
        print(f"Erro ao coletar os dados: {e}")


def iniciar_automacao(botao, progress_var):
    botao.config(text="Processando...", state="disabled")
    progress_var.set(10)  # Atualizar a barra de progresso (simulação)
    
    senha = 'initpass'
    
    # Obter os IPs do banco de dados
    ips = obter_ips_do_banco()
    
    # Inicializar o contador
    qt = 1
    
    index = 0
    while index < len(ips):
        ip = ips[index]
        print(f"Iniciando automação para o IP: {ip}")
        login_url = f"http://{ip}"  # Construir a URL com base no IP
        
        # Abrir o navegador e fazer login
        driver = abrir_navegador(login_url)
        fazer_login(driver, senha)
        
        # Acessar as páginas necessárias e coletar os dados
        acessar_pagina_admin(driver)
        acessar_pagina_controle(driver)
        
        # Obter o setor baseado no IP
        setor = obter_setor_por_ip(ip)
        
        # Coletar os dados e passar o setor como argumento
        coletar_dados(driver, qt, setor)
        
        # Executar a stored procedure após coletar os dados
        executar_stored_procedure(ip)
        
        # Fechar o navegador
        driver.quit()
        
        # Incrementar o índice e o contador
        index += 1
        qt += 1  # Incrementar o contador para o próximo arquivo ter um nome diferente
        progress_var.set(100 * (index / len(ips)))  # Atualizar a barra de progresso
    
    # Mostrar mensagem de conclusão
    messagebox.showinfo("Concluído", "Automação finalizada com sucesso!")
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
