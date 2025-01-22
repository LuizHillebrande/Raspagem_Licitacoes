import customtkinter as ctk
from tkinter import messagebox
from PIL import Image
import openpyxl
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import os
import re
import pdfplumber
import pandas as pd
import customtkinter as ctk
import json
import threading
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import tkinter as tk
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from openpyxl import load_workbook
from tkinter import filedialog
import cv2
import numpy as np

def converter_para_csv(pasta):
    # Verifica se a pasta existe
    if not os.path.exists(pasta):
        print(f"A pasta '{pasta}' não existe.")
        return
    
    # Lista todos os arquivos na pasta
    arquivos = os.listdir(pasta)
    
    for arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta, arquivo)
        
        # Verifica se o arquivo é do tipo Excel (pode ser .xls ou .xlsx)
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            try:
                # Carrega o arquivo Excel
                df = pd.read_excel(caminho_arquivo)
                
                # Define o nome do novo arquivo CSV
                nome_csv = arquivo.replace('.xlsx', '.csv').replace('.xls', '.csv')
                caminho_csv = os.path.join(pasta, nome_csv)
                
                # Salva como CSV
                df.to_csv(caminho_csv, index=False, encoding='utf-8')
                print(f"Arquivo {arquivo} convertido para CSV com sucesso!")
            except Exception as e:
                print(f"Erro ao converter o arquivo {arquivo}: {e}")

# Chama a função passando o caminho da pasta "emails"
pasta_emails = './emails'



def enviar_emails():
    driver = webdriver.Chrome()
    driver.get('https://marilirequiaseguros.com.br/mautic/s/login')
    driver.maximize_window()


    login = 'luiz.logika@gmail.com'
    password = 'Dev123@'
    logar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='username']"))
    )
    logar.send_keys(login)

    password_enter = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@type='password']"))
    )
    password_enter.send_keys(password)

    enter = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@class='btn btn-lg btn-primary btn-block']"))
    )
    enter.click()

    sleep(4)

    contacts = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//a[@data-menu-link='mautic_contact_index']"))
    )
    contacts.click()
    sleep(3)

    seta = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR,"i[class='fa fa-caret-down']"))
    )
    seta.click()
    sleep(2)

    importar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//a[@href='/mautic/s/contacts/import/new']"))
    )
    importar.click()
    sleep(2)


    limite = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='lead_import_batchlimit']"))
    )
    limite.clear()
    limite.send_keys('2000')
    sleep(2)

    import_csv = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='lead_import_file']"))
    )
    import_csv.send_keys(caminho_arquivo_selecionado)
    sleep(2)

    carregar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@id='lead_import_start']"))
    )
    carregar.click()

    propriedade_contato = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//div[@id='lead_field_import_owner_chosen']"))
    )
    propriedade_contato.click()
    sleep(1)

    opcao = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Scraping, Teste')]"))
    )
    opcao.click()

    segmento_contato = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//div[@id='lead_field_import_list_chosen']"))
    )
    segmento_contato.click()
    sleep(1)

    
    option_segmento = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(), 'Teste Mautic 2')]"))  # Ajuste o texto conforme necessário
    )
    option_segmento.click()

    # Selecionar a coluna de email
    email_coluna = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@id='lead_field_import_email_chosen']"))
    )
    email_coluna.click()
    sleep(1)

    # Selecionar a opção de email
    option_email = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(), 'Email')]"))  
    )
    option_email.click()

    # Selecionar a coluna de razão social
    razao_social_coluna = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@id='lead_field_import_razao_social_chosen']"))
    )
    razao_social_coluna.click()
    sleep(1)

    pyautogui.write('Company Name')
    sleep(1)

    # Enviar "Enter" para confirmar
    pyautogui.press('enter')
    sleep(2)

    #importacao_plano_fundo  = WebDriverWait(driver,5).until(
        #EC.element_to_be_clickable((By.XPATH,"//button[@id='lead_field_import_buttons_apply_toolbar']"))
    #)
    #importacao_plano_fundo.click()

    importacao_navegador = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@id='lead_field_import_buttons_save_toolbar']"))
    )
    importacao_navegador.click()
    sleep(2)

    nome_arquivo = os.path.basename(caminho_arquivo_selecionado)

    #link_arquivo = WebDriverWait(driver, 10).until(
        #EC.presence_of_element_located((By.XPATH, f"//a[contains(@href, 'view/') and contains(text(), '{nome_arquivo}')]"))
    #)

    # Agora, monitoramos o progresso
    monitorar_elemento(xpath, url, driver)

    
    sleep(15)
    driver.quit()

def monitorar_elemento(xpath, url, driver):
    while True:
        try:
            # Procura pelo elemento usando o XPath
            elemento = driver.find_element(By.XPATH, xpath)
            if elemento.is_displayed():
                print("Elemento encontrado! Texto:", elemento.text)
                break
        except NoSuchElementException:
            print("Elemento não encontrado, tentando novamente em 30 segundos...")

        # Aguarda 30 segundos antes de tentar novamente
        time.sleep(30)

# Configurações
url = "URL_DA_SUA_PAGINA"  # Substitua pela URL da página a ser monitorada
xpath = '//div[@class="panel-heading"]//h4[contains(text(), "Sucesso!")]'

def abrir_pasta_emails():
    # Abre o explorador de arquivos na pasta "emails"
    pasta_emails = './emails'
    if not os.path.exists(pasta_emails):
        print("A pasta 'emails' não existe.")
        return
    arquivo_selecionado = filedialog.askopenfilename(
    initialdir=pasta_emails,
    title="Escolha um arquivo CSV",
    filetypes=(("CSV Files", "*.csv"),)
)
    
    if arquivo_selecionado:
        global caminho_arquivo_selecionado
        caminho_arquivo_selecionado = arquivo_selecionado
        entry_arquivo.delete(0, ctk.END)  # Limpa a entrada anterior
        entry_arquivo.insert(0, arquivo_selecionado)  # Insere o caminho do arquivo selecionado

def fazer_upload():
    if caminho_arquivo_selecionado:
        print(f"Arquivo selecionado para upload: {caminho_arquivo_selecionado}")
        enviar_emails()  # Chama a função de envio dos emails
    else:
        print("Nenhum arquivo selecionado!")

root = ctk.CTk()

root.title("Upload de Arquivo Excel")
root.geometry("400x200")

# Label
label_titulo = ctk.CTkLabel(root, text="Escolha o arquivo Excel para upload", font=("Arial", 14))
label_titulo.pack(pady=10)

# Caixa de texto para mostrar o caminho do arquivo selecionado
entry_arquivo = ctk.CTkEntry(root, width=300)
entry_arquivo.pack(pady=5)

# Botão para abrir a pasta "emails"
botao_abrir_pasta = ctk.CTkButton(root, text="Abrir Pasta 'emails'", command=abrir_pasta_emails)
botao_abrir_pasta.pack(pady=5)

# Botão para fazer upload do arquivo (exemplo de ação)
botao_upload = ctk.CTkButton(root, text="Fazer Upload", command=fazer_upload)
botao_upload.pack(pady=10)

root.mainloop()



