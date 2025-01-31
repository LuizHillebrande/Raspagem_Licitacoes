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

caminho_arquivo_selecionado = ""

# Declaração global de entry_arquivo
entry_arquivo = None

def converter_para_csv(arquivo_excel):
    try:
        # Carrega o arquivo Excel
        df = pd.read_excel(arquivo_excel)

        # Cria a pasta "emails" no diretório atual, se não existir
        pasta_destino = "./emails"
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        # Define o nome do arquivo CSV dentro da pasta "emails"
        nome_arquivo = os.path.basename(arquivo_excel)  # Nome do arquivo Excel
        nome_csv = os.path.splitext(nome_arquivo)[0] + ".csv"
        caminho_csv = os.path.join(pasta_destino, nome_csv)

        # Salva o arquivo como CSV na pasta "emails"
        df.to_csv(caminho_csv, index=False, encoding='utf-8')
        messagebox.showinfo("Sucesso", f"Arquivo convertido para CSV com sucesso!\nSalvo em: {caminho_csv}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def selecionar_e_converter():
    # Abre o explorador de arquivos para selecionar o arquivo Excel
    arquivo_selecionado = filedialog.askopenfilename(
        title="Selecione um arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os Arquivos", "*.*"))
    )
    
    if arquivo_selecionado:
        # Chama a função para converter o arquivo selecionado
        converter_para_csv(arquivo_selecionado)

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
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(), 'teste_mautic')]"))  # Ajuste o texto conforme necessário
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
    sleep(3)

    # Selecionar a coluna de razão social
    #razao_social_coluna = WebDriverWait(driver, 5).until(
        #EC.presence_of_element_located((By.XPATH, "//div[@id='lead_field_import_razao_social_chosen']"))
    #)
    #razao_social_coluna.click()

    pyautogui.click(992,616, duration=1)
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
                sleep(5)
                driver.get('https://marilirequiaseguros.com.br/mautic/s/login')
                driver.maximize_window()

                sleep(2)
                canais_dropdown = WebDriverWait(driver,5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a#mautic_channels_root .arrow.pull-right.text-right"))
                )
                canais_dropdown.click()
                sleep(2)

                emails = WebDriverWait(driver,5).until(
                    EC.element_to_be_clickable((By.XPATH,"//a[@data-menu-link='mautic_email_index']"))
                )
                emails.click()
                sleep(5)

                tr_element = driver.find_element(By.XPATH, "//td[text()='39']/ancestor::tr")
                button = tr_element.find_element(By.CSS_SELECTOR, "button.dropdown-toggle")  # Encontra o botão dentro do tr
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                sleep(2)

                #link_enviar = driver.find_element(By.XPATH, "//td[text()='38']/ancestor::tr//a[@href='/mautic/s/emails/send/38']")

                # Garantir que o link seja clicável e clicar nele
                #WebDriverWait(driver, 10).until(EC.element_to_be_clickable(link_enviar)).click()
                pyautogui.click(337,497,duration=1)
                sleep(3)
                monitorar_elemento_envio_emails(xpath, url, driver)
                break
        except NoSuchElementException:
            print("Elemento não encontrado, tentando novamente em 30 segundos...")

        # Aguarda 30 segundos antes de tentar novamente
        time.sleep(30)

def monitorar_elemento_envio_emails(xpath, url, driver):
    while True:
        try:
            # Procura pelo elemento usando o XPath
            elemento = driver.find_element(By.XPATH, xpath)
            if elemento.is_displayed():
                print("Elemento encontrado! Texto:", elemento.text)
                messagebox.showinfo('Sucesso!', 'E-mails enviados!')
                driver.quit
                break
        except NoSuchElementException:
            print("Elemento não encontrado, tentando novamente em 30 segundos...")

        # Aguarda 30 segundos antes de tentar novamente
        time.sleep(30)

# Configurações
url = "URL_DA_SUA_PAGINA"  # Substitua pela URL da página a ser monitorada
xpath = '//div[@class="panel-heading"]//h4[contains(text(), "Sucesso!")]'

def abrir_pasta_emails():
    global entry_arquivo
    global caminho_arquivo_selecionado
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

def criar_interface_mautic():
    global entry_arquivo
    root = ctk.CTk()
    root.title("Upload de Arquivo Excel")
    root.geometry("600x400")

    label_titulo = ctk.CTkLabel(root, text="Escolha o arquivo Excel para upload", font=("Arial", 14))
    label_titulo.pack(pady=10)

    entry_arquivo = ctk.CTkEntry(root, width=300)
    entry_arquivo.pack(pady=5)

    botao_selecionar = ctk.CTkButton(
        root, 
        text="Converter para CSV", 
        command=selecionar_e_converter,
        font=("Arial", 14)
    )
    botao_selecionar.pack(pady=60)

    botao_abrir_pasta = ctk.CTkButton(root, text="Abrir Pasta 'emails'", command=abrir_pasta_emails)
    botao_abrir_pasta.pack(pady=5)

    botao_upload = ctk.CTkButton(root, text="Fazer Upload", command=fazer_upload)
    botao_upload.pack(pady=10)

    root.mainloop()

criar_interface_mautic()  # Chama a função para criar a interface



