import pyautogui
from time import sleep
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sys
import threading
import requests
from tkinter import messagebox
import customtkinter as ctk

# Configurar o tema dark
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# Lista de sites de licitação/contratos
sites_licitacao = ["Portal Nacional de Contratações Públicas", "Outros Sites de Licitação"]

# Função chamada ao selecionar o site
def selecionar_site():
    site = combo_sites.get()
    if site:
        if site == "Portal Nacional de Contratações Públicas":
            iniciar_raspagem_Portal_Nacional_de_Contratacoes_Publicas()
    else:
        messagebox.showwarning("Atenção", "Por favor, selecione um site.")

# Função para realizar a raspagem do Portal Nacional de Contratações Públicas
def iniciar_raspagem_Portal_Nacional_de_Contratacoes_Publicas():
    driver = webdriver.Chrome()
    driver.get('https://pncp.gov.br/app/contratos?q=&pagina=1')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Link", "CNPJ", "Razão Social"])  # Cabeçalho da planilha

    elemento = driver.find_element("css selector", "label[for='status-vigente']")
    conteudo_before = driver.execute_script(
        "return window.getComputedStyle(arguments[0], '::before').getPropertyValue('content');",
        elemento
    )

    pesquisar = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@class='br-button primary']"))
    )

    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "br-list")))

    # Encontrar os botões únicos (garantindo apenas os 10 primeiros)
    botoes = driver.find_elements(By.XPATH, "//a[contains(@class, 'br-item') and contains(@title, 'Acessar item.')]")
    botoes_unicos = list(dict.fromkeys([botao.get_attribute('href') for botao in botoes]))[:10]

    print(f"Total de botões únicos encontrados: {len(botoes_unicos)}")

    for i, link in enumerate(botoes_unicos, start=1):
        try:
            driver.get(link)
            sleep(2)  # Um pouco de espera para a página carregar

            try:
                cnpj_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'CNPJ/CPF:')]/following-sibling::span"))
                )
                cnpj = cnpj_element.text
                print(f"Botão {i}: CNPJ encontrado - {cnpj}")
            except Exception as e:
                print(f"Botão {i}: Erro ao extrair o CNPJ - {e}")
                cnpj = "Não encontrado"

            try:
                razao_social_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'Nome/Razão social:')]/following-sibling::span"))
                )
                razao_social = razao_social_element.text
                print(f"Botão {i}: Razão social encontrada - {razao_social}")
            except Exception as e:
                print(f"Botão {i}: Erro ao extrair a Razão Social - {e}")
                razao_social = "Não encontrado"

            ws.append([link, cnpj, razao_social])
        except Exception as e:
            print(f"Erro ao processar o botão {i}: {e}")

    wb.save("dados_fornecedores.xlsx")
    driver.quit()
    messagebox.showinfo("Concluído", "Raspagem concluída com sucesso!")

# Criar a janela principal
root = ctk.CTk()
root.title("Raspagem de Dados de Licitações")
root.geometry("500x350")
root.resizable(False, False)

# Centralizar componentes
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(0, weight=1)

# Combobox para selecionar o site
ctk.CTkLabel(root, text="Selecione o Site", font=("Arial", 16)).grid(row=0, column=0, pady=10)
combo_sites = ctk.CTkComboBox(root, values=sites_licitacao, width=200)
combo_sites.grid(row=1, column=0, pady=10)

# Botão para confirmar a seleção
btn_selecionar = ctk.CTkButton(root, text="Confirmar Seleção", command=selecionar_site)
btn_selecionar.grid(row=2, column=0, pady=10)

# Rodar o loop principal
root.mainloop()
