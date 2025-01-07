import openpyxl
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import messagebox
from time import sleep
import random
import customtkinter as ctk
import re

def iniciar_raspagem_compras_gov(ano):
    # Inicializa o navegador com undetected_chromedriver
    driver = uc.Chrome()
    driver.get('https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras')
    driver.maximize_window()
    sleep(random.uniform(1, 3))  # Delay aleatório

    # Criar a planilha para salvar os dados
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Link", "CNPJ", "Razão Social"])  

    # Interage com o filtro
    filtro = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='p-radiobutton-box']"))
    )
    filtro.click()
    sleep(random.uniform(1, 3))  # Delay aleatório


    input_ano = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@placeholder='Ex: 102021']"))
    )
    input_ano.send_keys(ano)
    sleep(random.uniform(1, 3))

        # Clica no botão pesquisar
    button_pesquisar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@class='br-button is-primary']"))
    )
    button_pesquisar.click()
    sleep(random.uniform(1, 3))  # Delay aleatório

    # Coleta os links únicos
    elements = driver.find_element(By.XPATH, "//i[@class='fa fa-tasks']")
    qtde_apps_card = len(driver.find_elements(By.XPATH, "//i[@class='fa fa-tasks']"))
                         
    print(qtde_apps_card)
    sleep(15)
    
    wb.save("dados_vencedores_portal_compras_publicas.xlsx")
    sleep(15)
    driver.quit()
    messagebox.showinfo("Concluído", "Raspagem concluída com sucesso!")

def iniciar():
    ano = entry_ano.get()  # Obtém o valor digitado na interface
    if ano:
        iniciar_raspagem_compras_gov(ano)  # Passa o ano para a função de raspagem
    else:
        messagebox.showwarning("Erro", "Por favor, insira um ano.")

root = ctk.CTk()
root.title("Raspagem Compras Gov")

# Campo de entrada para o ano
entry_ano = ctk.CTkEntry(root, placeholder_text="Digite o ano (Ex: 102021)")
entry_ano.pack(pady=10)

# Botão para iniciar o processo
button_iniciar = ctk.CTkButton(root, text="Iniciar Raspagem", command=iniciar)
button_iniciar.pack(pady=10)

# Rodar a interface
root.mainloop()
