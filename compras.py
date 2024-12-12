import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import customtkinter as ctk
from tkinter import messagebox
from time import sleep


def iniciar_raspagem_compras_gov():
    driver = webdriver.Chrome()
    driver.get('https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras')
    sleep(1)

    # Criar a planilha para salvar os dados
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Link", "CNPJ", "Razão Social"])  

    sleep(3)
    try:
        filtro = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH,"//div[@class='p-radiobutton-box']"))
        )
        filtro.click()

        sleep(3)
        button_pesquisar = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@class='br-button is-primary']"))
        )
        button_pesquisar.click()
        sleep(2)
        try:
            botoes_compras_gov = driver.find_elements(By.XPATH, "//button[contains(@class, 'button-class')]")
            botoes_unicos = list(dict.fromkeys([botao.get_attribute('href') for botao in botoes_compras_gov if botao.get_attribute('href')]))

            print(f"Total de botões encontrados: {len(botoes_unicos)}")

            for i, link in enumerate(botoes_unicos, start=1):
                try:
                    driver.get(link)
                except Exception as msg:
                    print(f"Erro ao acessar o link {link}: {msg}")
        except Exception as e:
            print(f"Erro ao localizar botões: {e}")
        

    except Exception as e:
        print(e)
    #wb.save("dados_vencedores_portal_compras_publicas.xlsx")
    driver.quit()
    messagebox.showinfo("Concluído", "Raspagem concluída com sucesso!")

iniciar_raspagem_compras_gov()