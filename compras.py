import openpyxl
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import messagebox
from time import sleep
import random

def iniciar_raspagem_compras_gov():
    # Inicializa o navegador com undetected_chromedriver
    driver = uc.Chrome()
    driver.get('https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras')
    sleep(random.uniform(1, 3))  # Delay aleatório

    # Criar a planilha para salvar os dados
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Link", "CNPJ", "Razão Social"])  

    try:
        # Interage com o filtro
        filtro = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@class='p-radiobutton-box']"))
        )
        filtro.click()
        sleep(random.uniform(1, 3))  # Delay aleatório

        # Clica no botão pesquisar
        button_pesquisar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@class='br-button is-primary']"))
        )
        button_pesquisar.click()
        sleep(random.uniform(1, 3))  # Delay aleatório

        # Coleta os links únicos
        botoes_compras_gov = driver.find_elements(By.XPATH, "//button[contains(@class, 'button-class')]")
        botoes_unicos = list(dict.fromkeys([botao.get_attribute('href') for botao in botoes_compras_gov if botao.get_attribute('href')]))

        print(f"Total de botões encontrados: {len(botoes_unicos)}")

        for i, link in enumerate(botoes_unicos, start=1):
            try:
                driver.get(link)
                sleep(random.uniform(2, 5))  # Simula tempo de leitura da página
            except Exception as msg:
                print(f"Erro ao acessar o link {link}: {msg}")

    except Exception as e:
        print(f"Erro na execução: {e}")

    finally:
        # Salva a planilha e encerra o navegador
        wb.save("dados_vencedores_portal_compras_publicas.xlsx")
        driver.quit()
        messagebox.showinfo("Concluído", "Raspagem concluída com sucesso!")

iniciar_raspagem_compras_gov()
