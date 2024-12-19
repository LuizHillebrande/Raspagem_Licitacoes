from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import customtkinter as ctk
from tkinter import messagebox
import os
import pyautogui

# Caminho para os arquivos .crx das extensões
extensao_buster = r'Buster Captcha Solver for Humans - Chrome Web Store 3.1.0.0.crx'
extensao_recaptcha = r'CAAHALKGHNHBABKNIPMCONMBICPKCOPL_0_0_0_2.crx'
# Configuração das opções do Chrome
chrome_options = Options()

# Adicionando as extensões ao Chrome
chrome_options.add_extension(extensao_buster)
chrome_options.add_extension(extensao_recaptcha)

# Caminho do ChromeDriver
driver = webdriver.Chrome(options=chrome_options)

# Acesse o site
driver.get('https://www.compras.rs.gov.br/egov2/acessarAtaEletronica.ctlx?idOffer=328141&siteContext=Celic')


time.sleep(5)
print('procurando')
try:
    # Tenta localizar a imagem na tela
    posicao = pyautogui.locateOnScreen('image.png', confidence=0.8)  # 'confidence' aumenta a tolerância em caso de pequenas diferenças
    if posicao:
        print('Imagem localizada!', posicao)
        
        # Obtém o centro da posição da imagem localizada
        centro = pyautogui.center(posicao)
        
        # Clica na posição central da imagem
        pyautogui.click(centro.x, centro.y)
        print('Clique realizado na imagem.')
    else:
        print('Imagem não encontrada na tela.')
except Exception as e:
    print(f'Erro ao localizar ou clicar na imagem: {e}')

time.sleep(40)

ver_ata = WebDriverWait(driver,5).until(
    EC.element_to_be_clickable((By.XPATH,"//button[@id='viewElectronicRecord']"))
)
ver_ata.click()

# Espera para garantir que as extensões sejam carregadas corretamente
time.sleep(10)

# Fechar o navegador
driver.quit()
