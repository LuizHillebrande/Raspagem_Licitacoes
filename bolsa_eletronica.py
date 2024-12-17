from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from time import sleep

def raspar_bolsa_sp():
    driver = webdriver.Chrome()
    driver.get('https://www.bec.sp.gov.br/bec_pregao_UI/OC/pesquisa_publica.aspx?chave=')

    # Aguarda e clica no dropdown (menu suspenso)
    situacao = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='fake_select']"))
    )
    situacao.click()

    # Move o mouse para "Encerrado" e clica
    encerrado = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//li[@data-descricao='Encerrado']"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(encerrado).click().perform()

    # Força o dropdown a permanecer aberto (se necessário)
    dropdown = driver.find_element(By.XPATH, "//ul[@class='grupo_status']")
    driver.execute_script("arguments[0].style.display = 'block';", dropdown)

    # Aguarda e clica na subopção "ENCERRADO COM VENCEDOR"
    encerrado_com_vencedor = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//optgroup[@label='Encerrado']//option[text()='ENCERRADO COM VENCEDOR']"))
    )
    actions.move_to_element(encerrado_com_vencedor).click().perform()
    print("Clicado em 'ENCERRADO COM VENCEDOR'")

    sleep(5)  # Para observar o resultado no navegador
    driver.quit()

# Executa a função
raspar_bolsa_sp()
