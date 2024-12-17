from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from time import sleep
import pyautogui

def limpa_campo():
    for _ in range(10):
        pyautogui.press('right')
    for _ in range(10):
        pyautogui.press('backspace')
def extrair_bllcompras():
    driver = webdriver.Chrome()
    driver.get('https://bllcompras.com/Process/ProcessSearchPublic?param1=0#')

    busca_avancada = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//a[@title='BUSCA AVANÇADA']"))
    )

    busca_avancada.click()

    public_inicio = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@data-target='#DateStart']"))
    )

    public_inicio.click()
    limpa_campo()
    public_inicio.send_keys('17/09/2024')



    public_final = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@data-target='#DateEnd']"))
    )

    public_final.click()
    limpa_campo()
    public_final.send_keys('20/10/2024')

    icon = driver.find_element(By.CSS_SELECTOR, "button i.fas.fa-search")
    icon.click()

    
    sleep(10)
    driver.quit
extrair_bllcompras()