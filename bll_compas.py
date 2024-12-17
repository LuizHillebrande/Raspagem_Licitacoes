from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException


def limpa_campo():
    for _ in range(10):
        pyautogui.press('right')
    for _ in range(10):
        pyautogui.press('backspace')

def extrair_bllcompras():
    driver = webdriver.Chrome()
    driver.get('https://bllcompras.com/Process/ProcessSearchPublic?param1=0#')

    # Aguarda até que o elemento de seleção de status esteja presente
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "fkStatus"))
    )
    select = Select(select_element)
    select.select_by_visible_text("HOMOLOGADO")

    # Aguarda até que o botão de busca avançada seja clicável
    busca_avancada = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@title='BUSCA AVANÇADA']"))
    )
    busca_avancada.click()

    # Aguarda até que o campo de data de início seja clicável
    public_inicio = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@data-target='#DateStart']"))
    )
    public_inicio.click()
    limpa_campo()
    public_inicio.send_keys('17/09/2024')

    # Aguarda até que o campo de data final seja clicável
    public_final = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@data-target='#DateEnd']"))
    )
    public_final.click()
    limpa_campo()
    public_final.send_keys('20/10/2024')

    # Aguarda até que o ícone de pesquisa seja clicável
    icon = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button i.fas.fa-search"))
    )
    icon.click()

    # Espera até que a lista de elementos de informações esteja carregada
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "i.fas.fa-info-circle"))
    )

    original_window = driver.current_window_handle

    elements = driver.find_elements(By.CSS_SELECTOR, "i.fas.fa-info-circle")
    for index, element in enumerate(elements, start=1):
        print(f"Elemento {index} de {len(elements)}")
        
        # Reencontrando o elemento antes de clicar para garantir que ele ainda é válido
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(element)
        )
        element.click()
        time.sleep(2)

        all_windows_before = driver.window_handles  # Pega todas as janelas antes de clicar
        WebDriverWait(driver, 5).until(
            lambda driver: len(driver.window_handles) > len(all_windows_before)  # Aguardar até que o número de janelas mude
        )
        
        relatorios_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//button[@title='Relatórios']"))
        )
        relatorios_button.click()


        

        # Agora, mude para a nova janela
        new_window = [window for window in driver.window_handles if window != original_window][0]
        driver.switch_to.window(new_window)
        
        download_buttons = driver.find_elements(By.CSS_SELECTOR, "i.fas.fa-download")

        for button in download_buttons:
            parent = button.find_element(By.XPATH, "./ancestor::a")  # Encontrando o link pai do ícone, por exemplo
            if "vencedores do processo" in parent.text.lower():
                parent.click()
                break

    # Espera um pouco mais para garantir que o download seja iniciado
    time.sleep(5)  # Pausa de 5 segundos antes de fechar o driver
    
    driver.quit()  # Fechamento do driver após a execução

extrair_bllcompras()
