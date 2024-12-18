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
    public_inicio.send_keys('17/05/2024')

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
    original_url = driver.current_url

    elements = driver.find_elements(By.CSS_SELECTOR, "i.fas.fa-info-circle")
    for index, element in enumerate(elements, start=1):
        print(f"Elemento {index} de {len(elements)}")

        elements = driver.find_elements(By.CSS_SELECTOR, "i.fas.fa-info-circle")
        element = elements[index]
        
        # Reencontrando o elemento antes de clicar para garantir que ele ainda é válido
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(element)
        )
        element.click()
        time.sleep(2)
        print("BOTAO 1 CLICADO")
        

        WebDriverWait(driver, 10).until(
        lambda driver: len(driver.window_handles) > 1  # Garantir que uma nova janela foi aberta
        )
        
        # Trocar para a nova janela
        new_window = [window for window in driver.window_handles if window != original_window][0]
        driver.switch_to.window(new_window)

        current_url = driver.current_url
        print("URL da nova janela:", current_url)

        try:
            # Exemplo: esperando por um botão que deve aparecer na nova janela
            relatorios_button = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//button[@title='Relatórios']"))
            )
            print("Elemento 'Relatórios' apareceu, a página foi carregada corretamente.")
            
            # Clicar no botão 'Relatórios' caso apareça
            relatorios_button.click()
            time.sleep(2)
            
        except Exception as e:
            print("O elemento 'Relatórios' não apareceu ou houve um erro:", e)
        
        try:
            # Aguarda a presença do elemento <b> e verifica o texto
            mensagem_elemento = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//b"))
            )
            mensagem_texto = mensagem_elemento.text

            # Verifica o conteúdo do texto
            if "relatórios ainda não estão disponíveis" in mensagem_texto:
                print("Relatórios não estão disponíveis. Pulando para o próximo item...")
                pyautogui.hotkey('ctrl','w')
                driver.switch_to.window(original_window)
                continue  # Passa para o próximo elemento da lista
        except Exception:
            print("Elemento <b> não encontrado. Continuando o download...")

        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, "tr"))
        )

        linhas = driver.find_elements(By.TAG_NAME, "tr")

        for linha in linhas:
            try:
                if "VENCEDORES DO PROCESSO" in linha.text:
                    print("Linha encontrada com 'VENCEDORES DO PROCESSO'.")

                    # Procura pelo ícone de download dentro da linha
                    botao_download = linha.find_element(By.CSS_SELECTOR, "i.fas.fa-download")
                    
                    # Clica no botão de download
                    driver.execute_script("arguments[0].scrollIntoView(true);", botao_download)
                    time.sleep(1)
                    
                    print("Download iniciado com sucesso.")
                    ActionChains(driver).move_to_element(botao_download).click().perform()

                    time.sleep(25)

                    break  # Sai do loop após encontrar e clicar no botão
            except Exception as e:
                 print(f"Erro ao processar a linha: {e}")
        
        pyautogui.hotkey('ctrl','w')
        driver.switch_to.window(original_window)

    # Espera um pouco mais para garantir que o download seja iniciado
    time.sleep(5)  # Pausa de 5 segundos antes de fechar o driver
    
    driver.quit()  # Fechamento do driver após a execução

extrair_bllcompras()
