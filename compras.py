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
from selenium.webdriver.common.action_chains import ActionChains

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
    elements = driver.find_elements(By.XPATH, "//i[@class='fa fa-tasks']")
    qtde_apps_card = len(driver.find_elements(By.XPATH, "//i[@class='fa fa-tasks']"))

    for index, element in enumerate(elements, start=1):
        print(f"Elemento {index} de {len(elements)}")
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        action = ActionChains(driver)
        action.move_to_element(element).click().perform()   
        print('clicado')  
        action.move_to_element(element).click().perform()  
        print('clicado dnv')
        print(qtde_apps_card)
        
        sleep(random.uniform(1, 3))
        elements_details = driver.find_elements(By.XPATH, "//i[@class='fa-tasks fas']")
        if elements_details:
            for index, elementt in enumerate(elements_details, start=1):
                print(f"Clicando no elemento {index} de {len(elements_details)}")

                # Rolando suavemente até o elemento
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", elementt)

                # Ação para mover até o elemento e clicar
                action = ActionChains(driver)
                action.move_to_element(elementt).click().perform()

                print(f'Elemento {index} clicado com sucesso.')
                

                elementos_situacao = WebDriverWait(driver, 10).until(
                    EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "div[data-test='situacao-proposta']"))
                )

                # Itera sobre os elementos encontrados
                for elemento_situacao in elementos_situacao:
                    # Obtém o texto do elemento
                    situacao_texto = elemento_situacao.text.strip()

                    # Verifica se o texto é "Adjudicada"
                    if situacao_texto.lower() == "adjudicada":
                        print("A proposta está adjudicada. Realizando ações subsequentes...")
                        chevron_icon = elemento_situacao.find_element(By.XPATH, ".//i[@class='fas fa-chevron-down']")
                        
                        # Clicar no ícone
                        chevron_icon.click()
                        print("Clicou no ícone de chevron.")
                        break  
                    

        else:
            print("Nenhum elemento encontrado.")
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
