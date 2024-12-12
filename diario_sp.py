import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import tkinter as tk
from tkinter import ttk, messagebox
import re

def iniciar_raspagem_compras_gov(dia_inicio, mes_inicio, ano_inicio, dia_fim, mes_fim, ano_fim):
    driver = webdriver.Chrome()
    driver.get('https://www.imprensaoficial.com.br/ENegocios/BuscaENegocios_14_1.aspx#12/12/2024')
    sleep(2)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Link", "CNPJ", "Razão Social"])

    # Preencher Status ENCERRADA
    encerrada = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//select[@id='content_content_content_Status_cboStatus']/option[text()='ENCERRADA']"))
    )
    encerrada.click()

    # Preencher Data Inicial
    xpath_dia_inicio = f"//select[@id='content_content_content_Status_cboAberturaSecaoInicioDia']/option[@value='{dia_inicio}']"
    xpath_mes_inicio = f"//select[@id='content_content_content_Status_cboAberturaSecaoInicioMes']/option[@value='{mes_inicio}']"
    xpath_ano_inicio = f"//select[@id='content_content_content_Status_cboAberturaSecaoInicioAno']/option[@value='{ano_inicio}']"
    
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_dia_inicio))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_mes_inicio))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_ano_inicio))).click()

    # Preencher Data Final
    xpath_dia_fim = f"//select[@id='content_content_content_Status_cboAberturaSecaoFimDia']/option[@value='{dia_fim}']"
    xpath_mes_fim = f"//select[@id='content_content_content_Status_cboAberturaSecaoFimMes']/option[@value='{mes_fim}']"
    xpath_ano_fim = f"//select[@id='content_content_content_Status_cboAberturaSecaoFimAno']/option[@value='{ano_fim}']"

    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_dia_fim))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_mes_fim))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_ano_fim))).click()

    sleep(3)
    
    buscar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@onclick='return verify();']"))
    )
    buscar.click()

    links = driver.find_elements(By.XPATH, "//a[contains(@id, 'ResultadoBusca_dtgResultadoBusca_hlkObjeto')]")

    # Coletar e remover duplicatas usando 'dict.fromkeys'
    links_unicos = list(dict.fromkeys([link.get_attribute('href') for link in links]))

    # Imprimir os resultados
    print(f"Total de links encontrados: {len(links_unicos)}")
    for i, link in enumerate(links_unicos, start=1):
        print(f"Acessando Link {i}: {link}")
        driver.get(link)

        homologacao_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'HOMOLOGAÇÃO')]")

        if homologacao_links:
            for homologacao_link in homologacao_links:
                homologacao_link.click()
                sleep(2)

                main_window = driver.current_window_handle

                
                WebDriverWait(driver, 10).until(EC.new_window_is_opened(driver.window_handles))

                #troca para janela do pop up
                all_windows = driver.window_handles
                popup_window = [window for window in all_windows if window != main_window][0]
                driver.switch_to.window(popup_window)

                try:
                    
                    detalhes_texto = driver.find_element(By.ID, "content_content_content_DetalheEvento_lblSintesePublicacao").text

                    cnpj_match = re.search(r'CNPJ\s*-\s*(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})', detalhes_texto)
                    razao_social_match = re.search(r'EMPRESA\s*-\s*(.*?)(?=\s*CNPJ)', detalhes_texto)

                    cnpj = cnpj_match.group(1) if cnpj_match else "Não encontrado"
                    razao_social = razao_social_match.group(1) if razao_social_match else "Não encontrado"

                    # Salva os dados na planilha
                    ws.append([link, cnpj, razao_social])
                    print(f"Dados capturados: CNPJ={cnpj}, Razão Social={razao_social}")

                except Exception as e:
                    print(f"Erro ao processar o link {link}: {e}")
                
                # Fechar o pop-up
                driver.close()

                # Voltar para a janela principal
                driver.switch_to.window(main_window)

                sleep(2)  # Tempo para garantir que a nova página carregue corretamente

        else:
            print(f"Não encontrado link de HOMOLOGAÇÃO na página {link}")
                
   
    wb.save("dados_vencedores_diario_sp.xlsx")
    print("Planilha salva como 'dados_vencedores_diario_sp.xlsx'")
    
    sleep(10)
    driver.quit()
    print("Navegador fechado.")

def capturar_datas():
    dia_inicio = combo_dia_inicio.get()
    mes_inicio = combo_mes_inicio.get()
    ano_inicio = combo_ano_inicio.get()

    dia_fim = combo_dia_fim.get()
    mes_fim = combo_mes_fim.get()
    ano_fim = combo_ano_fim.get()

    # Chama a função Selenium para preencher os dados
    try:
        iniciar_raspagem_compras_gov(dia_inicio, mes_inicio, ano_inicio, dia_fim, mes_fim, ano_fim)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Configuração inicial da janela principal
janela = tk.Tk()
janela.title("Seleção de Datas")
janela.geometry("400x400")

# Componentes para Data Inicial
ttk.Label(janela, text="Data Inicial:").pack(pady=5)
frame_inicio = tk.Frame(janela)
frame_inicio.pack(pady=5)

combo_dia_inicio = ttk.Combobox(frame_inicio, values=[str(i) for i in range(1, 32)], width=5)
combo_dia_inicio.set("1")
combo_dia_inicio.pack(side=tk.LEFT, padx=5)

combo_mes_inicio = ttk.Combobox(frame_inicio, values=[str(i) for i in range(1, 13)], width=5)
combo_mes_inicio.set("1")
combo_mes_inicio.pack(side=tk.LEFT, padx=5)

combo_ano_inicio = ttk.Combobox(frame_inicio, values=[str(i) for i in range(2004, 2026)], width=6)
combo_ano_inicio.set("2024")
combo_ano_inicio.pack(side=tk.LEFT, padx=5)

# Componentes para Data Final
ttk.Label(janela, text="Data Final:").pack(pady=5)
frame_fim = tk.Frame(janela)
frame_fim.pack(pady=5)

combo_dia_fim = ttk.Combobox(frame_fim, values=[str(i) for i in range(1, 32)], width=5)
combo_dia_fim.set("1")
combo_dia_fim.pack(side=tk.LEFT, padx=5)

combo_mes_fim = ttk.Combobox(frame_fim, values=[str(i) for i in range(1, 13)], width=5)
combo_mes_fim.set("1")
combo_mes_fim.pack(side=tk.LEFT, padx=5)

combo_ano_fim = ttk.Combobox(frame_fim, values=[str(i) for i in range(2004, 2026)], width=6)
combo_ano_fim.set("2024")
combo_ano_fim.pack(side=tk.LEFT, padx=5)

# Botão para capturar as datas e preencher no Selenium
btn_confirmar = ttk.Button(janela, text="Iniciar Raspagem", command=capturar_datas)
btn_confirmar.pack(pady=20)

# Inicia o loop da interface
tk.mainloop()
