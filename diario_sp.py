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
import pyautogui

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
    buscar = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@onclick='return verify();']"))
    )
    buscar.click()
    sleep(2)

    # Página inicial
    pagina_inicial_url = driver.current_url
    print('Página inicial:', pagina_inicial_url)

    # Loop processando páginas
    while True:
        print("Extraindo dados da página atual...")
        proxima_pagina = captura_link(driver, ws, wb, pagina_inicial_url)

        if not proxima_pagina:
            break  # Se não há mais próxima página, encerra o loop

    wb.save("dados_vencedores_diario_sp.xlsx")
    print("Planilha salva como 'dados_vencedores_diario_sp.xlsx'")
    driver.quit()
    print("Navegador fechado.")

def captura_link(driver, ws, wb, pagina_inicial_url):
    links = driver.find_elements(By.XPATH, "//a[contains(@id, 'ResultadoBusca_dtgResultadoBusca_hlkObjeto')]")
    links_unicos = list(dict.fromkeys([link.get_attribute('href') for link in links]))

    print(f"Total de links encontrados nesta página: {len(links_unicos)}")

    for i, link in enumerate(links_unicos, start=1):
        print(f"Acessando Link {i}: {link}")
        driver.get(link)

        homologacao_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'HOMOLOGAÇÃO')]")
        sleep(2)
        if homologacao_links:
            for homologacao_link in homologacao_links:
                onclick_content = homologacao_link.get_attribute('onclick')
                match = re.search(r"AbreJanelaDetalhes\((\d+),(\d+),\"HOMOLOGAÇÃO\"\)", onclick_content)
                if match:
                    id_licitacao, id_evento = match.groups()
                    homologacao_url = f"https://www.imprensaoficial.com.br/ENegocios/popup/pop_e-nego_detalhes.aspx?IdLicitacao={id_licitacao}&IdEventoLicitacao={id_evento}"
                    print(f"Abrindo link de HOMOLOGAÇÃO: {homologacao_url}")
                    driver.get(homologacao_url)
                    sleep(2)

                    try:
                        detalhes_texto = driver.find_element(By.ID, "content_content_content_DetalheEvento_lblSintesePublicacao").text
                        cnpj_match = re.search(r'CNPJ\s*(?:Nº|N.º|:|-)?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', detalhes_texto)
                        cnpj = cnpj_match.group(1) if cnpj_match else "Não encontrado"

                        razao_social_match = re.search(r'(?:EMPRESA VENCEDORA:|a favor da empresa|EMPRESA\s*[:-]\s*)(.*?)(?=\s*CNPJ)', detalhes_texto)
                        razao_social = razao_social_match.group(1).strip() if razao_social_match else "Não encontrado"

                        ws.append([homologacao_url, cnpj, razao_social])
                        wb.save("dados_vencedores_diario_sp.xlsx")
                    except Exception as e:
                        print(f"Erro ao capturar detalhes: {e}")

        # Voltar à página inicial
        driver.get(pagina_inicial_url)
        sleep(2)

    # Tentar clicar em "Próxima" ao final da página
    try:
        botao_proxima = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@id='content_content_content_ResultadoBusca_PaginadorCima_btnProxima']"))
        )
        botao_proxima.click()
        sleep(2)
        return True  # Retorna True para indicar que há próxima página
    except Exception as e:
        print(f"Não há mais próxima página: {e}")
        return False  # Retorna False quando não há mais páginas
    


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
