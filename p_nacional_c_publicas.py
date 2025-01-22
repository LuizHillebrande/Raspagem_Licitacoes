import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import customtkinter as ctk
from tkinter import messagebox
import os
import threading
from datetime import datetime
import logging

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


data_mais_recente = None
data_mais_antiga = None
intervalo_label = None

def atualizar_intervalo_datas(datas):
    global data_mais_recente, data_mais_antiga
    if datas:
        data_atual_mais_recente = max(datas)
        data_atual_mais_antiga = min(datas)
        
        # Atualizar as variáveis globais
        if not data_mais_recente or data_atual_mais_recente > data_mais_recente:
            data_mais_recente = data_atual_mais_recente
        if not data_mais_antiga or data_atual_mais_antiga < data_mais_antiga:
            data_mais_antiga = data_atual_mais_antiga

        # Atualizar o Label na interface gráfica
        intervalo_label.configure(
            text=f"Filtrado de {data_mais_antiga.strftime('%d/%m/%Y')} até {data_mais_recente.strftime('%d/%m/%Y')}"
        )

def extrair_datas_da_pagina(driver):
    datas_extradas = []
    try:
        elementos = driver.find_elements(By.XPATH, "//div[@class='datatable-body-cell-label']")
        for elemento in elementos:
            texto = elemento.text.strip()
            try:
                # Tentar converter o texto em uma data
                data = datetime.strptime(texto, "%d/%m/%Y")
                datas_extradas.append(data)
            except ValueError:
                # Ignorar valores que não sejam datas
                continue
    except Exception as e:
        logging.error(f"Erro ao extrair datas: {e}")
    return datas_extradas

def extrair_data_assinatura(driver):
    try:
        strong_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//strong[text()='Data de assinatura:']"))
        )

        # Encontrar o <span> imediatamente após o <strong> encontrado
        span_data_element = strong_element.find_element(By.XPATH, "following-sibling::span")

        # Extrair o texto da data
        data_texto = span_data_element.text.strip()

        # Retornar a data extraída
        return data_texto
    except Exception as e:
        logging.error(f"Erro ao extrair a data de assinatura: {e}")
        return "Não encontrada"

# Configurar o tema dark
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

cnpj_count = 0

# Lista de sites de licitação/contratos
sites_licitacao = ["Portal Nacional de Contratações Públicas", "Outros Sites de Licitação"]


# Função para realizar a raspagem do Portal Nacional de Contratações Públicas
def iniciar_raspagem_Portal_Nacional_de_Contratacoes_Publicas():
    global cnpj_count
    driver = webdriver.Chrome()

    data_atual = datetime.now().strftime("%d-%m-%Y")
    nome_arquivo = f"dados_fornecedores_{data_atual}.xlsx"  # Nome com data no final

    # Criar a planilha para salvar os dados
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Link", "CNPJ", "Razão Social", "Data de Divulgação no PNCP", "Data de Assinatura"])

    for pagina in range(1, 1000):  # De 1 a 999
        url = f'https://pncp.gov.br/app/contratos?q=&pagina={pagina}'
        driver.get(url)

        # Aguardar o carregamento dos elementos
        time.sleep(3)  # Espera de 3 segundos para garantir que a página carregue

        try:
            # Encontrar os botões únicos
            botoes = driver.find_elements(By.XPATH, "//a[contains(@class, 'br-item') and contains(@title, 'Acessar item.')]")
            botoes_unicos = list(dict.fromkeys([botao.get_attribute('href') for botao in botoes]))  # Remove duplicatas
            
            print(f"Página {pagina}: Total de botões encontrados: {len(botoes_unicos)}")

            for i, link in enumerate(botoes_unicos, start=1):
                try:
                    driver.get(link)
                    time.sleep(2)  # Espera um pouco para a página carregar

                    try:
                        # Extrair o CNPJ
                        cnpj_element = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'CNPJ/CPF:')]/following-sibling::span"))
                        )
                        cnpj = cnpj_element.text
                        print(f"Botão {i}: CNPJ encontrado - {cnpj}")
                        cnpj_count += 1
                        cnpj_label.configure(text=f"CNPJs extraídos: {cnpj_count}")
                    except Exception as e:
                        print(f"Botão {i}: Erro ao extrair o CNPJ - {e}")
                        cnpj = "Não encontrado"

                    try:
                        # Extrair a Razão Social
                        razao_social_element = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'Nome/Razão social:')]/following-sibling::span"))
                        )
                        razao_social = razao_social_element.text
                        print(f"Botão {i}: Razão social encontrada - {razao_social}")
                    except Exception as e:
                        print(f"Botão {i}: Erro ao extrair a Razão Social - {e}")
                        razao_social = "Não encontrado"

                    datas_pagina = extrair_datas_da_pagina(driver)
                    if datas_pagina:
                        data_divulgacao = datas_pagina[0].strftime("%d/%m/%Y")  # Usando a primeira data extraída
                    else:
                        data_divulgacao = "Não encontrada"
                    
                    data_texto = extrair_data_assinatura(driver)

                    # Atualizar intervalo de datas
                    atualizar_intervalo_datas(datas_pagina)

                    print(f"Data mais recente: {data_mais_recente.strftime('%d/%m/%Y')}")
                    print(f"Data mais antiga: {data_mais_antiga.strftime('%d/%m/%Y')}")
                    
                    
                    # Salvar os dados na planilha
                    ws.append([link, cnpj, razao_social, data_divulgacao, data_texto])
                    wb.save(nome_arquivo)

                except Exception as e:
                    print(f"Erro ao processar o botão {i}: {e}")
          
        except Exception as e:
            print(f"Erro ao acessar a página {pagina}: {e}")

    # Salvar a planilha
    
    driver.quit()
    messagebox.showinfo("Concluído", "Raspagem concluída com sucesso!")

def iniciar_raspagem():
    threading.Thread(target=iniciar_raspagem_Portal_Nacional_de_Contratacoes_Publicas, daemon=True).start()

# Criando a interface gráfica com CustomTkinter
def criar_interface_pncp():
    global cnpj_label
    global intervalo_label

    # Configuração da janela
    janela = ctk.CTk()
    janela.title("Raspagem de Dados de Licitação")
    janela.geometry("500x300")

    # Título
    titulo = ctk.CTkLabel(janela, text="Raspagem PNCP", font=("Arial", 18))
    titulo.pack(pady=20)

    # Botão para iniciar a raspagem
    btn_iniciar = ctk.CTkButton(janela, text="Iniciar Raspagem", command=iniciar_raspagem, font=("Arial", 16))
    btn_iniciar.pack(pady=20)

    # Label para mostrar o número de CNPJs extraídos
    cnpj_label = ctk.CTkLabel(janela, text="CNPJs extraídos: 0", font=("Arial", 14))
    cnpj_label.pack(pady=10)

    intervalo_label = ctk.CTkLabel(janela, text="Filtrado de --/--/---- até --/--/----", font=("Arial", 14))
    intervalo_label.pack(pady=10)

    # Iniciar a interface
    janela.mainloop()

criar_interface_pncp()

