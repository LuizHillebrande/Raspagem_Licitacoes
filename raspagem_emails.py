import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep

def consulta_cnpj_gratis():
    arquivos = [
        "dados_fornecedores.xlsx",
        "dados_vencedores_diario_sp.xlsx",
        "dados_fornecedores_3.xlsx",
        "dados_fornecedores_4.xlsx",
        "dados_fornecedores_5.xlsx"
    ]

    # Configuração do Selenium
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def limpar_cnpj(cnpj):
        """Remove caracteres indesejados do CNPJ."""
        return ''.join(filter(str.isdigit, str(cnpj)))  # Mantém apenas dígitos

    try:
        for arquivo in arquivos:
            if os.path.exists(arquivo):  # Verifica se o arquivo existe
                print(f"Lendo arquivo: {arquivo}")
                # Ler o arquivo
                df = pd.read_excel(arquivo)

                # Iterar sobre a coluna 1 (índice 1 no Python)
                for cnpj in df.iloc[:, 1]:  # Segunda coluna do arquivo
                    cnpj_limpo = limpar_cnpj(cnpj)  # Limpar o CNPJ
                    if cnpj_limpo:  # Garantir que o CNPJ não esteja vazio
                        link = f"https://consulta.guru/consultar-cnpj-gratis/{cnpj_limpo}"
                        print(f"Acessando: {link}")

                        # Acessar o link
                        driver.get(link)
                        sleep(4)  # Pausa para garantir o carregamento da página

            else:
                print(f"Arquivo não encontrado: {arquivo}")
    finally:
        driver.quit()

def consulta_cnpj_ja():
    arquivos = [
        "dados_fornecedores.xlsx",
        "dados_fornecedores_2.xlsx",
        "dados_fornecedores_3.xlsx",
        "dados_fornecedores_4.xlsx",
        "dados_fornecedores_5.xlsx"
    ]

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    def limpar_cnpj(cnpj):
        """Remove caracteres indesejados do CNPJ."""
        return ''.join(filter(str.isdigit, str(cnpj)))  # Mantém apenas dígitos

    resultados = []

    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        consulta_contador = 0

        for arquivo in arquivos:
            if os.path.exists(arquivo):  # Verifica se o arquivo existe
                print(f"Lendo arquivo: {arquivo}")
                df = pd.read_excel(arquivo)

                for cnpj in df.iloc[:, 1]:  # Segunda coluna do arquivo
                    cnpj_limpo = limpar_cnpj(cnpj)
                    if cnpj_limpo:
                        link = f"https://cnpja.com/office/{cnpj_limpo}"
                        print(f"Acessando: {link}")

                        driver.get(link)
                        sleep(2)  # Pausa para o carregamento da página

                        try:
                            # Espera o email aparecer na página, usando XPath para capturar texto com "@"
                            email_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '@')]"))
                            )
                            email = email_element.text.strip()  # Extrai o texto do email
                            print(f"Email encontrado para {cnpj_limpo}: {email}")
                        except Exception as e:
                            email = "Não encontrado"
                            print(f"Email não encontrado para {cnpj_limpo}: {e}")

                        # Adiciona o resultado à lista
                        resultados.append({"CNPJ": cnpj_limpo, "Email": email})

                        df_resultados = pd.DataFrame(resultados)
                        df_resultados.to_excel("resultado_cnpja.xlsx", index=False)
                        print("Resultados salvos em 'resultado_cnpja.xlsx'")

                        consulta_contador += 1

                        if consulta_contador % 5 == 0:  # A cada 5 consultas
                            print("Aguardando 60 segundos para evitar bloqueio...")
                            sleep(60)  # Aguarda 1 minuto
    finally:
        driver.quit()
    
    # Salvar os resultados em um arquivo Excel
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel("resultado_cnpja.xlsx", index=False)
    print("Resultados salvos em 'resultado_cnpja.xlsx'")

consulta_cnpj_ja()