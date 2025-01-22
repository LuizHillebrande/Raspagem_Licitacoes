import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import customtkinter as ctk
import glob  # Para buscar arquivos com padrão
from tkinter import messagebox

def criar_interface_raspagem_emails():
    janela = ctk.CTk()  # Janela principal
    janela.geometry("500x300")
    janela.title("Raspagem de E-mails")

    # Label para exibir o número de e-mails raspados
    label_contador = ctk.CTkLabel(janela, text="E-mails raspados: 0", font=("Arial", 16))
    label_contador.pack(pady=10)

    # Combobox para selecionar o site
    label_selecione = ctk.CTkLabel(janela, text="Selecione o site:", font=("Arial", 14))
    label_selecione.pack(pady=5)

    site_selecionado = ctk.StringVar(value="CNPJ Já")
    dropdown_site = ctk.CTkComboBox(janela, values=["Consulta Guru", "CNPJ Já"], variable=site_selecionado)
    dropdown_site.pack(pady=5)

    # Função de depuração para verificar a seleção
    def verificar_selecao():
        print(f"Site selecionado: {site_selecionado.get()}")

    # Botão para iniciar a raspagem
    def iniciar_raspagem():
        site = site_selecionado.get()
        print(f"Valor do site selecionado no botão: {site}")
        if site == "Consulta Guru":
            consulta_cnpj_gratis(label_contador, janela)
        elif site == "CNPJ Já":
            consulta_cnpj_ja(label_contador, janela)

    # Botão de depuração para verificar o valor da seleção
    botao_verificar = ctk.CTkButton(janela, text="Verificar Seleção", command=verificar_selecao)
    botao_verificar.pack(pady=5)

    botao_iniciar = ctk.CTkButton(janela, text="Iniciar Raspagem", command=iniciar_raspagem)
    botao_iniciar.pack(pady=20)

    janela.mainloop()

def salvar_emails(resultados):
    """
    Salva os resultados no arquivo de e-mails, criando novos arquivos se o número de linhas exceder 2000.
    """
    nome_arquivo_base = "emails_vencedores"
    arquivos_existentes = glob.glob(f"{nome_arquivo_base}_*.xlsx")
    indice_arquivo = len(arquivos_existentes) + 1

    if not arquivos_existentes:
        nome_arquivo = f"{nome_arquivo_base}.xlsx"
    else:
        nome_arquivo = f"{nome_arquivo_base}_{indice_arquivo}.xlsx"

    df_resultados = pd.DataFrame(resultados)

    if os.path.exists(nome_arquivo):
        df_existente = pd.read_excel(nome_arquivo)
        if len(df_existente) >= 2000:
            indice_arquivo += 1
            nome_arquivo = f"{nome_arquivo_base}_{indice_arquivo}.xlsx"

    df_resultados.to_excel(nome_arquivo, index=False)
    print(f"E-mails salvos no arquivo '{nome_arquivo}'.")

def consulta_cnpj_gratis(label_contador, janela):
    arquivos = glob.glob("cnpjs_unicos*.xlsx")
    
    if not arquivos:
        messagebox.showwarning("Atenção", "O arquivo 'cnpjs_unicos.xlsx' não foi encontrado! Elimine os CNPJ duplicados!")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def limpar_cnpj(cnpj):
        return ''.join(filter(str.isdigit, str(cnpj)))

    total_emails = 0
    resultados = []
    emails_unicos = set()

    try:
        for arquivo in arquivos:
            if os.path.exists(arquivo):
                df = pd.read_excel(arquivo)

                for cnpj in df.iloc[:, 0]:  # Segunda coluna
                    cnpj_limpo = limpar_cnpj(cnpj)
                    if cnpj_limpo:
                        link = f"https://consulta.guru/consultar-cnpj-gratis/{cnpj_limpo}"
                        driver.get(link)
                        sleep(5)

                        try:
                            email_element = driver.find_element(By.XPATH, '//p[contains(text(), "@")]')
                            email = email_element.text.strip()

                            # Capturando a razão social
                            razao_social_element = driver.find_element(By.XPATH, '//div[@id="overview"]//h1')
                            razao_social = razao_social_element.text.strip()

                            if email not in emails_unicos:
                                emails_unicos.add(email)
                                resultados.append({"Razao Social": razao_social, "Email": email, "CNPJ": cnpj_limpo})
                                total_emails += 1
                                label_contador.configure(text=f"E-mails raspados: {total_emails}")
                                janela.update()
                        except Exception:
                            print(f"Erro ao encontrar o e-mail ou razão social para o CNPJ: {cnpj_limpo}")
                            pass  # Se não achar, continua

                        salvar_emails(resultados)
                        sleep(12)

        salvar_emails(resultados)
    finally:
        driver.quit()

def consulta_cnpj_ja(label_contador, janela):
    arquivos = glob.glob("cnpjs_unicos*.xlsx")

    if not arquivos:
        messagebox.showwarning("Atenção", "O arquivo 'cnpjs_unicos.xlsx' não foi encontrado! Elimine os CNPJ duplicados!")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def limpar_cnpj(cnpj):
        return ''.join(filter(str.isdigit, str(cnpj)))

    total_emails = 0
    resultados = []
    emails_unicos = set()

    try:
        for arquivo in arquivos:
            if os.path.exists(arquivo):
                df = pd.read_excel(arquivo)

                for cnpj in df.iloc[:, 0]:  # Segunda coluna
                    cnpj_limpo = limpar_cnpj(cnpj)
                    if cnpj_limpo:
                        link = f"https://cnpja.com/office/{cnpj_limpo}"
                        driver.get(link)
                        sleep(3)

                        try:
                            email_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '@')]"))
                            )
                            email = email_element.text.strip()

                            # Capturando a razão social
                            razao_social_element = driver.find_element(By.XPATH, '//div[@class="inline cursor-copy "]/span')
                            razao_social = razao_social_element.text.strip()

                            if email not in emails_unicos:
                                if email != 'contato@cnpja.com':
                                    emails_unicos.add(email)
                                    resultados.append({"Razao Social": razao_social, "Email": email, "CNPJ": cnpj_limpo})
                                    total_emails += 1
                                    label_contador.configure(text=f"E-mails raspados: {total_emails}")
                                    janela.update()
                        except Exception:
                            print(f"Erro ao encontrar o e-mail ou razão social para o CNPJ: {cnpj_limpo}")
                            pass  # Se não achar, continua

                        salvar_emails(resultados)
                        sleep(12)

        salvar_emails(resultados)
    finally:
        driver.quit()

criar_interface_raspagem_emails()
