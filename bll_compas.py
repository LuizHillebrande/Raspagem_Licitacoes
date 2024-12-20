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
import os
import re
import pdfplumber
import pandas as pd
import customtkinter as ctk
import json

PROGRESSO_JSON = "progresso.json"

def salvar_progresso(pdfs_baixados, cnpjs_extraidos):
    progresso = {
        "pdfs_baixados": pdfs_baixados,
        "cnpjs_extraidos": cnpjs_extraidos
    }
    with open(PROGRESSO_JSON, "w") as file:
        json.dump(progresso, file)

# Função para carregar progresso
def carregar_progresso():
    if os.path.exists(PROGRESSO_JSON):
        with open(PROGRESSO_JSON, "r") as file:
            return json.load(file)
    return {"pdfs_baixados": 0, "cnpjs_extraidos": 0}

def limpa_campo():
    for _ in range(10):
        pyautogui.press('right')
    for _ in range(10):
        pyautogui.press('backspace')

def extrair_bllcompras(data_inicio, data_fim, status_processo,label_contador_pdfs):
    current_dir = os.getcwd()
    if status_processo == "HOMOLOGADO":
        pasta_destino = os.path.join(current_dir, "vencedores_bll_compras_homologado")
    elif status_processo == "ADJUDICADO":
        pasta_destino = os.path.join(current_dir, "vencedores_bll_compras_adjudicado")
    else:
        print(f"Erro: Status de processo inválido: {status_processo}")
        return

    # Verifica se a pasta existe, se não, cria a pasta
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
        print(f"Pasta criada: {pasta_destino}")
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": pasta_destino,  # Define o diretório atual
        "download.prompt_for_download": False,      # Não perguntar antes de baixar
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True                # Ativa o download seguro
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    driver.get('https://bllcompras.com/Process/ProcessSearchPublic?param1=0#')

    # Aguarda até que o elemento de seleção de status esteja presente
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "fkStatus"))
    )
    select = Select(select_element)
    select.select_by_visible_text(status_processo)

    # Aguarda até que o botão de busca avançada seja clicável
    busca_avancada = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@title='BUSCA AVANÇADA']"))
    )
    busca_avancada.click()

    public_inicio = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@data-target='#DateStart']"))
    )
    public_inicio.click()
    public_inicio.clear()
    public_inicio.send_keys(data_inicio)

    # Preencher a data final
    public_final = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@data-target='#DateEnd']"))
    )
    public_final.click()
    public_final.clear()
    public_final.send_keys(data_fim)

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
    progresso = carregar_progresso()
    pdfs_baixados = progresso["pdfs_baixados"]

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
                # Verifica se a linha contém o texto "VENCEDORES DO PROCESSO"
                if "VENCEDORES DO PROCESSO" in linha.text.upper():
                    print("Linha encontrada com 'VENCEDORES DO PROCESSO'.")

                    # Procura todas as tags <a> com o atributo 'download' dentro da linha
                    links_com_download = linha.find_elements(By.CSS_SELECTOR, "a[download]")
                    print(f"Total de links com 'download' encontrados: {len(links_com_download)}")

                    # Itera sobre os links com atributo 'download'
                    for link in links_com_download:
                        # Verifica se o atributo 'download' ou texto contém "VENCEDORES DO PROCESSO"
                        if "VencedoresProcesso" in link.get_attribute("download"):
                            print("Link correto encontrado com 'VencedoresProcesso' no atributo 'download'.")

                            try:
                                botao_download = link.find_element(By.CSS_SELECTOR, "i.fas.fa-download")
                                print("Botão de download encontrado. Tentando clicar...")

                                # Realiza o scroll até o botão para garantir que ele esteja visível
                                driver.execute_script("arguments[0].scrollIntoView(true);", botao_download)
                                time.sleep(1)

                                # Realiza o clique com ActionChains
                                ActionChains(driver).move_to_element(botao_download).click().perform()
                                print("Download iniciado com sucesso.")
                                pdfs_baixados += 1
                                label_contador_pdfs.configure(text=f"PDFs    Baixados: {pdfs_baixados}")
                                salvar_progresso(pdfs_baixados, progresso["cnpjs_extraidos"])
                                root.update()
                                time.sleep(5)  
                                download_realizado = True
                                break  

                            except Exception as e:
                                print(f"Erro ao tentar clicar no ícone de download: {e}")
                        if download_realizado:
                            print("Download concluído para esta linha. Encerrando busca na linha.")
                            break  
                        else:
                            print("Link ignorado, não corresponde ao 'VencedoresProcesso'.")

            except Exception as e:
                print(f"Erro ao processar a linha: {e}")

        # Fecha a aba atual e retorna à aba original
        pyautogui.hotkey('ctrl', 'w')
        driver.switch_to.window(original_window)

    # Espera um pouco mais para garantir que o download seja iniciado
    time.sleep(5)  # Pausa de 5 segundos antes de fechar o driver
    
    driver.quit()  # Fechamento do driver após a execução

def extrair_cnpjs_pasta(pasta, nome_arquivo_saida, label_erro):
    """
    Extrai CNPJs de todos os PDFs em uma pasta específica e salva em um arquivo Excel.
    
    :param pasta: Caminho da pasta contendo os PDFs
    :param nome_arquivo_saida: Nome do arquivo Excel de saída onde os CNPJs serão salvos
    :param label_erro: Widget de erro para exibir mensagens na interface
    """
    # Padrão regex para capturar CNPJs
    padrao_cnpj = r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b'
    progresso = carregar_progresso()
    cnpjs_encontrados = set()
    total_cnpjs = progresso["cnpjs_extraidos"]

    # Itera sobre os arquivos na pasta
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".pdf"):
            caminho_arquivo = os.path.join(pasta, arquivo)
            print(f"Processando: {caminho_arquivo}")

            # Abre o PDF e extrai texto
            try:
                with pdfplumber.open(caminho_arquivo) as pdf:
                    for pagina in pdf.pages:
                        texto = pagina.extract_text()
                        if texto:
                            # Procura por CNPJs no texto
                            cnpjs = re.findall(padrao_cnpj, texto)
                            cnpjs_encontrados.update(cnpjs)
                            total_cnpjs += len(cnpjs)
                            label_contador_cnpjs.configure(text=f"CNPJs Extraídos: {total_cnpjs}")
                            salvar_progresso(progresso["pdfs_baixados"], total_cnpjs)
                            root.update()
                            if cnpjs:
                                print(f"CNPJs encontrados no arquivo {arquivo}: {cnpjs}")
                                cnpjs_encontrados.extend(cnpjs)
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")

    # Remove duplicatas de CNPJs
    cnpjs_unicos = list(set(cnpjs_encontrados))

    # Exibe todos os CNPJs encontrados
    if cnpjs_unicos:
        print("\nCNPJs encontrados:")
        for cnpj in cnpjs_unicos:
            print(cnpj)
        
        # Salva os CNPJs em um arquivo Excel
        caminho_saida = os.path.join(os.getcwd(), nome_arquivo_saida)  # Caminho completo para o arquivo
        df = pd.DataFrame({
            "Índice": range(1, len(cnpjs_unicos) + 1),
            "CNPJs": cnpjs_unicos  # Insira os CNPJs na segunda coluna
        })
        df.to_excel(caminho_saida, index=False)
        print(f"\nCNPJs salvos em: {caminho_saida}")
        
        # Atualiza a interface com uma mensagem de sucesso
        label_erro.configure(text=f"CNPJs extraídos e salvos em {nome_arquivo_saida}", text_color="green")
    else:
        label_erro.configure(text="Nenhum CNPJ encontrado nos arquivos PDFs.", text_color="red")


# Função para ser chamada quando o botão "Extrair CNPJ" for clicado
def extrair_cnpjs(label_erro, tipo):
    if tipo == 'homologados':
        pasta_vencedores = os.path.join(os.getcwd(), "vencedores_bll_compras_homologado")
        nome_arquivo_saida = "dados_vencedores_homologados_bll_compras.xlsx"
    elif tipo == 'adjudicados':
        pasta_vencedores = os.path.join(os.getcwd(), "vencedores_bll_compras_adjudicado")
        nome_arquivo_saida = "dados_vencedores_adjudicados_bll_compras.xlsx"
    else:
        label_erro.configure(text="Tipo inválido.", text_color="red")
        return

    if os.path.exists(pasta_vencedores):
        extrair_cnpjs_pasta(pasta_vencedores, nome_arquivo_saida, label_erro)
    else:
        label_erro.configure(text="Pasta de vencedores não encontrada!", text_color="red")

def criar_interface():
    def iniciar_busca():
        # Coleta os valores da interface
        data_inicio = entry_data_inicio.get()
        data_fim = entry_data_fim.get()
        status_processo = combo_status.get()
        extrair_bllcompras(data_inicio, data_fim, status_processo, label_contador_pdfs)

        
        # Validação simples 
        if not data_inicio or not data_fim or not status_processo:
            label_erro.configure(text="Preencha todos os campos!", text_color="red")
            return
        else:
            label_erro.configure(text="Executando busca...", text_color="green")
            root.update()

        # Executa o programa principal
        extrair_bllcompras(data_inicio, data_fim, status_processo,label_contador_pdfs)

       
    global root
    # Configuração da interface
    root = ctk.CTk()
    root.title("Busca BLL Compras")
    root.geometry("400x300")
    ctk.set_appearance_mode("Light")

    # Widgets
    label_titulo = ctk.CTkLabel(root, text="Buscar Processos BLL Compras", font=("Arial", 16, "bold"))
    label_titulo.pack(pady=10)

    # Campo para data inicial
    label_data_inicio = ctk.CTkLabel(root, text="Data Inicial (DD/MM/AAAA):")
    label_data_inicio.pack()
    entry_data_inicio = ctk.CTkEntry(root, placeholder_text="17/05/2024")
    entry_data_inicio.pack(pady=5)

    # Campo para data final
    label_data_fim = ctk.CTkLabel(root, text="Data Final (DD/MM/AAAA):")
    label_data_fim.pack()
    entry_data_fim = ctk.CTkEntry(root, placeholder_text="20/10/2024")
    entry_data_fim.pack(pady=5)

    # Dropdown para selecionar "Homologado" ou "Adjudicado"
    label_status = ctk.CTkLabel(root, text="Status do Processo:")
    label_status.pack()
    combo_status = ctk.CTkComboBox(root, values=["HOMOLOGADO", "ADJUDICADO"])
    combo_status.pack(pady=5)

    # Botão de executar
    btn_executar = ctk.CTkButton(root, text="Executar Busca", command=iniciar_busca)
    btn_executar.pack(pady=10)

    btn_extrair_cnpj_homologados = ctk.CTkButton(root, text="Extrair CNPJ Homologados", 
                                                 command=lambda: extrair_cnpjs(label_erro, 'homologados'))
    btn_extrair_cnpj_homologados.pack(pady=10)

    btn_extrair_cnpj_adjudicados = ctk.CTkButton(root, text="Extrair CNPJ Adjudicados", 
                                                 command=lambda: extrair_cnpjs(label_erro, 'adjudicados'))
    btn_extrair_cnpj_adjudicados.pack(pady=10)

    global label_contador_pdfs, label_contador_cnpjs
    label_contador_pdfs = ctk.CTkLabel(root, text="PDFs Baixados: 0")
    label_contador_pdfs.pack()

    label_contador_cnpjs = ctk.CTkLabel(root, text="CNPJs Extraídos: 0")
    label_contador_cnpjs.pack()

    # Label de erro ou sucesso
    label_erro = ctk.CTkLabel(root, text="")
    label_erro.pack()

    # Inicia a interface
    root.mainloop()

# Executa a interface
if __name__ == "__main__":
    criar_interface()