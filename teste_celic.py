from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

# Caminho para os arquivos .crx das extensões
extensao_buster = r'C:\Users\Hillebrande\Downloads\buster-main'
extensao_recaptcha = r'CAAHALKGHNHBABKNIPMCONMBICPKCOPL_0_0_0_2.crx'
# Configuração das opções do Chrome
chrome_options = Options()

# Adicionando as extensões ao Chrome
chrome_options.add_extension(extensao_buster)
chrome_options.add_extension(extensao_recaptcha)

# Caminho do ChromeDriver
driver = webdriver.Chrome(options=chrome_options)

# Acesse o site
driver.get('https://www.compras.rs.gov.br/egov2/acessarAtaEletronica.ctlx?idOffer=328141&siteContext=Celic')

# Espera para garantir que as extensões sejam carregadas corretamente
time.sleep(10)

# Fechar o navegador
driver.quit()
