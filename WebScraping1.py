# Modulo para controlar o navegador
from selenium import webdriver

# Localizador de elementos
from selenium.webdriver.common.by import By

# Serviço para configurar o caminho do executável chromedriver
from selenium.webdriver.chrome.service import Service

# Classe que permite executar ações avançadas, como por exemplo o mover o mouse, o click e arrasta e etc..
from selenium.webdriver.common.action_chains import ActionChains

# Classe que espera de forma explicita até que uma condição seja satisfeita
# Condição (ex: Que um elemento apareça)
from selenium.webdriver.support.ui import WebDriverWait

# Condições esperadas usadas com WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

# Trabalhar com DataFrame
import pandas as pd

# Uso de funções relacionadas ao tempo
import time 

from selenium.common.exceptions import TimeoutException

# Definir o caminho do chromeDriver 
chrome_driver_path = "C:\Program Files\chromedriver-win64\chromedriver-win64\chromedriver.exe" # Onde esta armazenado o caminho do driver 

# configuração ao WebDriver
service = Service(chrome_driver_path) #navegador controlado pelo Selenium
options = webdriver.ChromeOptions() # configura opções do navegador
options.add_argument("--disable-gpu") # evita possíveis erros gráficos
options.add_argument("--window-size=1920,1080") # define uma resolução fixa

# incialização ao WebDriver
driver = webdriver.Chrome(service=service, options=options) # inicialização do navegador

# URl inicial
url_base = "https://masander.github.io/AlimenticiaLTDA-financeiro/"
driver.get(url_base)
time.sleep(10) # aguarda 5 segundos para garantir que a pág carregue

#criar um dicionário vazio para armazenar as marcas para armazenar as marcas e preços das cadeiras
dic_produtos = {"Despesa": [], "Data": [], "Tipo": [], "Setor": [], "Valor": [], "Fornecedor": []}

financeiro = driver.find_elements(By.XPATH, "//table/tbody//tr")

for financeiro in financeiro:
        try:
            Despesa = financeiro.find_element(By.CLASS_NAME, "td_id_despesa").text.strip()
            Data = financeiro.find_element(By.CLASS_NAME, "td_data").text.strip()
            Tipo = financeiro.find_element(By.CLASS_NAME, "td_tipo").text.strip()
            Setor = financeiro.find_element(By.CLASS_NAME, "td_setor").text.strip()
            Valor = financeiro.find_element(By.CLASS_NAME, "td_valor").text.strip()
            Fornecedor = financeiro.find_element(By.CLASS_NAME, "td_fornecedor").text.strip()


            print(f"{Despesa} - {Data} - {Tipo} - {Setor} - {Valor} - {Fornecedor}")

            dic_produtos["Despesa"].append(Despesa)
            dic_produtos["Data"].append(Data)
            dic_produtos["Tipo"].append(Tipo)
            dic_produtos["Setor"].append(Setor)
            dic_produtos["Valor"].append(Valor)
            dic_produtos["Fornecedor"].append(Fornecedor)

        except Exception as e:
            print("Erro ao coletar dados:", e)
            
# Fechar o navegador
driver.quit()

# DataFrame
df = pd.DataFrame(dic_produtos)

# Salvar os dados em excel
df.to_excel("Despesas.xlsx", index= False)
        
print(f"Arquivo 'Despesas' salvo com sucesso! {len(df)} produtos capturados") 
