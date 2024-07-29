import os
import sys
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from os.path import dirname

# Função de formatação de valores (supondo que esta função esteja definida)
def formata_valor(valor):
    if pd.notna(valor):
        return "{:,.2f}".format(valor)
    return None

# Função para procurar dados nas planilhas Excel
def procura(linha):
    df_icms = pd.read_excel('icms.xlsx', skiprows=3)
    valor_icms_g = pd.to_numeric(df_icms.iloc[linha, 6], errors='coerce')
    valor_icms_h = pd.to_numeric(df_icms.iloc[linha, 7], errors='coerce')

    df_feef = pd.read_excel('feef.xlsx', skiprows=3)

    try:
        valor_feef_g = pd.to_numeric(df_feef.iloc[linha, df_feef.columns.get_loc('ICMS NORMAL')], errors='coerce')
    except KeyError:
        valor_feef_g = pd.to_numeric(df_feef.iloc[linha, 6], errors='coerce')
    
    try:
        valor_feef_h = pd.to_numeric(df_feef.iloc[linha, 7], errors='coerce')    
    except KeyError:
        # Se a coluna 'ICMS NORMAL' não for encontrada, tenta pelo índice diretamente
        valor_feef_h = pd.to_numeric(df_feef.iloc[linha, 7], errors='coerce')
    
    # Convertendo CNPJ para string sem formatação, apenas dígitos
    cnpj = df_feef.iloc[linha, df_feef.columns.get_loc('CNPJ')]
    cnpj = f"{int(cnpj):014d}"  # Formata como uma string de dígitos contínuos sem formatação

    valor_icms_g = formata_valor(valor_icms_g)
    valor_icms_h = formata_valor(valor_icms_h)
    valor_feef_g = formata_valor(valor_feef_g)
    valor_feef_h = formata_valor(valor_feef_h)

    return valor_icms_g, valor_icms_h, valor_feef_g, valor_feef_h, cnpj

# Funções de espera e interação com elementos na página web
def espera(idElemento):
    wait = WebDriverWait(web, 12)
    wait.until(EC.presence_of_element_located((By.ID, idElemento)))
    return

def esperaXPATH(valor):
    wait = WebDriverWait(web, 12)
    expression1 = "//input[contains(@value, '"+valor+"')]"
    wait.until(EC.presence_of_element_located((By.XPATH, expression1)))
    return

def clicaElemento(idElemento):
    web.find_element(By.ID, idElemento).click()
    return

def clicaElementoXPATH(valor):
    expression2 = "//input[contains(@value, '"+valor+"')]"
    web.find_element(By.XPATH, expression2).click()
    return

def clicaElementoLink(idElementoLink):
    web.find_element(By.PARTIAL_LINK_TEXT, idElementoLink).click()
    return

def clicaElementoClass(idElemento):
    web.find_element(By.TAG_NAME, idElemento).click()
    return

def preencheElemento(idElemento, valor):
    web.find_element(By.ID, idElemento).send_keys(valor)
    return

def selecionaOpcao(idElemento, opcao):
    Select(web.find_element(By.ID, idElemento)).select_by_value(opcao)
    return

def limpaElemento(idElemento):
    web.find_element(By.ID, idElemento).clear()
    return

def pegaValor(idElemento):
    return web.find_element(By.ID, idElemento).get_attribute('value')

# Configurações de diretórios
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

config_path = os.path.join(application_path)
dirArquivo = dirname(config_path)
dirPrograma = dirArquivo
dirDownloadPdf = os.path.join(dirPrograma, "downloads")
diretorioPdfs = os.path.join(dirPrograma, "PDFs")

# Configurações do Chrome
prefs = {"download.default_directory": dirDownloadPdf}

chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)

chrome_driver_path = ChromeDriverManager().install()
chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
web = webdriver.Chrome(service=Service(chrome_driver_executable), options=chrome_options)
home = 'https://internet.sefaz.es.gov.br/agenciavirtual/area_publica/e-dua/icms.php'
web.get(home)

# Funções para manipulação de datas e execução do processo
mes_atual = datetime.now().month
# print(str(datetime.now().month))
# print(str(datetime.now().year))
mes_anterior = mes_atual - 1 if mes_atual > 1 else 12
mes_str = '{:02d}'.format(mes_anterior)
mes_str2 = '{:02d}'.format(mes_atual)
ano = datetime.now().year
periodo = str(mes_str) + str(ano)
periodo2 = str(mes_str2) + str(ano)

data = '18' + periodo2

def executa(linha):
    dados = procura(linha)

    # Primeira parte do processo
    selecionaOpcao("servico","1434") # 121-0
    preencheElemento('numIdentificacao', dados[4])
    preencheElemento('periodoReferencia', periodo)
    preencheElemento('dataVencimento', data)
    preencheElemento('valorImposto', dados[0])

    print('-----------------------------------')
    print(dados[0])
    print(dados[1])
    print(dados[2])
    print(dados[3])
    print(dados[4])


    print('-----------------------------------')
    input("Emitiu?")
    web.get(home)

    # Segunda parte do processo
    if dados[1] != 0 or dados[1] != 0.00:
        selecionaOpcao("servico","1440") # 128-7
        preencheElemento('numIdentificacao', dados[4])
        preencheElemento('periodoReferencia', periodo)
        preencheElemento('dataVencimento', data)
        preencheElemento('valorImposto', dados[1])
        input("Emitiu?")
        web.get(home)

    # Terceira parte do processo
    selecionaOpcao("servico","1446") # 472-3
    preencheElemento('numIdentificacao', dados[4])
    preencheElemento('periodoReferencia', periodo)
    preencheElemento('dataVencimento', data)
    preencheElemento('valorImposto', dados[3])
    input("Emitiu?")
    web.get(home)
    return

# Loop para execução do processo
x = 0
j = 0
z = 5
while j <= z:
    print(f"Linha atual: {j}")
    try:
        exec = executa(j)
        print("------------------Acabei a linha", j)
        j = j+1
    except Exception as error:
        print("Erro na linha ", j, "tentando de novo")
        print(error)
        j = j-1

print("End")