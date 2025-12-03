from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pyautogui import moveTo, click
import datetime
import time

navegador = webdriver.Chrome()
navegador.maximize_window() 
xpathEntrarForm = "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[2]/button"

# Aguardar elemento ficar visível
def esperarElemento(navegador, xpath, time):
    WebDriverWait(navegador, time).until(EC.presence_of_element_located((By.XPATH, xpath)))

# Aguardar elemento tornar-se clicável
def esperarClick(navegador, xpath, time):
    WebDriverWait(navegador, time).until(EC.element_to_be_clickable((By.XPATH, xpath)))

# Selecionar e Clicar no Elemento
def selecionarEClicarNoElemento(navegador, xpath):
    elemento = navegador.find_element(By.XPATH, xpath)
    elemento.click()


### X --- Acessar Site e Efetuar o Login --- X ###
def acessarSite(navegador):
    navegador.get("https://onvio.com.br/login/#/")

    #Seleciona o botão de entrar e clica nele
    esperarElemento(navegador, "/html/body/div/div[1]/div/main/div/onvio-login/div/div[1]/fieldset/div/div/section/form[3]/div/button", 15)
    esperarClick(navegador, "/html/body/div/div[1]/div/main/div/onvio-login/div/div[1]/fieldset/div/div/section/form[3]/div/button", 15)
    selecionarEClicarNoElemento(navegador, "/html/body/div/div[1]/div/main/div/onvio-login/div/div[1]/fieldset/div/div/section/form[3]/div/button", 15)

    # Preenche o campo email
    esperarElemento(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[1]/div/div/div/input")
    navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[1]/div/div/div/input").send_keys("tiago.italo.c.moura@gmail.com")

    # Clica no botão entrar do formulário
    esperarElemento(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[2]/button", 10)
    esperarClick(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[2]/button", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[2]/button")

    # Preenche o campo senha
    esperarElemento(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/form/div[1]/div/div[2]/div/input", 10)
    navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/form/div[1]/div/div[2]/div/input").send_keys("Italo@office")

    selecionarEClicarNoElemento(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[2]/button")



### X --- Acessar Empresa e Setor Fiscal --- X ###
def acessarEmpresa(navegador):
    # Acessando a área de documentos
    esperarElemento(navegador, "/html/body/bm-optional-header/bm-staff-custom-header/bm-header/header/ul[2]/li[2]/li/a", 10)
    esperarClick(navegador, "/html/body/bm-optional-header/bm-staff-custom-header/bm-header/header/ul[2]/li[2]/li/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-optional-header/bm-staff-custom-header/bm-header/header/ul[2]/li[2]/li/a")

    time.sleep(10)
    abas = navegador.window_handles
    navegador.switch_to.window(abas[1])

    # Selecionando a Empresa Exemplo 9999
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input", 10)
    navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input").send_keys("9999")

    # Aguarda até o elemento desejado estar na tela
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]")

    # Retornando para a pasta do Fiscal e preparando para copiar a pasta do ano seguite em todas as empresas
    retonarParaSetorFiscal(navegador)





### X --- Criar Pasta Ano Seguinte --- X ###
def criarPasta(navegador):
    # Clica no botão novo para criar uma nova pasta
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/a", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/a")

    # Adiciona uma nova pasta
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a")

    # Pegando a data atual e depois seleciona o ano da data
    dataAtual = datetime.date.today()
    anoSeguinte = dataAtual.year + 1

    # Sempre escrevará o ano atual como nome do arquivo
    esperarElemento(navegador, "/html/body/div[1]/div/div/div/div[1]/form/div[1]/input", 10)
    navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[1]/form/div[1]/input").send_keys(f"{anoSeguinte}")

    # Salvando a nova pasta criada
    esperarElemento(navegador, "/html/body/div[1]/div/div/div/div[2]/button[1]", 10)
    esperarClick(navegador, "/html/body/div[1]/div/div/div/div[2]/button[1]", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/div[1]/div/div/div/div[2]/button[1]")



### X --- Copiar e Colar Subpastas do Ano Anterior --- X ###
def copiarEColarSubpastas(navegador):
    # Selecionar as subpastas do ano anterior com todos os meses e copiá-las no ano atual
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[2]/div[1]/div[2]/div[2]/div/span/dms-grid-text-cell/div/span[1]/a", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[2]/div[1]/div[2]/div[2]/div/span/dms-grid-text-cell/div/span[1]/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[2]/div[1]/div[2]/div[2]/div/span/dms-grid-text-cell/div/span[1]/a")
    
    # Seleciona e copia todas as subpastas da pasta 2025
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[7]/div/div/div/div/i", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[7]/div/div/div/div/i", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[7]/div/div/div/div/i")

    # Clicando no gerenciador de arquivos
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a")

    # Copiando subpastas do ano atual
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/ul/li[4]/a", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/ul/li[4]/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/ul/li[4]/a")

    # Colando subpastas na pasta do ano seguinte
    esperarElemento(navegador, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/div/span[1]", 10)
    esperarClick(navegador, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/div/span[1]", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/div/span[1]")

    # Selecionando o caminho para colar as subpastas
    esperarElemento(navegador, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/ul/li[2]/span[2]", 10)
    esperarClick(navegador, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/ul/li[2]/span[2]", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/ul/li[2]/span[2]")

    # Colando as subpastas na pasta do ano seguinte
    esperarElemento(navegador, "/html/body/div[1]/div/div/div/div[3]/button[1]", 10)
    esperarClick(navegador, "/html/body/div[1]/div/div/div/div[3]/button[1]", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/div[1]/div/div/div/div[3]/button[1]")



### X --- Retornar Para o Setor Fiscal --- X ###
def retonarParaSetorFiscal(navegador):
    # Retornando para a pasta do Fiscal e preparando para copiar a pasta do ano seguite em todas as empresas
    moveTo(140, 565)
    click()



### X --- Selecionar, Copiar e Colar a Pasta Ano Seguinte em Todas as Empresas --- X ###
def copiarPastaAnoSeguinteEmTodasAsEmpresas(navegador):
    # Selecionando a pasta do ano seguinte
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[4]/div/div[1]/div/div/i", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[4]/div/div[1]/div/div/i", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[4]/div/div[1]/div/div/i")

    # Clicando no gerenciador de arquivos
    esperarElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a", 10)
    esperarClick(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a", 10)
    selecionarEClicarNoElemento(navegador, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a")
    
    # Copiando a Pasta do ano seguinte
    esperarElemento(navegador, '//*[@id="dms-fe-legacy-components-client-documents-copy-button"]', 10)
    esperarClick(navegador, '//*[@id="dms-fe-legacy-components-client-documents-copy-button"]', 10)
    selecionarEClicarNoElemento(navegador, '//*[@id="dms-fe-legacy-components-client-documents-copy-button"]')



### X --- Funções que Percorrem Todas as Empresas --- X ###
def percorrerEmpresas(navegador):
    pass

def percorrerEmpresasCom2Duplicatas(navegador):
    pass

def percorrerEmpresasCom3Duplicatas(navegador):
    pass

def percorrerEmpresasCom4Duplicatas(navegador):
    pass

def percorrerEmpresasCom5Duplicatas(navegador):
    pass

def percorrerEmpresasCom7Duplicatas(navegador):
    pass

def percorrerEmpresasCom11Duplicatas(navegador):
    pass



### X --- Fluxo Principal do Código --- X ###

acessarSite(navegador, xpathEntrarForm)