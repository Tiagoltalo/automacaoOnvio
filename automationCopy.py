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
wdw = WebDriverWait(navegador, 10) 

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

    esperarClick()

    #Seleciona o botão de entrar e clica nele
    botaoEntrarInicial = navegador.find_element(By.XPATH, "/html/body/div/div[1]/div/main/div/onvio-login/div/div[1]/fieldset/div/div/section/form[3]/div/button")
    botaoEntrarInicial.click()

    esperarElemento(navegador, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[1]/div/div/div/input", 15)

    # Preenche o campo email
    navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[1]/div/div/div/input").send_keys("tiago.italo.c.moura@gmail.com")

    botaoEntrarForm = navegador.find_element(By.XPATH, xpath)
    botaoEntrarForm.click()

    

    # Preenche o campo senha
    navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/form/div[1]/div/div[2]/div/input").send_keys("Italo@office")

    botaoEntrarForm.click()





### X --- Acessar Empresa e Setor Fiscal --- X ###
def acessarEmpresa(navegador):
    botaoDocumentos = navegador.find_element(By.XPATH, "/html/body/bm-optional-header/bm-staff-custom-header/bm-header/header/ul[2]/li[2]/li/a")
    botaoDocumentos.click()

    time.sleep(12)
    abas = navegador.window_handles
    navegador.switch_to.window(abas[1])

    time.sleep(5)

    # Selecionando a Empresa Exemplo 9999
    navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input").send_keys("9999")

    time.sleep(4)

    # Aguarda até o elemento desejado estar na tela
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]")))

    selecionarEmpresa = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]")
    selecionarEmpresa.click()

    retonarParaSetorFiscal(navegador)





### X --- Criar Pasta Ano Seguinte --- X ###
def criarPasta(navegador):
    WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/a")))

    botaoNovo = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/a")
    botaoNovo.click()

    # Adiciona uma nova pasta
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a")))

    novaPasta = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a")
    novaPasta.click()

    # Pegando a data atual e depois seleciona o ano da data
    dataAtual = datetime.date.today()
    anoSeguinte = dataAtual.year + 1

    # Sempre escrevará o ano atual como nome do arquivo
    navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[1]/form/div[1]/input").send_keys(f"{anoSeguinte}")

    # Salvamento
    salvar = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[2]/button[1]")
    salvar.click()




### X --- Copiar e Colar Subpastas do Ano Anterior --- X ###
def copiarEColarSubpastas(navegador):
    # Selecionar as subpastas do ano anterior com todos os meses e copiá-las no ano atual
    acessarAnoAnterior = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[2]/div[1]/div[2]/div[2]/div/span/dms-grid-text-cell/div/span[1]/a")
    acessarAnoAnterior.click()

    print(acessarAnoAnterior.text)

    time.sleep(2)

    # Seleciona e copia todas as subpastas da pasta 2025
    selecionarSubpastas = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[7]/div/div/div/div/i")
    selecionarSubpastas.click()

    time.sleep(2)

    clicarGerenciador = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a")
    clicarGerenciador.click()

    time.sleep(2)

    # Copiando subpastas do ano atual
    copiarSubPastas = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/ul/li[4]/a")
    copiarSubPastas.click()

    time.sleep(5)

    # Colando subpastas na pasta do ano seguinte
    caminho = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/div/span[1]")
    caminho.click()

    time.sleep(2.5)

    caminho = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/ul/li[2]/span[2]")
    caminho.click()

    colarSubPastas = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[3]/button[1]")
    colarSubPastas.click()




### X --- Retornar Para o Setor Fiscal --- X ###
def retonarParaSetorFiscal(navegador):
    # Retornando para a pasta do Fiscal e preparando para copiar a pasta do ano seguite em todas as empresas
    moveTo(140, 565)
    click()





### X --- Selecionar, Copiar e Colar a Pasta Ano Seguinte em Todas as Empresas --- X ###
    selecionarPastaAnoSeguinte = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[4]/div/div[1]/div/div/i")
    selecionarPastaAnoSeguinte.click()

    clicarGerenciador = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a")
    clicarGerenciador.click()

    copiarPastaAnoSeguinte = navegador.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-copy-button"]')
    copiarPastaAnoSeguinte.click()

    time.sleep(4)

    # Copiando a Pasta do ano seguinte em todas as empresas
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(Keys.CONTROL, 'a')
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(Keys.DELETE)

    time.sleep(5)

    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[1]/input').send_keys("2WV CONSTRUCOES E REFORMAS LTDA")

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="option-0"]/bento-combobox-row-template/span[2]')))

    time.sleep(10)

    selecionarEmpresa = navegador.find_element(By.XPATH, '//*[@id="option-0"]/bento-combobox-row-template/span[2]')
    selecionarEmpresa.click()

    time.sleep(3)

    fiscalTest = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div/div[1]/div/div/div/ul/li[5]/span[2]")
    fiscalTest.click()

    time.sleep(10)    




### X --- Fluxo Principal do Código --- X ###

acessarSite(navegador, xpathEntrarForm)