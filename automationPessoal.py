from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains as AC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from pyautogui import moveTo, click
import empresas
import datetime
import time
import sys


listaEmpresas = empresas.listaEmpresasSemDuplicatas
listaEmpresasCom2Duplicatas = empresas.empresasCom2Duplicatas
listaEmpresasCom3Duplicatas = empresas.empresasCom3Duplicatas
listaEmpresasCom4Duplicatas = empresas.empresasCom4Duplicatas
listaEmpresasCom5Duplicatas = empresas.empresasCom5Duplicatas
listaEmpresasCom7Duplicatas = empresas.empresasCom7Duplicatas
listaEmpresasCom11Duplicatas = empresas.empresasCom11Duplicatas


navegador = webdriver.Chrome()

navegador.maximize_window()

actionChains = AC(navegador)

### Processo de Login ###
### ----------------- ###
# Acessa o site da Onvio
navegador.get("https://onvio.com.br/login/#/")

time.sleep(5)

#Seleciona o botão de entrar, após isso, clica nele
botaoEntrar = navegador.find_element(By.XPATH, "/html/body/div/div[1]/div/main/div/onvio-login/div/div[1]/fieldset/div/div/section/form[3]/div/button")
botaoEntrar.click()

time.sleep(5)

# Seleciona os campos de email e senha, após isso, os preenche
navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[1]/div/div/div/input").send_keys("tiago.italo.c.moura@gmail.com")

botaoEntrar = navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/div[1]/div/form/div[2]/button")
botaoEntrar.click()

time.sleep(5)

navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/form/div[1]/div/div[2]/div/input").send_keys("Italo@office")

botaoEntrar = navegador.find_element(By.XPATH, "/html/body/div/div/div/div[2]/main/section/div/div/div/form/div[2]/button")
botaoEntrar.click()

time.sleep(12)

### ----------------- ###

# Acessando seção de documentos
botaoDocumentos = navegador.find_element(By.XPATH, "/html/body/bm-optional-header/bm-staff-custom-header/bm-header/header/ul[2]/li[2]/li/a")
botaoDocumentos.click()

time.sleep(12)
abas = navegador.window_handles
navegador.switch_to.window(abas[1])

time.sleep(6)

# Selecionando a Empresa Exemplo 9999
navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input").send_keys("9999")

time.sleep(6)

# Aguarda até o elemento desejado estar na tela
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]")))

selecionarEmpresa = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]")
selecionarEmpresa.click()

# Acessa o setor Fiscal da empresa
WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-storage-grid/div/div/div[9]/div[2]/div[1]/div[7]/div[3]/div/div/dms-grid-text-cell/div/span[1]/a")))

setorPessoal = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-storage-grid/div/div/div[9]/div[2]/div[1]/div[7]/div[3]/div/div/dms-grid-text-cell/div/span[1]/a")
setorPessoal.click()

time.sleep(5)

# Clica no botão de adicionar um novo elemento
botaoNovo = navegador.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-new-menu-button"]')
botaoNovo.click()

# Adiciona uma nova pasta
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a")))

novaPasta = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[1]/ul/li[1]/a")
novaPasta.click()

time.sleep(4)

# Pegando a data atual e depois seleciona o ano da data
dataAtual = datetime.date.today()
anoSeguinte = dataAtual.year + 1

# Sempre escrevará o ano atual como nome do arquivo
navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[1]/form/div[1]/input").send_keys(f"{anoSeguinte}")

# Salvamento
salvar = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[2]/button[1]")
salvar.click()

time.sleep(5)

# Selecionar as subpastas do ano anterior com todos os meses e copiá-las no ano atual
acessarAnoAnterior = navegador.find_element(By.XPATH, '//*[@id="documentsTreeNavSplitter"]/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[2]/div[1]/div[2]/div[2]/div/span/dms-grid-text-cell/div/span[1]/a')
acessarAnoAnterior.click()

print(acessarAnoAnterior.text)

time.sleep(5)

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

time.sleep(12)

# Retornando para a pasta do Fiscal e preparando para copiar a pasta do ano seguite em todas as empresas
moveTo(145, 630)
click()

time.sleep(5)

selecionarPastaAnoSeguinte = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/div/dms-document-grid/div/div/div[14]/div[4]/div/div[1]/div/div/i")
selecionarPastaAnoSeguinte.click()


time.sleep(4)

### ----------------- ###

# Copiando a Pasta do ano seguinte em todas as empresas

for empresa in listaEmpresas:
    print(f"{empresa} - OK")

    time.sleep(2)

    clicarGerenciador = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[2]/div/section/div/documents-detail-pane/div/dms-document-grid-toolbar/dms-toolbar/div/ul/li[4]/a")
    clicarGerenciador.click()

    copiarPastaAnoSeguinte = navegador.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-copy-button"]')
    copiarPastaAnoSeguinte.click()

    time.sleep(2)

    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(Keys.CONTROL, 'a')
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(Keys.DELETE)
    
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(empresa)

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]')))

    empresaSelecionada = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]')

    WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[1]/bento-combobox-row-template/span[2]')))

    empresaSelecionada.click()

    time.sleep(2)

    pastasEmpresa = navegador.find_elements(By.CLASS_NAME, "item")

    for pasta in pastasEmpresa:
        if pasta.text == "Pessoal":
            fiscal = True
            local = pasta
            break
        else:
            continue

    if fiscal:
        actionChains.double_click(local).perform()

    colarPastaAnoSeguinte = navegador.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[3]/button[1]")
    colarPastaAnoSeguinte.click()

    time.sleep(1.5)

time.sleep(10)