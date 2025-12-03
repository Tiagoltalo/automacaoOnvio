from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains as AC
from pyautogui import press
import time

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

time.sleep(10)

### ----------------- ###

# Acessando seção de documentos
botaoDocumentos = navegador.find_element(By.XPATH, "/html/body/bm-optional-header/bm-staff-custom-header/bm-header/header/ul[2]/li[2]/li/a")
botaoDocumentos.click()

time.sleep(10)
abas = navegador.window_handles
navegador.switch_to.window(abas[1])

time.sleep(5)

listaEmpresas = []

pesquisa = navegador.find_element(By.XPATH, "/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input")
pesquisa.click()


# navegador.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(Keys.CONTROL, 'a')
# navegador.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[1]/input').send_keys(Keys.DELETE)

time.sleep(5)

for i in range(1, 709):
    time.sleep(.3)
    mapearEmpresas = navegador.find_element(By.XPATH, f"/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/div[2]/div/ul/li[{i}]/bento-combobox-row-template/span[2]")

    time.sleep(.3)
    listaEmpresas.append(mapearEmpresas.text)
    print(mapearEmpresas.text)

    press('down', presses=1)

print(listaEmpresas)

time.sleep(10)