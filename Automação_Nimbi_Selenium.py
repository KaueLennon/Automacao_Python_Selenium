from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import Select
# import pandas as pd
import time

# senha = pd.read_excel('Arquivo_auxiliar.xlsx', sheet_name='PERIODO')

navegador = webdriver.Chrome()
wait = WebDriverWait(navegador, 20)

navegador.get("https://ss001.nimbi.com.br/Login/")

#Realizar logue
campo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wtLayout_Normal_2Col_wtLoginElements_wtLoginWB_V2_wt16_wtUsernameField"]')))
campo_usuario.send_keys('email@email.com.br')

#Clicar em entrar
botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="wtLayout_Normal_2Col_wtLoginElements_wtLoginWB_V2_wt16_wtFirstStep"]')))
botao_entrar.click()

time.sleep(5)

#Preencher senha
campo_senha = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wtLayout_Normal_2Col_wtLoginElements_wtLoginWB_V2_wt16_wtinpPassword"]')))
campo_senha.send_keys('senhausuario')

time.sleep(40)

try:
    #Clicar em analitycs
    analitycs = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wt2_Huge_WebbBaseTheme_wt18_block_wtAppSwitcher_CComponentsWebb_wt11_block_RichWidgets_wtDropDownMenuWB14_block_wtMenuItem_wtlnkAnalytics"]')))
    analitycs.click()
except Exception as e:
    print("Erro ao tentar logar")
    navegador.quit()

time.sleep(3)

#Clicar em Relatórios
relatorio = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wt13_Huge_WebbBaseTheme_wt5_block_wtAppSwitcher_CComponentsWebb_wt2_block_RichWidgets_wtDropDownMenuWB14_block_wtMenuSubItems_wtMenuAnalytics_wt23_wtlnkMenuIReports"]')))
relatorio.click()

time.sleep(2)

#Pesquisa no campo de pesquisa "Relatório de Pedidos"
pesquisa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wt18_Huge_WebbBaseTheme_wt5_block_wtMainContent_wtMainContent_wt14_Huge_WebbBaseTheme_wt35_block_wtContent_Huge_WebbBaseTheme_wt30_block_wtArea_Col_1_Huge_WebbBaseTheme_wt14_block_wtField_WebPatterns_wt37_block_wtSearch_wrapper_wtSearchInput"]')))
pesquisa.send_keys('Relatório de Pedidos')
time.sleep(2)

#Clicar para realizar a pesquisa do relatório digitado no campo de pesquisa
pesquisa_click = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wt18_Huge_WebbBaseTheme_wt5_block_wtMainContent_wtMainContent_wt14_Huge_WebbBaseTheme_wt35_block_wtContent_Huge_WebbBaseTheme_wt30_block_wtArea_Col_1_Huge_WebbBaseTheme_wt14_block_wtButton1_wt28"]')))
pesquisa_click.click()
time.sleep(2)

#Acessar o link do relatório de pedidos
relatorio_pedidos = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="wt18_Huge_WebbBaseTheme_wt5_block_wtMainContent_wtMainContent_wt14_Huge_WebbBaseTheme_wt35_block_wtContent_wtIndicatorTable_ctl03_wt21"]')))
relatorio_pedidos.click()

time.sleep(60) #Tempo para a tela de relatório carregar

#Inserir data inicio
try:
    # Aguardar até que o campo de data de início seja visível e interagível
    campo_data_inicio = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="react-select-2-input"]')))
    campo_data_inicio.send_keys("01/01/2024")
    print("Data inserida com sucesso")
except Exception as e:
    print(f"Erro ao tentar inserir a data: {e}")
finally:
    navegador.quit()
