from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
from urllib.parse import quote
import pyautogui

planilha = openpyxl.load_workbook('Book 1.xlsx')
pagina = planilha["Sheet1"]

# Setup do Selenium WebDriver
driver = webdriver.Edge()  # Certifique-se de que o WebDriver do Edge está instalado

# Abrir WhatsApp Web
driver.get('https://web.whatsapp.com')
time.sleep(20)  # Tempo para escanear o código QR e carregar o WhatsApp Web

# Iterar sobre as linhas da planilha
for linha in pagina.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = f'Olá {nome}, eu estou fazendo uma automacao'
    link_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    
    # Navegar para o link
    driver.get(link_whatsapp)
    time.sleep(10)
    pyautogui.press('enter')
    time.sleep(5)  # Tempo para a página carregar completamente

driver.quit()
pyautogui.alert("Automação concluída")