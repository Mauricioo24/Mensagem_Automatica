import pyautogui
import openpyxl
from urllib.parse import quote
import time

pyautogui.PAUSE = 1

planilha = openpyxl.load_workbook('Book 1.xlsx')
pagina = planilha["Sheet1"]

pyautogui.press('win')
pyautogui.write('edge')
pyautogui.press('enter')
time.sleep(5)

for linha in pagina.iter_rows(min_row=2):
    nome = linha[0].value
    telefone  =linha[1].value
    mensagem = f'Olá {nome}, eu estou fazendo uma automacao'
    #link_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    link_whatsapp = f"https://web.whatsapp.com/send?phone={telefone} "
    pyautogui.hotkey('ctrl','l')
    pyautogui.write(link_whatsapp)
    print(link_whatsapp)
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(10)
    #pyautogui.write(mensagem)
    #pyautogui.press('enter')
    

#https://web.whatsapp.com/send?phone5511932368610

