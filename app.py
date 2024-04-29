import openpyxl
import webbrowser
import pyautogui
from time import sleep

wb = openpyxl.load_workbook('planilha.xltx')





area_de_trabalho = wb['Planilha1']

for conteudo in area_de_trabalho.iter_rows(min_row=2):

    telefone = conteudo[0].value
    telefone = str(telefone)

    nome = conteudo[1].value
    nome = str(nome)



    mensagem = f'Olá, {nome}. :)'
    sleep(2)
    webbrowser.open(f'https://web.whatsapp.com/send?phone=55{telefone}&text={mensagem}')
    sleep(7)
    pyautogui.hotkey('enter')
    sleep(1.5)
    pyautogui.hotkey('ctrl' , 'w')
    sleep(1)

    