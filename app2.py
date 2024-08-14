# encoding: utf-8
import pyautogui
import openpyxl
from urllib.parse import quote
from time import sleep


pyautogui.press('winleft')
sleep(2)
pyautogui.write('whatsapp')
pyautogui.press('enter')
sleep(5)

planilha = openpyxl.load_workbook('Excel_app2.xlsx')
mes_setembro = planilha['Planilha1']

for linha in mes_setembro.iter_rows(min_row=2):
    nome = linha[0].value
    valor = linha[1].value
    
    valor_formatado = f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

    mensagem = f'Olá {nome}, td bem? Segue o valor do seu boleto com vencimento no fim do mês\n\nTotal é R${valor_formatado}\n\nFavor pagar na chave pix (Nubank): clauthucio@gmail.com'

    pyautogui.hotkey('ctrl', 'f')
    pyautogui.write(nome)
    sleep(2)
    pyautogui.press('tab')
    sleep(2)
    pyautogui.press('enter')
    sleep(2)
    pyautogui.typewrite(f'Olá {nome}, td bem? Segue o valor do seu boleto com vencimento no fim do mês')
    pyautogui.hotkey('shift', 'enter')
    pyautogui.typewrite(f'Total é {valor_formatado}')
    pyautogui.hotkey('shift', 'enter')
    pyautogui.typewrite('Favor pagar na chave pix (Nubank): clauthucio@gmail.com')
