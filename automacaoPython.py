import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

planilha = openpyxl.load_workbook('Excel_automacaoPython.xlsx')
mes_setembro = planilha['Planilha1']

for linha in mes_setembro.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    valor = linha[2].value

    valor_formatado = f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    mensagem = f'Olá {nome}, tudo bem? Segue o valor da sua conta com vencimento no fim do mês\n\nTotal é {valor_formatado}\n\nFavor pagar na chave pix (Nubank): clauthucio@gmail.com\n\nCaso queira a conta mais detalhada (naquele formato antigo), me avisa que te envio.'
    try:
        link_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_whatsapp)
        sleep(5)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone},{valor_formatado}')