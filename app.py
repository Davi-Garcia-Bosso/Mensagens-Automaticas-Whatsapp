import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


webbrowser.open('https://web.whatsapp.com/')
sleep(30)
workbook = openpyxl.load_workbook('Clientes.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):

    Nome = linha[0].value
    Telefone = linha[1].value
    Vencimento = linha[2].value

    mensagem = f'Olá {Nome} seu boleto vence no dia {
        Vencimento.strftime('%d/%m/%Y')}. Favor pagar no link: https://openai.com/chatgpt/'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={
            Telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.PNG')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {Nome}')
        with open('erros.csv', 'a', newline='', encoding='utf=8') as arquivo:
            arquivo.write(f'{Nome}, {Telefone}')
