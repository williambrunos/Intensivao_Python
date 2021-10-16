import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 0.75


def download_data_set():
    # Step 1: Open a new chrome tab on windows
    #pyautogui.press('win')
    #pyautogui.write('google')
    #pyautogui.press('enter')

    time.sleep(2)
    pyautogui.hotkey('ctrl', 't')
    # This link redirect us to the link with the dataset of the sales
    pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
    pyautogui.press('enter')
    time.sleep(5)

    # Step 2: Double clicking on the position of "exportar" folder, obtained by using pyautogui.position() function
    pyautogui.doubleClick(x=337, y=299)
    time.sleep(1.5)

    # Step 3: downloading the xlsx file
    pyautogui.click(x=418, y=360)
    pyautogui.rightClick()
    pyautogui.click(x=577, y=863)
    time.sleep(5)


def copy_paste_data_set():
    # explorador
    pyautogui.click(x=474, y=1045)
    # downloads
    pyautogui.click(x=81, y=257)
    # download file
    pyautogui.click(x=359, y=243)
    # copy
    pyautogui.hotkey('ctrl', 'x')
    # Este computador
    pyautogui.click(x=100, y=463)
    # Disco local C:
    pyautogui.doubleClick(x=287, y=384)
    # Usuários
    pyautogui.doubleClick(x=339, y=392)
    # User
    pyautogui.doubleClick(x=294, y=266)
    # Pycharm Projects
    pyautogui.doubleClick(x=304, y=675)
    # Intensivão Dia - 1
    pyautogui.doubleClick(x=279, y=227)
    # paste
    pyautogui.hotkey('ctrl', 'v')


# Step 4: analysing dataset excel; 'r' -> reading special chars: '\'
def analyse_data_set():
    table = pd.read_excel(r'.\Vendas - Dez.xlsx')
    faturamento = table['Valor Final'].sum()
    quantidade_de_produtos = table['Quantidade'].sum()

    return faturamento, quantidade_de_produtos


# Step 5: write the e-mail to some random one
def open_email():
    pyautogui.click(x=474, y=1045)
    pyautogui.hotkey('ctrl', 't')
    pyautogui.write(r'https://mail.google.com/mail/u/0/#inbox')
    pyautogui.press('enter')
    time.sleep(5)


def write_and_send_email():
    pyautogui.click(x=89, y=196)
    time.sleep(1)
    pyautogui.write(r'williambrunosalesdepaulalima@gmai.com')
    pyautogui.press('tab')
    pyautogui.write(r'williambruno172@gmail.com')
    time.sleep(1)
    pyautogui.press('tab') # new e-mail
    time.sleep(1)
    pyautogui.press('tab') # Assunto
    time.sleep(1)

    pyperclip.copy('Relatório de vendas')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    pyautogui.press('tab')
    faturamento, quantidade_de_produtos_vendidos = analyse_data_set()
    texto = f'Prezados \n\nSegue o resumo de vendas referente ao mês de dezembro.\n\nFaturamento total do mês: ' \
            f'R$ {faturamento:,.2f} ;\nQuantidade total de produtos vendidos: {quantidade_de_produtos_vendidos:,} produtos ;' \
            f'\n\nAtenciosamente.\nWilliam, analista de dados'
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'enter')


if __name__ == '__main__':
    download_data_set()
    copy_paste_data_set()
    analyse_data_set()
    open_email()
    write_and_send_email()