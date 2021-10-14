import pyautogui
import pyperclip
import time
import pandas as pd
import numpy as np
import openpyxl

pyautogui.PAUSE = 0.5

"""
# Step 1: Open a new chrome tab on windows
pyautogui.press('win')
pyautogui.write('google')
pyautogui.press('enter')

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
"""

time.sleep(3)
print(pyautogui.position())
# explorador -> x=474, y=1045
# downloads -> x=81, y=257
# download file -> x=359, y=243
# CTRL + x => Este computador => Disclo local C: => Users => User => Pycharm Projects => Intensivão => CTRL + V

# Step 4: analysing dataset excel; 'r' -> reading special chars: '\'
table = pd.read_excel(r'C:\Users\user\Downloads\Vendas - Dez.xlsx')
faturamento = table['Valor Final'].sum()
quantidade_de_produtos = table['Quantidade'].sum()
print(f'Faturamento: R$ {faturamento}')
print(f'Quantidade de Produtos: R$ {quantidade_de_produtos}')

