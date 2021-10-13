import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 0.5

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