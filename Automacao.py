import pyautogui
import time
import pandas as pd
import openpyxl
import os
import pyperclip

x = pd.read_excel(r"C:\Users\rafael.oliveira\OneDrive\EstudosPython\CadFunc.xlsx")

#for i in range(len(x)):
    #print(x['Número_NF'][i])
#print(pyautogui.position())

pyautogui.moveTo(135,214)
time.sleep(.9)
pyautogui.click()
time.sleep(.9)


for i in range(len(x)):
    pyperclip.copy(str(x['Número_NF'][i]))
    if (str(x['Número_NF'][i])) != 'nan':
        time.sleep(.9)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(.9)
        pyautogui.hotkey('enter')


