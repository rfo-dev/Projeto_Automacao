#################################################################################
# Programa de cadastro automatico utilizando dados de origem de uma planilha.   #
# Clica automaticamente nos campos.                                             #
#################################################################################

import pyautogui
import time
import pandas as pd
import openpyxl
import os
import pyperclip

x = pd.read_excel(r"C:\Users\rafael.oliveira\Documents\Projeto_Automacao\CadFunc.xlsx")

#for i in range(len(x)):
    #print(x['Número_NF'][i])
#print(pyautogui.position())

pyautogui.moveTo(135,214)
time.sleep(.9)
pyautogui.click()
time.sleep(.9)


for i in range(len(x)):
    pyperclip.copy(str(x['Nome'][i]))
    if (str(x['Nome'][i])) != 'nan':
        time.sleep(.9)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(.9)
        pyautogui.hotkey('enter')


