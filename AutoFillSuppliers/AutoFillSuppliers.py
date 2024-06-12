import pyautogui
import time
import openpyxl

default_time = 1

# Nome do arquivo Excel
excel_file = "suppliers_db.xlsx"

# Carregar o arquivo Excel
workbook = openpyxl.load_workbook(excel_file)

# Selecionar a planilha desejada
sheet = workbook["sheet4"]

pyautogui.hotkey('alt', 'tab')
pyautogui.sleep(default_time)

# Percorrer as linhas na planilha
for row in sheet.iter_rows(min_row=2, values_only=True):

    pyautogui.moveTo(2581,918)
    pyautogui.sleep(default_time)
    pyautogui.click()    
    pyautogui.sleep(default_time)

    # RAZAO SOCIAL
    pyautogui.write(row[0])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # APELIDO
    pyautogui.write(row[1])
    pyautogui.sleep(default_time)
    pyautogui.press('enter')
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # CNPJ
    pyautogui.write(row[2])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # IE
    pyautogui.write(row[3])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # LOCALIZACAO
    pyautogui.write(row[4])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # VENDEDOR
    pyautogui.write(row[5])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # TELEFONE
    pyautogui.write(row[6])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # EMAIL
    pyautogui.write(row[7])
    pyautogui.sleep(default_time)
    pyautogui.press('tab')
    pyautogui.sleep(default_time)

    # CATEGORIAS 
    pyautogui.write(row[8])
    pyautogui.sleep(default_time)
    pyautogui.press('enter')
    pyautogui.sleep(default_time)

    pyautogui.press('esc')
    pyautogui.press('esc')

    pyautogui.sleep(default_time)