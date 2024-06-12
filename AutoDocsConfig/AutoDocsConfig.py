from docxtpl import DocxTemplate
from datetime import datetime
from docx2pdf import convert
import calendar
import pyautogui as pag
import os

def PortugueseMonth(month_name):
    english_months = list(calendar.month_name)
    portuguese_months = ['','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

    return portuguese_months[english_months.index(month_name)]

files_folder1 = [ "anx3.docx", 
                 "anx4.docx", 
                 "anx5.docx", 
                 "anx6.docx", 
                 "anx7.docx", 
                 "anx8.docx",
                 "anx9.docx", 
                 "anx10.docx"
                 ]

files_folder2 = ["anx3.docx", 
                 "anx4.docx", 
                 "anx5.docx", 
                 "anx6.docx", 
                 "anx7.docx", 
                 "anx9.docx", 
                 "anx10.docx",
                 "anx-spc-1.docx",
                 "anx-spc-2.docx" 
                 ]


pdf_names = []

certame = str(input("Número de certame '0000/0000': "))
setor = str(input("Laranjal (S/N): ")).upper()
company = "CEDAE - Companhia Estadual de Águas e Esgoto do rio de Janeiro"
object = str(input("Objeto do certame: "))

current_date = datetime.now()
day = current_date.day
month = current_date.month
year = current_date.year

month = PortugueseMonth(current_date.strftime('%B'))

new_files_name1 = [ "DECLARACAO DE ATENDIMENTO AO DISPOSTO NO ART 7.docx", 
                    "DECLARACAO DE ELABORACAO INDEPENDENTE.docx", 
                    "DECLARACAO DE ENQUADRAMENTO OU NAO NOS REQUISITOS PREVISTOS NA LEI COMPLEMENTAR.docx", 
                    "FORMULARIO DE SOLICITACAO DE CADASTRO DE CREDOR.docx", 
                    "DECLARACAO DE QUE NAO ADOTA TRABALHO FORCADO-ESCRAVO.docx", 
                    "DECLARACAO DE INEXISTENCIA DE FATO IMPEDITIVO.docx", 
                    "DECLARACAO QUE NAO SE ENCONTRA EM FALENCIA, SOLVENCIA OU CONCORDATA.docx",
                    "DECLARACAO DE AUTENTICIDADE.docx" ]

new_files_name2 = [ "DECLARACAO DE ATENDIMENTO AO DISPOSTO NO ART 7.docx", 
                    "DECLARACAO DE ELABORACAO INDEPENDENTE.docx", 
                    "DECLARACAO DE ENQUADRAMENTO OU NAO NOS REQUISITOS PREVISTOS NA LEI COMPLEMENTAR.docx", 
                    "FORMULARIO DE SOLICITACAO DE CADASTRO DE CREDOR.docx", 
                    "DECLARACAO DE QUE NAO ADOTA TRABALHO FORCADO-ESCRAVO.docx", 
                    "DECLARACAO QUE NAO SE ENCONTRA EM FALENCIA, SOLVENCIA OU CONCORDATA.docx",
                    "DECLARACAO DE AUTENTICIDADE.docx",
                    "1- DECLARAÇÃO DE INEXISTÊNCIA DE FATO IMPEDITIVO.docx",
                    "2- Minuta - Declaração de impedimento art. 38 e 44.docx" ]

cont = 0
if setor == "N":
    for file in files_folder1:
        doc = DocxTemplate(file)
        correction = { 'certame_number' : certame,
                    'day' : day,
                    'month' : month,
                    'year' : year,
                    'company' : company,
                    'object' : object}

        doc.render(correction)

        new_file_name = '[' + certame.replace('/', '-') + ']_' + new_files_name1[cont]
        doc.save(new_file_name)
        new_file_name_pdf = new_file_name.replace('docx', 'pdf')
        convert(new_file_name, new_file_name_pdf)

        pdf_names.append(new_file_name_pdf)

        cont+=1

cont=0
if setor == "S":
    for file in files_folder2:
        doc = DocxTemplate(file)
        correction = { 'certame_number' : certame,
                    'day' : day,
                    'month' : month,
                    'year' : year,
                    'company' : company,
                    'object' : object}

        doc.render(correction)

        new_file_name = '[' + certame.replace('/', '-') + ']_' + new_files_name2[cont]
        doc.save(new_file_name)
        new_file_name_pdf = new_file_name.replace('docx', 'pdf')
        convert(new_file_name, new_file_name_pdf)

        pdf_names.append(new_file_name_pdf)

        cont+=1


# Auto_Signature 
default_time = 1.7

for file in pdf_names:
    os.startfile(file) 
    pag.sleep(default_time)
    #if file==new_files_name_pdf[0]:
    #    pag.hotkey('win', 'up') #Maximizar janela 
    pag.sleep(default_time)

    pag.moveTo(x=1101, y=1132) #Ponto para descida
    pag.sleep(default_time)
    pag.scroll(-20000)
    pag.sleep(default_time)

    pag.moveTo(x=290, y=307) #Usar um certtificado
    pag.sleep(default_time)
    pag.click()
    pag.sleep(default_time)

    pag.moveTo(x=181, y=337) #Assinar digitalmente
    pag.sleep(default_time)
    pag.click()
    pag.sleep(default_time)

    pag.moveTo(x=986, y=1272) #Ponto inicial assinatura
    pag.sleep(default_time+2)

    pag.dragTo(2448, 1587, 2, button='left') #Arrastar até o ponto final assinatura
    pag.sleep(default_time)

    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('enter') #Escolhe o certificado e clica em ASSINAR
    pag.sleep(default_time)

    pag.write("30103010") #Escreve a senha
    pag.sleep(default_time) 

    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('enter') #Clica em ASSINAR
    pag.sleep(5)

    pag.press('right')
    pag.sleep(default_time) 
    pag.write('_ASS') #Renomear arquivo
    pag.sleep(default_time) 

    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('tab')
    pag.sleep(0.2)
    pag.press('enter') #Salvar arquivo
    pag.sleep(default_time)