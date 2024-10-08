from docxtpl import DocxTemplate
from datetime import datetime
from docx2pdf import convert
import zipfile
import calendar
import pyautogui as pag
import os
import subprocess
import keyboard

def MesesPortugues(month_name):
    english_months = list(calendar.month_name)
    portuguese_months = ['','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
    return portuguese_months[english_months.index(month_name)]

def CorrigirDocumentos(setor, modelo_padrao, nome_arquivos_padrao, modelo_laranjal, nome_arquivos_laranjal, nomes_pdf):
    cont = 0
    if setor == "N":
        for arquivo in modelo_padrao:
            doc = DocxTemplate(arquivo)
            correction = { 'certame_number' : certame,
                           'day' : day,
                           'month' : month,
                           'year' : year,
                           'company' : company,
                           'object' : object
                        }

            doc.render(correction)

            nome_arquivos_padrao[cont] = '[' + certame.replace('/', '-') + ']_' + nome_arquivos_padrao[cont]
            doc.save(nome_arquivos_padrao[cont])
            current_nome_arquivo_padrao_pdf = nome_arquivos_padrao[cont].replace('docx', 'pdf')
            convert(nome_arquivos_padrao[cont], current_nome_arquivo_padrao_pdf)

            nomes_pdf.append(current_nome_arquivo_padrao_pdf)

            cont += 1

    cont = 0
    if setor == "S":
        for arquivo in modelo_laranjal:
            doc = DocxTemplate(arquivo)
            correction = { 'certame_number' : certame,
                           'day' : day,
                           'month' : month,
                           'year' : year,
                           'company' : company,
                           'object' : object
                        }

            doc.render(correction)

            nome_arquivos_laranjal[cont] = '[' + certame.replace('/', '-') + ']_' + nome_arquivos_laranjal[cont]
            doc.save(nome_arquivos_laranjal[cont])
            current_nome_arquivo_laranjal_pdf = nome_arquivos_laranjal[cont].replace('docx', 'pdf')
            convert(nome_arquivos_laranjal[cont], current_nome_arquivo_laranjal_pdf)

            nomes_pdf.append(current_nome_arquivo_laranjal_pdf)

            cont += 1

def ApagarDocumentosWord(setor, nome_arquivos_padrao, nome_arquivos_laranjal):
    if setor == 'N':
        nome_arquivos = nome_arquivos_padrao
    if setor == 'S':
        nome_arquivos = nome_arquivos_laranjal

    for arquivo in nome_arquivos:
        try:
            caminho_arquivo = os.path.join(os.getcwd(), arquivo)
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)
                print(f"Deleted: {caminho_arquivo}")
            else:
                print(f"File not found: {caminho_arquivo}")
        except Exception as e:
            print(f"Error deleting {arquivo}: {e}")

def AssinarDocumentos(setor, nome_arquivos_padrao, nome_arquivos_laranjal):  
    if setor == 'N':
        nome_arquivos = nome_arquivos_padrao
    if setor == 'S':
        nome_arquivos = nome_arquivos_laranjal

    default_time = 1.7

    for arquivo in nome_arquivos:
        arquivo = arquivo.replace('.docx', '.pdf')
        os.startfile(arquivo) 
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
        pag.press('enter') # Escolhe o certificado e clica em CONTINUAR
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
        pag.press('enter') # Clica em ASSINAR
        pag.sleep(default_time)
        pag.sleep(default_time)

        # pag.write("30103010") #Escreve a senha
        # pag.sleep(default_time) 

        # pag.press('tab')
        # pag.sleep(0.2)
        # pag.press('tab')
        # pag.sleep(0.2)
        # pag.press('enter') #Clica em ASSINAR

        if keyboard.is_pressed('caps lock'):
            # Desativa o Caps Lock
            keyboard.press_and_release('caps lock')
        pag.press('caps lock')
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

def ApagarDocumentosPDF():

    try:
        subprocess.run(["taskkill", "/F", "/IM", "Acrobat.exe"], check=True)
    except subprocess.CalledProcessError:
        print("Erro ao tentar fechar o Adobe Acrobat ou ele não está aberto.")

    pag.sleep(3)

    caminho_arquivo = os.getcwd()
    
    arquivos = os.listdir(caminho_arquivo)
    
    lista_arquivos = {}
    
    for nome_arquivo in arquivos:
        if nome_arquivo.endswith('_ASS.pdf'):
            nome_base = nome_arquivo[:-8]  
            lista_arquivos[nome_base] = True  
        elif nome_arquivo.endswith('.pdf'):
            lista_arquivos[nome_arquivo] = False
    
    for nome_base, has_ass in lista_arquivos.items():
        if not has_ass and nome_base.endswith('.pdf'):
            arquivo_para_apagar = nome_base
            if os.path.isfile(arquivo_para_apagar):
                os.remove(arquivo_para_apagar)
                print(f"Arquivo removido: {arquivo_para_apagar}")

def JuntarAnexosLaranjal(certame):
    lista_arquivos= [
    '[' + certame.replace('/', '-') + ']_' + "1- DECLARAÇÃO DE INEXISTÊNCIA DE FATO IMPEDITIVO_ASS.pdf",
    '[' + certame.replace('/', '-') + ']_' + "2- Minuta - Declaração de impedimento art. 38 e 44_ASS.pdf"
    ]

    nome_zip = '[' + certame.replace('/', '-') + ']_' + "DECLARAÇÃO DE INEXISTÊNCIA DE FATO IMPEDITIVO_ASS.zip"

    with zipfile.ZipFile(nome_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for arquivo in lista_arquivos:
            if os.path.isfile(arquivo):
                zipf.write(arquivo, os.path.basename(arquivo))
                print(f"Adicionado ao zip: {arquivo}")
            else:
                print(f"Arquivo não encontrado: {arquivo}")
            
        for arquivo in lista_arquivos:
            os.remove(arquivo)


modelo_padrao = ["anx3.docx", 
                 "anx4.docx", 
                 "anx5.docx", 
                 "anx6.docx", 
                 "anx7.docx", 
                 "anx8.docx",
                 "anx9.docx", 
                 "anx10.docx"
                ]
nome_arquivos_padrao = [ "DECLARACAO DE ATENDIMENTO AO DISPOSTO NO ART 7.docx", 
                         "DECLARACAO DE ELABORACAO INDEPENDENTE.docx", 
                         "DECLARACAO DE ENQUADRAMENTO OU NAO NOS REQUISITOS PREVISTOS NA LEI COMPLEMENTAR.docx", 
                         "FORMULARIO DE SOLICITACAO DE CADASTRO DE CREDOR.docx", 
                         "DECLARACAO DE QUE NAO ADOTA TRABALHO FORCADO-ESCRAVO.docx", 
                         "DECLARACAO DE INEXISTENCIA DE FATO IMPEDITIVO.docx", 
                         "DECLARACAO QUE NAO SE ENCONTRA EM FALENCIA, SOLVENCIA OU CONCORDATA.docx",
                         "DECLARACAO DE AUTENTICIDADE.docx" 
                        ]

modelo_laranjal = ["anx3.docx", 
                    "anx4.docx", 
                    "anx5.docx", 
                    "anx6.docx", 
                    "anx7.docx", 
                    "anx9.docx", 
                    "anx10.docx",
                    "anx-spc-1.docx",
                    "anx-spc-2.docx" 
                    ]
nome_arquivos_laranjal = [  "DECLARACAO DE ATENDIMENTO AO DISPOSTO NO ART 7.docx", 
                            "DECLARACAO DE ELABORACAO INDEPENDENTE.docx", 
                            "DECLARACAO DE ENQUADRAMENTO OU NAO NOS REQUISITOS PREVISTOS NA LEI COMPLEMENTAR.docx", 
                            "FORMULARIO DE SOLICITACAO DE CADASTRO DE CREDOR.docx", 
                            "DECLARACAO DE QUE NAO ADOTA TRABALHO FORCADO-ESCRAVO.docx", 
                            "DECLARACAO QUE NAO SE ENCONTRA EM FALENCIA, SOLVENCIA OU CONCORDATA.docx",
                            "DECLARACAO DE AUTENTICIDADE.docx",
                            "1- DECLARAÇÃO DE INEXISTÊNCIA DE FATO IMPEDITIVO.docx",
                            "2- Minuta - Declaração de impedimento art. 38 e 44.docx" 
                        ]


certame = str(input("Número de certame '0000/0000': "))
company = "CEDAE - Companhia Estadual de Águas e Esgoto do Rio de Janeiro"
object = str(input("Objeto do certame: "))
setor = str(input("Laranjal (S/N): ")).upper()

current_date = datetime.now()
day = current_date.day
month = current_date.month
month = MesesPortugues(current_date.strftime('%B'))
year = current_date.year

nomes_pdf = []
CorrigirDocumentos(setor, modelo_padrao, nome_arquivos_padrao, modelo_laranjal, nome_arquivos_laranjal, nomes_pdf)
ApagarDocumentosWord(setor, nome_arquivos_padrao, nome_arquivos_laranjal)
AssinarDocumentos(setor, nome_arquivos_padrao, nome_arquivos_laranjal)
ApagarDocumentosPDF()
if setor == 'S':
    JuntarAnexosLaranjal(certame)