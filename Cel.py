import sys

def relevamiento():
    folder_path, folder, fecha = pdf()
    procesar(folder_path,folder, fecha)

def pdf():
    import subprocess
    import pyautogui
    import time
    import os
    import pandas as pd
    from datetime import datetime 
    import calendar
    import locale
    import easygui

    def validate_date(date_str):
        try:
            fecha = datetime.strptime(date_str, '%d-%m-%y')
            return fecha
        except ValueError:
            return None

    while True:
        fecha_str = input('Fecha dd-mm-aa: ')
        fecha = validate_date(fecha_str)
        if fecha:
            break
        else:
            print('Formato de fecha incorrecto.')

    def validate_visita(VISITA):
        try:
            visita = int(VISITA)
            if visita in [1, 2]:
                return visita
            else:
                return None
        except ValueError:
            return None
        
    while True:
        VISITA = input('Número de visita (1 o 2): ')
        visita = validate_visita(VISITA)
        if visita:
            break
        else:
            print('Número de visita incorrecto. Debe ser 1 o 2.')

    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

    mes = calendar.month_name[int(fecha.strftime('%m'))].upper()

    folder_path = r'C:\Users\usuario\Desktop\Proyectos\Celular'

    folder_to_create = os.path.join(folder_path, mes)

    if not os.path.exists(folder_to_create):      
        os.makedirs(folder_to_create)

    folder_path = folder_to_create

    folder='VISITA ' + str(visita)

    folder_to_create = os.path.join(folder_path, folder)

    if os.path.exists(folder_to_create):
        print('Error:  La Carpeta ya existe.')
    else:

        os.makedirs(folder_to_create)

        def descargar(url, archivo, compañia):
            if compañia != 'PERSONAL':
                subprocess.Popen(["C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", 
                    "--start-fullscreen", url])
                time.sleep(5)
                
                pyautogui.hotkey('ctrl', 'p')

                time.sleep(3)

                pyautogui.press('enter')

                time.sleep(1)
                dir= folder_to_create +  '\\' + compañia + ' ' + archivo + '.pdf'

                pyautogui.typewrite(dir)
                pyautogui.press('enter')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'w')
            else: 
                subprocess.Popen(["C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", 
                    "--start-fullscreen", url])
                time.sleep(5)
                
                
                
                pyautogui.hotkey('ctrl', 'f')
                time.sleep(0.5)
                pyautogui.typewrite('¿Necesitas ayuda?')
                pyautogui.press('enter')
                time.sleep(0.5)
                pyautogui.moveTo(200, 895) 
                pyautogui.click()
                pyautogui.hotkey('esc')
                time.sleep(1)

                text_location = pyautogui.locateOnScreen('LEGALES.png')
                pyautogui.moveTo(text_location[0], text_location[1])  
                pyautogui.click()

                time.sleep(0.5)

                pyautogui.hotkey('ctrl', 'p')

                time.sleep(3)

                pyautogui.press('enter')

                time.sleep(1)

                dir= folder_to_create +  '\\' + compañia + ' ' + archivo + '.pdf'

                pyautogui.typewrite(dir)
                pyautogui.press('enter')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'w')

        data=pd.read_excel(r'C:/Users/usuario/Desktop/Proyectos/Celular/RELEVAMIENTOcelular.xlsx')
        nombre = data.iloc[:, 1].tolist()
        compañia = data.iloc[:, 0].tolist()
        pagina = data.iloc[:, 2].tolist()

        if data.duplicated().any():
            print('Hay planes repetidos, por favor controlar excel.')
            data=data.drop_duplicates(inplace=True)

        easygui.msgbox('Controle que tenga abierta una pestaña del navegador Chrome', title='Relevamiento Precios Planes Celular', ok_button="LISTO")

        for i in range(len(pagina)):
            descargar(pagina[i],nombre[i],compañia[i])
        
        return folder_path, folder, fecha

def procesar(folder_path,folder, fecha):
    import pdfplumber
    import os
    import pandas as pd
    
    Carpeta= folder_path + '\\' + folder

    compañia=[]
    plan=[]
    precio=[]
    validez=[]
    sms=[]
    giga=[]
    llamada=[]

    if os.path.exists(Carpeta):
        files = os.listdir(Carpeta)
        
        Clarofiles= [Archivo for Archivo in files if 'CLARO' in Archivo] 
        for archivo in Clarofiles:
            all_text = []
            with pdfplumber.open(Carpeta  + '\\' + archivo) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    lines = text.split("\n")  
                    all_text.extend(lines)
        
            nombre=archivo[6:-4]
            compañia.append('AMX')
            plan.append(nombre)
            

            linea1 = [linea for linea in all_text if nombre in linea]
            if linea1 is not None:
                linea=linea1[0]   
                c=len(nombre)
                descripcion2 = False
                while descripcion2 == False and c< len(linea) - len(nombre):                        
                    if not  linea[c].isdigit():
                        c+=1
                    else:
                        descripcion2 = True
                        p=c
                        descripcion2 = False
                        while descripcion2 == False and p< len(linea):                        
                            if not linea[p] == ':':
                                p+=1
                            else:
                                descripcion2 = True
                                precio.append(linea[c:p])
                                linea2=linea[p:]
                                linea3 = linea2[linea2.find('capacidad') + len('capacidad'):linea2.find('gig')].strip()
                                giga.append(linea3)
            else: 
                precio.append('')
                giga.append('')
            
            linea2 = [linea for linea in all_text if 'Oferta válida en Argentina' in linea]
            if linea2 is not None:
                linea=linea2[0]   
                descripcion2 = False
                c=1
                while descripcion2 == False and c< len(linea) -1:                        
                    if not  linea[-c] .isdigit():
                        c+=1
                    else:
                        descripcion2 = True
                        validez.append(linea[:-c+1])
            else: 
                validez.append('')

            linea3 = [linea for linea in all_text if 'SMS Excedente' in linea]
            if linea3 is not None:
                linea= [linea.split('SMS Excedente')[1].strip() for linea in linea3][0]   
                descripcion = False
                c=0
                while descripcion == False and c< len(linea):                        
                    if not  linea[c]== '.':
                        c+=1
                    else:
                        descripcion = True
                        sms.append('Ilimitados. SMS Excedente' + linea[:c])
            else: 
                sms.append('')

            linea4 = [linea for linea in all_text if 'Establecimiento de llamada Excedente' in linea]

            if linea4 is not None:
                linea= [linea.split('Establecimiento de llamada Excedente')[1].strip() for linea in linea4][0]
                descripcion = False
                c=0
                while descripcion == False:
                    if c < len(linea):                       
                        if not  linea[c]== '.':
                            c+=1
                        else:
                            descripcion = True
                            k=linea.rfind('$')
                            llamada.append(linea[k+1:c])
                    else: 
                        descripcion = True
                        k=linea.rfind('$')
                        llamada.append(linea[k+1:])
            else: 
                llamada.append('')

        Movistarfiles= [Archivo for Archivo in files if 'MOVISTAR' in Archivo]    
        for archivo in Movistarfiles:
            all_text = []
            with pdfplumber.open(Carpeta  + '\\' + archivo) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    lines = text.split("\n")  
                    all_text.extend(lines)
            
            nombre=archivo[9:-4]
            compañia.append('TMA')
            plan.append(nombre)
            

            linea1 = next((linea for linea in all_text if linea.startswith('$')), None)

            if linea1:
                k=linea1.rfind('$')
                precio.append(linea1[k+1:])
            else: 
                precio.append('')

            linea2= [linea for linea in all_text if 'Desde el' in linea]
            if linea2:
                validez.append(linea2[0])
            else:
                validez.append('')

            linea3= [index for index, linea in enumerate(all_text) if 'Navegá' == linea]
            if linea3:
                k=linea3[0]
                linea3=all_text[k+1]
                l=linea3.find('GB')
                giga.append(linea3[:l])
            else:
                giga.append('')
            
            llamada.append('Libre')
            sms.append('Libre')

        Personalfiles= [Archivo for Archivo in files if 'PERSONAL' in Archivo]   
        for archivo in Personalfiles:
            all_text = []
            with pdfplumber.open(Carpeta  + '\\' + archivo) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    lines = text.split("\n")  
                    all_text.extend(lines)
            legales_index = None
            for i, line in enumerate(all_text):
                if "legales" in line.lower():
                    legales_index = i
                    break


            if legales_index is not None:
                legales_text = " ".join(all_text[legales_index + 1:])

                l=legales_text.find('PLAN DE ALCANCE')
                
                linea=legales_text[l:]

                all_text = all_text[:legales_index + 1]

                nombre=archivo[9:-4]
                compañia.append('Personal')
                plan.append(nombre)

                j1=linea.find('VIGENTE DESDE')
                validez.append(linea[j1:j1 + 47])

                l1=linea.find('PRECIO DE LISTA DEL SEGUNDO')
                linea2=linea[l1:]
                l2=linea2.find('$')
                l3=linea2.find('.')
                llamada.append(linea2[l2+1:l3])

                k1=linea.find('PRECIO DE LISTA DEL BLOQUE INICIAL')
                linea2=linea[k1:]
                k2=linea2.find('$')
                k3=linea2.find('.')
                sms.append(linea2[k2+1:k3])


            else:
                nombre=archivo[9:-4]
                compañia.append('Personal')
                plan.append(nombre)
                validez.append('')
                llamada.append('')
                sms.append('')

            if any('Precio de lista' in linea for linea in all_text):
                linea2= [linea for linea in all_text if 'Precio de lista' in linea][0]
                k=linea2.find('$')
                precio.append(linea2[k +1:])

            else: 
                precio_linea = None
                for linea in all_text:
                    if linea.startswith('$'):
                        precio_linea = linea[1:]
                        break
                print(precio_linea)
                
                t = None
                for index, caracter in enumerate(precio_linea):
                    if not caracter.isdigit() and caracter != '.':
                        t = index
                        break
                precio.append(precio_linea[0:t])          
            
            giga.append('')
    

    df= pd.DataFrame({'Empresa':compañia, 'Nombre del plan':plan, 'Precio de lista':precio, 'Precio del bloque de los primeros 30 segundos':llamada, 'SMS':sms,'Giga':giga, 'Comentarios': validez})

    fecha = fecha.strftime('%d-%m-%y')
    Archivo='Back Up Cel ' + fecha+'.xlsx'
    df.to_excel(Archivo, index=False)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "relevamiento":
            relevamiento()
