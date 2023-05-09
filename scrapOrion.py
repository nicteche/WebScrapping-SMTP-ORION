import os
import time
import chromedriver_autoinstaller
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import datetime as dt
import shutil
from pathlib import Path
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from PIL import Image
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

fecha_actual = dt.date.today()
#Path (Downloads, Root)
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
downloads_path = str(Path.home() / "Downloads")
#Options to chromeDriver 
chromedriver_autoinstaller.install()
options_cs = webdriver.ChromeOptions()
options_cs.add_argument("--disable-gpu")
options_cs.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ""Chrome/95.0.4638.54 Safari/537.36")
options_cs.add_experimental_option("excludeSwitches", ["enable-automation"])
options_cs.add_argument("--disable-blink-features=AutomationControlled")
options_cs.add_experimental_option("useAutomationExtension", False)
options_cs.add_experimental_option("excludeSwitches", ["enable-automation"])
options_cs.add_argument('--disable-extensions')
options_cs.add_argument('--ignore-certificate-errors')
options_cs.add_argument('--print-to-pdf')  # Habilitar la impresión a PDF
options_cs.add_argument('--no-sandbox')
options_cs.add_argument('--disable-dev-shm-usage')
options_cs.add_argument('--enable-print-browser')
options_cs.add_argument('--kiosk-printing')

browserOrion = webdriver.Chrome('chromedriver.exe',options=options_cs)
browserOrion.maximize_window()
#Url To conecte the page
url = 'https://sbp0100tcm14.pacifico.bpgf/orion/netperfmon/events.aspx'
browserOrion.get(url)
#Make a main window
main_window = browserOrion.current_window_handle
time.sleep(15)

#Login in to he page
try:
    username = browserOrion.find_element(By.CLASS_NAME, "password-policy-username")
    username.clear()
    username.send_keys("nechever@pacifico")
    time.sleep(2)
    password = browserOrion.find_element(By.CLASS_NAME, "password-input")
    password.clear()
    password.send_keys("1710180033Te")
    time.sleep(2)
    btnlogin = browserOrion.find_element(By.ID,"ctl00_BodyContent_LoginButton")
    btnlogin.click()
    time.sleep(10)
except: 
    print("Ya esta logeado")

#This is the configuration to select the schedule that we will take
btnTime= browserOrion.find_element(By.ID,"ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList")
btnTime.click()
btnCustom = browserOrion.find_element(By.XPATH , '//*[@id="ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList"]/option[13]')
btnCustom.click()
time.sleep(2)
txtDateIn = browserOrion.find_element(By.ID,"ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtDatePicker")
time.sleep(2)
#The date can change depending on the configuration of the variables of the computer where it is executed.
fecha_actual_str = fecha_actual.strftime('%d/%m/%Y')
horaF = "8:00"
horaI = "20:00"
fecha_anterior = fecha_actual - dt.timedelta(days=1)
fecha_anterior_str =  fecha_anterior.strftime('%d/%m/%Y')
txtDateIn.send_keys(fecha_anterior_str)

txtDateInHora = browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtTimePicker")
txtDateInHora.send_keys(horaI)
time.sleep(1)
txtDateON = browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtDatePicker")
txtDateON.send_keys(fecha_actual_str)
time.sleep(1)
txtDateOnHora = browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtTimePicker")
txtDateOnHora.send_keys(horaF)
time.sleep(1)
btnRefresh = browserOrion.find_element(By.ID,"ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
btnRefresh.click()

#Through Beautifulsoup we take the HTML information of the Page and look for the table of events, which is made up of Date, Status and description.
soup = BeautifulSoup(browserOrion.page_source, 'html.parser')
elementosTime = soup.find_all("td")
fechaL = []
notificacionL = []

#Here we will go through the obtained records, where we extract only the date and the description
try:
    for y in elementosTime:
        dato = str(y.text)
        dato = dato.strip().split("\n")
        if len(dato[0]) > 1:
            if "/" in dato[0] and len(dato[0]) < 20:
                fechaL.append(dato[0])
            if len(dato[0]) > 28:
                notificacionL.append(dato[0])
            
except:
    print("No se encontro")

#We create a dataframe with the information obtained and we filter the records that contain the words that we are looking for.
fedic={"Fecha" : fechaL,"Notificacion" : notificacionL}

df = pd.DataFrame(fedic)
#We perform a filter where only the records containing 'has stopped|is responding again|reboot' are presented
filtro = df['Notificacion'].str.contains('has stopped|is responding again|reboot', case=False)
df= df[filtro]

diposL= []
tempo = []
#We transform the date strings, by datetime
df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M')

# We proceed to obtain the names of all the devices within the table that were affected
dispositivos = df['Notificacion'].str.extract(r'^(.+?) (?:has stopped)').dropna()
dispositivos = dispositivos[0].str.split("has stopped").str[0].unique()

#Using the obtained list of affected devices, we proceed to filter the table for each one of them.
for dispositivo in dispositivos:
    df_filtered1 = df[df['Notificacion'].str.contains(f"{dispositivo}")]
    enlace = dispositivo.split(" ")[0]
    
    tiene_reboot = df_filtered1['Notificacion'].str.contains("reboot").any()
    tiene_rebootL = df_filtered1['Notificacion'].str.contains("reboot").tolist()
    revisarI = df_filtered1['Notificacion'].str.contains("is responding").tolist()
    revisarF = df_filtered1['Notificacion'].str.contains("has stopped responding").tolist()
    
    
    try:
        for x in revisarF:
            if x:
                index_STOP = revisarF.index(x)
                notificacionF = str(df_filtered1.iloc[index_STOP]['Notificacion'])
                if notificacionF.startswith(enlace):
                    print(notificacionF)
                else:
                    df_filtered1 = df_filtered1.drop(index=df_filtered1.index[index_STOP])
    except:
        print("No hay registros")
    
    
    try:
        for x in revisarI:
            if x:
                index_INI = revisarI.index(x)
                notificacionS = str(df_filtered1.iloc[index_INI]['Notificacion'])
                if notificacionS.startswith(enlace):
                    print(notificacionS)
                else:
                    df_filtered1 = df_filtered1.drop(index=df_filtered1.index[index_INI])
    except:
        print("No hay registros")


    # It is verified that the device is not a switch
    if not str(enlace).startswith("sw"):
        #In case they have been rebooted, the records up to that date are not taken into account
        if tiene_reboot:
            index_reboot = tiene_rebootL.index(True)
            df_filtered1 = df_filtered1.drop(index=df_filtered1.index[index_reboot:])
        #The date will sorted 
        df_filtered1 = df_filtered1.sort_values('Fecha', ascending=False)
        ff = len(df_filtered1[df_filtered1['Notificacion'].str.contains("is responding again.")])
        fi = len(df_filtered1[df_filtered1['Notificacion'].str.contains("has stopped responding")])
        
        i = 1
        registros = len(df_filtered1)
        while registros >= i :
            #We proceed to validate the number of times affected and the times it was recovered.
            ff = len(df_filtered1[df_filtered1['Notificacion'].str.contains("is responding again.")])
            fi = len(df_filtered1[df_filtered1['Notificacion'].str.contains("has stopped responding")])
            #validation if you do not have notifications
            if  ff == 0 and fi == 0:
                registros=0
            #Validation if only the recovered device notification appears    
            if ff == 1 and fi == 0:
                registros= 0
            #Validation if it is only affected and the information is added to the list  
            if ff == 0 and fi == 1:
                diposL.append(dispositivo)
                tempo.append("CAIDO")
                registros = 0
                
            #Validation if it is infected 2 times and the oldest notification is deleted for further analysis
            if ff == 1 and fi == 2:
                index = df_filtered1[df_filtered1['Notificacion'].str.contains("has stopped responding")].index[-1]
                fecha = df_filtered1.loc[index, 'Fecha']
                #Deleted the extra description
                df_filtered1 = df_filtered1.drop(df_filtered1[df_filtered1['Fecha'] == fecha].index)
                
            if ff !=0 and fi != 0:
                try:
                    #Get the information when we recovery device
                    fecha_inicioF = df_filtered1[df_filtered1['Notificacion'].str.contains("is responding again.")]['Fecha'].iloc[-1]
                    fecha_inicio = pd.to_datetime(df_filtered1[df_filtered1['Notificacion'].str.contains("is responding again.")]['Fecha'].iloc[-1])
                    # Get the information when the device was affected
                    fecha_finF = df_filtered1[df_filtered1['Notificacion'].str.contains("has stopped responding")]['Fecha'].iloc[-1]
                    fecha_fin = pd.to_datetime(df_filtered1[df_filtered1['Notificacion'].str.contains("has stopped responding")]['Fecha'].iloc[-1])
                    
                    # We calculate the time between recovery and affect the device.
                    diferencia_tiempo = (fecha_inicio - fecha_fin).total_seconds() / 60
                    diferencia_tiempo = diferencia_tiempo
                    print(diferencia_tiempo)
                    
                except:
                    print("No hay registro")
                    
                #When the time exceeds 5 minutes, it enters the list of affected devices, otherwise the records are deleted
                if diferencia_tiempo > 5:
                    
                    print("SI hay registro")
                    diposL.append(dispositivo)
                    tempo.append(diferencia_tiempo)
                    registros = 0
                    
                    #tiempoF.append(diferencia_tiempo)
                else:
                    print("No hay registro")
                    try:
                        # We get the index of the first "has stopped responding"
                        index = df_filtered1[df_filtered1['Notificacion'].str.contains("has stopped responding")].index[-1]
                        fechaS = df_filtered1.loc[index, 'Fecha']
                        #Deleted the record
                        df_filtered1 = df_filtered1.drop(df_filtered1[df_filtered1['Fecha'] == fechaS].index)
                        #We found the index of the first "has stopped responding"
                        index = df_filtered1[df_filtered1['Notificacion'].str.contains("is responding again.")].index[-1]

                        fechaF = df_filtered1.loc[index, 'Fecha']
                        #Deleted the record
                        df_filtered1 = df_filtered1.drop(df_filtered1[df_filtered1['Fecha'] == fechaF].index)

                    except:
                        print("No hay indices") 
                        registros = 0    

#Dataframe with the affeted device and how long take to recovery
dictFINAL = {"Enlace": diposL, "Tiempo Minutos": tempo}
final = pd.DataFrame(dictFINAL)
fecha_actual = dt.date.today() 
mes = fecha_actual.month
dia = fecha_actual.day
archivo = f"Reporte{mes}-{dia}prueba1.xlsx"
final.to_excel(archivo, index=False)

# With the dataframe we search the page for the events of the affected device within the schedule
nodo = browserOrion.find_element(By.CSS_SELECTOR,"#ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects")
nodo.click()

#We obtain the information from the page where we look for each affected device
soup = BeautifulSoup(browserOrion.page_source, 'html.parser')
nodos = soup.find_all("option")
for x in nodos:
    print(str(x.text))
    Validation of the device found, obtaining the ID of each one
    if str(x.text) in final["Enlace"].values:
        nombre= str(x.text)
        try:
            partes = nombre.split(" ")
            indice = partes.index("P:")
            resultado = " ".join(partes[:indice+1])
        except:
            print("No se pudo")
            resultado = nombre
            
            
        enlace = x.attrs
        print("Encontrado  ", enlace, "  ", enlace["value"])
        time.sleep(5)
        select_element = Select(browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects"))
        select_element.select_by_value(str(enlace['value']))
        select = browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects")
        select.click()
        time.sleep(5)
        #We click on each one and proceed to obtain a screenshot for each table shown on the page
        btnRefresh = browserOrion.find_element(By.ID,"ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
        btnRefresh.click()
        time.sleep(5)
        table_element = browserOrion.find_element(By.CLASS_NAME,"sw-pg-events")
        screenshot = browserOrion.get_screenshot_as_png()

        # Make a screenshot 
        location = table_element.location
        size = table_element.size
        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
        table_screenshot = Image.open(BytesIO(screenshot)).crop((left, top, right, bottom))

        # Guarda la imagen recortada en un archivo
        table_screenshot.save(f'{resultado}.png')
        time.sleep(5)
        
# We will send the Dataframe and each picture about the affected device
#Write the smtp server and port
smtp_server = "smtp.office365.com"
smtp_port = 587
#write email credencials
smtp_user = "nechever@domain.com"
smtp_password = "-"
#write the email to whom you will send the information
correos = 'use1@domain.com,user2@domain.com'
correos = correos.split(", ")
#Make the loop to sent each email to each person
for correo in correos:
    message = MIMEMultipart()
    #Who sent this
    message['From'] = "nechever@domain.com"
    #who recive the email
    message['To'] = str(correo)
    #the Subjest of the email
    message['Subject'] = f'INCIDENCIA DE ENLACE {fecha_actual}'
    #The message of the email
    mensaje = f"Por favor revisar incidencia  de este enlace {final['Enlace'].values} \n \n Capture adjunta en el correo"
    #attach the mensaje
    message.attach(MIMEText(mensaje))
    #attach the image with validation about the name of the files 
    for x in final["Enlace"].values:
        nombre= str(x)
        try:
            partes = nombre.split(" ")
            indice = partes.index("CNT")
            resultado = " ".join(partes[:indice-2])
        except:
            print("No se pudo")
            resultado = nombre
        archivoIMG = f"{resultado}.png"
        try:
            with open(archivoIMG, 'rb') as f:
                img_data = f.read()
            img = MIMEImage(img_data, name= archivoIMG)
            message.attach(img)
        except :
            print("No se adjunto")
        # Conexión al servidor SMTP
    #attach the Excel, it have a dataframe with the information
    try:    
        with open(archivo, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='octet-stream')
            attachment.add_header('Content-Disposition', 'attachment', filename= archivo)

        message.attach(attachment)
    except:
        print("No se adjunto")
    #Start connection with SMTP
    smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
    smtp_connection.starttls()
    smtp_connection.login(smtp_user, smtp_password)

    #Send the email 
    smtp_connection.sendmail(message['From'], message['To'], message.as_string())

    #Close the conection
    smtp_connection.quit()

