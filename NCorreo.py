#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Import modules

import ssl
import logging
import os
import sqlite3
from json import load
from time import sleep
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

configFile = "./config.json"
sendFail = 0 


logging.basicConfig(filename='info.log', encoding='utf-8', format='%(levelname)s: %(asctime)s %(message)s', datefmt='%d/%m/%Y %I:%M:%S %p', level=logging.DEBUG)

try:
    with open(configFile, "r") as read_file:
        dataApp = load(read_file)
except:
    logging.critical(": Problemas en el Archivo de Configuración")
    exit()

try:
    with sqlite3.connect(dataApp["appBaseDir"] + dataApp["dataBaseFile"]) as conn:
        cursor = conn.execute("SELECT * from LISTADO")
        lista = cursor.fetchall()
except:
    logging.critical(": Problemas en la BD")
    exit()

   
emailUser = dataApp["emailUser"]
emailPsw = dataApp["emailPsw"]
emailResponderA = dataApp["emailResponse"]
logging.info(": Iniciando Envío")

def adjuntarArchivo(emailMsg, rutaAdjunto, encabezados=None):
    try:
        with open(rutaAdjunto, "rb") as f:
            archivoAdjunto = MIMEApplication(f.read())
    except:
        logging.warning(": Problemas al Abrir Adjunto:" + rutaAdjunto)
        return()       

    if encabezados is not None:

        for name, value in encabezados.items():
            archivoAdjunto.add_header(name, value)

    emailMsg.attach(archivoAdjunto)

for familia in lista:

    emailMsg = MIMEMultipart()
    emailMsg['From'] = dataApp["emailUserAlias"]
    emailMsg['To'] = familia[3]
    emailMsg['Reply-To'] = emailResponderA
    emailMsg['Subject'] = dataApp["emailAsuntoGeneral"] + f' - Flia. {familia[2]}'

    try:
        with open(dataApp["appBaseDir"] + dataApp["htmlToSend"], "br") as f:
            mt = MIMEText(f.read().decode(), "html")
            emailMsg.attach(mt)
    except:
        logging.error(": Problemas al Abrir HTML")       
        exit()           

    for archivo in dataApp["archivosAdjuntos"]:
        adjuntarArchivo(emailMsg, archivo["pathFile"] + archivo["nameFile"] + archivo["extensionFile"], archivo["headers"])

    if dataApp["cuponesSend"]:

        cuotasDirs = os.scandir(dataApp["appBaseDir"] + dataApp["cuponesAdjuntos"]["pathFile"])
        cuotasFileType =  dataApp["cuponesAdjuntos"]["extensionFile"]
        cuotasFileHeader = dataApp["cuponesAdjuntos"]["headers"]
      
        for dir in cuotasDirs:
            cuponActual = os.path.abspath(dir.path + "/" + str(familia[1]) + cuotasFileType)
            headerActual = cuotasFileHeader.copy()
            headerActual["Content-Disposition"] = dataApp["cuponesAdjuntos"]["headers"]["Content-Disposition"] + dir.name + cuotasFileType

            if os.path.isfile(cuponActual):
                adjuntarArchivo(emailMsg, cuponActual, headerActual)


    email_string = emailMsg.as_string()
    context = ssl.create_default_context()


    while (sendFail < dataApp["maxSendFails"]):
        try:
            with SMTP_SSL("mail.colegioparroquialsantarosa.com.ar", 465, context=context) as server:
                server.login(emailUser, emailPsw)
                server.sendmail(emailUser, familia[3], email_string) 
                sleep(dataApp["delaySend"])
                logging.info(f': Envío Realizado a: {str(familia[1])} - Flia. {familia[2]}')
                break
        except:
            sendFail += 1
            logging.warning(": Fallo al Enviar Intento " + str(sendFail) + " a: " + familia[3])
            sleep(sendFail * dataApp["delayFailSend"])
     
    emailMsg = None

    if (sendFail < dataApp["maxSendFails"]):
        sendFail = 0 
    else:
        logging.error(": Falló la Conexión")
        exit()

#   print(datetime.now().strftime('%Y-%m-%d %H:%M:%S') +": Fallo al Enviar Intento " + str(sendFail) + " a: " + familia[3])