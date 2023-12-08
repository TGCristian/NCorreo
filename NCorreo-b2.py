#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Import modules
import smtplib
import ssl
import os
import sqlite3
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

configFile = "./config.json"

with open(configFile, "r") as read_file:
    dataApp = json.load(read_file)

with sqlite3.connect(dataApp["appBaseDir"] + dataApp["dataBaseFile"]) as conn:
    cursor = conn.execute("SELECT * from LISTADO")
    lista = cursor.fetchall()
    
    
emailUser = dataApp["emailUser"]
emailPsw = dataApp["emailPsw"]
emailResponderA = dataApp["emailResponse"]

file_name = os.path.basename('./Barbijo.pdf')

encabezadoArchivo = {"Content-Disposition" : f"attachment; filename= {file_name}"}
encabezadoImagen = {
    "Content-Disposition" : f"attachment; filename='",
    'Content-ID': '<logo-png>'
    }

def adjuntarArchivo(emailMsg, rutaAdjunto, encabezados=None):

    with open(rutaAdjunto, "rb") as f:
        archivoAdjunto = MIMEApplication(f.read())

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

    with open(dataApp["appBaseDir"] + dataApp["htmlToSend"], "br") as f:
        mt = MIMEText(f.read().decode(), "html")
        emailMsg.attach(mt)

    adjuntarArchivo(emailMsg, './images/image-1.png', encabezadoImagen)
    adjuntarArchivo(emailMsg, './Barbijo.pdf', encabezadoArchivo)

    email_string = emailMsg.as_string()
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("mail.colegioparroquialsantarosa.com.ar", 465, context=context) as server:
        server.login(emailUser, emailPsw)
        server.sendmail(emailUser, familia[3], email_string) 

    emailMsg = None
    