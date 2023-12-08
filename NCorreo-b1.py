#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Import modules
import smtplib
import ssl
import os
import sqlite3
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
#from email import charset
#import pandas as pd

"""
html = '''
    <html>
        <body>
            <h1>Daily S&P 500 prices report</h1>
            <p>Hello, welcome to your report!</p>
            <img src='cid:myimageid' width="700">
        </body>
    </html>
    '''
"""
dataBase = "./cuotas.db"
with sqlite3.connect(dataBase) as conn:
    cursor = conn.execute("SELECT * from LISTADO")
    lista = cursor.fetchall()

def attach_file_to_email(emailMsg, rutaAdjunto, extra_headers=None):
    file_name = os.path.basename(rutaAdjunto)
    with open(rutaAdjunto, "rb") as f:
        archivoAdjunto = MIMEApplication(f.read())
    archivoAdjunto.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_name}",
    )
    if extra_headers is not None:
        archivoAdjunto.replace_header(
            "Content-Disposition",
            f"inline; filename='",
        )

        for name, value in extra_headers.items():
            archivoAdjunto.add_header(name, value)

    emailMsg.attach(archivoAdjunto)


emailUser = "tesoreria@colegioparroquialsantarosa.com.ar"
emailPsw = 'Lacasadepapel2020'
emailResponderA = "administracion@cpsantarosa.edu.ar"

for familia in lista:

    emailDestinatario = familia[3]
    emailAsunto = f'Cuota Septiembre - Flia. {familia[2]}'
 
#date_str = pd.Timestamp.today().strftime('%Y-%m-%d')

emailMsg = MIMEMultipart()
emailMsg['From'] = "CPSR Cupones <tesoreria@colegioparroquialsantarosa.com.ar"
emailMsg['To'] = emailDestinatario
emailMsg['Reply-To'] = emailResponderA
emailMsg['Subject'] = emailAsunto

with open("index.html", "br") as f:
    mt = MIMEText(f.read().decode(), "html")
    emailMsg.attach(mt)
    """    
    cs = charset.Charset('utf-8')
    cs.body_encoding = charset.QP
    mt = MIMEText(f.read().decode(), 'html', cs)
    """

attach_file_to_email(emailMsg, './images/image-1.png', {'Content-ID': '<logo-png>'})
attach_file_to_email(emailMsg, './Barbijo.pdf')

email_string = emailMsg.as_string()

context = ssl.create_default_context()
with smtplib.SMTP_SSL("mail.colegioparroquialsantarosa.com.ar", 465, context=context) as server:
    server.login(emailUser, emailPsw)
    server.sendmail(emailUser, emailDestinatario, email_string) 