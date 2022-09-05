# -*- coding: utf-8 -*-

# Proyecto Final de curso Python-ProgramBI
# Alumno : Christian Farnast Contardo
# Profesor: Emanuel Berrocal Zapata

#%%
# Utilizando el codigo final, utilizar todas las funciones vistas, para generar
# el siguiente Reporte en python:

# Muestre las utilidades de cada producto y por día.
# Muestre un resumen de la suma de utilidades por productos vendidos
# Muestre un resumen por ventas por mes-año de todo.
#%%
import pandas as pd
utilidades = pd.read_excel('Utilidades_anuales.xlsx')

def nom_prod(x):
    if x == 1:
        return(str('Televisor'))
    elif x == 2:
        return(str('Refrigerador'))
    elif x == 3:
        return(str('Secadora'))
    elif x == 4:
        return(str('Computador'))
    else:
        return(str('Sofá'))



#Aplicacion de la funcion "nombre"
utilidades['Nom_Producto'] = utilidades['Cod_Producto'].apply(nom_prod)
#%%
# Muestre las utilidades de cada producto y por día.
utilidades['Utilidades'] = (utilidades['Ventas'] - utilidades['Costos'])* utilidades['Cantidad_vendidas']
Total = utilidades['Utilidades'].sum()

# Muestre un resumen de la suma de utilidades por productos vendidos
utilidades_por_producto = utilidades.groupby('Nom_Producto')['Utilidades'].sum()

# Muestre un resumen por ventas por mes-año de todo.
ventas_mes_anio = utilidades.groupby(pd.Grouper(key='Fecha', axis=0, 
                      freq='M'))['Utilidades'].sum()
#%%
# Instalación paquete xlwings
#! pip install xlwings

import xlwings as xw
# Incorporando utilidades al informe
wb = xw.Book('C:/Users/Administrador/Desktop/ProgramBI/PYTHON/PROYECTO FINAL/Informe_Python_ProgramBI.xlsx')
sht1 = wb.sheets(1) 
sht1.range("A1").value = utilidades

# Incorporando utilidades_por_producto al informe
sht1.range("J1").value = utilidades_por_producto
sht1.range("J7").value = 'Total'
sht1.range("K7").value = Total

# Incorporando ventas_mes_anio al informe
sht1.range("J10").value = ventas_mes_anio
sht1.range("J19").value = 'Total'
sht1.range("K19").value = Total
#%%
# Envio de correo

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

username = "christian.farnast@edu-ccs.cl"
password = "XXXXXX"
mail_from = "christian.farnast@edu-ccs.cl"
mail_to = "eberrocal@programbi.cl"
mail_cc = 'moliva@programbi.cl', 'christian.farnast@edu-ccs.cl'
mail_subject = "Entrega proyecto final Python"
mail_body = "Estimado Emanuel, junto con saludar envío adjunto codigo y proyecto final. Gracias, Christian."
mail_attachment='C:/Users/Administrador/Desktop/ProgramBI/PYTHON/PROYECTO FINAL/PROYECTO FINAL.rar'
mail_attachment_name="PROYECTO FINAL.rar"

mimemsg = MIMEMultipart()
mimemsg['From']=mail_from
mimemsg['To']=mail_to
mimemsg['Cc']=mail_cc
mimemsg['Subject']=mail_subject
mimemsg.attach(MIMEText(mail_body, 'plain'))

with open(mail_attachment, "rb") as attachment:
    mimefile = MIMEBase('application', 'octet-stream')
    mimefile.set_payload((attachment).read())
    encoders.encode_base64(mimefile)
    mimefile.add_header('Content-Disposition', "attachment; filename= %s" % mail_attachment_name)
    mimemsg.attach(mimefile)
    connection = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()