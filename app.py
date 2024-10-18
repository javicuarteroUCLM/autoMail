import os
import base64
import pandas as pd
import mailchimp_transactional as MailchimpTransactional

from mailchimp_transactional.api_client import ApiClientError
from random import choice


# Mailchimp Transactional API key
API_KEY = 'putyourapikeyhere'
# API_KEY= 'putyourapikey-test-here'  #API KEY FOR TEST
mailchimp = MailchimpTransactional.Client(API_KEY)

# Asocia cada hoja de Excel con su directorio de PDFs/Associates each Excel sheet with its PDF directory
hojas_pdf_dirs = {
    "GENERAL": "pdfs/BonosGeneral",
    "GENERAL (HASTA 28 AÑOS)": "pdfs/BonosGeneralConDescuento",
    "SOCIOS ORO": "pdfs/BonosOro",
    "LEGADO": "pdfs/BonosLegado",
    "LEYENDA": "pdfs/BonosLeyenda",
}
# Control de archivos enviados
archivos_enviados = set()
correos_ok = []
correos_fail = []

def send_mail(receiver_email, pdf_path, nombre_abonado):
    with open(pdf_path, 'rb') as f:
        file_data = f.read()
        encoded_file_data = base64.b64encode(file_data).decode('utf-8')
    
    attachment = {
        'type': 'application/pdf',
        'name': os.path.basename(pdf_path),
        'content': encoded_file_data
    }

    message = {
        'from_email': 'putyouremail@email.com',
        'subject': 'PUT YOUR SUBJECT HERE',
        #'text': message_text,
        'to': [{'email': receiver_email, 'type': 'to'}],
        'attachments': [attachment],
        'merge_vars': [{'rcpt': receiver_email, 'vars': [{'name': 'FNAME', 'content': nombre_abonado}]}] # Para personalizar el mensaje/For customizing the message
    }

    try:
        #response = mailchimp.messages.send({"message": message}) # Para enviar sin plantilla/For sending without template
        response = mailchimp.messages.send_template({"template_name": "your_template_name",
        "template_content": [],
        "message": message
        })
        print(f'Correo enviado a {receiver_email}: {response}')
        correos_ok.append(receiver_email)
    except ApiClientError as error:
        print(f"Error al enviar email a {receiver_email}: {error.text}")
        correos_fail.append(receiver_email)



def procesar_hoja(nombre_hoja, directorio_pdf):
    df = pd.read_excel('YourFile.xlsx', sheet_name=nombre_hoja)
    pdf_files = os.listdir(directorio_pdf)

    # Iterar sobre cada fila del Excel y enviar un correo/ Iterate over each row of the Excel and send an email
    for _, row in df.iterrows():
        email = row['EMAIL']
        nombre_abonado = row['ABONADO']

        archivos_disponibles = [pdf for pdf in pdf_files if pdf not in archivos_enviados]
        if not archivos_disponibles:
            print(f"No hay suficientes archivos PDF en {directorio_pdf} para enviar a {email}.")
            return
        
        pdf_seleccionado = choice(archivos_disponibles)
        archivos_enviados.add(pdf_seleccionado)

        pdf_path = os.path.join(directorio_pdf, pdf_seleccionado)
        send_mail(email, pdf_path, nombre_abonado)


for hoja, directorio in hojas_pdf_dirs.items():
    procesar_hoja(hoja, directorio)

# Track de correos enviados
print("\n=== Resumen del envío de correos ===")
print("Correos enviados con éxito:")
for correo in correos_ok:
    print(correo)

print("\nCorreos fallidos:")
for correo in correos_fail:
    print(correo)

print("\nTodos los correos procesados.")




    
