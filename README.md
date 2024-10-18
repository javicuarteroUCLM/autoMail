# ES ~ 🇪🇸

## Funcionalidad
Este script en Python automatiza el envío de correos electrónicos permitiendo enviar un archivo adjunto (único para cada usuario) dentro del mail.

El script lee los diferentes emails de diferentes hojas de un archivo excel (.xlxs) y según en la hoja que esté ubicado el email se le envía un tipo de archivo u otro.

Se ha usado el servicio SMTP de MandrillApp, asociado con Mailchimp.

Se incluye un archivo .xlxs (CorreosBonos.xlxs) de ejemplo para que se pueda ver el formato del excel. 

## Caso de uso del script
En este caso, se ha usado la automatización para enviar a todos los socios de un club deportivo todos sus cupones de descuento como recompensa por ser socio del club.

Según el tipo de abono, el archivo que había que enviarle se seleccionaba desde un directorio u otro.

Por ejemplo; 
Para los socios de tipo 'General', se les enviaba un archivo que contenía los descuentos ubicado en la carpeta 'BonosGeneral'.
Para los socios de tipo 'Leyenda', se les enviaba un archivo que contenía los descuentos ubicado en la carpeta 'BonosLeyenda'.

## Atención ⚠️
Este script manda un PDF *ÚNICO Y DISTINTO* a cada usuario.
Este script *NO* mandar el mismo PDF a todos los usuarios.

# EN ~ 🇬🇧

## Functionality
This Python script automates the sending of emails, allowing you to attach a file (unique for each user) to the email.

The script reads different emails from different sheets of an Excel file (.xlsx) and, depending on which sheet the email is located, a specific file type is sent.

MandrillApp's SMTP service, associated with Mailchimp, has been used.

An example .xlsx file (CorreosBonos.xlsx) is included to show the format of the Excel file.

## Script use case
In this case, the automation was used to send all members of a sports club their discount coupons as a reward for being club members.

Depending on the type of membership, the file to be sent was selected from one directory or another.

For example: For 'General' members, a file containing the discounts located in the 'BonosGeneral' folder was sent. For 'Leyenda' members, a file containing the discounts located in the 'BonosLeyenda' folder was sent.

## Atenttion ⚠️
This script sends a UNIQUE AND DIFFERENT PDF to each user.
This script does NOT send the same PDF to all users.



