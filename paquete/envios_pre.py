import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
from tkinter import messagebox


def envio_correos_previo(correo_electronico, password, usuario):
    try:
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        smtp_username = correo_electronico
        smtp_password = password

        # Cargar un archivo de Excel
        tabla_destinatarios = pd.read_excel('./excels/db_fuente.xlsx', sheet_name='Sheet1')

        # Conectar y autenticar con el servidor de correo
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)

        # Iterar a través de la tabla de destinatarios y enviar correos individuales
        for index, row in tabla_destinatarios.iterrows():
            emails = row['CORREOS'].split(';')  
            codigo_compania = row['CIA']
            nombre_compañia = row['NOMBRE']

            for correo in emails:
                msg = MIMEMultipart()
                msg['From'] = correo_electronico
                msg['To'] = correo.strip()  # Eliminar espacios en blanco alrededor de la dirección de correo
                msg['Subject'] = f'RECORDATORIO - {nombre_compañia}'

                # Mensaje que deseas enviar
                mensaje = f"""
Estimado/a:

Se les recuerda a las compañías que el día 20 del corriente mes vence el plazo para subir a SINENSUP lo correspondiente a la Información de Producción Mensual.
Estamos a disposición ante cualquier duda puede comunicarse mediante este mail al teléfono 4338-4000 interno 1559.

{usuario}
Gerencia de Estudios y Estadísticas
{smtp_username}
Tel: +54 11 4338-4000
www.argentina.gob.ar/ssn
Av. Julio A. Roca 721 | (C1067ABC) | CABA | República Argentina
                """
                # Agregar el mensaje al cuerpo del correo
                msg.attach(MIMEText(mensaje, 'plain'))

                # Envía el correo al destinatario actual
                server.sendmail(smtp_username, correo.strip(), msg.as_string())

                print(f"Correo enviado a {correo} para {codigo_compania}")

        # Cierra la conexión con el servidor de correo
        server.quit()

        print("Proceso terminado.")
        messagebox.showinfo("Éxito", "Correos enviados exitosamente.")
    except Exception as e:
        print(f"Error al enviar correos: {e}")
        messagebox.showerror("Error", f"Error al enviar correos: {e}")