import re
import pdfplumber
import win32com.client as win32
from clientes import clientes
import datetime
import locale
import os

imagen_path = "C:/Users/Alessandro/PycharmProjects/automated_sending_mail/gamnic.png"

def extraer_seccion_pdf(pdf_path, inicio, fin, password=None):
    with pdfplumber.open(pdf_path, password=password) as pdf:
        texto_completo = ""
        for page in pdf.pages:
            texto_completo += page.extract_text() + "\n"

        # Buscar inicio y fin de la sección
        inicio_idx = texto_completo.find(inicio)
        fin_idx = texto_completo.find(fin, inicio_idx)

        if inicio_idx != -1 and fin_idx != -1:
            return texto_completo[inicio_idx:fin_idx + len(fin)]
        else:
            return "⚠ Sección no encontrada en el PDF."

def poner_en_negrita_despues_de_es(texto, palabra_negrita):
    lineas = texto.split('\n')
    resultado = []
    for linea in lineas:
        linea_modificada = re.sub(r'\bes\b(.*)', r'es<b>\1</b>', linea)
        linea_modificada = re.sub(re.escape(palabra_negrita), f'<b>{palabra_negrita}</b>', linea_modificada)
        resultado.append(linea_modificada)
    return '\n'.join(resultado)

def formatear_a_html(texto, font_family="Calibri", line_height="1"):
    lineas = texto.split('\n')
    resultado = []
    for i, linea in enumerate(lineas):
        if i == 0:
            # Add a line break after the title
            resultado.append(f'<p style="font-family: {font_family}; line-height: {line_height};"><b>{linea}</b></p>')
        else:
            # Indent the rest of the text
            resultado.append(f'<p style="font-family: {font_family}; line-height: {line_height}; margin-left: 20px;">{linea}</p>')
    return ''.join(resultado)

def enviar_correo(destinatario, nombre, asunto, cuerpo, adjunto=None, imagen_path=None):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    mail.To = destinatario
    mail.Subject = asunto
    mail.BodyFormat = 1  # Establecer el formato del cuerpo como HTML

    # Cuerpo del correo con HTML, incluyendo la imagen embebida
    cuerpo_html = f"""
    <html>
    <body>
        <p>Hola {nombre},</p>
        <p>{cuerpo}</p>
        <p>Saludos,</p>
        <p>Mario Ubillús</p>
        <img src="cid:gamnic_image" alt="Gamnic Logo"/>
        <p>T: +51 1 437 6494 Ext. 109</p>
        <p>C: +51 989875041<br>www.gamnic.com</p>
    </body>
    </html>
    """

    mail.HTMLBody = cuerpo_html

    if imagen_path:
        # Primero, agregar la imagen como adjunto
        attachment = mail.Attachments.Add(imagen_path)
        # Establecer ContentID en el adjunto
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "gamnic_image")

    if adjunto:
        mail.Attachments.Add(adjunto)

    mail.Send()
    print(f"📧 Correo enviado a {nombre} ({destinatario}).")

# Obtener el mes anterior
locale.setlocale(locale.LC_TIME, "Spanish_Spain.1252")
today = datetime.date.today()
first = today.replace(day=1)
last_month = first - datetime.timedelta(days=1)
last_month_str = last_month.strftime('%B').capitalize()

# Enviar los correos
for codigo, datos in clientes.items():
    pdf_path = f"C:/Users/Alessandro/Desktop/portfolios/{codigo.zfill(3)} - 2024 12 - Estado de Cuenta.pdf"
    password = f"gamnic{codigo.zfill(3)}"  # Replace with the actual logic to get the password
    seccion_extraida = extraer_seccion_pdf(pdf_path, "Rentabilidad del Portafolio", "anual)", password)

    seccion_extraida_modificada = poner_en_negrita_despues_de_es(seccion_extraida, "Rentabilidad del Portafolio")
    seccion_extraida_html = formatear_a_html(seccion_extraida_modificada, font_family="Calibri", line_height="1")

    enviar_correo(
        destinatario=datos["email"],
        nombre=datos["nombre"],
        asunto="Extracto del Estado de Cuenta",
        cuerpo=f"Esperamos que te encuentres bien.<br><br> Te adjuntamos el estado de cuenta consolidado al cierre de {last_month_str}. La clave es la de siempre.<br><br>{seccion_extraida_html}<br>Quedamos a tu disposición para reunirnos y revisar los resultados, el detalle del portafolio, la estrategia a seguir y las ideas de inversión.<br>",
        adjunto=pdf_path,
        imagen_path=imagen_path
    )

print("✅ Se enviaron todos los correos exitosamente.")