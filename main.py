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

        # Buscar inicio de la sección
        inicio_idx = texto_completo.find(inicio)
        if inicio_idx == -1:
            return "⚠ Sección no encontrada en el PDF."

        # Buscar el final de la sección
        fin_idx = texto_completo.find(fin, inicio_idx)
        if fin_idx == -1:
            # Si no se encuentra el final, extraer hasta el final del texto
            fin_idx = len(texto_completo)

        # Extraer la sección
        seccion = texto_completo[inicio_idx:fin_idx]

        # Obtener la parte relevante del nombre del archivo
        filename = os.path.basename(pdf_path)
        relevant_part = re.search(r'\((.*?)\)', filename)
        if relevant_part:
            relevant_part = relevant_part.group(1)
        else:
            relevant_part = filename.split('Cuenta')[-1].strip().replace('.pdf', '')

        if relevant_part:
            seccion = seccion.replace(inicio, f"{inicio} - {relevant_part}", 1)
        else:
            seccion = seccion.replace(inicio, f"{inicio}", 1)
        return seccion

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
        for file in adjunto:
            mail.Attachments.Add(file)

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
    pdf_dir = "C:/Users/Alessandro/Desktop/portfolios/"
    pdf_files = [f for f in os.listdir(pdf_dir) if f.startswith(codigo.zfill(3)) and f.endswith(".pdf")]
    password = f"gamnic{codigo.zfill(3)}"
    secciones_extraidas = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_dir, pdf_file)
        seccion_extraida = extraer_seccion_pdf(pdf_path, "Rentabilidad del Portafolio", "Comentario de Mercado", password)
        seccion_extraida = seccion_extraida.replace("Portafolio", "Mes")
        seccion_extraida_modificada = poner_en_negrita_despues_de_es(seccion_extraida, "Rentabilidad del Portafolio")
        seccion_extraida_html = formatear_a_html(seccion_extraida_modificada, font_family="Calibri", line_height="1")
        secciones_extraidas.append(seccion_extraida_html)

    cuerpo_completo = "<br><br>".join(secciones_extraidas)

    enviar_correo(
        destinatario=datos["email"],
        nombre=datos["nombre"],
        asunto="Extracto del Estado de Cuenta",
        cuerpo=f"Esperamos que te encuentres bien.<br><br> Te adjuntamos el estado de cuenta consolidado al cierre de {last_month_str}. La clave es la de siempre.<br><br>{cuerpo_completo}<br>Quedamos a tu disposición para reunirnos y revisar los resultados, el detalle del portafolio, la estrategia a seguir y las ideas de inversión.<br>",
        adjunto=[os.path.join(pdf_dir, pdf_file) for pdf_file in pdf_files],
        imagen_path=imagen_path
    )

print("✅ Se enviaron todos los correos exitosamente.")