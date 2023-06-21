from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter
from datetime import datetime
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
import time
from email.message import EmailMessage
from reportlab.pdfgen import canvas
import tkinter as tk


def buscar_archivo_excel():
    ruta_archivo_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=(
        ("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
    ruta_archivo_excel_entry.delete(0, tk.END)
    ruta_archivo_excel_entry.insert(tk.END, ruta_archivo_excel)


def buscar_carpeta_destino():
    carpeta_destino = filedialog.askdirectory(
        title="Seleccionar carpeta de destino")
    carpeta_destino_entry.delete(0, tk.END)
    carpeta_destino_entry.insert(tk.END, carpeta_destino)


def generar_comprobante(ruta_comprobante, correo_destino, asunto, cuerpo, adjunto):
    doc = SimpleDocTemplate(ruta_comprobante, pagesize=letter, rightMargin=1 *
                            inch, leftMargin=1*inch, topMargin=1*inch, bottomMargin=1*inch)
    elements = []

    styles = getSampleStyleSheet()

    # NOTAS# Encabezado del comprobante
    header = Paragraph(
        "<b>Comprobante de Envío de Correo</b>", styles["Heading1"])
    elements.append(header)

    # NOTAS# Información del correo
    elements.append(Spacer(1, 0.2*inch))
    elements.append(Paragraph(
        "<b>Correo enviado a:</b> {}".format(correo_destino), styles["Normal"]))
    elements.append(Paragraph(
        "<b>Fecha:</b> {}".format(datetime.now().strftime("%Y-%m-%d")), styles["Normal"]))
    elements.append(Paragraph(
        "<b>Hora:</b> {}".format(datetime.now().strftime("%H:%M:%S")), styles["Normal"]))
    elements.append(
        Paragraph("<b>Asunto:</b> {}".format(asunto), styles["Normal"]))

    # NOTAS# Cuerpo del correo
    elements.append(Paragraph("<b>Cuerpo del correo:</b>", styles["Normal"]))
    cuerpo_paragraphs = cuerpo.split("\n")
    cuerpo_style = ParagraphStyle(
        name='CuerpoCorreoStyle',
        parent=styles['Normal'],
        spaceAfter=0.2*inch,
        alignment=1,  # NOTAS# Alineación justificada
        fontSize=12,
    )
    for p in cuerpo_paragraphs:
        elements.append(Paragraph(p, cuerpo_style))
    elements.append(Spacer(1, 0.2*inch))

    # NOTAS# Información de adjunto
    if adjunto:
        elements.append(Paragraph("<b>Adjunto:</b> Sí", styles["Normal"]))
    else:
        elements.append(Paragraph("<b>Adjunto:</b> No", styles["Normal"]))

    # NOTAS# Comprobante generado por ReportLab
    elements.append(Spacer(1, 0.4*inch))
    elements.append(
        Paragraph("<b>Comprobante generado por ReportLab</b>", styles["Heading2"]))

    doc.build(elements)


def enviar_correos():
    ruta_archivo_excel = ruta_archivo_excel_entry.get()
    carpeta_destino = carpeta_destino_entry.get()

    # NOTAS# Leer el archivo Excel
    try:
        df = pd.read_excel(ruta_archivo_excel)
    except pd.errors.EmptyDataError:
        messagebox.showerror("Error", "El archivo Excel está vacío.")
        return
    except pd.errors.ParserError:
        messagebox.showerror("Error", "Error al leer el archivo Excel.")
        return
    except FileNotFoundError:
        messagebox.showerror("Error", "No se encontró el archivo Excel.")
        return

    # NOTAS# Configurar servidor SMTP
    server = smtplib.SMTP_SSL('mail.serviciosespeciales.com.co', 465)
    username = 'curaduria@serviciosespeciales.com.co'
    password = 'Servicios2023'
    try:
        server.login(username, password)
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror(
            "Error", "No se pudo iniciar sesión en el servidor SMTP. Verifica las credenciales de acceso.")
        return

    # NOTAS# Recorrer filas del DataFrame
    for index, row in df.iterrows():
        asunto = row['ASUNTO']
        cuerpo = row['CUERPO']
        correo_destino = row['CORREO']
        ruta_adjunto = row['RUTA CARPETA']

        # NOTAS# Crear objeto EmailMessage
        mensaje = EmailMessage()
        mensaje['Subject'] = asunto
        mensaje['From'] = username
        mensaje['To'] = correo_destino
        # NOTAS# Agregar destinatario en copia
        mensaje['Cc'] = 'curaduria@serviciosespeciales.com.co'
        mensaje.set_content(cuerpo)

        # NOTAS# Adjuntar archivo PDF
        with open(ruta_adjunto, 'rb') as f:
            contenido_adjunto = f.read()
        mensaje.add_attachment(
            contenido_adjunto, maintype='application', subtype='pdf', filename='adjunto.pdf')

        try:
            # NOTAS# Enviar correo
            server.send_message(mensaje)

            # NOTAS# Generar comprobante de envío
            nombre_comprobante = str(row['ASSET']) + '.pdf'
            ruta_comprobante = carpeta_destino + '/' + nombre_comprobante
            adjunto = True  # NOTAS# O False, según corresponda
            generar_comprobante(
                ruta_comprobante, correo_destino, asunto, cuerpo, adjunto)

            # NOTAS# Esperar un momento antes de continuar
            time.sleep(0)

        except smtplib.SMTPException as e:
            messagebox.showerror(
                "Error", "Error al enviar el correo: " + str(e))

    # NOTAS# Cerrar conexión con el servidor SMTP
    server.quit()

    messagebox.showinfo("Información", "Envío de correos completado.")


# NOTAS# Crear ventana
ventana = tk.Tk()
ventana.title("MASSIVE PRO | CURADURIA VERS")
ventana.configure(bg='dark olive green')
ventana.geometry("475x215")  # NOTAS# Cambio de tamaño de la ventana

ruta_archivo_excel_label = tk.Label(
    ventana, text="DAME LA RUTA DE TU ARCHIVO:", bg='dark olive green', fg='white')
ruta_archivo_excel_label.pack()

ruta_archivo_excel_entry = tk.Entry(ventana)
ruta_archivo_excel_entry.pack()

ruta_archivo_excel_button = tk.Button(
    ventana, text="Buscar", command=buscar_archivo_excel, bg='orange')
ruta_archivo_excel_button.pack()

carpeta_destino_label = tk.Label(
    ventana, text="DONDE DESEAS GUARDAR LOS COMPROBANTES:", bg='dark olive green', fg='white')
carpeta_destino_label.pack()

carpeta_destino_entry = tk.Entry(ventana)
carpeta_destino_entry.pack()

buscar_carpeta_destino_button = tk.Button(
    ventana, text="Buscar", command=buscar_carpeta_destino, bg='orange')
buscar_carpeta_destino_button.pack()

enviar_correos_button = tk.Button(
    ventana, text="Enviar correos", command=enviar_correos, bg='orange')
enviar_correos_button.pack()

etiqueta_final = tk.Label(
    ventana, text="  \u00a9 Rambo Programmer", bg='dark olive green', fg='white')
etiqueta_final.pack()

ventana.mainloop()
