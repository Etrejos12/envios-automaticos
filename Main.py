"""ENVIO_DE_RESPUESTAS_POR_CORREO"""

import os
import re
import smtplib
import win32com.client as win32
import pdfplumber
from docx import Document
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import pandas as pd
import json
import time
import sys
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

def cargar_configuracion():
    with open("C:/Configuracion/Configuracion.json", "r", encoding="utf-8") as f:
        return json.load(f)

configuracion = cargar_configuracion()
# ===========================
# CONFIGURACIÓN
# ===========================

RUTA_PRINCIPAL = configuracion["RUTA_PRINCIPAL"]
REMITENTE = configuracion["REMITENTE"]

# ===========================
# FUNCIONES
# ===========================

def extraer_nombre_y_correo(file_path, max_lineas=20):
    """Lee PDF o DOCX y extrae nombre + correo buscando solo en las primeras `max_lineas` líneas."""
    text = ""

    # Leer archivo según extensión
    if file_path.endswith(".pdf"):
        with pdfplumber.open(file_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    elif file_path.endswith(".docx"):
        doc = Document(file_path)
        text = "\n".join(p.text for p in doc.paragraphs)

    # Limpiar y dividir en líneas
    lineas = [l.strip() for l in text.split("\n") if l.strip()]

    # Limitar a las primeras `max_lineas`
    lineas = lineas[:max_lineas]

    # Patrón para correo
    patron_correo = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"

    nombre = None
    correo = None

    # Buscar correo y nombre solo dentro del bloque recortado
    for i, linea in enumerate(lineas):
        m = re.search(patron_correo, linea)
        if m:
            correo = m.group()
            # Buscar nombre en líneas anteriores (máx 5 arriba dentro del recorte)
            for j in range(i - 1, max(i - 5, -1), -1):
                if re.match(r"^[A-ZÁÉÍÓÚÑ ]{5,}$", lineas[j]):
                    nombre = lineas[j].title()
                    break
            break

    return nombre, correo

def registrar_envio_excel(destinatario, asunto, entry_id):
    archivo = configuracion["ARCHIVO_EXCEL"]

    nueva_fila = {
        "Fecha_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Destinatario": destinatario,
        "Asunto": asunto,
        "ID": entry_id
    }

    if os.path.exists(archivo):
        df = pd.read_excel(archivo)
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
        df.to_excel(archivo, index=False)
    else:
        df = pd.DataFrame([nueva_fila])
        df.to_excel(archivo, index=False)

def enviar_mail_outlook(destinatario, asunto, CUERPO_HTML, adjuntos):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    mail.Subject = asunto
    mail.HTMLbody = CUERPO_HTML
    mail.To = destinatario
    mail.OriginatorDeliveryReportRequested = True
    entry_id = mail.EntryID

    mail.SentOnBehalfOfName = REMITENTE

    for archivo in adjuntos:
        mail.Attachments.Add(archivo)

    mail.Save()
    entry_id = mail.EntryID
    mail.Send()

    registrar_envio_excel(destinatario, asunto, entry_id)

# --------------------------------------------------------
# VALIDACIÓN 1: Verificar que todas las cartas tengan correo (robusta)
# Usa el MISMO extractor de tu código (extraer_nombre_y_correo)
# --------------------------------------------------------
def validar_correos():
    errores = []

    for carpeta in os.listdir(RUTA_PRINCIPAL):
        if not carpeta.startswith("AS_"):
            continue

        codigo = carpeta.replace("AS_", "")
        ruta_carpeta = os.path.join(RUTA_PRINCIPAL, carpeta)

        cr_file = None
        for archivo in os.listdir(ruta_carpeta):
            # Para ser simétrico con procesar(), validamos solo PDF
            if archivo.startswith(f"CR_{codigo}") and archivo.endswith(".pdf"):
                cr_file = os.path.join(ruta_carpeta, archivo)
                break

        if not cr_file:
            errores.append(f"No existe CR_{codigo} en {carpeta}")
            continue

        try:
            nombre, correo = extraer_nombre_y_correo(cr_file)
        except Exception as e:
            errores.append(f"CR_{codigo}: ERROR al leer el archivo ({type(e).__name__}: {e})")
            continue

        if not correo:
            errores.append(f"CR_{codigo}: SIN CORREO")

    return errores

# --------------------------------------------------------
# VALIDACIÓN 2: Verificar que el ZIP pese menos de 13 MB
# --------------------------------------------------------
def validar_pesos_zip():
    errores = []

    for carpeta in os.listdir(RUTA_PRINCIPAL):
        if not carpeta.startswith("AS_"):
            continue

        codigo = carpeta.replace("AS_", "")
        zip_file = os.path.join(RUTA_PRINCIPAL, carpeta, f"AS_{codigo}.zip")

        if os.path.exists(zip_file):
            tamaño_mb = os.path.getsize(zip_file) / (1024 * 1024)

            if tamaño_mb > 18:
                errores.append(f"AS_{codigo}.zip pesa {tamaño_mb:.2f} MB (supera 18 MB)")

    return errores

# --------------------------------------------------------
# FUNCIÓN PARA MOSTRAR ALERTA (siempre visible) Y PREGUNTAR SI CONTINUAR
# Botones: "Corregido" / "Cancelar"
# --------------------------------------------------------
def mostrar_alerta_errores(lista_errores):
    """
    Muestra un diálogo modal con la lista de errores y dos botones:
    - "Corregido": devuelve True
    - "Cancelar": devuelve False
    """
    root = tk.Tk()
    root.withdraw()
    root.update_idletasks()
    try:
        root.attributes("-topmost", True)  # traer al frente
    except tk.TclError:
        pass
    root.lift()
    root.focus_force()

    # Crear ventana de diálogo
    dialog = tk.Toplevel(root)
    dialog.title("VALIDACIONES FALLIDAS")
    dialog.geometry("720x480")
    dialog.resizable(True, True)

    # Modal
    dialog.transient(root)
    dialog.grab_set()
    try:
        dialog.attributes("-topmost", True)
    except tk.TclError:
        pass

    # Contenedor
    frm = ttk.Frame(dialog, padding=12)
    frm.pack(fill="both", expand=True)

    lbl = ttk.Label(frm, text="Se encontraron los siguientes puntos para corregir:", anchor="w")
    lbl.pack(fill="x", pady=(0, 8))

    # Texto con scroll
    txt_frame = ttk.Frame(frm)
    txt_frame.pack(fill="both", expand=True)

    txt = tk.Text(txt_frame, wrap="word")
    txt_scroll = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=txt_scroll.set)

    txt.pack(side="left", fill="both", expand=True)
    txt_scroll.pack(side="right", fill="y")

    # Contenido
    if lista_errores:
        txt.insert("1.0", "\n".join(f"- {e}" for e in lista_errores))
    else:
        txt.insert("1.0", "No se encontraron errores.")
    txt.configure(state="disabled")

    # Botones
    btns = ttk.Frame(frm)
    btns.pack(fill="x", pady=(12, 0))

    resultado = {"valor": False}

    def on_corregido():
        resultado["valor"] = True
        dialog.destroy()

    def on_cancelar():
        resultado["valor"] = False
        dialog.destroy()

    btn_corregido = ttk.Button(btns, text="Corregido", command=on_corregido)
    btn_cancelar = ttk.Button(btns, text="Cancelar", command=on_cancelar)

    btn_corregido.pack(side="right", padx=(8, 0))
    btn_cancelar.pack(side="right")

    # Centrar en pantalla
    dialog.update_idletasks()
    w = dialog.winfo_width()
    h = dialog.winfo_height()
    sw = dialog.winfo_screenwidth()
    sh = dialog.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    dialog.geometry(f"+{x}+{y}")

    # Esperar interacción
    root.deiconify()
    root.wait_window(dialog)

    try:
        root.attributes("-topmost", False)
    except tk.TclError:
        pass
    root.destroy()

    return resultado["valor"]

# --------------------------------------------------------
# ALERTA DE ÉXITO FINAL
# --------------------------------------------------------
def alerta_exito():
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    messagebox.showinfo("ÉXITO", "Todos los envíos se realizaron con éxito.", parent=root)
    try:
        root.attributes("-topmost", False)
    except tk.TclError:
        pass
    root.destroy()

# ===========================
# PROCESO PRINCIPAL
# ===========================
def procesar():
    for carpeta in os.listdir(RUTA_PRINCIPAL):
        if not carpeta.startswith("AS_"):
            continue

        codigo = carpeta.replace("AS_", "")
        ruta_carpeta = os.path.join(RUTA_PRINCIPAL, carpeta)

        # Buscar CR_codigo# (solo PDF, igual que la validación)
        cr_file = None
        for archivo in os.listdir(ruta_carpeta):
            if archivo.startswith(f"CR_{codigo}") and archivo.endswith(".pdf"):
                cr_file = os.path.join(ruta_carpeta, archivo)
                break

        if not cr_file:
            print(f"[ERROR] No se encontró CR_{codigo} en {carpeta}")
            continue

        # Buscar ZIP (opcional)
        zip_file = os.path.join(ruta_carpeta, f"AS_{codigo}.zip")
        zip_existe = os.path.exists(zip_file)

        # Extraer info de destinatario
        nombre, correo = extraer_nombre_y_correo(cr_file)

        if not correo:
            print(f"[ERROR] No se detectó correo en {cr_file}")
            continue

        print(f"Procesando código {codigo} → destinatario {correo}")

        # Asunto
        asunto = f"CONFIDENCIAL: Respuesta Requerimiento Legal Bancolombia Cod No {codigo}"

        # Evitar TypeError si nombre es None
        nombre = nombre or "Estimado/a"

        CUERPO_HTML = f"""
<html>
<body style="font-family:Arial; font-size:14px;">
    <p style="color:#666666;">Hola {nombre} </p>
    <p style="color:#666666;">
        En atención a su solicitud, le adjuntamos la respectiva carta de respuesta correspondiente a la gestión realizada por parte de nuestra entidad,
        <b>con código interno No {codigo} (Favor citar en el asunto al responder).</b>
    </p>
    <p style="color:#666666;">
        <b>Por reserva bancaria, el correo electrónico se encuentra encriptado. Para abrirlo por favor dar clic, digitar su correo electrónico y la contraseña de su buzón y con el código generado por el sistema, podrá acceder a la información.</b>
    </p>
    <p style="color:#666666;">
        Con el envío de este correo y el acuse automático de lectura y recibo que genera el mismo, damos por entendido la recepción en su Entidad.
        <b>Agradecemos no copiar el correo de Requerinf en las gestiones internas que se realizan, debido a que esto está afectando la capacidad de nuestro buzón.</b>
    </p>
    <p style="color:#666666;">
        NOTA DE CONFIDENCIALIDAD: la información contenida en este correo electrónico y sus anexos son confidenciales y contienen información privilegiada.
        Está dirigido para el uso exclusivo del destinatario. Si ha recibido este mensaje por error, queda notificado que cualquier divulgación del correo electrónico o sus anexos, está estrictamente prohibido.
        Si usted no es el destinatario del correo electrónico por favor borre todas las copias del mensaje y sus anexos y notifíquenos inmediatamente.
    </p>
    """

        # Cuerpo con reemplazo
        cuerpo_final = CUERPO_HTML.replace("{{nombre}}", nombre).replace("{{codigo}}", codigo)

        # Adjuntos según regla:
        adjuntos = [cr_file]
        if zip_existe:
            adjuntos.append(zip_file)

        # Enviar
        enviar_mail_outlook(correo, asunto, cuerpo_final, adjuntos)

        
        # Esperar 5 segundos antes de continuar con el siguiente correo
        time.sleep(configuracion["TIEMPO_ESPERA"])


        if zip_existe:
            print(f"[OK] Enviado PDF + ZIP → {correo}")
        else:
            print(f"[OK] Enviado solo PDF (sin ZIP disponible) → {correo}")

if __name__ == "__main__":
    # Revalidación automática hasta que no haya errores o el usuario cancele
    while True:
        errores_correos = validar_correos()
        errores_pesos = validar_pesos_zip()
        todos_errores = errores_correos + errores_pesos

        if not todos_errores:
            break  # No hay errores → enviar

        continuar = mostrar_alerta_errores(todos_errores)
        if not continuar:
            print("Envío cancelado por el usuario.")
            sys.exit(0)
        # Si elige "Corregido", el bucle vuelve a validar

    # Si todo bien → ejecutar envíos
    procesar()

    # Mensaje final de éxito
    alerta_exito()
