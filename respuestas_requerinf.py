# -*- coding: utf-8 -*-
"""
respuestas_requerinf.py  (v2 – alineado con Main.py del repo envios-automaticos)
================================================================================
- El destinatario (y el nombre) se extraen de la carta CR_<codigo> (PDF o DOCX),
  primeras 20 líneas, igual que en Main.py.
- Correo NUEVO desde el buzón delegado, asunto fijo con el código.
- Adjunta carta CR_<codigo> y, si existe, AS_<codigo>.zip.
- La carta puede estar suelta o dentro de la carpeta AS_<codigo>; el zip igual.
- Valida antes de enviar (carta, correo, peso) con diálogo Corregido / Cancelar.
- Solo acuse de ENTREGA; lo busca en Bandeja de entrada y Eliminados (tu buzón
  y el delegado), lo descarga como .msg y lo registra en el log CSV.

Requisitos: pip install pywin32 pdfplumber python-docx   |  Outlook Classic abierto

Uso:
    python respuestas_requerinf.py enviar [--simular]
    python respuestas_requerinf.py acuses
    python respuestas_requerinf.py todo
"""
import argparse
import csv
import json
import re
import shutil
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

# =====================================================================
# CONFIGURACIÓN  (si existe Configuracion.json se usan sus claves)
# =====================================================================
CONFIG_JSON = Path(r"C:\Configuracion\Configuracion.json")


def _leer_json():
    try:
        with open(CONFIG_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


_cfg = _leer_json()

CARPETA_ENTRADA = Path(_cfg.get("RUTA_PRINCIPAL", r"C:\Respuestas\Entrada"))
CARPETA_ENVIADOS = Path(_cfg.get("RUTA_ENVIADOS", r"C:\Respuestas\Enviados"))
CARPETA_ACUSES = Path(_cfg.get("RUTA_ACUSES", r"C:\Respuestas\Acuses"))
ARCHIVO_LOG = Path(_cfg.get("ARCHIVO_LOG", r"C:\Respuestas\log_envios.csv"))
REMITENTE = _cfg.get("REMITENTE", "respuestasrequerinf@bancolombia.com.co")
PAUSA_ENTRE_ENVIOS_SEG = _cfg.get("TIEMPO_ESPERA", 2)

ASUNTO = "CONFIDENCIAL: Respuesta Requerimiento Legal Bancolombia Cod No {codigo}"

# Límites de peso en MB (el adjunto crece ~37 % al codificarse en el correo).
MAX_ARCHIVO_MB = 18   # por archivo (Main.py validaba el zip a 18 MB)
MAX_TOTAL_MB = 18     # carta + zip juntos

MOVER_ENVIADOS = True   # mueve lo enviado a CARPETA_ENVIADOS
MAX_LINEAS_CARTA = 20   # líneas de la carta donde se busca el correo

# Carpetas de Outlook donde buscar el acuse: 6 = Bandeja de entrada, 3 = Eliminados
CARPETAS_ACUSE = [6, 3]
INCLUIR_SUBCARPETAS = True

# ---------------------------------------------------------------------
# CUERPO DEL CORREO (HTML). Pega tu HTML entre las comillas triples.
# Se reemplazan {nombre} y {codigo}.
# ---------------------------------------------------------------------
CUERPO_HTML = """
<html>
<body style="font-family:Arial; font-size:14px;">
<!-- ===== PEGA AQUÍ TU HTML ===== -->
<p>Hola {nombre}</p>
<p>Respuesta al requerimiento con código interno No {codigo}.</p>
<!-- ===== FIN DEL HTML ===== -->
</body>
</html>
"""

# =====================================================================
# LOG
# =====================================================================
CAMPOS_LOG = [
    "codigo", "fecha_envio", "destinatario", "asunto", "carta", "zip",
    "peso_mb", "estado_envio", "estado_acuse", "fecha_acuse", "ruta_acuse",
    "ultima_busqueda_acuse", "observacion",
]


def leer_log():
    if not ARCHIVO_LOG.exists():
        return {}
    with open(ARCHIVO_LOG, newline="", encoding="utf-8-sig") as f:
        return {r["codigo"]: r for r in csv.DictReader(f, delimiter=";")}


def guardar_log(log):
    ARCHIVO_LOG.parent.mkdir(parents=True, exist_ok=True)
    try:
        _escribir(ARCHIVO_LOG, log)
    except PermissionError:  # el CSV está abierto en Excel
        alterno = ARCHIVO_LOG.with_name(f"{ARCHIVO_LOG.stem}_{datetime.now():%H%M%S}.csv")
        _escribir(alterno, log)
        print(f"AVISO: el log estaba abierto; se guardó en {alterno}")


def _escribir(destino, log):
    with open(destino, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_LOG, delimiter=";", extrasaction="ignore")
        w.writeheader()
        for fila in log.values():
            w.writerow({c: fila.get(c, "") for c in CAMPOS_LOG})


# =====================================================================
# ESCANEO DE LA CARPETA DE ENTRADA
# =====================================================================
RE_CR = re.compile(r"^CR_(.+)$", re.I)
RE_AS = re.compile(r"^AS_(.+)$", re.I)


def escanear_entrada(base: Path):
    """{codigo: {"cartas": [Path], "zips": [Path], "origenes": {Path}}}"""
    casos = {}

    def caso(cod):
        return casos.setdefault(cod.upper(), {"cartas": [], "zips": [], "origenes": set()})

    def clasificar(f: Path, origen: Path):
        if f.name.startswith("~$"):
            return
        m_cr, m_as = RE_CR.match(f.stem), RE_AS.match(f.stem)
        if m_cr:
            c = caso(m_cr.group(1))
            c["cartas"].append(f)
            c["origenes"].add(origen)
        elif m_as and f.suffix.lower() == ".zip":
            c = caso(m_as.group(1))
            c["zips"].append(f)
            c["origenes"].add(origen)

    for p in sorted(base.iterdir()):
        if p.name.startswith(("_", ".")):
            continue
        if p.is_dir():
            m = RE_AS.match(p.name)
            if not m:
                continue
            caso(m.group(1))["origenes"].add(p)
            for f in p.rglob("*"):
                if f.is_file():
                    clasificar(f, p)
        elif p.is_file():
            clasificar(p, p)
    return casos


def mb(n_bytes):
    return n_bytes / (1024 * 1024)


# =====================================================================
# EXTRACCIÓN DE NOMBRE Y CORREO DE LA CARTA (misma lógica de Main.py)
# =====================================================================
def extraer_nombre_y_correo(ruta: Path, max_lineas=MAX_LINEAS_CARTA):
    ext = ruta.suffix.lower()
    if ext == ".pdf":
        import pdfplumber
        with pdfplumber.open(str(ruta)) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    elif ext == ".docx":
        from docx import Document
        texto = "\n".join(p.text for p in Document(str(ruta)).paragraphs)
    else:
        raise ValueError(f"formato no soportado ({ext}); use PDF o DOCX")

    lineas = [l.strip() for l in texto.split("\n") if l.strip()][:max_lineas]
    patron = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    nombre = correo = None
    for i, linea in enumerate(lineas):
        m = re.search(patron, linea)
        if m:
            correo = m.group()
            for j in range(i - 1, max(i - 5, -1), -1):
                if re.match(r"^[A-ZÁÉÍÓÚÑ ]{5,}$", lineas[j]):
                    nombre = lineas[j].title()
                    break
            break
    return nombre, correo


# =====================================================================
# VALIDACIÓN PREVIA
# =====================================================================
def validar_caso(codigo, d):
    """Devuelve (error|None, info). info = nombre, correo, adjuntos, peso."""
    if not d["cartas"]:
        return f"CR_{codigo}: no existe la carta", {}
    adjuntos = d["cartas"] + d["zips"]
    peso = mb(sum(p.stat().st_size for p in adjuntos))

    for p in adjuntos:
        if mb(p.stat().st_size) > MAX_ARCHIVO_MB:
            return f"{p.name} pesa {mb(p.stat().st_size):.2f} MB (supera {MAX_ARCHIVO_MB} MB)", {}
    if peso > MAX_TOTAL_MB:
        return f"{codigo}: anexos suman {peso:.2f} MB (supera {MAX_TOTAL_MB} MB)", {}

    try:
        nombre, correo = extraer_nombre_y_correo(d["cartas"][0])
    except Exception as e:
        return f"{d['cartas'][0].name}: ERROR al leer ({type(e).__name__}: {e})", {}
    if not correo:
        return f"{d['cartas'][0].name}: SIN CORREO", {}

    return None, {"nombre": nombre, "correo": correo, "adjuntos": adjuntos, "peso": peso}


def preguntar_corregir(errores):
    """Diálogo Corregido (Reintentar) / Cancelar. True = revalidar."""
    import tkinter as tk
    from tkinter import messagebox
    lista = errores[:25] + ([f"... y {len(errores) - 25} más"] if len(errores) > 25 else [])
    texto = ("Se encontraron los siguientes puntos para corregir:\n\n- " + "\n- ".join(lista) +
             "\n\nReintentar = ya corregí / Cancelar = salir")
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    r = messagebox.askretrycancel("VALIDACIONES FALLIDAS", texto, parent=root)
    root.destroy()
    return r


# =====================================================================
# OUTLOOK
# =====================================================================
def conectar():
    import win32com.client  # import tardío: --simular funciona sin Outlook
    app = win32com.client.Dispatch("Outlook.Application")
    return app, app.GetNamespace("MAPI")


def enviar_mail(app, destinatario, asunto, html, adjuntos):
    mail = app.CreateItem(0)
    mail.To = destinatario
    mail.Subject = asunto
    mail.HTMLBody = html
    mail.SentOnBehalfOfName = REMITENTE            # envía desde el buzón delegado
    mail.OriginatorDeliveryReportRequested = True   # SOLO acuse de entrega
    mail.ReadReceiptRequested = False
    for a in adjuntos:
        mail.Attachments.Add(str(a))
    mail.Send()


def mover_a_enviados(origenes):
    CARPETA_ENVIADOS.mkdir(parents=True, exist_ok=True)
    for o in origenes:
        destino = CARPETA_ENVIADOS / o.name
        if destino.exists():
            destino = CARPETA_ENVIADOS / f"{o.stem}_{datetime.now():%Y%m%d%H%M%S}{o.suffix}"
        shutil.move(str(o), str(destino))


# =====================================================================
# ENVIAR
# =====================================================================
def proceso_enviar(simular):
    if not CARPETA_ENTRADA.exists():
        sys.exit(f"No existe la carpeta de entrada: {CARPETA_ENTRADA}")
    log = leer_log()

    # --- validación con revalidación hasta que todo esté bien o se cancele ---
    while True:
        casos = {c: d for c, d in escanear_entrada(CARPETA_ENTRADA).items()
                 if log.get(c, {}).get("estado_envio") != "ENVIADO"}
        if not casos:
            print("No hay casos pendientes en la carpeta de entrada.")
            return
        validos, errores = {}, {}
        for cod, d in sorted(casos.items()):
            err, info = validar_caso(cod, d)
            if err:
                errores[cod] = err
            else:
                validos[cod] = info
        if not errores:
            break
        if simular:
            for e in errores.values():
                print("[ERROR]", e)
            break
        if not preguntar_corregir(list(errores.values())):
            for cod, e in errores.items():
                log[cod] = {**log.get(cod, {}), "codigo": cod, "estado_envio": "NO_ENVIADO", "observacion": e}
            guardar_log(log)
            print("Envío cancelado por el usuario.")
            return

    app = None if simular else conectar()[0]
    for cod, info in sorted(validos.items()):
        d = casos[cod]
        asunto = ASUNTO.format(codigo=cod)
        fila = {
            "codigo": cod, "destinatario": info["correo"], "asunto": asunto,
            "carta": " | ".join(p.name for p in d["cartas"]),
            "zip": " | ".join(p.name for p in d["zips"]),
            "peso_mb": f"{info['peso']:.2f}", "estado_acuse": "", "observacion": "",
        }
        if simular:
            fila["estado_envio"] = "SIMULADO"
            print(f"[{cod}] SIMULADO → {info['correo']} ({info['peso']:.2f} MB)")
        else:
            html = CUERPO_HTML.replace("{nombre}", info["nombre"] or "Estimado/a").replace("{codigo}", cod)
            try:
                enviar_mail(app, info["correo"], asunto, html, info["adjuntos"])
                fila.update(estado_envio="ENVIADO", estado_acuse="PENDIENTE",
                            fecha_envio=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                if MOVER_ENVIADOS:
                    mover_a_enviados(d["origenes"])
                print(f"[{cod}] ENVIADO → {info['correo']}")
                time.sleep(PAUSA_ENTRE_ENVIOS_SEG)
            except Exception as e:
                fila.update(estado_envio="ERROR", observacion=f"{type(e).__name__}: {e}")
                print(f"[{cod}] ERROR {e}")
            log[cod] = {**log.get(cod, {}), **fila}
            guardar_log(log)   # guarda tras cada caso por si se corta a mitad


# =====================================================================
# ACUSES DE ENTREGA
# =====================================================================
def carpetas_de(ns, tipos):
    """Carpetas por defecto de tu buzón y del buzón delegado."""
    res = []
    for tipo in tipos:
        try:
            res.append(ns.GetDefaultFolder(tipo))
        except Exception as e:
            print(f"AVISO: no se pudo abrir carpeta {tipo} propia: {e}")
        try:
            rec = ns.CreateRecipient(REMITENTE)
            rec.Resolve()
            if rec.Resolved:
                res.append(ns.GetSharedDefaultFolder(rec, tipo))
        except Exception as e:
            print(f"AVISO: no se pudo abrir carpeta {tipo} del buzón delegado: {e}")
    return res


def recorrer(carpeta):
    yield carpeta
    if INCLUIR_SUBCARPETAS:
        try:
            for sub in carpeta.Folders:
                yield from recorrer(sub)
        except Exception:
            pass


def naive(dt):
    return dt.replace(tzinfo=None) if getattr(dt, "tzinfo", None) else dt


def proceso_acuses():
    log = leer_log()
    pendientes = [r for r in log.values()
                  if r.get("estado_envio") == "ENVIADO"
                  and r.get("estado_acuse") not in ("ENCONTRADO", "REBOTE")]
    if not pendientes:
        print("No hay acuses pendientes.")
        return
    _, ns = conectar()
    CARPETA_ACUSES.mkdir(parents=True, exist_ok=True)

    desde = min(datetime.strptime(r["fecha_envio"], "%Y-%m-%d %H:%M:%S") for r in pendientes) - timedelta(minutes=5)

    # Una sola pasada: junta los reportes (entrega / rebote) recientes
    reportes = []
    for raiz in carpetas_de(ns, CARPETAS_ACUSE):
        for carpeta in recorrer(raiz):
            try:
                items = carpeta.Items
                items.Sort("[ReceivedTime]", True)
                for it in items:
                    try:
                        recibido = naive(it.ReceivedTime)
                        if recibido < desde:
                            break
                        if it.Class != 46:      # 46 = informe (acuse / NDR)
                            continue
                        mc = it.MessageClass.upper()
                        if mc.startswith("REPORT.IPM.NOTE.DR"):
                            tipo = "ENTREGA"
                        elif ".NDR" in mc:
                            tipo = "REBOTE"
                        else:
                            continue
                        reportes.append({
                            "item": it, "tipo": tipo, "recibido": recibido,
                            "asunto": (it.Subject or "").lower(),
                            "cuerpo": (it.Body or "")[:6000].lower(),
                        })
                    except Exception:
                        continue
            except Exception:
                continue
    print(f"Reportes (entrega/rebote) encontrados: {len(reportes)}")

    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for r in pendientes:
        cod = r["codigo"].lower()
        asunto = (r.get("asunto") or "").lower()
        envio = datetime.strptime(r["fecha_envio"], "%Y-%m-%d %H:%M:%S") - timedelta(minutes=1)

        hallados = [x for x in reportes
                    if x["recibido"] >= envio and
                    ((asunto and asunto in x["asunto"]) or cod in x["asunto"] or cod in x["cuerpo"])]

        r["ultima_busqueda_acuse"] = ahora
        if not hallados:
            r["estado_acuse"] = "NO_ENCONTRADO"
            print(f"[{r['codigo']}] acuse NO encontrado")
            continue

        rutas = []
        for i, x in enumerate(hallados, 1):
            nombre = re.sub(r"[^\w\-]", "_", r["codigo"])
            ruta = CARPETA_ACUSES / f"ACUSE_{x['tipo']}_{nombre}_{i}_{x['recibido']:%Y%m%d_%H%M%S}.msg"
            try:
                x["item"].SaveAs(str(ruta), 3)   # 3 = olMSG
                rutas.append(str(ruta))
            except Exception as e:
                rutas.append(f"(no se pudo guardar: {e})")

        r["estado_acuse"] = "REBOTE" if any(x["tipo"] == "REBOTE" for x in hallados) else "ENCONTRADO"
        r["fecha_acuse"] = min(x["recibido"] for x in hallados).strftime("%Y-%m-%d %H:%M:%S")
        r["ruta_acuse"] = " | ".join(rutas)
        r["observacion"] = f"{len(hallados)} acuse(s)"
        print(f"[{r['codigo']}] {r['estado_acuse']} ({len(hallados)})")

    guardar_log(log)


# =====================================================================
# MAIN
# =====================================================================
def main():
    ap = argparse.ArgumentParser(description="Respuestas por Outlook con control de acuses de entrega")
    ap.add_argument("modo", choices=["enviar", "acuses", "todo"])
    ap.add_argument("--simular", action="store_true", help="no usa Outlook ni envía; solo valida y muestra")
    a = ap.parse_args()

    if a.modo in ("enviar", "todo"):
        proceso_enviar(a.simular)
    if a.modo in ("acuses", "todo") and not a.simular:
        proceso_acuses()


if __name__ == "__main__":
    main()
