# alerta_plan_trabajo.py

import io
import os
import smtplib
import requests
import traceback
import pandas as pd

from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ==========================
# CONFIGURACIÓN
# ==========================

EXCEL_URL = "https://valserindustriales-my.sharepoint.com/:x:/p/tecnicodeservicios/IQCQVdRLOnqXTolHkf6rqgNzAU0sWBihplDqt4zCfrm16IA?e=GEgKMd&download=1"

HOJA = 0
DIAS_AVISO = 45

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

CORREOS_DESTINO = [
    "asesorcomercial@valserindustriales.com",
    "tecnicodeservicios@valserindustriales.com"
]

MESES = {
    1: ("ENE", 5, 6),
    2: ("FEB", 7, 8),
    3: ("MAR", 9, 10),
    4: ("ABR", 11, 12),
    5: ("MAY", 13, 14),
    6: ("JUN", 15, 16),
    7: ("JUL", 17, 18),
    8: ("AGO", 19, 20),
    9: ("SEP", 21, 22),
    10: ("OCT", 23, 24),
    11: ("NOV", 25, 26),
    12: ("DIC", 27, 28)
}

COL_ACTIVIDAD = 2
COL_PERIODICIDAD = 4
COL_RESPONSABLE = 29
COL_ALERTA = 34
FILA_INICIO = 11


# ==========================
# DESCARGAR EXCEL
# ==========================

def descargar():

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/138.0.0.0 Safari/537.36"
        )
    }

    r = requests.get(
        EXCEL_URL,
        headers=headers,
        allow_redirects=True,
        timeout=60
    )

    print("Status:", r.status_code)
    print("URL final:", r.url)
    print("Content-Type:", r.headers.get("Content-Type"))
    print("Primeros 300 bytes:")
    print(r.text[:300])

    r.raise_for_status()

    return r.content
# ==========================
# HTML
# ==========================

def tabla_html(df):

    if df.empty:
        return "<p>No hay registros.</p>"

    return df.to_html(index=False, border=1)


# ==========================
# ENVIAR CORREO
# ==========================

def enviar(html):

    print("Preparando correo...")

    if not SMTP_USER:
        raise Exception("SMTP_USER no existe.")

    if not SMTP_PASS:
        raise Exception("SMTP_PASS no existe.")

    mensaje = MIMEMultipart("alternative")

    mensaje["Subject"] = "Alerta Plan de Trabajo SST"
    mensaje["From"] = SMTP_USER
    mensaje["To"] = ", ".join(CORREOS_DESTINO)

    mensaje.attach(MIMEText(html, "html"))

    servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)

    servidor.starttls()

    print("Conectado a Gmail...")

    servidor.login(SMTP_USER, SMTP_PASS)

    print("Login correcto")

    servidor.sendmail(
        SMTP_USER,
        CORREOS_DESTINO,
        mensaje.as_string()
    )

    servidor.quit()

    print("Correo enviado correctamente.")


# ==========================
# PROGRAMA PRINCIPAL
# ==========================

def main():

    contenido = descargar()

    print("Leyendo Excel...")

    df = pd.read_excel(
        io.BytesIO(contenido),
        sheet_name=HOJA,
        header=None,
        engine="openpyxl"
    )

    for i in range(FILA_INICIO, len(df)):

    print("-" * 60)

    print("Fila:", i)

    print("Actividad :", df.iat[i, COL_ACTIVIDAD])
    print("Alerta    :", repr(df.iat[i, COL_ALERTA]))
    print("Resp      :", repr(df.iat[i, COL_RESPONSABLE]))
    print("Periodic. :", repr(df.iat[i, COL_PERIODICIDAD]))

    for mes, (nom, p, e) in MESES.items():

        print(
            nom,
            "Plan:", repr(df.iat[i, p]),
            "Ejecutado:", repr(df.iat[i, e])
        )
    hoy = datetime.now().date()

    criticas = []
    pendientes = []
    proximas = []

    for i in range(FILA_INICIO, len(df)):

        if str(df.iat[i, COL_ALERTA]).strip().upper() != "X":
            continue

        for i in range(FILA_INICIO, len(df)):

    print(
        i,
        df.iat[i, COL_ACTIVIDAD],
        "| ALERTA:",
        repr(df.iat[i, COL_ALERTA]),
    )

    if str(df.iat[i, COL_ALERTA]).strip().upper() != "X":
        continue

    print("Fila con alerta:", i)

    for mes, (nom, p, e) in MESES.items():

        print(
            nom,
            "Plan:",
            repr(df.iat[i, p]),
            "Ejecutado:",
            repr(df.iat[i, e]),
        )
            fecha = datetime(hoy.year, mes, 1).date()

            dias = (fecha - hoy).days

            registro = {
                "Actividad": actividad,
                "Responsable": responsable,
                "Periodicidad": periodicidad,
                "Mes": nombre
            }

            if dias < -30:
                registro["Estado"] = f"Crítica ({abs(dias)} días)"
                criticas.append(registro)

            elif dias < 0:
                registro["Estado"] = f"Pendiente ({abs(dias)} días)"
                pendientes.append(registro)

            elif dias <= DIAS_AVISO:
                registro["Estado"] = f"Próxima ({dias} días)"
                proximas.append(registro)

    print("Críticas:", len(criticas))
    print("Pendientes:", len(pendientes))
    print("Próximas:", len(proximas))

    if not (criticas or pendientes or proximas):
        print("No existen actividades para notificar.")
        return

    html = f"""
    <h2>Plan de Trabajo SST</h2>

    <h3>🔴 Críticas</h3>
    {tabla_html(pd.DataFrame(criticas))}

    <h3>🟠 Pendientes</h3>
    {tabla_html(pd.DataFrame(pendientes))}

    <h3>🟡 Próximas</h3>
    {tabla_html(pd.DataFrame(proximas))}
    """

    enviar(html)


# ==========================
# INICIO
# ==========================

if __name__ == "__main__":

    try:

        print("====================================")
        print("INICIANDO ALERTA PLAN DE TRABAJO SST")
        print("====================================")

        main()

        print("Proceso terminado correctamente.")

    except Exception as e:

        print("ERROR:")
        print(e)
        traceback.print_exc()

        raise
