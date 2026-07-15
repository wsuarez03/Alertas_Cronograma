# alerta_plan_trabajo.py
import io, os, requests, smtplib
from datetime import datetime
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

EXCEL_URL="https://valserindustriales-my.sharepoint.com/personal/tecnicodeservicios_valserindustriales_com/_layouts/15/download.aspx?UniqueId=4bd45590%2D7a3a%2D4e97%2D8947%2D91feabaa0373"
HOJA=0
DIAS_AVISO=45
SMTP_SERVER="smtp.gmail.com"
SMTP_PORT=587
SMTP_USER=os.getenv("SMTP_USER")
SMTP_PASS=os.getenv("SMTP_PASS")
CORREOS_DESTINO=["correo@empresa.com"]

MESES={
1:("ENE",5,6),2:("FEB",7,8),3:("MAR",9,10),4:("ABR",11,12),
5:("MAY",13,14),6:("JUN",15,16),7:("JUL",17,18),8:("AGO",19,20),
9:("SEP",21,22),10:("OCT",23,24),11:("NOV",25,26),12:("DIC",27,28)
}
COL_ACTIVIDAD=2
COL_PERIODICIDAD=4
COL_RESPONSABLE=29
COL_ALERTA=34
FILA_INICIO=11

def descargar():
    return requests.get(EXCEL_URL).content

def html(df):
    if df.empty:return "<p>No hay registros.</p>"
    return df.to_html(index=False)

def enviar(htmltxt):
    m=MIMEMultipart("alternative")
    m["Subject"]="Alerta Plan de Trabajo SST"
    m["From"]=SMTP_USER
    m["To"]=", ".join(CORREOS_DESTINO)
    m.attach(MIMEText(htmltxt,"html"))
    s=smtplib.SMTP(SMTP_SERVER,SMTP_PORT)
    s.starttls(); s.login(SMTP_USER,SMTP_PASS)
    s.sendmail(SMTP_USER,CORREOS_DESTINO,m.as_string()); s.quit()

def main():
    df=pd.read_excel(io.BytesIO(descargar()),sheet_name=HOJA,header=None)
    hoy=datetime.now().date()
    criticas=[]; pendientes=[]; proximas=[]
    for i in range(FILA_INICIO,len(df)):
        if str(df.iat[i,COL_ALERTA]).strip().upper()!="X": continue
        act=df.iat[i,COL_ACTIVIDAD]
        resp=df.iat[i,COL_RESPONSABLE]
        per=df.iat[i,COL_PERIODICIDAD]
        for mes,(nom,p,e) in MESES.items():
            if str(df.iat[i,p]).strip().upper()!="P": continue
            if str(df.iat[i,e]).strip()!="": continue
            fecha=datetime(hoy.year,mes,1).date()
            dias=(fecha-hoy).days
            reg={"Actividad":act,"Responsable":resp,"Periodicidad":per,"Mes":nom}
            if dias< -30:
                reg["Estado"]=f"Crítica ({abs(dias)} días atraso)"; criticas.append(reg)
            elif dias<0:
                reg["Estado"]=f"Pendiente ({abs(dias)} días atraso)"; pendientes.append(reg)
            elif dias<=DIAS_AVISO:
                reg["Estado"]=f"Próxima ({dias} días)"; proximas.append(reg)
    if not(criticas or pendientes or proximas):
        print("Sin alertas"); return
    htmlmsg=f"<h2>Plan Trabajo SST</h2><h3>🔴 Críticas</h3>{html(pd.DataFrame(criticas))}<h3>🟠 Pendientes</h3>{html(pd.DataFrame(pendientes))}<h3>🟡 Próximas</h3>{html(pd.DataFrame(proximas))}"
    enviar(htmlmsg)
if __name__=="__main__":
    main()
