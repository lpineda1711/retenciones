import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Retenciones SRI", layout="wide")

st.title("Generador de Retenciones SRI")

uploaded_files = st.file_uploader(
    "Subir comprobantes PDF",
    type="pdf",
    accept_multiple_files=True
)

columnas = [
"FECHA","IFIS","N FACTURA","RUC","DOC IFIS","AUTORIZACION",
"NO OBJETO","EXCENTO IVA","BASE 0%","BASE 15%","PROPINA","IVA",
"TOTAL","N° RETENCION","0% R.FTE","RETE 10%","RETE 100%",
"2% R.FTE","TOTAL RETENCION","valor retenido"
]


def extraer_texto(pdf):

    texto=""

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            t=page.extract_text()
            if t:
                texto+=t+"\n"

    return texto


def extraer_ruc(texto):

    ruc=re.search(r'RUC[:\s]*([0-9]{13})',texto)

    if ruc:
        return ruc.group(1)

    alt=re.search(r'\b[0-9]{13}\b',texto)

    if alt:
        return alt.group(0)

    return ""


def buscar(texto,patron):

    m=re.search(patron,texto,re.IGNORECASE)

    if m:
        return m.group(1)

    return ""


def extraer_empresa(texto):

    lineas=texto.split("\n")

    for l in lineas[:10]:

        if "S.A" in l.upper() or "CIA" in l.upper() or "LTDA" in l.upper():
            return l.strip()

    return ""


def leer_tabla_retenciones(texto):

    base0=0
    base15=0

    rete10=0
    rete2=0
    rete100=0

    patron=r"([0-9]+\.[0-9]+)\s+(Impuesto a la Renta|IVA)\s+([0-9]+)"

    matches=re.findall(patron,texto,re.IGNORECASE)

    for base,impuesto,porcentaje in matches:

        base=float(base)
        porcentaje=float(porcentaje)

        if "RENTA" in impuesto.upper():

            base0+=base

        if "IVA" in impuesto.upper():

            base15+=base

        valor_retencion=round(base*(porcentaje/100),2)

        if porcentaje==10:
            rete10+=valor_retencion

        elif porcentaje==2:
            rete2+=valor_retencion

        elif porcentaje==100:
            rete100+=valor_retencion

    return base0,base15,rete10,rete2,rete100


def procesar_pdf(pdf):

    texto=extraer_texto(pdf)

    fecha=buscar(texto,r"Fecha[:\s]*([0-9/\-]+)")

    if fecha=="":
        fecha=datetime.today().strftime("%Y-%m-%d")

    empresa=extraer_empresa(texto)

    factura=buscar(texto,r"No\.?\s*([0-9\-]+)")

    ruc=extraer_ruc(texto)

    autorizacion=buscar(texto,r"Autorizaci[oó]n[:\s]*([0-9]{10,})")

    base0,base15,rete10,rete2,rete100=leer_tabla_retenciones(texto)

    propina=0
    iva=0

    if base15>0:
        iva=round(base15*0.15,2)

    total=base0+base15+propina+iva

    total_retencion=rete10+rete2+rete100

    fila={
        "FECHA":fecha,
        "IFIS":empresa,
        "N FACTURA":factura,
        "RUC":ruc,
        "DOC IFIS":"",
        "AUTORIZACION":autorizacion,
        "NO OBJETO":"",
        "EXCENTO IVA":"",
        "BASE 0%":base0,
        "BASE 15%":base15,
        "PROPINA":propina,
        "IVA":iva,
        "TOTAL":total,
        "N° RETENCION":"",
        "0% R.FTE":"",
        "RETE 10%":rete10,
        "RETE 100%":rete100,
        "2% R.FTE":rete2,
        "TOTAL RETENCION":total_retencion,
        "valor retenido":total_retencion
    }

    return fila


if uploaded_files:

    datos=[]

    for file in uploaded_files:
        datos.append(procesar_pdf(file))

    df=pd.DataFrame(datos,columns=columnas)

    st.dataframe(df)

    output=BytesIO()

    with pd.ExcelWriter(output,engine="xlsxwriter") as writer:

        df.to_excel(writer,index=False,sheet_name="RETENCIONES")

    output.seek(0)

    st.download_button(
        "Descargar Excel",
        data=output,
        file_name="retenciones_sri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
