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


def leer_retencion(pdf):

    base0=0
    base15=0
    porcentaje=0

    with pdfplumber.open(pdf) as pdf_file:

        for page in pdf_file.pages:

            tablas=page.extract_tables()

            for tabla in tablas:

                for fila in tabla:

                    if not fila:
                        continue

                    texto=" ".join([str(x) for x in fila if x])

                    base_match=re.search(r"\d+\.\d+",texto)

                    if base_match:

                        base=float(base_match.group())

                        if "RENTA" in texto.upper():

                            base0+=base

                        if "IVA" in texto.upper():

                            base15+=base

                        porc=re.search(r"\b(10|2|100)\b",texto)

                        if porc:
                            porcentaje=int(porc.group())

    return base0,base15,porcentaje


def procesar_pdf(pdf):

    texto=extraer_texto(pdf)

    fecha=buscar(texto,r"Fecha[:\s]*([0-9/\-]+)")

    if fecha=="":
        fecha=datetime.today().strftime("%Y-%m-%d")

    empresa=extraer_empresa(texto)

    factura=buscar(texto,r"No\.?\s*([0-9\-]+)")

    ruc=extraer_ruc(texto)

    autorizacion=buscar(texto,r"Autorizaci[oó]n[:\s]*([0-9]{10,})")

    base0,base15,porcentaje=leer_retencion(pdf)

    propina=0
    iva=0

    if base15>0:
        iva=round(base15*0.15,2)

    total=base0+base15+propina+iva

    rete10=0
    rete2=0
    rete100=0

    if porcentaje==10:
        rete10=round(total*0.10,2)

    if porcentaje==2:
        rete2=round(total*0.02,2)

    if porcentaje==100:
        rete100=round(iva*1,2)

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

    df["FECHA"]=pd.to_datetime(df["FECHA"],errors="coerce")

    st.dataframe(df)

    output=BytesIO()

    with pd.ExcelWriter(output,engine="xlsxwriter") as writer:

        df.to_excel(writer,index=False,sheet_name="RETENCIONES")

        workbook=writer.book
        worksheet=writer.sheets["RETENCIONES"]

        header_format=workbook.add_format({
            "bold":True,
            "align":"center",
            "border":1,
            "bg_color":"#FFFF00"
        })

        for col,col_name in enumerate(columnas):
            worksheet.write(0,col,col_name,header_format)

        filas=len(df)+1

        total_format=workbook.add_format({
            "bold":True,
            "border":1,
            "bg_color":"#FFFF00"
        })

        worksheet.write(filas,0,"TOTAL",total_format)

        for i in range(8,20):

            letra=chr(65+i)

            formula=f"=SUM({letra}2:{letra}{filas})"

            worksheet.write_formula(filas,i,formula,total_format)

        worksheet.set_column(0,20,18)

        # --------- NUEVO: TABLAS POR MES ---------

        fila_inicio=filas+4

        meses=df.groupby(df["FECHA"].dt.to_period("M"))

        for mes,datos_mes in meses:

            worksheet.write(fila_inicio,0,f"MES {mes}",header_format)

            datos_mes.to_excel(
                writer,
                sheet_name="RETENCIONES",
                startrow=fila_inicio+1,
                index=False
            )

            fila_inicio+=len(datos_mes)+5

    output.seek(0)

    st.download_button(
        "Descargar Excel",
        data=output,
        file_name="retenciones_sri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
