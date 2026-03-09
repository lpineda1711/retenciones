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


def leer_tabla_retencion(pdf):

    base0=""
    base15=""

    rete10=""
    rete2=""
    rete100=""
    valor_retenido=""

    with pdfplumber.open(pdf) as pdf_file:

        for page in pdf_file.pages:

            tablas=page.extract_tables()

            for tabla in tablas:

                for fila in tabla:

                    if not fila:
                        continue

                    texto=" ".join([str(x) for x in fila if x])

                    numeros=re.findall(r"\d+\.\d+",texto)

                    # detectar porcentaje correctamente
                    porc=re.search(r"\b(1|2|8|10|20|30|70|100)\b",texto)

                    if numeros:

                        base=float(numeros[0])

                        if "RENTA" in texto.upper():
                            base0=base

                        if "IVA" in texto.upper() and base>0:
                            base15=base

                    if porc and len(numeros)>=2:

                        porcentaje=int(porc.group())
                        valor=float(numeros[-1])

                        valor_retenido=valor

                        if porcentaje==10:
                            rete10=valor

                        elif porcentaje==2:
                            rete2=valor

                        elif porcentaje==100:
                            rete100=valor

    return base0,base15,rete10,rete2,rete100,valor_retenido


def procesar_pdf(pdf):

    texto=extraer_texto(pdf)

    fecha=buscar(texto,r"Fecha[:\s]*([0-9/\-]+)")

    if fecha=="":
        fecha=datetime.today().strftime("%Y-%m-%d")

    empresa=extraer_empresa(texto)

    factura=buscar(texto,r"No\.?\s*([0-9\-]+)")

    ruc=extraer_ruc(texto)

    autorizacion=buscar(texto,r"Autorizaci[oó]n[:\s]*([0-9]{10,})")

    base0,base15,rete10,rete2,rete100,valor_retenido=leer_tabla_retencion(pdf)

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
        "PROPINA":"",
        "IVA":"",
        "TOTAL":"",
        "N° RETENCION":"",
        "0% R.FTE":"",
        "RETE 10%":rete10,
        "RETE 100%":rete100,
        "2% R.FTE":rete2,
        "TOTAL RETENCION":valor_retenido,
        "valor retenido":valor_retenido
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

        workbook=writer.book
        worksheet=workbook.add_worksheet("RETENCIONES")
        writer.sheets["RETENCIONES"]=worksheet

        header_format=workbook.add_format({
            "bold":True,
            "align":"center",
            "border":1,
            "bg_color":"#FFFF00"
        })

        total_format=workbook.add_format({
            "bold":True,
            "border":1,
            "bg_color":"#FFFF00"
        })

        fila_excel=0

        meses=df.groupby(df["FECHA"].dt.to_period("M"))

        for mes,datos_mes in meses:

            worksheet.write(fila_excel,0,f"MES {mes}",header_format)

            fila_excel+=1

            for col,col_name in enumerate(columnas):
                worksheet.write(fila_excel,col,col_name,header_format)

            fila_excel+=1

            inicio_datos=fila_excel

            for i,row in datos_mes.iterrows():

                for col,col_name in enumerate(columnas):
                    worksheet.write(fila_excel,col,row[col_name])

                fila_excel+=1

            # FILA TOTAL AMARILLA COMPLETA
            for col in range(len(columnas)):
                worksheet.write(fila_excel,col,"",total_format)

            worksheet.write(fila_excel,0,"TOTAL",total_format)

            for col in range(8,20):

                letra=chr(65+col)

                formula=f"=SUM({letra}{inicio_datos+1}:{letra}{fila_excel})"

                worksheet.write_formula(fila_excel,col,formula,total_format)

            fila_excel+=3

        worksheet.set_column(0,20,18)

    output.seek(0)

    st.download_button(
        "Descargar Excel",
        data=output,
        file_name="retenciones_sri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
