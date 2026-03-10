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

    base0=0
    base15=0
    rete10=0
    rete2=0
    rete100=0

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:

            tablas=page.extract_tables()

            for tabla in tablas:
                for fila in tabla:

                    if not fila:
                        continue

                    texto=" ".join([str(x) for x in fila if x]).upper()

                    numeros=re.findall(r"\d+[.,]\d+",texto)

                    porcentaje=re.search(r"\b(100|10|2)\b",texto)

                    if len(numeros)>=2 and porcentaje:

                        base=float(numeros[0].replace(",","."))
                        valor=float(numeros[-1].replace(",","."))

                        porc=int(porcentaje.group())

                        if porc==10:
                            rete10=valor

                        elif porc==2:
                            rete2=valor

                        elif porc==100:
                            rete100=valor

                        if "RENTA" in texto:
                            base0=base

                        if "IVA" in texto:
                            base15=base

    total_retencion = rete10 + rete2 + rete100

    return base0,base15,rete10,rete2,rete100,total_retencion


def procesar_pdf(pdf):

    texto=extraer_texto(pdf)

    fecha=buscar(texto,r"Fecha[:\s]*([0-9/\-]+)")
    if fecha=="":
        fecha=datetime.today().strftime("%d/%m/%Y")

    empresa=extraer_empresa(texto)
    factura=buscar(texto,r"No\.?\s*([0-9\-]+)")
    ruc=extraer_ruc(texto)
    autorizacion=buscar(texto,r"Autorizaci[oó]n[:\s]*([0-9]{10,})")

    base0,base15,rete10,rete2,rete100,total_retencion = leer_tabla_retencion(pdf)

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
        "TOTAL RETENCION":total_retencion,
        "valor retenido":total_retencion
    }

    return fila


if uploaded_files:

    datos=[]

    for file in uploaded_files:
        fila=procesar_pdf(file)
        datos.append(fila)

    df=pd.DataFrame(datos,columns=columnas)

    df["FECHA"]=pd.to_datetime(df["FECHA"],dayfirst=True,errors="coerce")

    df["TOTAL RETENCION"]=df["TOTAL RETENCION"].fillna(0)
    df["valor retenido"]=df["valor retenido"].fillna(0)
    df["RETE 10%"]=df["RETE 10%"].fillna(0)
    df["RETE 100%"]=df["RETE 100%"].fillna(0)
    df["2% R.FTE"]=df["2% R.FTE"].fillna(0)

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

        date_format=workbook.add_format({'num_format':'dd/mm/yyyy'})

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

                    if col_name=="FECHA":
                        worksheet.write_datetime(
                            fila_excel,col,row[col_name],date_format
                        )

                    elif col_name=="TOTAL":
                        formula=f"=SUM(I{fila_excel+1}:L{fila_excel+1})"
                        worksheet.write_formula(fila_excel,col,formula)

                    elif col_name=="TOTAL RETENCION":
                        formula=f"=SUM(P{fila_excel+1}:R{fila_excel+1})"
                        worksheet.write_formula(fila_excel,col,formula)

                    elif col_name=="valor retenido":
                        formula=f"=S{fila_excel+1}"
                        worksheet.write_formula(fila_excel,col,formula)

                    else:
                        worksheet.write(fila_excel,col,row[col_name])

                fila_excel+=1

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
