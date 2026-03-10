import os
import pdfplumber
import streamlit as st
import pandas as pd
from datetime import datetime
from collections import defaultdict

st.title("Procesador de Retenciones")

carpeta_pdfs = "pdfs"

def extraer_datos(pdf_path):
    datos = {
        "fecha": "",
        "base0": 0,
        "base15": 0,
        "propina": 0,
        "iva": 0,
        "porcentaje": 0,
        "base_retencion": 0
    }

    with pdfplumber.open(pdf_path) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text()

    lineas = texto.split("\n")

    for l in lineas:

        if "Fecha de Emisión" in l:
            try:
                fecha = l.split()[-1]
                datos["fecha"] = datetime.strptime(fecha,"%d/%m/%Y").date()
            except:
                pass

        if "Base Imponible para la Retención" in l:
            try:
                datos["base_retencion"] = float(l.split()[-1])
            except:
                pass

        if "Porcentaje Retención" in l:
            try:
                datos["porcentaje"] = float(l.split()[-1])
            except:
                datos["porcentaje"] = 0

        if "IVA" in l:
            try:
                datos["iva"] = float(l.split()[-1])
            except:
                pass

    if datos["iva"] > 0:
        datos["base15"] = datos["base_retencion"]
    else:
        datos["base0"] = datos["base_retencion"]

    return datos


datos_por_mes = defaultdict(list)

if os.path.exists(carpeta_pdfs):

    for archivo in os.listdir(carpeta_pdfs):

        if archivo.endswith(".pdf"):

            ruta = os.path.join(carpeta_pdfs, archivo)

            datos = extraer_datos(ruta)

            if datos["fecha"] != "":
                mes = datos["fecha"].strftime("%B")
                datos_por_mes[mes].append(datos)

else:
    st.error("La carpeta 'pdfs' no existe")


columnas = [
"FECHA","BASE 0%","BASE 15%","PROPINA","IVA","TOTAL",
"RETE 100%","RETE 10%","RETE 2%","TOTAL RETENCION"
]

for mes, registros in datos_por_mes.items():

    st.markdown(f"### 🟨 {mes.upper()}")

    filas = []

    for r in registros:

        base0 = r["base0"]
        base15 = r["base15"]
        propina = r["propina"]
        iva = r["iva"]

        porcentaje = r["porcentaje"]
        base_ret = r["base_retencion"]

        rete100 = 0
        rete10 = 0
        rete2 = 0

        if porcentaje == 100:
            rete100 = base_ret * 1

        elif porcentaje == 10:
            rete10 = base_ret * 0.10

        elif porcentaje == 2:
            rete2 = base_ret * 0.02

        else:
            rete100 = 0
            rete10 = 0
            rete2 = 0

        total_ret = rete100 + rete10 + rete2

        filas.append([
            r["fecha"],
            base0,
            base15,
            propina,
            iva,
            "", 
            rete100,
            rete10,
            rete2,
            total_ret
        ])

    df = pd.DataFrame(filas, columns=columnas)

    for i in range(len(df)):
        df.loc[i,"TOTAL"] = f"=SUMA(B{i+2}:E{i+2})"

    fila_final = [
        "TOTAL",
        "=SUMA(B2:B100)",
        "=SUMA(C2:C100)",
        "=SUMA(D2:D100)",
        "=SUMA(E2:E100)",
        "=SUMA(F2:F100)",
        "=SUMA(G2:G100)",
        "=SUMA(H2:H100)",
        "=SUMA(I2:I100)",
        "=SUMA(J2:J100)"
    ]

    df.loc[len(df)] = fila_final

    st.dataframe(df)
