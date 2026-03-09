import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill

carpeta_pdfs = "pdfs"
archivo_excel = "retenciones.xlsx"

datos_por_mes = {}

def extraer_valor(texto, patron):
    match = re.search(patron, texto)
    return match.group(1) if match else ""

for archivo in os.listdir(carpeta_pdfs):

    if not archivo.endswith(".pdf"):
        continue

    ruta = os.path.join(carpeta_pdfs, archivo)

    with pdfplumber.open(ruta) as pdf:
        texto = ""
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"

    fecha = extraer_valor(texto, r"Fecha de Emisión\s*([\d/:-]+)")
    base_ret = extraer_valor(texto, r"Base Imponible para la Retención\s*([\d.]+)")
    porcentaje = extraer_valor(texto, r"Porcentaje Retención\s*([\d.]+)")
    impuesto = extraer_valor(texto, r"Impuesto\s*([A-Za-z ]+)")

    base_ret = float(base_ret) if base_ret else 0
    porcentaje = float(porcentaje) if porcentaje else 0

    try:
        fecha_obj = datetime.strptime(fecha.split()[0], "%d/%m/%Y")
        fecha = fecha_obj.strftime("%Y-%m-%d")
        mes = fecha_obj.strftime("%B").upper()
    except:
        mes = "SIN_MES"

    base0 = 0
    base15 = 0

    if "RENTA" in impuesto.upper():
        base0 = base_ret
    elif "IVA" in impuesto.upper():
        base15 = base_ret

    propina = 0
    iva = base15 * 0.15

    if mes not in datos_por_mes:
        datos_por_mes[mes] = []

    datos_por_mes[mes].append({
        "fecha": fecha,
        "base0": base0,
        "base15": base15,
        "propina": propina,
        "iva": iva,
        "base_ret": base_ret,
        "porcentaje": porcentaje
    })

wb = Workbook()
ws = wb.active
ws.title = "RETENCIONES"

amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

fila = 1

for mes, registros in datos_por_mes.items():

    ws.cell(row=fila, column=1, value=mes)
    ws.cell(row=fila, column=1).fill = amarillo
    fila += 1

    titulos = [
        "FECHA",
        "BASE 0%",
        "BASE 15%",
        "PROPINA",
        "IVA",
        "TOTAL",
        "RETE 2%",
        "RETE 10%",
        "RETE 100%",
        "TOTAL RETENCION"
    ]

    for col, titulo in enumerate(titulos, 1):
        celda = ws.cell(row=fila, column=col, value=titulo)
        celda.fill = amarillo

    fila += 1
    inicio_tabla = fila

    for r in registros:

        ws.cell(row=fila, column=1, value=r["fecha"])
        ws.cell(row=fila, column=2, value=r["base0"])
        ws.cell(row=fila, column=3, value=r["base15"])
        ws.cell(row=fila, column=4, value=r["propina"])
        ws.cell(row=fila, column=5, value=r["iva"])

        ws.cell(row=fila, column=6,
            value=f"=SUM(B{fila}:E{fila})")

        base_ret = r["base_ret"]
        porcentaje = r["porcentaje"]

        rete2 = 0
        rete10 = 0
        rete100 = 0

        if porcentaje == 2:
            rete2 = base_ret * 0.02

        elif porcentaje == 10:
            rete10 = base_ret * 0.10

        elif porcentaje == 100:
            rete100 = base_ret * 1

        ws.cell(row=fila, column=7, value=rete2)
        ws.cell(row=fila, column=8, value=rete10)
        ws.cell(row=fila, column=9, value=rete100)

        ws.cell(row=fila, column=10,
            value=f"=SUM(G{fila}:I{fila})")

        fila += 1

    fila_suma = fila

    ws.cell(row=fila_suma, column=1, value="TOTAL")

    for col in range(2, 11):
        letra = chr(64 + col)
        ws.cell(row=fila_suma, column=col,
            value=f"=SUM({letra}{inicio_tabla}:{letra}{fila_suma-1})")

    for col in range(1, 11):
        ws.cell(row=fila_suma, column=col).fill = amarillo

    fila += 3

wb.save(archivo_excel)

print("Excel generado correctamente")
