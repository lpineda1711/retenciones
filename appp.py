import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os

carpeta_pdfs = "pdfs"
archivo_excel = "retenciones.xlsx"

encabezados = [
    "MES",
    "BASE 0%",
    "BASE 15%",
    "PROPINA",
    "IVA",
    "TOTAL",
    "RETE 1%",
    "RETE 2%",
    "RETE 10%",
    "RETE 100%",
    "TOTAL RETENIDO"
]

meses = [
    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
    "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
]

fill_amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

wb = Workbook()
ws = wb.active
ws.title = "RETENCIONES"

fila_excel = 1

for mes in meses:

    ws.cell(row=fila_excel, column=1, value=mes)

    for col in range(1,len(encabezados)+1):
        ws.cell(row=fila_excel,column=col).fill = fill_amarillo

    fila_excel += 1

    for col,enc in enumerate(encabezados,1):
        ws.cell(row=fila_excel,column=col,value=enc)
        ws.cell(row=fila_excel,column=col).fill = fill_amarillo

    fila_excel += 1

    for archivo in os.listdir(carpeta_pdfs):

        if not archivo.lower().endswith(".pdf"):
            continue

        ruta = os.path.join(carpeta_pdfs,archivo)

        base0 = ""
        base15 = ""
        propina = 0
        iva = 0
        total = 0

        rete1=""
        rete2=""
        rete10=""
        rete100=""

        porcentaje = None
        base_imponible = None
        impuesto_tipo = ""

        with pdfplumber.open(ruta) as pdf:
            texto = ""
            for pagina in pdf.pages:
                texto += pagina.extract_text()

        lineas = texto.split("\n")

        for l in lineas:

            if "Base Imponible para la Retención" in l:
                try:
                    base_imponible = float(l.split()[-1])
                except:
                    pass

            if "Impuesto" in l:
                impuesto_tipo = l.lower()

            if "Porcentaje Retención" in l:
                try:
                    porcentaje = float(l.split()[-1])
                except:
                    pass

            if "IVA" in l and iva == 0:
                try:
                    iva = float(l.split()[-1])
                except:
                    pass

        if base_imponible != None:

            if "iva" in impuesto_tipo:
                base15 = base_imponible
                base0 = ""
            else:
                base0 = base_imponible
                base15 = ""

        if isinstance(base0,float) or isinstance(base15,float):
            total = (base0 if base0!="" else 0) + (base15 if base15!="" else 0) + propina + iva

        valor_retenido = ""

        if porcentaje != None and base_imponible != None:

            valor_retenido = round(base_imponible * porcentaje / 100 ,2)

            if porcentaje == 1:
                rete1 = valor_retenido

            elif porcentaje == 2:
                rete2 = valor_retenido

            elif porcentaje == 10:
                rete10 = valor_retenido

            elif porcentaje == 100:
                rete100 = valor_retenido

        total_retenido = sum([
            rete1 if isinstance(rete1,float) else 0,
            rete2 if isinstance(rete2,float) else 0,
            rete10 if isinstance(rete10,float) else 0,
            rete100 if isinstance(rete100,float) else 0
        ])

        datos = [
            mes,
            base0,
            base15,
            propina,
            iva,
            total,
            rete1,
            rete2,
            rete10,
            rete100,
            total_retenido
        ]

        for col,val in enumerate(datos,1):
            ws.cell(row=fila_excel,column=col,value=val)

        fila_excel += 1

    fila_sumas = fila_excel

    for col in range(2,len(encabezados)+1):

        letra = ws.cell(row=1,column=col).column_letter

        ws.cell(row=fila_sumas,column=col,
        value=f"=SUM({letra}{fila_excel-10}:{letra}{fila_excel-1})")

        ws.cell(row=fila_sumas,column=col).fill = fill_amarillo

    fila_excel += 2

wb.save(archivo_excel)

print("Excel generado correctamente")
