import pdfplumber
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# carpeta donde están los pdf
carpeta_pdfs = "pdfs"

# crear carpeta si no existe
if not os.path.exists(carpeta_pdfs):
    os.makedirs(carpeta_pdfs)

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

amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

wb = Workbook()
ws = wb.active
ws.title = "RETENCIONES"

fila = 1

for mes in meses:

    # fila del mes
    ws.cell(row=fila, column=1, value=mes)
    for c in range(1, len(encabezados)+1):
        ws.cell(row=fila, column=c).fill = amarillo
    fila += 1

    # encabezados
    for c,enc in enumerate(encabezados,1):
        ws.cell(row=fila, column=c, value=enc)
        ws.cell(row=fila, column=c).fill = amarillo
    fila += 1

    # leer pdfs
    for archivo in os.listdir(carpeta_pdfs):

        if not archivo.lower().endswith(".pdf"):
            continue

        ruta = os.path.join(carpeta_pdfs, archivo)

        base0=""
        base15=""
        propina=0
        iva=0
        total=0

        rete1=""
        rete2=""
        rete10=""
        rete100=""

        base_imponible=None
        porcentaje=None
        impuesto=""

        with pdfplumber.open(ruta) as pdf:
            texto=""
            for pagina in pdf.pages:
                texto+=pagina.extract_text()

        lineas=texto.split("\n")

        for l in lineas:

            if "Base Imponible para la Retención" in l:
                try:
                    base_imponible=float(l.split()[-1])
                except:
                    pass

            if "Porcentaje Retención" in l:
                try:
                    porcentaje=float(l.split()[-1])
                except:
                    pass

            if "Impuesto" in l:
                impuesto=l.lower()

            if "IVA" in l and iva==0:
                try:
                    iva=float(l.split()[-1])
                except:
                    pass

        if base_imponible:

            if "iva" in impuesto:
                base15=base_imponible
            else:
                base0=base_imponible

        total=(base0 if base0!="" else 0)+(base15 if base15!="" else 0)+propina+iva

        if porcentaje and base_imponible:

            valor=round(base_imponible*porcentaje/100,2)

            if porcentaje==1:
                rete1=valor

            elif porcentaje==2:
                rete2=valor

            elif porcentaje==10:
                rete10=valor

            elif porcentaje==100:
                rete100=valor

        total_retenido=sum([
            rete1 if isinstance(rete1,float) else 0,
            rete2 if isinstance(rete2,float) else 0,
            rete10 if isinstance(rete10,float) else 0,
            rete100 if isinstance(rete100,float) else 0
        ])

        datos=[
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

        for c,v in enumerate(datos,1):
            ws.cell(row=fila,column=c,value=v)

        fila+=1

    # fila de sumatoria
    fila_suma=fila

    for c in range(2,len(encabezados)+1):

        letra=ws.cell(row=1,column=c).column_letter

        ws.cell(row=fila_suma,column=c,
        value=f"=SUM({letra}3:{letra}{fila-1})")

        ws.cell(row=fila_suma,column=c).fill=amarillo

    fila+=2

wb.save(archivo_excel)

print("Excel generado correctamente")
