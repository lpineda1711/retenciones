import os
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# carpeta de PDFs
carpeta_pdfs = "pdfs"

# crear carpeta si no existe
if not os.path.isdir(carpeta_pdfs):
    os.makedirs(carpeta_pdfs)

archivo_excel = "retenciones.xlsx"

# encabezados
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

# meses
meses = [
    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
    "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
]

# color amarillo
amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

wb = Workbook()
ws = wb.active
ws.title = "RETENCIONES"

fila_excel = 1

for mes in meses:

    # fila del mes
    ws.cell(row=fila_excel, column=1, value=mes)

    for c in range(1, len(encabezados)+1):
        ws.cell(row=fila_excel, column=c).fill = amarillo

    fila_excel += 1

    # encabezados
    for c, titulo in enumerate(encabezados, 1):
        ws.cell(row=fila_excel, column=c, value=titulo)
        ws.cell(row=fila_excel, column=c).fill = amarillo

    fila_excel += 1

    fila_inicio_datos = fila_excel

    # recorrer PDFs
    for archivo in os.listdir(carpeta_pdfs):

        if not archivo.lower().endswith(".pdf"):
            continue

        ruta_pdf = os.path.join(carpeta_pdfs, archivo)

        base0 = ""
        base15 = ""
        propina = ""
        iva = ""
        total = ""

        rete1 = ""
        rete2 = ""
        rete10 = ""
        rete100 = ""

        base_imponible = None
        porcentaje = None
        impuesto_tipo = ""

        texto = ""

        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                t = pagina.extract_text()
                if t:
                    texto += t + "\n"

        lineas = texto.split("\n")

        for l in lineas:

            if "Base Imponible para la Retención" in l:
                try:
                    base_imponible = float(l.split()[-1])
                except:
                    pass

            if "Porcentaje Retención" in l:
                try:
                    porcentaje = float(l.split()[-1])
                except:
                    pass

            if "Impuesto" in l:
                impuesto_tipo = l.lower()

            if "IVA" in l and iva == "":
                try:
                    iva = float(l.split()[-1])
                except:
                    pass

        # decidir base
        if base_imponible is not None:

            if "iva" in impuesto_tipo:
                base15 = base_imponible
            else:
                base0 = base_imponible

        # calcular total
        total = (
            (base0 if isinstance(base0,float) else 0) +
            (base15 if isinstance(base15,float) else 0) +
            (propina if isinstance(propina,float) else 0) +
            (iva if isinstance(iva,float) else 0)
        )

        # calcular retención
        if porcentaje is not None and base_imponible is not None:

            valor_retenido = round(base_imponible * porcentaje / 100, 2)

            if porcentaje == 1:
                rete1 = valor_retenido

            elif porcentaje == 2:
                rete2 = valor_retenido

            elif porcentaje == 10:
                rete10 = valor_retenido

            elif porcentaje == 100:
                rete100 = valor_retenido

        total_retenido = (
            (rete1 if isinstance(rete1,float) else 0) +
            (rete2 if isinstance(rete2,float) else 0) +
            (rete10 if isinstance(rete10,float) else 0) +
            (rete100 if isinstance(rete100,float) else 0)
        )

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

        for c, valor in enumerate(datos, 1):
            ws.cell(row=fila_excel, column=c, value=valor)

        fila_excel += 1

    # fila sumatoria
    fila_suma = fila_excel

    for c in range(2, len(encabezados)+1):

        letra = ws.cell(row=1, column=c).column_letter

        formula = f"=SUM({letra}{fila_inicio_datos}:{letra}{fila_excel-1})"

        ws.cell(row=fila_suma, column=c, value=formula)
        ws.cell(row=fila_suma, column=c).fill = amarillo

    fila_excel += 2

wb.save(archivo_excel)

print("Archivo Excel creado correctamente")
