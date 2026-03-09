import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("Generador de Retenciones")

st.write("Sube tus PDFs de retenciones")

archivos = st.file_uploader(
    "Subir PDFs",
    type="pdf",
    accept_multiple_files=True
)

datos = []

if archivos:

    for archivo in archivos:

        texto = ""

        with pdfplumber.open(archivo) as pdf:
            for pagina in pdf.pages:
                t = pagina.extract_text()
                if t:
                    texto += t + "\n"

        lineas = texto.split("\n")

        base0 = ""
        base15 = ""
        iva = 0
        propina = 0

        rete1 = ""
        rete2 = ""
        rete10 = ""
        rete100 = ""

        base_imponible = None
        porcentaje = None
        impuesto_tipo = ""

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

            if "IVA" in l and iva == 0:
                try:
                    iva = float(l.split()[-1])
                except:
                    pass

        if base_imponible:

            if "iva" in impuesto_tipo:
                base15 = base_imponible
            else:
                base0 = base_imponible

        total = (
            (base0 if isinstance(base0,float) else 0) +
            (base15 if isinstance(base15,float) else 0) +
            iva +
            propina
        )

        if porcentaje and base_imponible:

            valor_retenido = round(base_imponible * porcentaje / 100,2)

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

        datos.append({
            "BASE 0%": base0,
            "BASE 15%": base15,
            "PROPINA": propina,
            "IVA": iva,
            "TOTAL": total,
            "RETE 1%": rete1,
            "RETE 2%": rete2,
            "RETE 10%": rete10,
            "RETE 100%": rete100,
            "TOTAL RETENIDO": total_retenido
        })

    df = pd.DataFrame(datos)

    st.dataframe(df)

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        label="Descargar Excel",
        data=buffer.getvalue(),
        file_name="retenciones.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
