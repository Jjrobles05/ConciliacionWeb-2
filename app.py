import io

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Conciliación Bancaria",
    layout="wide"
)

st.title("📑 Conciliación Bancaria")


# =========================================================
# FUNCIONES
# =========================================================

def total_columna(df):
    """Suma la columna Valor de forma segura."""
    if "Valor" not in df.columns:
        return 0.0

    valores = pd.to_numeric(
        df["Valor"],
        errors="coerce"
    ).fillna(0)

    return float(valores.sum())


def limpiar_dataframe(df):
    """Deja únicamente las filas que tengan algún dato."""
    if df.empty:
        return df.copy()

    resultado = df.copy()

    resultado = resultado.replace(
        r"^\s*$",
        None,
        regex=True
    )

    columnas = list(resultado.columns)

    if columnas:
        resultado = resultado.dropna(
            how="all",
            subset=columnas
        )

    return resultado.reset_index(drop=True)


def preparar_excel(
    empresa,
    nit,
    mes,
    banco,
    cuenta,
    tipo,
    saldo_extracto,
    saldo_libros,
    diferencia_inicial,
    diferencia_final,
    salidas_extracto,
    salidas_libros,
    entradas_libros,
    entradas_extracto,
    total_salidas_extracto,
    total_salidas_libros,
    total_entradas_libros,
    total_entradas_extracto
):
    """Construye el archivo Excel completo en memoria."""

    output = io.BytesIO()

    salidas_extracto = limpiar_dataframe(salidas_extracto)
    salidas_libros = limpiar_dataframe(salidas_libros)
    entradas_libros = limpiar_dataframe(entradas_libros)
    entradas_extracto = limpiar_dataframe(entradas_extracto)

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # -------------------------------------------------
        # HOJA 1 - INFORMACIÓN GENERAL
        # -------------------------------------------------
        informacion = pd.DataFrame([
            ["INFORMACIÓN GENERAL", ""],
            ["Empresa", empresa],
            ["NIT", nit],
            ["Mes y Año", mes],
            ["Banco", banco],
            ["Número de Cuenta", cuenta],
            ["Tipo de Cuenta", tipo],
        ])

        informacion.to_excel(
            writer,
            sheet_name="Información General",
            index=False,
            header=False
        )

        # -------------------------------------------------
        # HOJA 2 - RESUMEN
        # -------------------------------------------------
        resumen = pd.DataFrame([
            ["RESUMEN DE LA CONCILIACIÓN", ""],
            ["Saldo según Extracto Bancario", saldo_extracto],
            ["Saldo según Libros", saldo_libros],
            ["Diferencia Inicial", diferencia_inicial],
            ["", ""],
            ["Salidas no registradas en Extracto", total_salidas_extracto],
            ["Salidas Bancarias no contabilizadas en Libros", total_salidas_libros],
            ["Entradas Bancarias no contabilizadas en Libros", total_entradas_libros],
            ["Entradas no evidenciadas en Extractos", total_entradas_extracto],
            ["", ""],
            ["Diferencia Final", diferencia_final],
            [
                "Estado",
                "CONCILIADA"
                if abs(diferencia_final) < 0.005
                else "CON DIFERENCIA"
            ],
        ])

        resumen.to_excel(
            writer,
            sheet_name="Resumen",
            index=False,
            header=False
        )

        # -------------------------------------------------
        # HOJAS DE DETALLE
        # -------------------------------------------------
        salidas_extracto.to_excel(
            writer,
            sheet_name="Salidas no registradas",
            index=False,
            startrow=1
        )

        salidas_libros.to_excel(
            writer,
            sheet_name="Salidas no contabilizadas",
            index=False,
            startrow=1
        )

        entradas_libros.to_excel(
            writer,
            sheet_name="Entradas no contabilizadas",
            index=False,
            startrow=1
        )

        entradas_extracto.to_excel(
            writer,
            sheet_name="Entradas no evidenciadas",
            index=False,
            startrow=1
        )

    output.seek(0)

    # -----------------------------------------------------
    # FORMATO PROFESIONAL DEL EXCEL
    # -----------------------------------------------------

    workbook = load_workbook(output)

    titulo_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    encabezado_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    titulo_font = Font(
        bold=True,
        color="FFFFFF",
        size=14
    )

    encabezado_font = Font(
        bold=True,
        color="000000"
    )

    borde = Border(
        bottom=Side(
            style="thin",
            color="B7B7B7"
        )
    )

    # -----------------------------------------------------
    # INFORMACIÓN GENERAL Y RESUMEN
    # -----------------------------------------------------

    for nombre in [
        "Información General",
        "Resumen"
    ]:

        ws = workbook[nombre]

        ws.column_dimensions["A"].width = 48
        ws.column_dimensions["B"].width = 28

        ws.merge_cells("A1:B1")

        ws["A1"].fill = titulo_fill
        ws["A1"].font = titulo_font
        ws["A1"].alignment = Alignment(
            horizontal="center"
        )

        for row in ws.iter_rows(
            min_row=2,
            max_row=ws.max_row,
            min_col=1,
            max_col=2
        ):
            for cell in row:
                cell.border = borde

        if nombre == "Resumen":

            for fila in range(
                2,
                ws.max_row + 1
            ):

                if fila in [
                    2,
                    3,
                    4,
                    6,
                    7,
                    8,
                    9,
                    11
                ]:

                    ws.cell(
                        row=fila,
                        column=2
                    ).number_format = '#,##0.00'

    # -----------------------------------------------------
    # HOJAS DE DETALLE
    # -----------------------------------------------------

    hojas_detalle = [
        "Salidas no registradas",
        "Salidas no contabilizadas",
        "Entradas no contabilizadas",
        "Entradas no evidenciadas",
    ]

    for nombre in hojas_detalle:

        ws = workbook[nombre]

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=ws.max_column
        )

        ws.cell(
            row=1,
            column=1
        ).value = nombre

        ws.cell(
            row=1,
            column=1
        ).fill = titulo_fill

        ws.cell(
            row=1,
            column=1
        ).font = titulo_font

        ws.cell(
            row=1,
            column=1
        ).alignment = Alignment(
            horizontal="center"
        )

        # Encabezados
        for cell in ws[2]:

            cell.fill = encabezado_fill

            cell.font = encabezado_font

            cell.border = borde

            cell.alignment = Alignment(
                horizontal="center"
            )

        # Anchos
        for columna in range(
            1,
            ws.max_column + 1
        ):

            letra = get_column_letter(
                columna
            )

            valor = ws.cell(
                row=2,
                column=columna
            ).value

            if valor == "Fecha":
                ancho = 15

            elif valor in [
                "Beneficiario",
                "Concepto"
            ]:
                ancho = 40

            elif valor == "Documento":
                ancho = 20

            elif valor == "Valor":
                ancho = 18

            else:
                ancho = 20

            ws.column_dimensions[
                letra
            ].width = ancho

        # Formato monetario
        for columna in range(
            1,
            ws.max_column + 1
        ):

            if ws.cell(
                row=2,
                column=columna
            ).value == "Valor":

                for fila in range(
                    3,
                    ws.max_row + 1
                ):

                    ws.cell(
                        row=fila,
                        column=columna
                    ).number_format = '#,##0.00'

        ws.freeze_panes = "A3"

        ws.auto_filter.ref = ws.dimensions

    workbook[
        "Información General"
    ].freeze_panes = "A2"

    workbook[
        "Resumen"
    ].freeze_panes = "A2"

    # -----------------------------------------------------
    # NOMBRE DEL ARCHIVO
    # -----------------------------------------------------

    nombre_mes = (
        str(mes).strip()
        if str(mes).strip()
        else "Sin_mes"
    )

    nombre_empresa = (
        str(empresa).strip()
        if str(empresa).strip()
        else "Empresa"
    )

    caracteres = '<>:"/\\|?*'

    for caracter in caracteres:

        nombre_mes = nombre_mes.replace(
            caracter,
            "-"
        )

        nombre_empresa = nombre_empresa.replace(
            caracter,
            "-"
        )

    nombre_archivo = (
        f"CONCILIACION_{nombre_empresa}_{nombre_mes}.xlsx"
    )

    final_output = io.BytesIO()

    workbook.save(final_output)

    final_output.seek(0)

    return (
        final_output.getvalue(),
        nombre_archivo
    )


# =========================================================
# INFORMACIÓN GENERAL
# =========================================================

st.subheader("Información General")

col1, col2 = st.columns(2)

with col1:

    empresa = st.text_input(
        "Nombre de la Empresa"
    )

    nit = st.text_input(
        "NIT"
    )

    mes = st.text_input(
        "Mes y Año"
    )

with col2:

    banco = st.text_input(
        "Nombre del Banco"
    )

    cuenta = st.text_input(
        "Número de Cuenta"
    )

    tipo = st.selectbox(
        "Tipo de Cuenta",
        ["Ahorros", "Corriente"]
    )


# =========================================================
# SALDOS
# =========================================================

st.divider()

st.subheader("Saldos")

col1, col2 = st.columns(2)

with col1:

    saldo_extracto = st.number_input(
        "Saldo según Extracto Bancario",
        value=0.0,
        format="%.2f"
    )

with col2:

    saldo_libros = st.number_input(
        "Saldo según Libros",
        value=0.0,
        format="%.2f"
    )

diferencia_inicial = (
    saldo_extracto - saldo_libros
)

st.metric(
    "Diferencia Inicial",
    f"${diferencia_inicial:,.2f}"
)


# =========================================================
# SALIDAS NO REGISTRADAS
# =========================================================

st.divider()

st.subheader(
    "Salidas no Registradas en Extracto"
)

salidas_extracto = st.data_editor(
    pd.DataFrame(
        columns=[
            "Fecha",
            "Beneficiario",
            "Documento",
            "Valor"
        ]
    ),
    num_rows="dynamic",
    width="stretch",
    key="tabla1"
)


# =========================================================
# SALIDAS BANCARIAS
# =========================================================

st.divider()

st.subheader(
    "Salidas Bancarias no Contabilizadas en Libros"
)

salidas_libros = st.data_editor(
    pd.DataFrame(
        columns=[
            "Fecha",
            "Concepto",
            "Valor"
        ]
    ),
    num_rows="dynamic",
    width="stretch",
    key="tabla2"
)


# =========================================================
# ENTRADAS BANCARIAS
# =========================================================

st.divider()

st.subheader(
    "Entradas Bancarias no Contabilizadas en Libros"
)

entradas_libros = st.data_editor(
    pd.DataFrame(
        columns=[
            "Fecha",
            "Concepto",
            "Valor"
        ]
    ),
    num_rows="dynamic",
    width="stretch",
    key="tabla3"
)


# =========================================================
# ENTRADAS NO EVIDENCIADAS
# =========================================================

st.divider()

st.subheader(
    "Entradas no Evidenciadas en Extractos"
)

entradas_extracto = st.data_editor(
    pd.DataFrame(
        columns=[
            "Fecha",
            "Concepto",
            "Valor"
        ]
    ),
    num_rows="dynamic",
    width="stretch",
    key="tabla4"
)


# =========================================================
# TOTALES
# =========================================================

total_salidas_extracto = total_columna(
    salidas_extracto
)

total_salidas_libros = total_columna(
    salidas_libros
)

total_entradas_libros = total_columna(
    entradas_libros
)

total_entradas_extracto = total_columna(
    entradas_extracto
)

st.divider()

st.subheader(
    "Resumen de Justificaciones"
)

c1, c2 = st.columns(2)

with c1:

    st.metric(
        "Salidas no registradas",
        f"${total_salidas_extracto:,.2f}"
    )

    st.metric(
        "Salidas no contabilizadas",
        f"${total_salidas_libros:,.2f}"
    )

with c2:

    st.metric(
        "Entradas no contabilizadas",
        f"${total_entradas_libros:,.2f}"
    )

    st.metric(
        "Entradas no evidenciadas",
        f"${total_entradas_extracto:,.2f}"
    )


# =========================================================
# DIFERENCIA FINAL
# =========================================================

diferencia_final = (
    diferencia_inicial
    - total_salidas_extracto
    - total_salidas_libros
    + total_entradas_libros
    + total_entradas_extracto
)

st.divider()

st.subheader(
    "Resultado de la Conciliación"
)

st.metric(
    "Diferencia Final",
    f"${diferencia_final:,.2f}"
)

conciliada = (
    abs(diferencia_final) < 0.005
)

if conciliada:

    st.success(
        "✅ La conciliación está cuadrada."
    )

else:

    st.warning(
        "⚠️ La conciliación presenta diferencias."
    )


# =========================================================
# EXPORTACIÓN A EXCEL
# =========================================================

st.divider()

st.subheader(
    "Exportar Conciliación"
)

excel_data, nombre_archivo = preparar_excel(
    empresa=empresa,
    nit=nit,
    mes=mes,
    banco=banco,
    cuenta=cuenta,
    tipo=tipo,
    saldo_extracto=saldo_extracto,
    saldo_libros=saldo_libros,
    diferencia_inicial=diferencia_inicial,
    diferencia_final=diferencia_final,
    salidas_extracto=salidas_extracto,
    salidas_libros=salidas_libros,
    entradas_libros=entradas_libros,
    entradas_extracto=entradas_extracto,
    total_salidas_extracto=total_salidas_extracto,
    total_salidas_libros=total_salidas_libros,
    total_entradas_libros=total_entradas_libros,
    total_entradas_extracto=total_entradas_extracto,
)

st.download_button(
    label="💾 Descargar Conciliación en Excel",
    data=excel_data,
    file_name=nombre_archivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch"
)

st.caption(
    "El archivo Excel incluye la información general, "
    "el resumen y las cuatro hojas de detalle "
    "de la conciliación."
)
