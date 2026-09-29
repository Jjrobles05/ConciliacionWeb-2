import io
import re

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Conciliación Bancaria",
    layout="wide"
)

st.title("📑 Conciliación Bancaria")


# =========================================================
# FUNCIONES
# =========================================================

def total_columna(df, columna="Valor"):
    """Suma una columna numérica de forma segura."""
    if columna not in df.columns:
        return 0.0

    return float(
        pd.to_numeric(
            df[columna],
            errors="coerce"
        ).fillna(0).sum()
    )


def limpiar_dataframe(df):
    """Elimina filas completamente vacías."""
    if df is None or df.empty:
        return df.copy()

    resultado = df.copy()

    for columna in resultado.columns:
        if resultado[columna].dtype == "object":
            resultado[columna] = resultado[columna].replace(
                r"^\s*$",
                None,
                regex=True
            )

    return resultado.dropna(
        how="all"
    ).reset_index(drop=True)


def limpiar_nombre_archivo(texto):
    """Elimina caracteres no permitidos en nombres de archivos."""
    texto = str(texto).strip()

    if not texto:
        texto = "Empresa"

    return re.sub(
        r'[<>:"/\\|?*]',
        "-",
        texto
    )


def estilo_titulo(celda, tamano=12):
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )
    celda.font = Font(
        bold=True,
        color="FFFFFF",
        size=tamano
    )
    celda.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )


def estilo_encabezado(celda):
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )
    celda.font = Font(
        bold=True
    )
    celda.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )


def escribir_seccion(ws, fila, titulo, columnas=4):
    """Crea el título de una sección."""
    ws.merge_cells(
        start_row=fila,
        start_column=1,
        end_row=fila,
        end_column=columnas
    )

    celda = ws.cell(
        row=fila,
        column=1
    )

    celda.value = titulo
    estilo_titulo(celda)

    ws.row_dimensions[fila].height = 22

    return fila + 1


def escribir_tabla(ws, fila, titulo, df):
    """
    Escribe una tabla completa en la misma hoja:
    título + encabezados + datos + total.
    """
    df = limpiar_dataframe(df)

    numero_columnas = max(
        4,
        len(df.columns)
    )

    fila = escribir_seccion(
        ws,
        fila,
        titulo,
        numero_columnas
    )

    # Encabezados
    for columna, nombre in enumerate(
        df.columns,
        start=1
    ):
        celda = ws.cell(
            row=fila,
            column=columna
        )
        celda.value = nombre
        estilo_encabezado(celda)

    fila += 1

    # Datos
    for _, registro in df.iterrows():

        for columna, nombre in enumerate(
            df.columns,
            start=1
        ):
            valor = registro[nombre]

            if pd.isna(valor):
                valor = ""

            celda = ws.cell(
                row=fila,
                column=columna
            )
            celda.value = valor

            if nombre == "Valor":
                celda.number_format = '#,##0.00'

            if nombre in [
                "4 x 1000",
                "Cuota de manejo",
                "IVA",
                "Rte. fuente",
                "Comisión",
                "Ing. x intereses"
            ]:
                celda.number_format = '#,##0.00'

        fila += 1

    # Si no hay registros, dejar una fila visual
    if df.empty:
        fila += 1

    # Total
    columna_valor = None

    for columna, nombre in enumerate(
        df.columns,
        start=1
    ):
        if nombre == "Valor":
            columna_valor = columna
            break

    if columna_valor is not None:

        ws.cell(
            row=fila,
            column=columna_valor - 1
        ).value = "TOTAL"

        ws.cell(
            row=fila,
            column=columna_valor - 1
        ).font = Font(
            bold=True
        )

        ws.cell(
            row=fila,
            column=columna_valor
        ).value = total_columna(
            df,
            "Valor"
        )

        ws.cell(
            row=fila,
            column=columna_valor
        ).number_format = '#,##0.00'

        ws.cell(
            row=fila,
            column=columna_valor
        ).font = Font(
            bold=True
        )

    return fila + 2


def escribir_gastos_bancarios(ws, fila, df):
    """Escribe la tabla especial de gastos bancarios."""
    df = limpiar_dataframe(df)

    columnas = [
        "Fecha",
        "4 x 1000",
        "Cuota de manejo",
        "IVA",
        "Rte. fuente",
        "Comisión",
        "Ing. x intereses"
    ]

    fila = escribir_seccion(
        ws,
        fila,
        "GASTOS BANCARIOS",
        len(columnas)
    )

    for columna, nombre in enumerate(
        columnas,
        start=1
    ):
        celda = ws.cell(
            row=fila,
            column=columna
        )
        celda.value = nombre
        estilo_encabezado(celda)

    fila += 1

    for _, registro in df.iterrows():

        for columna, nombre in enumerate(
            columnas,
            start=1
        ):
            valor = registro.get(
                nombre,
                ""
            )

            if pd.isna(valor):
                valor = ""

            celda = ws.cell(
                row=fila,
                column=columna
            )
            celda.value = valor

            if nombre != "Fecha":
                celda.number_format = '#,##0.00'

        fila += 1

    if df.empty:
        fila += 1

    # Totales de gastos
    fila_total = fila

    ws.cell(
        row=fila_total,
        column=1
    ).value = "TOTAL"

    ws.cell(
        row=fila_total,
        column=1
    ).font = Font(
        bold=True
    )

    for columna, nombre in enumerate(
        columnas[1:],
        start=2
    ):
        total = float(
            pd.to_numeric(
                df.get(
                    nombre,
                    pd.Series(dtype=float)
                ),
                errors="coerce"
            ).fillna(0).sum()
        )

        ws.cell(
            row=fila_total,
            column=columna
        ).value = total

        ws.cell(
            row=fila_total,
            column=columna
        ).number_format = '#,##0.00'

        ws.cell(
            row=fila_total,
            column=columna
        ).font = Font(
            bold=True
        )

    return fila_total + 2


def preparar_excel(
    empresa,
    nit,
    mes,
    fecha_elaboracion,
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
    gastos_bancarios,
    preparado_por,
    revisado_por
):
    """
    Genera UN SOLO archivo Excel con UNA SOLA HOJA,
    siguiendo la estructura del formato suministrado.
    """

    salidas_extracto = limpiar_dataframe(
        salidas_extracto
    )
    salidas_libros = limpiar_dataframe(
        salidas_libros
    )
    entradas_libros = limpiar_dataframe(
        entradas_libros
    )
    entradas_extracto = limpiar_dataframe(
        entradas_extracto
    )
    gastos_bancarios = limpiar_dataframe(
        gastos_bancarios
    )

    workbook = Workbook()

    ws = workbook.active
    ws.title = "Conciliación Bancaria"

    # -----------------------------------------------------
    # CONFIGURACIÓN DE PÁGINA
    # -----------------------------------------------------

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.40
    ws.page_margins.bottom = 0.40

    # -----------------------------------------------------
    # ANCHOS
    # -----------------------------------------------------

    anchos = {
        "A": 15,
        "B": 28,
        "C": 16,
        "D": 16,
        "E": 16,
        "F": 16,
        "G": 16
    }

    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    # -----------------------------------------------------
    # TÍTULO
    # -----------------------------------------------------

    ws.merge_cells("A1:G1")

    ws["A1"] = "CONCILIACIÓN BANCARIA"

    estilo_titulo(
        ws["A1"],
        16
    )

    ws.row_dimensions[1].height = 28

    # -----------------------------------------------------
    # INFORMACIÓN GENERAL
    # -----------------------------------------------------

    fila = 3

    fila = escribir_seccion(
        ws,
        fila,
        "INFORMACIÓN GENERAL",
        7
    )

    datos_generales = [
        ("Empresa", empresa),
        ("NIT", nit),
        ("Mes y Año", mes),
        ("Fecha de elaboración", fecha_elaboracion),
        ("Banco", banco),
        ("Cuenta No.", cuenta),
        ("Tipo", tipo),
    ]

    for nombre, valor in datos_generales:

        ws.cell(
            row=fila,
            column=1
        ).value = nombre

        ws.cell(
            row=fila,
            column=1
        ).font = Font(
            bold=True
        )

        ws.merge_cells(
            start_row=fila,
            start_column=2,
            end_row=fila,
            end_column=4
        )

        ws.cell(
            row=fila,
            column=2
        ).value = valor

        fila += 1

    fila += 1

    # -----------------------------------------------------
    # SALDOS
    # -----------------------------------------------------

    fila = escribir_seccion(
        ws,
        fila,
        "SALDOS",
        4
    )

    saldos = [
        (
            "SALDO SEGÚN EXTRACTO BANCARIO",
            saldo_extracto
        ),
        (
            "SALDO SEGÚN LIBROS",
            saldo_libros
        ),
        (
            "DIFERENCIA A JUSTIFICAR",
            diferencia_inicial
        ),
    ]

    for nombre, valor in saldos:

        ws.cell(
            row=fila,
            column=1
        ).value = nombre

        ws.cell(
            row=fila,
            column=1
        ).font = Font(
            bold=True
        )

        ws.cell(
            row=fila,
            column=2
        ).value = valor

        ws.cell(
            row=fila,
            column=2
        ).number_format = '#,##0.00'

        fila += 1

    fila += 1

    # -----------------------------------------------------
    # JUSTIFICACIÓN - RESUMEN
    # -----------------------------------------------------

    fila = escribir_seccion(
        ws,
        fila,
        "JUSTIFICACIÓN",
        4
    )

    justificaciones = [
        (
            "SALIDAS NO REGISTRADAS EN EXTRACTO",
            total_columna(
                salidas_extracto
            )
        ),
        (
            "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
            total_columna(
                salidas_libros
            )
        ),
        (
            "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
            total_columna(
                entradas_libros
            )
        ),
        (
            "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS",
            total_columna(
                entradas_extracto
            )
        ),
    ]

    for nombre, valor in justificaciones:

        ws.cell(
            row=fila,
            column=1
        ).value = nombre

        ws.cell(
            row=fila,
            column=1
        ).font = Font(
            bold=True
        )

        ws.cell(
            row=fila,
            column=2
        ).value = valor

        ws.cell(
            row=fila,
            column=2
        ).number_format = '#,##0.00'

        fila += 1

    fila += 1

    # -----------------------------------------------------
    # DETALLE DE MOVIMIENTOS
    # -----------------------------------------------------

    fila = escribir_tabla(
        ws,
        fila,
        "SALIDAS NO REGISTRADAS EN EXTRACTO",
        salidas_extracto
    )

    fila = escribir_tabla(
        ws,
        fila,
        "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
        salidas_libros
    )

    fila = escribir_tabla(
        ws,
        fila,
        "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
        entradas_libros
    )

    fila = escribir_tabla(
        ws,
        fila,
        "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS",
        entradas_extracto
    )

    # -----------------------------------------------------
    # GASTOS BANCARIOS
    # -----------------------------------------------------

    fila = escribir_gastos_bancarios(
        ws,
        fila,
        gastos_bancarios
    )

    # -----------------------------------------------------
    # RESULTADO
    # -----------------------------------------------------

    fila = escribir_seccion(
        ws,
        fila,
        "RESULTADO DE LA CONCILIACIÓN",
        4
    )

    ws.cell(
        row=fila,
        column=1
    ).value = "DIFERENCIA CONCILIADA"

    ws.cell(
        row=fila,
        column=1
    ).font = Font(
        bold=True
    )

    ws.cell(
        row=fila,
        column=2
    ).value = diferencia_final

    ws.cell(
        row=fila,
        column=2
    ).number_format = '#,##0.00'

    fila += 1

    estado = (
        "CONCILIACIÓN BANCARIA CORRECTA"
        if abs(diferencia_final) < 0.005
        else "CONCILIACIÓN CON DIFERENCIA"
    )

    ws.cell(
        row=fila,
        column=1
    ).value = "ESTADO"

    ws.cell(
        row=fila,
        column=1
    ).font = Font(
        bold=True
    )

    ws.merge_cells(
        start_row=fila,
        start_column=2,
        end_row=fila,
        end_column=4
    )

    ws.cell(
        row=fila,
        column=2
    ).value = estado

    ws.cell(
        row=fila,
        column=2
    ).font = Font(
        bold=True
    )

    fila += 2

    # -----------------------------------------------------
    # PREPARADO / REVISADO
    # -----------------------------------------------------

    ws.cell(
        row=fila,
        column=1
    ).value = "Preparado por:"
    ws.cell(
        row=fila,
        column=1
    ).font = Font(
        bold=True
    )

    ws.cell(
        row=fila,
        column=2
    ).value = preparado_por

    ws.cell(
        row=fila,
        column=5
    ).value = "Revisado por:"
    ws.cell(
        row=fila,
        column=5
    ).font = Font(
        bold=True
    )

    ws.cell(
        row=fila,
        column=6
    ).value = revisado_por

    fila += 2

    ws.cell(
        row=fila,
        column=1
    ).value = "Firma:"
    ws.cell(
        row=fila,
        column=5
    ).value = "Firma:"

    # -----------------------------------------------------
    # BORDES
    # -----------------------------------------------------

    borde = Border(
        left=Side(
            style="thin",
            color="D9D9D9"
        ),
        right=Side(
            style="thin",
            color="D9D9D9"
        ),
        top=Side(
            style="thin",
            color="D9D9D9"
        ),
        bottom=Side(
            style="thin",
            color="D9D9D9"
        )
    )

    for fila_excel in ws.iter_rows():

        for celda in fila_excel:

            if celda.value is not None:
                celda.border = borde
                celda.alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

    # Volver a centrar títulos después del borde general
    for fila_excel in ws.iter_rows():

        for celda in fila_excel:

            if (
                celda.fill
                and celda.fill.fill_type == "solid"
                and celda.fill.fgColor.rgb
                and "1F4E78" in str(
                    celda.fill.fgColor.rgb
                )
            ):
                celda.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

    ws.freeze_panes = "A3"

    # -----------------------------------------------------
    # EXPORTAR A MEMORIA
    # -----------------------------------------------------

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    nombre_archivo = (
        f"CONCILIACION_"
        f"{limpiar_nombre_archivo(empresa)}_"
        f"{limpiar_nombre_archivo(mes)}.xlsx"
    )

    return (
        output.getvalue(),
        nombre_archivo
    )


# =========================================================
# INFORMACIÓN GENERAL - APLICACIÓN
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

    fecha_elaboracion = st.date_input(
        "Fecha de elaboración"
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
    "Diferencia a Justificar",
    f"${diferencia_inicial:,.2f}"
)


# =========================================================
# MOVIMIENTOS
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
# GASTOS BANCARIOS
# =========================================================

st.divider()

st.subheader(
    "Gastos Bancarios"
)

st.caption(
    "Esta sección permite registrar los conceptos "
    "que aparecen en el formato de conciliación "
    "suministrado."
)

gastos_bancarios = st.data_editor(
    pd.DataFrame(
        columns=[
            "Fecha",
            "4 x 1000",
            "Cuota de manejo",
            "IVA",
            "Rte. fuente",
            "Comisión",
            "Ing. x intereses"
        ]
    ),
    num_rows="dynamic",
    width="stretch",
    key="tabla5"
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
    "Diferencia Conciliada",
    f"${diferencia_final:,.2f}"
)

conciliada = (
    abs(diferencia_final) < 0.005
)

if conciliada:

    st.success(
        "✅ CONCILIACIÓN BANCARIA CORRECTA"
    )

else:

    st.warning(
        "⚠️ La conciliación presenta diferencias."
    )


# =========================================================
# PREPARADO / REVISADO
# =========================================================

st.divider()

st.subheader(
    "Responsables"
)

col1, col2 = st.columns(2)

with col1:

    preparado_por = st.text_input(
        "Preparado por"
    )

with col2:

    revisado_por = st.text_input(
        "Revisado por"
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
    fecha_elaboracion=fecha_elaboracion,
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
    gastos_bancarios=gastos_bancarios,
    preparado_por=preparado_por,
    revisado_por=revisado_por
)

st.download_button(
    label="💾 Descargar Conciliación en Excel",
    data=excel_data,
    file_name=nombre_archivo,
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    width="stretch"
)

st.caption(
    "El Excel se genera en una sola hoja, "
    "siguiendo la estructura del formato suministrado."
)
