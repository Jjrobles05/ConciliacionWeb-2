import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Conciliación Bancaria",
    layout="wide"
)

st.title("📑 Conciliación Bancaria")

# =========================
# INFORMACIÓN GENERAL
# =========================

st.subheader("Información General")

col1, col2 = st.columns(2)

with col1:
    empresa = st.text_input("Nombre de la Empresa")
    nit = st.text_input("NIT")
    mes = st.text_input("Mes y Año")

with col2:
    banco = st.text_input("Nombre del Banco")
    cuenta = st.text_input("Número de Cuenta")
    tipo = st.selectbox(
        "Tipo de Cuenta",
        ["Ahorros", "Corriente"]
    )

# =========================
# SALDOS
# =========================

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

diferencia_inicial = saldo_extracto - saldo_libros

st.metric(
    "Diferencia Inicial",
    f"${diferencia_inicial:,.2f}"
)

# =========================
# SALIDAS NO REGISTRADAS
# =========================

st.divider()

st.subheader("Salidas no Registradas en Extracto")

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

# =========================
# SALIDAS BANCARIAS
# =========================

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

# =========================
# ENTRADAS BANCARIAS
# =========================

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

# =========================
# ENTRADAS NO EVIDENCIADAS
# =========================

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

# =========================
# TOTALES
# =========================

def total_columna(df):
    if "Valor" in df.columns:
        valores = pd.to_numeric(
            df["Valor"],
            errors="coerce"
        ).fillna(0)
        return valores.sum()
    return 0

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

st.subheader("Resumen de Justificaciones")

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

# =========================
# DIFERENCIA FINAL
# =========================

diferencia_final = (
    diferencia_inicial
    - total_salidas_extracto
    - total_salidas_libros
    + total_entradas_libros
    + total_entradas_extracto
)

st.divider()

st.subheader("Resultado de la Conciliación")

st.metric(
    "Diferencia Final",
    f"${diferencia_final:,.2f}"
)

if diferencia_final == 0:
    st.success(
        "✅ La conciliación está cuadrada."
    )
else:
    st.warning(
        "⚠️ La conciliación presenta diferencias."
    )

st.divider()

if st.button("💾 Guardar Conciliación"):

    hoja = pd.DataFrame([
        ["CONCILIACIÓN BANCARIA", ""],
        ["", ""],
        ["Empresa", empresa],
        ["NIT", nit],
        ["Mes", mes],
        ["Banco", banco],
        ["Cuenta", cuenta],
        ["Tipo", tipo],
        ["", ""],
        ["Saldo según Extracto", saldo_extracto],
        ["Saldo según Libros", saldo_libros],
        ["Diferencia Final", diferencia_final],
        ["", ""],
        ["SALIDAS NO REGISTRADAS EN EXTRACTO", ""],
        ["", ""],
        ["SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS", ""],
        ["", ""],
        ["ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS", ""],
        ["", ""],
        ["ENTRADAS NO EVIDENCIADAS EN EXTRACTOS", ""]
    ])

    with pd.ExcelWriter(
        "conciliacion_guardada.xlsx",
        engine="openpyxl"
    ) as writer:

        hoja.to_excel(
            writer,
            sheet_name="Conciliacion",
            index=False,
            header=False
        )

    st.success(
        "Archivo Excel guardado correctamente"
    )