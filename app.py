import io
import re
import json
import sqlite3
import hashlib
import hmac
import secrets
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Sistema de Conciliación Bancaria",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# BASE DE DATOS Y TABLAS (SQLite)
# =========================================================

DB_FILE = "conciliaciones.db"


def conectar_db():
    return sqlite3.connect(DB_FILE)


def inicializar_db():
    with conectar_db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS conciliaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_guardado TEXT NOT NULL,
                empresa TEXT,
                nit TEXT,
                mes TEXT,
                fecha_elaboracion TEXT,
                banco TEXT,
                cuenta TEXT,
                tipo TEXT,
                saldo_extracto REAL NOT NULL,
                saldo_libros REAL NOT NULL,
                diferencia_inicial REAL NOT NULL,
                diferencia_conciliada REAL NOT NULL,
                resultado_final REAL NOT NULL,
                estado TEXT NOT NULL,
                datos_json TEXT NOT NULL,
                excel BLOB NOT NULL,
                workflow_status TEXT NOT NULL DEFAULT 'Pendiente de revisión',
                revisado_por_usuario TEXT,
                fecha_revision TEXT,
                motivo_correccion TEXT,
                usuario_ultima_accion TEXT
            )
            """
        )
        columnas = {fila[1] for fila in conn.execute("PRAGMA table_info(conciliaciones)").fetchall()}
        migraciones = {
            "workflow_status": "ALTER TABLE conciliaciones ADD COLUMN workflow_status TEXT NOT NULL DEFAULT 'Pendiente de revisión'",
            "revisado_por_usuario": "ALTER TABLE conciliaciones ADD COLUMN revisado_por_usuario TEXT",
            "fecha_revision": "ALTER TABLE conciliaciones ADD COLUMN fecha_revision TEXT",
            "motivo_correccion": "ALTER TABLE conciliaciones ADD COLUMN motivo_correccion TEXT",
            "usuario_ultima_accion": "ALTER TABLE conciliaciones ADD COLUMN usuario_ultima_accion TEXT",
        }
        for nombre, sql in migraciones.items():
            if nombre not in columnas:
                conn.execute(sql)
        conn.execute("UPDATE conciliaciones SET workflow_status = 'Pendiente de revisión' WHERE workflow_status IS NULL OR workflow_status = ''")

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS empresas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                nit TEXT NOT NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS cuentas_bancarias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                banco TEXT NOT NULL,
                numero_cuenta TEXT NOT NULL,
                tipo_cuenta TEXT NOT NULL,
                empresa_id INTEGER,
                FOREIGN KEY(empresa_id) REFERENCES empresas(id)
            )
            """
        )
        conn.commit()


# =========================================================
# FUNCIONES MAESTRAS
# =========================================================

def obtener_empresas():
    with conectar_db() as conn:
        return pd.read_sql_query("SELECT id, nombre, nit FROM empresas ORDER BY nombre", conn)

def guardar_empresa(nombre, nit):
    with conectar_db() as conn:
        conn.execute("INSERT INTO empresas (nombre, nit) VALUES (?, ?)", (nombre.strip(), nit.strip()))
        conn.commit()

def eliminar_empresa(id_empresa):
    with conectar_db() as conn:
        conn.execute("DELETE FROM empresas WHERE id = ?", (int(id_empresa),))
        conn.commit()

def obtener_cuentas():
    with conectar_db() as conn:
        query = """
        SELECT c.id, c.banco, c.numero_cuenta, c.tipo_cuenta, e.nombre as empresa_nombre
        FROM cuentas_bancarias c
        LEFT JOIN empresas e ON c.empresa_id = e.id
        ORDER BY c.banco, c.numero_cuenta
        """
        return pd.read_sql_query(query, conn)

def guardar_cuenta(banco, numero_cuenta, tipo_cuenta, empresa_id):
    with conectar_db() as conn:
        conn.execute(
            "INSERT INTO cuentas_bancarias (banco, numero_cuenta, tipo_cuenta, empresa_id) VALUES (?, ?, ?, ?)",
            (banco.strip(), numero_cuenta.strip(), tipo_cuenta, empresa_id)
        )
        conn.commit()

def eliminar_cuenta(id_cuenta):
    with conectar_db() as conn:
        conn.execute("DELETE FROM cuentas_bancarias WHERE id = ?", (int(id_cuenta),))
        conn.commit()


# =========================================================
# OPERACIONES CON CONCILIACIONES
# =========================================================

def guardar_conciliacion_historial(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial,
    diferencia_conciliada, resultado_final, salidas_extracto,
    salidas_libros, entradas_libros, entradas_extracto,
    gastos_bancarios, preparado_por, revisado_por, excel_data
):
    estado = (
        "CONCILIACIÓN BANCARIA CORRECTA"
        if abs(resultado_final) < 0.005
        else "CONCILIACIÓN CON DIFERENCIA"
    )

    def dataframe_a_registros(df):
        limpio = limpiar_dataframe(df)
        if limpio is None or limpio.empty:
            return []
        return json.loads(limpio.to_json(orient="records", date_format="iso"))

    datos = {
        "salidas_extracto": dataframe_a_registros(salidas_extracto),
        "salidas_libros": dataframe_a_registros(salidas_libros),
        "entradas_libros": dataframe_a_registros(entradas_libros),
        "entradas_extracto": dataframe_a_registros(entradas_extracto),
        "gastos_bancarios": dataframe_a_registros(gastos_bancarios),
        "preparado_por": preparado_por,
        "revisado_por": revisado_por,
    }

    fecha_guardado = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fecha_elaboracion_texto = (
        fecha_elaboracion.isoformat()
        if hasattr(fecha_elaboracion, "isoformat")
        else str(fecha_elaboracion)
    )

    with conectar_db() as conn:
        cursor = conn.execute(
            """
            INSERT INTO conciliaciones (
                fecha_guardado, empresa, nit, mes, fecha_elaboracion,
                banco, cuenta, tipo, saldo_extracto, saldo_libros,
                diferencia_inicial, diferencia_conciliada, resultado_final,
                estado, datos_json, excel, workflow_status, usuario_ultima_accion
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                fecha_guardado, empresa, nit, mes, fecha_elaboracion_texto,
                banco, cuenta, tipo, float(saldo_extracto), float(saldo_libros),
                float(diferencia_inicial), float(diferencia_conciliada),
                float(resultado_final), estado,
                json.dumps(datos, ensure_ascii=False),
                sqlite3.Binary(excel_data),
                "Pendiente de revisión",
                preparado_por or None,
            )
        )
        conn.commit()
        return cursor.lastrowid


def obtener_historial():
    with conectar_db() as conn:
        return pd.read_sql_query(
            """
            SELECT id, fecha_guardado, empresa, nit, mes, banco, cuenta,
                   resultado_final, estado, workflow_status, revisado_por_usuario,
                   fecha_revision, motivo_correccion
            FROM conciliaciones
            ORDER BY id DESC
            """, conn
        )


def obtener_conciliacion(id_conciliacion):
    with conectar_db() as conn:
        return conn.execute(
            """
            SELECT id, fecha_guardado, empresa, nit, mes, fecha_elaboracion,
                   banco, cuenta, tipo, saldo_extracto, saldo_libros,
                   diferencia_inicial, diferencia_conciliada, resultado_final,
                   estado, datos_json, excel, workflow_status, revisado_por_usuario,
                   fecha_revision, motivo_correccion, usuario_ultima_accion
            FROM conciliaciones WHERE id = ?
            """, (int(id_conciliacion),)
        ).fetchone()


def eliminar_conciliacion(id_conciliacion):
    with conectar_db() as conn:
        conn.execute("DELETE FROM conciliaciones WHERE id = ?", (int(id_conciliacion),))
        conn.commit()


inicializar_db()


# =========================================================
# AUTENTICACIÓN Y USUARIOS
# =========================================================

def preparar_tabla_usuarios():
    with conectar_db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT NOT NULL UNIQUE,
                nombre TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                rol TEXT NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1,
                fecha_creacion TEXT NOT NULL
            )
            """
        )
        conn.commit()


def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000
    ).hex()
    return salt, digest


def verificar_password(password, salt, password_hash):
    _, digest = hash_password(password, salt)
    return hmac.compare_digest(digest, password_hash)


def contar_usuarios():
    with conectar_db() as conn:
        return conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]


def crear_usuario(usuario, nombre, password, rol):
    salt, password_hash = hash_password(password)
    with conectar_db() as conn:
        cursor = conn.execute(
            """
            INSERT INTO usuarios
            (usuario, nombre, password_hash, salt, rol, activo, fecha_creacion)
            VALUES (?, ?, ?, ?, ?, 1, ?)
            """,
            (usuario.strip(), nombre.strip(), password_hash, salt, rol,
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        return cursor.lastrowid


def autenticar_usuario(usuario, password):
    with conectar_db() as conn:
        fila = conn.execute(
            """
            SELECT id, usuario, nombre, password_hash, salt, rol
            FROM usuarios WHERE usuario = ? AND activo = 1
            """,
            (usuario.strip(),)
        ).fetchone()
    if not fila or not verificar_password(password, fila[4], fila[3]):
        return None
    return {"id": fila[0], "usuario": fila[1], "nombre": fila[2], "rol": fila[5]}


def obtener_usuarios():
    with conectar_db() as conn:
        return pd.read_sql_query(
            "SELECT id, usuario, nombre, rol, activo, fecha_creacion FROM usuarios ORDER BY id",
            conn
        )


def tiene_permiso(rol, permiso):
    permisos = {
        "Administrador": {"editar", "historial", "eliminar", "usuarios", "maestros"},
        "Preparador": {"editar", "historial", "maestros"},
        "Revisor": {"historial"},
    }
    return permiso in permisos.get(rol, set())


def iniciar_autenticacion():
    preparar_tabla_usuarios()
    if "usuario_autenticado" not in st.session_state:
        st.session_state.usuario_autenticado = None

    if contar_usuarios() == 0:
        st.title("🔐 Configuración inicial")
        st.info("Crea el primer usuario con rol Administrador para comenzar.")
        with st.form("form_primer_admin"):
            usuario = st.text_input("Usuario")
            nombre = st.text_input("Nombre completo")
            password = st.text_input("Contraseña", type="password")
            confirmar = st.text_input("Confirmar contraseña", type="password")
            crear = st.form_submit_button("Crear administrador", type="primary", width="stretch")
        if crear:
            if not usuario.strip() or not nombre.strip() or not password:
                st.error("Completa todos los campos.")
            elif len(password) < 8:
                st.error("La contraseña debe tener al menos 8 caracteres.")
            elif password != confirmar:
                st.error("Las contraseñas no coinciden.")
            else:
                crear_usuario(usuario, nombre, password, "Administrador")
                datos_admin = autenticar_usuario(usuario, password)
                st.session_state.usuario_autenticado = datos_admin
                st.success(f"Administrador creado correctamente.")
                st.rerun()
        st.stop()

    if st.session_state.usuario_autenticado is None:
        st.title("🔐 Inicio de sesión")
        st.caption("Ingresa tus credenciales para acceder a Conciliación Bancaria.")
        with st.form("form_login"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Iniciar sesión", type="primary", width="stretch")
        if entrar:
            datos = autenticar_usuario(usuario, password)
            if datos:
                st.session_state.usuario_autenticado = datos
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.stop()


iniciar_autenticacion()
usuario_actual = st.session_state.usuario_autenticado
rol_actual = usuario_actual["rol"]


# =========================================================
# MENÚ LATERAL Y NAVEGACIÓN
# =========================================================

with st.sidebar:
    st.image("https://img.icons8.com/color/96/bank-building.png", width=70)
    st.title("Conciliación Web")
    st.markdown(f"👤 **{usuario_actual['nombre']}**")
    st.caption(f"Rol: {rol_actual}")
    st.divider()

    opciones_menu = ["📊 Dashboard"]
    if tiene_permiso(rol_actual, "editar"):
        opciones_menu.append("📝 Nueva Conciliación")
    
    opciones_menu.extend(["📚 Historial", "🏢 Empresas", "🏦 Bancos y Cuentas", "📈 Reportes"])
    if tiene_permiso(rol_actual, "usuarios"):
        opciones_menu.append("👥 Usuarios")

    menu_seleccionado = st.radio("Navegación principal", opciones_menu)
    st.divider()
    if st.button("🚪 Cerrar sesión", width="stretch"):
        st.session_state.usuario_autenticado = None
        st.rerun()


# =========================================================
# FUNCIONES AUXILIARES PARA EXCEL Y FORMATO
# =========================================================

def total_columna(df, columna="Valor"):
    if columna not in df.columns:
        return 0.0
    return float(pd.to_numeric(df[columna], errors="coerce").fillna(0).sum())


def limpiar_dataframe(df):
    if df is None or df.empty:
        return df.copy()
    resultado = df.copy()
    for columna in resultado.columns:
        if resultado[columna].dtype == "object":
            resultado[columna] = resultado[columna].replace(r"^\s*$", None, regex=True)
    return resultado.dropna(how="all").reset_index(drop=True)


def limpiar_nombre_archivo(texto):
    texto = str(texto).strip()
    if not texto:
        texto = "Empresa"
    return re.sub(r'[<>:"/\\|?*]', "-", texto)


def estilo_titulo(celda, tamano=12):
    celda.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    celda.font = Font(bold=True, color="FFFFFF", size=tamano)
    celda.alignment = Alignment(horizontal="center", vertical="center")


def estilo_encabezado(celda):
    celda.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    celda.font = Font(bold=True)
    celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def escribir_seccion(ws, fila, titulo, columnas=4):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=columnas)
    celda = ws.cell(row=fila, column=1)
    celda.value = titulo
    estilo_titulo(celda)
    ws.row_dimensions[fila].height = 22
    return fila + 1


def escribir_tabla(ws, fila, titulo, df):
    df = limpiar_dataframe(df)
    numero_columnas = max(4, len(df.columns))
    fila = escribir_seccion(ws, fila, titulo, numero_columnas)

    for columna, nombre in enumerate(df.columns, start=1):
        celda = ws.cell(row=fila, column=columna)
        celda.value = nombre
        estilo_encabezado(celda)

    fila += 1
    for _, registro in df.iterrows():
        for columna, nombre in enumerate(df.columns, start=1):
            valor = registro[nombre]
            if pd.isna(valor):
                valor = ""
            celda = ws.cell(row=fila, column=columna)
            celda.value = valor
            if nombre == "Valor":
                celda.number_format = '#,##0.00'
        fila += 1

    if df.empty:
        fila += 1

    columna_valor = None
    for columna, nombre in enumerate(df.columns, start=1):
        if nombre == "Valor":
            columna_valor = columna
            break

    if columna_valor is not None:
        ws.cell(row=fila, column=columna_valor - 1).value = "TOTAL"
        ws.cell(row=fila, column=columna_valor - 1).font = Font(bold=True)
        ws.cell(row=fila, column=columna_valor).value = total_columna(df, "Valor")
        ws.cell(row=fila, column=columna_valor).number_format = '#,##0.00'
        ws.cell(row=fila, column=columna_valor).font = Font(bold=True)

    return fila + 2


def escribir_gastos_bancarios(ws, fila, df):
    df = limpiar_dataframe(df)
    columnas = ["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"]
    fila = escribir_seccion(ws, fila, "GASTOS BANCARIOS", len(columnas))

    for columna, nombre in enumerate(columnas, start=1):
        celda = ws.cell(row=fila, column=columna)
        celda.value = nombre
        estilo_encabezado(celda)

    fila += 1
    for _, registro in df.iterrows():
        for columna, nombre in enumerate(columnas, start=1):
            valor = registro.get(nombre, "")
            if pd.isna(valor):
                valor = ""
            celda = ws.cell(row=fila, column=columna)
            celda.value = valor
            if nombre != "Fecha":
                celda.number_format = '#,##0.00'
        fila += 1

    if df.empty:
        fila += 1

    fila_total = fila
    ws.cell(row=fila_total, column=1).value = "TOTAL"
    ws.cell(row=fila_total, column=1).font = Font(bold=True)

    for columna, nombre in enumerate(columnas[1:], start=2):
        total = float(pd.to_numeric(df.get(nombre, pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        ws.cell(row=fila_total, column=columna).value = total
        ws.cell(row=fila_total, column=columna).number_format = '#,##0.00'
        ws.cell(row=fila_total, column=columna).font = Font(bold=True)

    return fila_total + 2


def preparar_excel(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
    resultado_final, salidas_extracto, salidas_libros, entradas_libros,
    entradas_extracto, gastos_bancarios, preparado_por, revisado_por,
    nombres_titulos
):
    salidas_extracto = limpiar_dataframe(salidas_extracto)
    salidas_libros = limpiar_dataframe(salidas_libros)
    entradas_libros = limpiar_dataframe(entradas_libros)
    entradas_extracto = limpiar_dataframe(entradas_extracto)
    gastos_bancarios = limpiar_dataframe(gastos_bancarios)

    workbook = Workbook()
    ws = workbook.active
    ws.title = "Conciliación Bancaria"

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    anchos = {"A": 15, "B": 28, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    ws.merge_cells("A1:G1")
    ws["A1"] = f"CONCILIACIÓN - {tipo.upper()}"
    estilo_titulo(ws["A1"], 16)
    ws.row_dimensions[1].height = 28

    fila = escribir_seccion(ws, 3, "INFORMACIÓN GENERAL", 7)
    datos_generales = [
        ("Empresa", empresa), ("NIT", nit), ("Mes y Año", mes),
        ("Fecha de elaboración", fecha_elaboracion), ("Banco", banco),
        ("Cuenta No.", cuenta), ("Tipo", tipo),
    ]

    for nombre, valor in datos_generales:
        ws.cell(row=fila, column=1).value = nombre
        ws.cell(row=fila, column=1).font = Font(bold=True)
        ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=4)
        ws.cell(row=fila, column=2).value = valor
        fila += 1

    fila += 1
    fila = escribir_seccion(ws, fila, "SALDOS", 4)
    saldos = [
        ("SALDO SEGÚN EXTRACTO BANCARIO", saldo_extracto),
        ("SALDO SEGÚN LIBROS", saldo_libros),
        ("DIFERENCIA A JUSTIFICAR", diferencia_inicial),
    ]

    for nombre, valor in saldos:
        ws.cell(row=fila, column=1).value = nombre
        ws.cell(row=fila, column=1).font = Font(bold=True)
        ws.cell(row=fila, column=2).value = valor
        ws.cell(row=fila, column=2).number_format = '#,##0.00'
        fila += 1

    fila += 1
    fila = escribir_seccion(ws, fila, "JUSTIFICACIÓN", 4)
    
    justificaciones = [
        (nombres_titulos["t1"], total_columna(salidas_extracto)),
        (nombres_titulos["t2"], total_columna(salidas_libros)),
        (nombres_titulos["t3"], total_columna(entradas_libros)),
        (nombres_titulos["t4"], total_columna(entradas_extracto)),
    ]

    for nombre, valor in justificaciones:
        ws.cell(row=fila, column=1).value = nombre
        ws.cell(row=fila, column=1).font = Font(bold=True)
        ws.cell(row=fila, column=2).value = valor
        ws.cell(row=fila, column=2).number_format = '#,##0.00'
        fila += 1

    fila += 1
    fila = escribir_seccion(ws, fila, "CÁLCULO DE LA CONCILIACIÓN", 4)
    calculo = [
        ("DIFERENCIA A JUSTIFICAR", diferencia_inicial),
        ("DIFERENCIA CONCILIADA", diferencia_conciliada),
        ("RESULTADO FINAL", resultado_final),
    ]

    for nombre, valor in calculo:
        ws.cell(row=fila, column=1).value = nombre
        ws.cell(row=fila, column=1).font = Font(bold=True)
        ws.cell(row=fila, column=2).value = valor
        ws.cell(row=fila, column=2).number_format = '#,##0.00'
        fila += 1

    fila += 1
    fila = escribir_tabla(ws, fila, nombres_titulos["t1"], salidas_extracto)
    fila = escribir_tabla(ws, fila, nombres_titulos["t2"], salidas_libros)
    fila = escribir_tabla(ws, fila, nombres_titulos["t3"], entradas_libros)
    fila = escribir_tabla(ws, fila, nombres_titulos["t4"], entradas_extracto)
    fila = escribir_gastos_bancarios(ws, fila, gastos_bancarios)

    fila = escribir_seccion(ws, fila, "RESULTADO DE LA CONCILIACIÓN", 4)
    ws.cell(row=fila, column=1).value = "DIFERENCIA A JUSTIFICAR"
    ws.cell(row=fila, column=2).value = diferencia_inicial
    ws.cell(row=fila, column=2).number_format = '#,##0.00'
    fila += 1
    ws.cell(row=fila, column=1).value = "DIFERENCIA CONCILIADA"
    ws.cell(row=fila, column=2).value = diferencia_conciliada
    ws.cell(row=fila, column=2).number_format = '#,##0.00'
    fila += 1
    ws.cell(row=fila, column=1).value = "RESULTADO FINAL"
    ws.cell(row=fila, column=2).value = resultado_final
    ws.cell(row=fila, column=2).number_format = '#,##0.00'

    estado = "CONCILIACIÓN BANCARIA CORRECTA" if abs(resultado_final) < 0.005 else "CONCILIACIÓN CON DIFERENCIA"
    fila += 1
    ws.cell(row=fila, column=1).value = "ESTADO"
    ws.cell(row=fila, column=1).font = Font(bold=True)
    ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=4)
    ws.cell(row=fila, column=2).value = estado
    ws.cell(row=fila, column=2).font = Font(bold=True)

    fila += 2
    ws.cell(row=fila, column=1).value = "Preparado por:"
    ws.cell(row=fila, column=1).font = Font(bold=True)
    ws.cell(row=fila, column=2).value = preparado_por
    ws.cell(row=fila, column=5).value = "Revisado por:"
    ws.cell(row=fila, column=5).font = Font(bold=True)
    ws.cell(row=fila, column=6).value = revisado_por

    borde = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
    )
    for fila_excel in ws.iter_rows():
        for celda in fila_excel:
            if celda.value is not None:
                celda.border = borde
                celda.alignment = Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = "A3"
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    nombre_archivo = f"CONCILIACION_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.xlsx"
    return output.getvalue(), nombre_archivo


# =========================================================
# VISTAS
# =========================================================

if menu_seleccionado == "📊 Dashboard":
    st.title("📊 Dashboard de Conciliaciones")
    historial = obtener_historial()
    if len(historial) == 0:
        st.info("Aún no hay conciliaciones registradas.")
    else:
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Total Conciliaciones", len(historial))
        kpi2.metric("Aprobadas 🔒", len(historial[historial["workflow_status"] == "Aprobada"]))
        kpi3.metric("Pendientes 🕒", len(historial[historial["workflow_status"] == "Pendiente de revisión"]))
        st.divider()
        st.dataframe(historial[["id", "empresa", "mes", "banco", "resultado_final", "workflow_status"]].head(5), width="stretch", hide_index=True)


elif menu_seleccionado == "🏢 Empresas":
    st.title("🏢 Maestro de Empresas")
    if tiene_permiso(rol_actual, "maestros"):
        with st.expander("➕ Registrar nueva empresa"):
            with st.form("form_empresa_maestro"):
                nom = st.text_input("Nombre de la empresa")
                nit = st.text_input("NIT")
                if st.form_submit_button("Guardar Empresa", type="primary"):
                    guardar_empresa(nom, nit)
                    st.rerun()
    st.dataframe(obtener_empresas(), width="stretch", hide_index=True)


elif menu_seleccionado == "🏦 Bancos y Cuentas":
    st.title("🏦 Maestro de Bancos y Cuentas Bancarias")
    empresas_df = obtener_empresas()
    if tiene_permiso(rol_actual, "maestros"):
        with st.expander("➕ Registrar nueva cuenta bancaria"):
            with st.form("form_cuenta_maestro"):
                banco = st.text_input("Banco")
                num = st.text_input("Número de Cuenta / Tarjeta")
                tipo = st.selectbox("Tipo de Cuenta", ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito"])
                emp_id = None
                if not empresas_df.empty:
                    emp_id = st.selectbox("Asociar a Empresa", empresas_df["id"].tolist(), format_func=lambda x: empresas_df.loc[empresas_df["id"]==x, "nombre"].values[0])
                if st.form_submit_button("Guardar Cuenta", type="primary"):
                    guardar_cuenta(banco, num, tipo, emp_id)
                    st.rerun()
    st.dataframe(obtener_cuentas(), width="stretch", hide_index=True)


elif menu_seleccionado == "📝 Nueva Conciliación":
    st.title("📝 Captura de Conciliación Bancaria")

    if "tabla1" not in st.session_state:
        st.session_state.tabla1 = pd.DataFrame(columns=["Fecha", "Beneficiario", "Documento", "Valor"])
    if "tabla2" not in st.session_state:
        st.session_state.tabla2 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    if "tabla3" not in st.session_state:
        st.session_state.tabla3 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    if "tabla4" not in st.session_state:
        st.session_state.tabla4 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    if "tabla5" not in st.session_state:
        st.session_state.tabla5 = pd.DataFrame(columns=["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"])

    empresas_df = obtener_empresas()
    cuentas_df = obtener_cuentas()

    st.subheader("Información General")
    col1, col2 = st.columns(2)

    with col1:
        if not empresas_df.empty:
            empresa_obj = st.selectbox("Empresa Registrada", empresas_df["nombre"].tolist())
            empresa = empresa_obj
            nit = empresas_df.loc[empresas_df["nombre"] == empresa_obj, "nit"].values[0]
            st.text_input("NIT", value=nit, disabled=True)
        else:
            empresa = st.text_input("Nombre de la Empresa", key="form_empresa")
            nit = st.text_input("NIT", key="form_nit")

        mes = st.text_input("Mes y Año (ej. SEPTIEMBRE 2026)", key="form_mes")
        fecha_elaboracion = st.date_input("Fecha de elaboración", key="form_fecha_elaboracion")

    with col2:
        if not cuentas_df.empty:
            cta_sel = st.selectbox(
                "Cuenta Bancaria Registrada",
                cuentas_df["id"].tolist(),
                format_func=lambda x: f"{cuentas_df.loc[cuentas_df['id']==x, 'banco'].values[0]} - {cuentas_df.loc[cuentas_df['id']==x, 'numero_cuenta'].values[0]} ({cuentas_df.loc[cuentas_df['id']==x, 'tipo_cuenta'].values[0]})"
            )
            banco = cuentas_df.loc[cuentas_df["id"] == cta_sel, "banco"].values[0]
            cuenta = cuentas_df.loc[cuentas_df["id"] == cta_sel, "numero_cuenta"].values[0]
            tipo = cuentas_df.loc[cuentas_df["id"] == cta_sel, "tipo_cuenta"].values[0]
        else:
            banco = st.text_input("Nombre del Banco", key="form_banco")
            cuenta = st.text_input("Número de Cuenta", key="form_cuenta")
            tipo = st.selectbox("Tipo de Cuenta", ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito"], key="form_tipo")

    # CONFIGURACIÓN DINÁMICA DE TÍTULOS Y FÓRMULAS
    es_tc = "tarjeta" in tipo.lower() or "crédito" in tipo.lower() or "credito" in tipo.lower()

    if es_tc:
        st.info("💳 Modo activado: Conciliación de Tarjeta de Crédito.")
        nombres_titulos = {
            "t1": "COMPRAS NO EVIDENCIADAS EN EXTRACTOS",
            "t2": "COMPRAS NO CONTABILIZADAS EN LIBROS",
            "t3": "DÉBITOS BANCARIOS NO CONTABILIZADOS EN LIBROS",
            "t4": "ABONOS NO REGISTRADOS EN EXTRACTO"
        }
    else:
        nombres_titulos = {
            "t1": "SALIDAS NO REGISTRADAS EN EXTRACTO",
            "t2": "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
            "t3": "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
            "t4": "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS"
        }

    st.divider()
    st.subheader("Saldos")
    c1, c2 = st.columns(2)
    with c1:
        saldo_extracto = st.number_input("Saldo según Extracto", format="%.2f", key="form_saldo_extracto")
    with c2:
        saldo_libros = st.number_input("Saldo según Libros", format="%.2f", key="form_saldo_libros")

    diferencia_inicial = saldo_extracto - saldo_libros
    st.metric("Diferencia a Justificar", f"${diferencia_inicial:,.2f}")

    # TABLAS DE MOVIMIENTOS CON TÍTULOS DINÁMICOS
    st.divider()
    st.subheader(f"1. {nombres_titulos['t1']}")
    salidas_extracto = st.data_editor(st.session_state.tabla1, num_rows="dynamic", width="stretch", key="editor_tabla1")
    st.session_state.tabla1 = salidas_extracto

    st.divider()
    st.subheader(f"2. {nombres_titulos['t2']}")
    salidas_libros = st.data_editor(st.session_state.tabla2, num_rows="dynamic", width="stretch", key="editor_tabla2")
    st.session_state.tabla2 = salidas_libros

    st.divider()
    st.subheader(f"3. {nombres_titulos['t3']}")
    entradas_libros = st.data_editor(st.session_state.tabla3, num_rows="dynamic", width="stretch", key="editor_tabla3")
    st.session_state.tabla3 = entradas_libros

    st.divider()
    st.subheader(f"4. {nombres_titulos['t4']}")
    entradas_extracto = st.data_editor(st.session_state.tabla4, num_rows="dynamic", width="stretch", key="editor_tabla4")
    st.session_state.tabla4 = entradas_extracto

    st.divider()
    st.subheader("Gastos Bancarios")
    gastos_bancarios = st.data_editor(st.session_state.tabla5, num_rows="dynamic", width="stretch", key="editor_tabla5")
    st.session_state.tabla5 = gastos_bancarios

    # CÁLCULO DE LA DIFERENCIA SEGÚN TIPO DE CUENTA
    m1 = total_columna(salidas_extracto)
    m2 = total_columna(salidas_libros)
    m3 = total_columna(entradas_libros)
    m4 = total_columna(entradas_extracto)

    if es_tc:
        # FÓRMULA TARJETA DE CRÉDITO: - t1 + t2 - t3 + t4
        diferencia_conciliada = - m1 + m2 - m3 + m4
    else:
        # FÓRMULA ESTÁNDAR BANCO: + t1 - t2 + t3 - t4
        diferencia_conciliada = m1 - m2 + m3 - m4

    resultado_final = diferencia_inicial - diferencia_conciliada

    st.divider()
    st.subheader("Resultado de la Conciliación")
    r1, r2, r3 = st.columns(3)
    r1.metric("Diferencia Inicial", f"${diferencia_inicial:,.2f}")
    r2.metric("Diferencia Conciliada", f"${diferencia_conciliada:,.2f}")
    r3.metric("Resultado Final", f"${resultado_final:,.2f}")

    preparado_por = st.text_input("Preparado por", value=usuario_actual["nombre"], key="form_preparado_por")
    revisado_por = st.text_input("Revisado por", key="form_revisado_por")

    excel_data, nombre_archivo = preparar_excel(
        empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
        saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
        resultado_final, salidas_extracto, salidas_libros, entradas_libros,
        entradas_extracto, gastos_bancarios, preparado_por, revisado_por,
        nombres_titulos
    )

    st.divider()
    if st.button("💾 Guardar Conciliación en Historial", type="primary", width="stretch"):
        id_g = guardar_conciliacion_historial(
            empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
            saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
            resultado_final, salidas_extracto, salidas_libros, entradas_libros,
            entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data
        )
        st.success(f"✅ Conciliación #{id_g} guardada con éxito.")


elif menu_seleccionado == "📚 Historial":
    st.title("📚 Historial de Conciliaciones")
    historial = obtener_historial()
    if historial.empty:
        st.info("No hay conciliaciones guardadas.")
    else:
        st.dataframe(historial, width="stretch", hide_index=True)


elif menu_seleccionado == "📈 Reportes":
    st.title("📈 Reportes")
    historial = obtener_historial()
    if not historial.empty:
        csv = historial.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Todo el Historial (CSV)", data=csv, file_name="historial_conciliaciones.csv", mime="text/csv", width="stretch")


elif menu_seleccionado == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios")
    st.dataframe(obtener_usuarios(), width="stretch", hide_index=True)
