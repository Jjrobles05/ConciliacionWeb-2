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

from reportlab.lib.pagesizes import letter, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


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
        columnas_conc = {fila[1] for fila in conn.execute("PRAGMA table_info(conciliaciones)").fetchall()}
        migraciones_conc = {
            "workflow_status": "ALTER TABLE conciliaciones ADD COLUMN workflow_status TEXT NOT NULL DEFAULT 'Pendiente de revisión'",
            "revisado_por_usuario": "ALTER TABLE conciliaciones ADD COLUMN revisado_por_usuario TEXT",
            "fecha_revision": "ALTER TABLE conciliaciones ADD COLUMN fecha_revision TEXT",
            "motivo_correccion": "ALTER TABLE conciliaciones ADD COLUMN motivo_correccion TEXT",
            "usuario_ultima_accion": "ALTER TABLE conciliaciones ADD COLUMN usuario_ultima_accion TEXT",
        }
        for nombre, sql in migraciones_conc.items():
            if nombre not in columnas_conc:
                conn.execute(sql)

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS empresas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                nit TEXT NOT NULL,
                logo BLOB
            )
            """
        )
        cols_emp = {fila[1] for fila in conn.execute("PRAGMA table_info(empresas)").fetchall()}
        if "logo" not in cols_emp:
            conn.execute("ALTER TABLE empresas ADD COLUMN logo BLOB")

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
                fecha_creacion TEXT NOT NULL,
                empresa_id INTEGER,
                FOREIGN KEY(empresa_id) REFERENCES empresas(id)
            )
            """
        )
        cols_usr = {fila[1] for fila in conn.execute("PRAGMA table_info(usuarios)").fetchall()}
        if "empresa_id" not in cols_usr:
            conn.execute("ALTER TABLE usuarios ADD COLUMN empresa_id INTEGER")

        conn.commit()


# =========================================================
# FUNCIONES MAESTRAS
# =========================================================

def obtener_empresas():
    with conectar_db() as conn:
        return pd.read_sql_query("SELECT id, nombre, nit, logo FROM empresas ORDER BY nombre", conn)

def obtener_empresa_por_id(empresa_id):
    with conectar_db() as conn:
        res = conn.execute("SELECT id, nombre, nit, logo FROM empresas WHERE id = ?", (int(empresa_id),)).fetchone()
        if res:
            return {"id": res[0], "nombre": res[1], "nit": res[2], "logo": res[3]}
    return None

def obtener_logo_empresa(nombre_empresa):
    with conectar_db() as conn:
        res = conn.execute("SELECT logo FROM empresas WHERE nombre = ?", (nombre_empresa,)).fetchone()
        if res and res[0]:
            return res[0]
    return None

def guardar_empresa(nombre, nit, logo_bytes=None):
    with conectar_db() as conn:
        if logo_bytes:
            conn.execute("INSERT INTO empresas (nombre, nit, logo) VALUES (?, ?, ?)", (nombre.strip(), nit.strip(), sqlite3.Binary(logo_bytes)))
        else:
            conn.execute("INSERT INTO empresas (nombre, nit) VALUES (?, ?)", (nombre.strip(), nit.strip()))
        conn.commit()

def actualizar_empresa_db(empresa_id, nombre, nit, logo_bytes=None):
    with conectar_db() as conn:
        if logo_bytes:
            conn.execute("UPDATE empresas SET nombre = ?, nit = ?, logo = ? WHERE id = ?", (nombre.strip(), nit.strip(), sqlite3.Binary(logo_bytes), int(empresa_id)))
        else:
            conn.execute("UPDATE empresas SET nombre = ?, nit = ? WHERE id = ?", (nombre.strip(), nit.strip(), int(empresa_id)))
        conn.commit()

def eliminar_empresa_db(empresa_id):
    with conectar_db() as conn:
        conn.execute("DELETE FROM empresas WHERE id = ?", (int(empresa_id),))
        conn.commit()

def eliminar_conciliacion_db(conciliacion_id):
    with conectar_db() as conn:
        conn.execute("DELETE FROM conciliaciones WHERE id = ?", (int(conciliacion_id),))
        conn.commit()

def obtener_cuentas(empresa_id=None):
    with conectar_db() as conn:
        if empresa_id:
            query = """
            SELECT c.id, c.banco, c.numero_cuenta, c.tipo_cuenta, c.empresa_id, e.nombre as empresa_nombre
            FROM cuentas_bancarias c
            LEFT JOIN empresas e ON c.empresa_id = e.id
            WHERE c.empresa_id = ?
            ORDER BY c.banco, c.numero_cuenta
            """
            return pd.read_sql_query(query, conn, params=(int(empresa_id),))
        else:
            query = """
            SELECT c.id, c.banco, c.numero_cuenta, c.tipo_cuenta, c.empresa_id, e.nombre as empresa_nombre
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


# =========================================================
# ROTACIÓN AUTOMÁTICA MENSUAL DE CUENTAS (EXCLUSIVO PREPARADORES ACTIVOS)
# =========================================================

def obtener_cuentas_rotadas_por_usuario(empresa_id, mes_num, usuario_id):
    if not empresa_id:
        return pd.DataFrame()

    with conectar_db() as conn:
        usuarios = pd.read_sql_query(
            "SELECT id, nombre FROM usuarios WHERE empresa_id = ? AND activo = 1 AND rol = 'Preparador' ORDER BY id",
            conn, params=(int(empresa_id),)
        )
        cuentas = pd.read_sql_query(
            "SELECT id, banco, numero_cuenta, tipo_cuenta FROM cuentas_bancarias WHERE empresa_id = ? ORDER BY id",
            conn, params=(int(empresa_id),)
        )

    if cuentas.empty:
        return pd.DataFrame()

    if usuarios.empty or usuario_id not in usuarios["id"].tolist():
        return cuentas

    num_usuarios = len(usuarios)
    ids_usuarios = usuarios["id"].tolist()

    cuentas_asignadas = []
    for idx_cuenta, fila_cuenta in cuentas.iterrows():
        idx_usuario_asignado = (idx_cuenta + mes_num) % num_usuarios
        usuario_asignado_id = ids_usuarios[idx_usuario_asignado]

        if usuario_asignado_id == usuario_id:
            cuentas_asignadas.append(fila_cuenta)

    if cuentas_asignadas:
        return pd.DataFrame(cuentas_asignadas)
    return pd.DataFrame(columns=cuentas.columns)


# =========================================================
# FUNCIÓN PARA PEGAR ÍTEMS DE EXCEL DIRECTAMENTE
# =========================================================

def parsear_texto_pegado(texto, columnas_esperadas):
    if not texto or not texto.strip():
        return None
    lines = [l.strip() for l in texto.strip().splitlines() if l.strip()]
    rows = []
    for line in lines:
        parts = line.split('\t') if '\t' in line else line.split(',')
        parts = [p.strip() for p in parts]
        
        if len(parts) < len(columnas_esperadas):
            parts += [""] * (len(columnas_esperadas) - len(parts))
        else:
            parts = parts[:len(columnas_esperadas)]
            
        row_dict = {}
        for idx, col in enumerate(columnas_esperadas):
            val = parts[idx]
            if col == "Valor" or col not in ["Fecha", "Beneficiario", "Documento", "Concepto"]:
                val_limpio = val.replace("$", "").replace(".", "").replace(",", ".").replace(" ", "")
                try:
                    row_dict[col] = float(val_limpio)
                except ValueError:
                    row_dict[col] = 0.0
            else:
                row_dict[col] = val
        rows.append(row_dict)
    return pd.DataFrame(rows)


# =========================================================
# OPERACIONES DE CONCILIACIÓN Y AUDITORÍA
# =========================================================

def guardar_conciliacion_historial(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial,
    diferencia_conciliada, resultado_final, salidas_extracto,
    salidas_libros, entradas_libros, entradas_extracto,
    gastos_bancarios, preparado_por, revisado_por, excel_data,
    workflow_status="Pendiente de revisión", id_edicion=None
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
        if id_edicion:
            conn.execute(
                """
                UPDATE conciliaciones
                SET fecha_guardado=?, empresa=?, nit=?, mes=?, fecha_elaboracion=?,
                    banco=?, cuenta=?, tipo=?, saldo_extracto=?, saldo_libros=?,
                    diferencia_inicial=?, diferencia_conciliada=?, resultado_final=?,
                    estado=?, datos_json=?, excel=?, workflow_status=?,
                    usuario_ultima_accion=?
                WHERE id=?
                """,
                (
                    fecha_guardado, empresa, nit, mes, fecha_elaboracion_texto,
                    banco, cuenta, tipo, float(saldo_extracto), float(saldo_libros),
                    float(diferencia_inicial), float(diferencia_conciliada),
                    float(resultado_final), estado,
                    json.dumps(datos, ensure_ascii=False),
                    sqlite3.Binary(excel_data), workflow_status, preparado_por or None, int(id_edicion)
                )
            )
            conn.commit()
            return id_edicion
        else:
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
                    workflow_status,
                    preparado_por or None,
                )
            )
            conn.commit()
            return cursor.lastrowid


def obtener_historial(empresa_nombre=None):
    with conectar_db() as conn:
        if empresa_nombre and empresa_nombre != "Todas las empresas":
            return pd.read_sql_query(
                """
                SELECT id, fecha_guardado, empresa, nit, mes, banco, cuenta, tipo,
                       saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                       resultado_final, estado, workflow_status, revisado_por_usuario,
                       fecha_revision, motivo_correccion, datos_json, fecha_elaboracion, excel
                FROM conciliaciones
                WHERE empresa = ?
                ORDER BY id DESC
                """, conn, params=(empresa_nombre,)
            )
        else:
            return pd.read_sql_query(
                """
                SELECT id, fecha_guardado, empresa, nit, mes, banco, cuenta, tipo,
                       saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                       resultado_final, estado, workflow_status, revisado_por_usuario,
                       fecha_revision, motivo_correccion, datos_json, fecha_elaboracion, excel
                FROM conciliaciones
                ORDER BY id DESC
                """, conn
            )


def obtener_conciliacion_por_id(id_conciliacion):
    with conectar_db() as conn:
        fila = conn.execute(
            "SELECT * FROM conciliaciones WHERE id = ?", (int(id_conciliacion),)
        ).fetchone()
        if not fila:
            return None
        cols = [col[1] for col in conn.execute("PRAGMA table_info(conciliaciones)").fetchall()]
        return dict(zip(cols, fila))


def actualizar_estado_auditoria(id_conciliacion, nuevo_estado, revisado_por, motivo_correccion=None):
    fecha_rev = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with conectar_db() as conn:
        conn.execute(
            """
            UPDATE conciliaciones
            SET workflow_status = ?, revisado_por_usuario = ?, fecha_revision = ?, motivo_correccion = ?
            WHERE id = ?
            """,
            (nuevo_estado, revisado_por, fecha_rev, motivo_correccion, int(id_conciliacion))
        )
        conn.commit()


inicializar_db()


# =========================================================
# AUTENTICACIÓN Y USUARIOS
# =========================================================

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


def crear_usuario(usuario, nombre, password, rol, empresa_id=None):
    salt, password_hash = hash_password(password)
    with conectar_db() as conn:
        cursor = conn.execute(
            """
            INSERT INTO usuarios
            (usuario, nombre, password_hash, salt, rol, activo, fecha_creacion, empresa_id)
            VALUES (?, ?, ?, ?, ?, 1, ?, ?)
            """,
            (usuario.strip(), nombre.strip(), password_hash, salt, rol,
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"), empresa_id)
        )
        conn.commit()
        return cursor.lastrowid


def actualizar_empresa_usuario(usuario_id, empresa_id):
    with conectar_db() as conn:
        conn.execute("UPDATE usuarios SET empresa_id = ? WHERE id = ?", (empresa_id, int(usuario_id)))
        conn.commit()


def actualizar_rol_usuario(usuario_id, nuevo_rol):
    with conectar_db() as conn:
        conn.execute("UPDATE usuarios SET rol = ? WHERE id = ?", (nuevo_rol, int(usuario_id)))
        conn.commit()


def autenticar_usuario(usuario, password):
    with conectar_db() as conn:
        fila = conn.execute(
            """
            SELECT u.id, u.usuario, u.nombre, u.password_hash, u.salt, u.rol, u.empresa_id, e.nombre as empresa_nombre, e.nit as empresa_nit
            FROM usuarios u
            LEFT JOIN empresas e ON u.empresa_id = e.id
            WHERE u.usuario = ? AND u.activo = 1
            """,
            (usuario.strip(),)
        ).fetchone()
    if not fila or not verificar_password(password, fila[4], fila[3]):
        return None
    return {
        "id": fila[0],
        "usuario": fila[1],
        "nombre": fila[2],
        "rol": fila[5],
        "empresa_id": fila[6],
        "empresa_nombre": fila[7],
        "empresa_nit": fila[8]
    }


def obtener_usuarios():
    with conectar_db() as conn:
        query = """
        SELECT u.id, u.usuario, u.nombre, u.rol, u.activo, u.fecha_creacion, e.nombre as empresa_nombre
        FROM usuarios u
        LEFT JOIN empresas e ON u.empresa_id = e.id
        ORDER BY u.id
        """
        return pd.read_sql_query(query, conn)


def obtener_opciones_menu(rol):
    if rol == "Preparador":
        return ["📊 Dashboard", "📝 Nueva Conciliación", "📚 Historial"]
    elif rol == "Revisor":
        return ["🔍 Auditoría y Revisiones", "📊 Dashboard", "📚 Historial", "🏢 Empresas", "🏦 Bancos y Cuentas", "📈 Reportes"]
    elif rol == "Administrador":
        return ["📊 Dashboard", "🔍 Auditoría y Revisiones", "📝 Nueva Conciliación", "📚 Historial", "🏢 Empresas", "🏦 Bancos y Cuentas", "📈 Reportes", "👥 Usuarios"]
    return ["📊 Dashboard"]


def iniciar_autenticacion():
    if "usuario_autenticado" not in st.session_state:
        st.session_state.usuario_autenticado = None

    params = st.query_params
    if "registro" in params and params["registro"] == "true":
        st.title("📝 Registro de Nuevo Usuario")
        st.caption("Completa tus datos para crear tu cuenta de acceso.")
        
        rol_invitado = params.get("rol", "Preparador")
        empresa_id_invitada = params.get("empresa_id", None)
        if empresa_id_invitada:
            try:
                empresa_id_invitada = int(empresa_id_invitada)
            except ValueError:
                empresa_id_invitada = None

        with st.form("form_autoregistro"):
            nombre = st.text_input("Nombre completo")
            usuario = st.text_input("Nombre de usuario")
            password = st.text_input("Contraseña", type="password")
            confirmar = st.text_input("Confirmar contraseña", type="password")
            registro_btn = st.form_submit_button("Crear mi cuenta", type="primary", width="stretch")

        if registro_btn:
            if not usuario.strip() or not nombre.strip() or not password:
                st.error("Completa todos los campos.")
            elif len(password) < 8:
                st.error("La contraseña debe tener al menos 8 caracteres.")
            elif password != confirmar:
                st.error("Las contraseñas no coinciden.")
            else:
                try:
                    crear_usuario(usuario, nombre, password, rol_invitado, empresa_id_invitada)
                    st.success("¡Cuenta creada con éxito! Ya puedes iniciar sesión.")
                    st.query_params.clear()
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("El nombre de usuario ya existe. Por favor elige otro.")
        st.stop()

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
                st.success("Administrador creado correctamente.")
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
# MENÚ LATERAL Y SELECCIÓN DE EMPRESA
# =========================================================

empresas_df = obtener_empresas()

with st.sidebar:
    st.image("https://img.icons8.com/color/96/bank-building.png", width=70)
    st.title("Conciliación Web")
    st.markdown(f"👤 **{usuario_actual['nombre']}**")
    st.caption(f"Rol: **{rol_actual}**")
    st.divider()

    if rol_actual in ["Administrador", "Revisor"]:
        st.subheader("🏢 Selección de Empresa")
        opciones_emp = ["Todas las empresas"] + empresas_df["nombre"].tolist() if not empresas_df.empty else ["Todas las empresas"]
        empresa_activa_nombre = st.selectbox("Empresa a Auditar / Revisar", opciones_emp)
        
        if empresa_activa_nombre != "Todas las empresas" and not empresas_df.empty:
            empresa_activa_id = empresas_df.loc[empresas_df["nombre"] == empresa_activa_nombre, "id"].values[0]
            empresa_activa_nit = empresas_df.loc[empresas_df["nombre"] == empresa_activa_nombre, "nit"].values[0]
        else:
            empresa_activa_id = None
            empresa_activa_nit = ""
    else:
        empresa_activa_id = usuario_actual.get("empresa_id")
        empresa_activa_nombre = usuario_actual.get("empresa_nombre") or "Sin asignar"
        empresa_activa_nit = usuario_actual.get("empresa_nit") or ""
        st.info(f"🏢 **Empresa:** {empresa_activa_nombre}")

    st.divider()

    opciones_menu = obtener_opciones_menu(rol_actual)

    # Redirección dinámica si el usuario hace clic en Editar en el Historial
    if "menu_override" in st.session_state:
        menu_seleccionado = st.session_state.pop("menu_override")
    else:
        menu_seleccionado = st.radio("Navegación principal", opciones_menu)
    
    st.divider()
    if st.button("🚪 Cerrar sesión", width="stretch"):
        st.session_state.usuario_autenticado = None
        st.rerun()


# =========================================================
# FUNCIONES AUXILIARES DE FORMATO Y EXCEL / PDF
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


# GENERACIÓN DE PDF VERTICAL CON LOGO Y FORMATO LIMPIO
def generar_pdf_conciliacion(c_data, datos, nombres_titulos):
    buffer = io.BytesIO()
    # Hoja vertical (portrait)
    doc = SimpleDocTemplate(buffer, pagesize=portrait(letter), rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], alignment=1, textColor=colors.white, backColor=colors.HexColor("#1F4E78"), fontSize=13, spaceAfter=8, leading=16
    )
    sec_style = ParagraphStyle(
        'SecStyle', parent=styles['Heading2'], textColor=colors.HexColor("#1F4E78"), fontSize=10, spaceBefore=6, spaceAfter=3
    )
    normal_style = ParagraphStyle('NormStyle', parent=styles['Normal'], fontSize=8, leading=10)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold")

    consecutivo_str = f"CONC-{int(c_data.get('id', 0)):06d}"
    tipo_cta = c_data.get("tipo") or "Cuenta"

    # Encabezado con título y Logo
    logo_bytes = obtener_logo_empresa(c_data.get("empresa"))
    p_header_text = Paragraph(f"<b>CONCILIACIÓN - {tipo_cta.upper()}</b><br/><font size=8>Consecutivo No: {consecutivo_str}</font>", title_style)

    if logo_bytes:
        try:
            img_stream = io.BytesIO(logo_bytes)
            img_logo = Image(img_stream, width=80, height=45)
            header_table = Table([[p_header_text, img_logo]], colWidths=[440, 100])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (1,0), (1,0), 'RIGHT'),
            ]))
            story.append(header_table)
        except Exception:
            story.append(p_header_text)
    else:
        story.append(p_header_text)

    story.append(Spacer(1, 6))

    # Datos Generales
    data_gen = [
        [Paragraph("<b>Empresa:</b>", normal_style), Paragraph(str(c_data.get("empresa", "")), normal_style), Paragraph("<b>NIT:</b>", normal_style), Paragraph(str(c_data.get("nit", "")), normal_style)],
        [Paragraph("<b>Mes/Año:</b>", normal_style), Paragraph(str(c_data.get("mes", "")), normal_style), Paragraph("<b>Elaboración:</b>", normal_style), Paragraph(str(c_data.get("fecha_elaboracion", "")), normal_style)],
        [Paragraph("<b>Banco:</b>", normal_style), Paragraph(str(c_data.get("banco", "")), normal_style), Paragraph("<b>Cuenta No:</b>", normal_style), Paragraph(str(c_data.get("cuenta", "")), normal_style)],
    ]
    t_gen = Table(data_gen, colWidths=[80, 190, 80, 190])
    t_gen.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F2F2F2")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(t_gen)
    story.append(Spacer(1, 8))

    # Saldos y Cálculo Limpio
    story.append(Paragraph("SALDOS Y CÁLCULO DE LA CONCILIACIÓN", sec_style))

    res_final_val = c_data.get('resultado_final', 0.0)
    res_final_txt = f"$ {res_final_val:,.2f}"

    data_sal = [
        [Paragraph("SALDO SEGÚN EXTRACTO BANCARIO", normal_style), f"$ {c_data.get('saldo_extracto', 0):,.2f}"],
        [Paragraph("SALDO SEGÚN LIBROS", normal_style), f"$ {c_data.get('saldo_libros', 0):,.2f}"],
        [Paragraph("DIFERENCIA A JUSTIFICAR", normal_style), f"$ {c_data.get('diferencia_inicial', 0):,.2f}"],
        [Paragraph("DIFERENCIA CONCILIADA", normal_style), f"$ {c_data.get('diferencia_conciliada', 0):,.2f}"],
        [Paragraph("RESULTADO FINAL", bold_style), Paragraph(res_final_txt, bold_style)],
        [Paragraph("ESTADO", bold_style), Paragraph(str(c_data.get('estado', '')), bold_style)],
    ]
    t_sal = Table(data_sal, colWidths=[320, 220])
    t_sal.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#D9EAF7")),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(t_sal)
    story.append(Spacer(1, 8))

    # Tablas Auxiliares
    def agregar_tabla_pdf(titulo, lista_datos, cols_keys, cols_names, col_widths):
        story.append(Paragraph(titulo, sec_style))
        if not lista_datos:
            story.append(Paragraph("<i>Sin movimientos registrados</i>", normal_style))
            story.append(Spacer(1, 4))
            return

        header = [Paragraph(f"<b>{col}</b>", normal_style) for col in cols_names]
        rows = [header]
        for reg in lista_datos:
            r = []
            for k in cols_keys:
                val = reg.get(k, "")
                if k == "Valor" or k not in ["Fecha", "Beneficiario", "Documento", "Concepto"]:
                    try:
                        val = f"$ {float(val):,.2f}"
                    except (ValueError, TypeError):
                        pass
                r.append(Paragraph(str(val), normal_style))
            rows.append(r)

        t_m = Table(rows, colWidths=col_widths)
        t_m.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(t_m)
        story.append(Spacer(1, 4))

    agregar_tabla_pdf(f"1. {nombres_titulos['t1']}", datos.get("salidas_extracto", []), ["Fecha", "Beneficiario", "Documento", "Valor"], ["Fecha", "Beneficiario", "Documento", "Valor"], [90, 210, 120, 120])
    agregar_tabla_pdf(f"2. {nombres_titulos['t2']}", datos.get("salidas_libros", []), ["Fecha", "Concepto", "Valor"], ["Fecha", "Concepto", "Valor"], [100, 310, 130])
    agregar_tabla_pdf(f"3. {nombres_titulos['t3']}", datos.get("entradas_libros", []), ["Fecha", "Concepto", "Valor"], ["Fecha", "Concepto", "Valor"], [100, 310, 130])
    agregar_tabla_pdf(f"4. {nombres_titulos['t4']}", datos.get("entradas_extracto", []), ["Fecha", "Concepto", "Valor"], ["Fecha", "Concepto", "Valor"], [100, 310, 130])
    agregar_tabla_pdf("GASTOS BANCARIOS", datos.get("gastos_bancarios", []), ["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"], ["Fecha", "4x1000", "Cuota Manejo", "IVA", "Rte.Fuente", "Comisión", "Intereses"], [70, 78, 78, 78, 78, 78, 80])

    story.append(Spacer(1, 12))
    firmas = [
        [Paragraph(f"<b>Preparado por:</b> {datos.get('preparado_por', 'N/A')}", normal_style), Paragraph(f"<b>Revisado por:</b> {c_data.get('revisado_por_usuario', 'N/A')}", normal_style)]
    ]
    t_firmas = Table(firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('LINEABOVE', (0, 0), (-1, -1), 1, colors.HexColor("#1F4E78")),
    ]))
    story.append(t_firmas)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# VISTAS DE LA APLICACIÓN
# =========================================================

if menu_seleccionado == "🔍 Auditoría y Revisiones":
    st.title("🔍 BANDEJA DE AUDITORÍA Y REVISIONES")
    st.caption("Módulo exclusivo para la revisión técnica y aprobación de Conciliaciones Bancarias.")

    historial = obtener_historial(empresa_activa_nombre)
    pendientes = historial[historial["workflow_status"] == "Pendiente de revisión"]

    k1, k2, k3 = st.columns(3)
    k1.metric("Pendientes de Revisar 🕒", len(pendientes))
    k2.metric("Aprobadas 🔒", len(historial[historial["workflow_status"] == "Aprobada"]))
    k3.metric("Devueltas ❌", len(historial[historial["workflow_status"] == "Requiere corrección"]))

    st.divider()

    if pendientes.empty:
        st.success("🎉 ¡Excelente! No hay conciliaciones pendientes por auditar.")
    else:
        st.subheader("📋 Conciliaciones Pendientes de Revisión")
        
        for idx, fila in pendientes.iterrows():
            cuenta_txt = fila.get("cuenta") or "N/A"
            banco_txt = fila.get("banco") or "N/A"
            res_fin = fila.get("resultado_final", 0.0)
            consecutivo_str = f"CONC-{int(fila['id']):06d}"
            
            with st.expander(f"📌 {consecutivo_str} | {fila['empresa']} - {banco_txt} ({cuenta_txt}) | Mes: {fila['mes']} | Resultado: ${res_fin:,.2f}"):
                
                c_data = obtener_conciliacion_por_id(fila["id"])
                
                try:
                    datos = json.loads(c_data.get("datos_json", "{}"))
                except Exception:
                    datos = {}

                tipo_cta = c_data.get("tipo") or "Cuenta de ahorros"
                es_tc = "tarjeta" in tipo_cta.lower() or "crédito" in tipo_cta.lower() or "credito" in tipo_cta.lower()

                if es_tc:
                    t1_nombre, t2_nombre, t3_nombre, t4_nombre = (
                        "COMPRAS NO EVIDENCIADAS EN EXTRACTOS",
                        "COMPRAS NO CONTABILIZADAS EN LIBROS",
                        "DÉBITOS BANCARIOS NO CONTABILIZADOS EN LIBROS",
                        "ABONOS NO REGISTRADOS EN EXTRACTO"
                    )
                else:
                    t1_nombre, t2_nombre, t3_nombre, t4_nombre = (
                        "SALIDAS NO REGISTRADAS EN EXTRACTO",
                        "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
                        "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
                        "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS"
                    )

                col_tit, col_lg = st.columns([4, 1])
                with col_tit:
                    st.markdown(
                        f"""
                        <div style="background-color: #1F4E78; color: white; padding: 12px; text-align: center; border-radius: 5px; font-weight: bold; font-size: 18px;">
                            CONCILIACIÓN - {tipo_cta.upper()} ({consecutivo_str})
                        </div>
                        """, unsafe_allow_html=True
                    )
                with col_lg:
                    logo_aud = obtener_logo_empresa(c_data.get("empresa"))
                    if logo_aud:
                        st.image(logo_aud, width=110)

                col_inf1, col_inf2 = st.columns(2)
                with col_inf1:
                    st.write(f"🏢 **Empresa:** {c_data.get('empresa', 'N/A')} | **NIT:** {c_data.get('nit', 'N/A')}")
                    st.write(f"📅 **Mes/Año:** {c_data.get('mes', 'N/A')} | **Elaboración:** {c_data.get('fecha_elaboracion', 'N/A')}")
                with col_inf2:
                    st.write(f"🏦 **Banco:** {c_data.get('banco', 'N/A')} | **Cuenta No:** {c_data.get('cuenta', 'N/A')}")
                    st.write(f"👤 **Preparado por:** {datos.get('preparado_por', 'N/A')}")

                st.divider()

                st.markdown("**💰 SALDOS Y CÁLCULO DE LA CONCILIACIÓN**")
                df_saldos = pd.DataFrame([
                    {"Concepto": "SALDO SEGÚN EXTRACTO BANCARIO", "Valor": f"${c_data.get('saldo_extracto', 0):,.2f}"},
                    {"Concepto": "SALDO SEGÚN LIBROS", "Valor": f"${c_data.get('saldo_libros', 0):,.2f}"},
                    {"Concepto": "DIFERENCIA A JUSTIFICAR", "Valor": f"${c_data.get('diferencia_inicial', 0):,.2f}"},
                    {"Concepto": "DIFERENCIA CONCILIADA", "Valor": f"${c_data.get('diferencia_conciliada', 0):,.2f}"},
                    {"Concepto": "RESULTADO FINAL", "Valor": f"${c_data.get('resultado_final', 0):,.2f}"},
                    {"Concepto": "ESTADO", "Valor": c_data.get('estado', 'N/A')}
                ])
                st.table(df_saldos)

                st.markdown(f"**1. {t1_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("salidas_extracto", [])), width="stretch", hide_index=True)

                st.markdown(f"**2. {t2_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("salidas_libros", [])), width="stretch", hide_index=True)

                st.markdown(f"**3. {t3_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("entradas_libros", [])), width="stretch", hide_index=True)

                st.markdown(f"**4. {t4_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("entradas_extracto", [])), width="stretch", hide_index=True)

                st.markdown("**GASTOS BANCARIOS**")
                st.dataframe(pd.DataFrame(datos.get("gastos_bancarios", [])), width="stretch", hide_index=True)

                st.divider()

                st.subheader("⚡ Decisión del Auditor")
                btn_col1, btn_col2 = st.columns(2)

                with btn_col1:
                    if st.button("🔒 APROBAR CONCILIACIÓN", key=f"aprob_{fila['id']}", type="primary", width="stretch"):
                        actualizar_estado_auditoria(fila["id"], "Aprobada", usuario_actual["nombre"])
                        st.success("✅ Conciliación aprobada con éxito.")
                        st.rerun()

                with btn_col2:
                    motivo = st.text_input("Observaciones / Motivo de Corrección (Obligatorio si se devuelve)", key=f"mot_{fila['id']}")
                    if st.button("❌ DEVOLVER PARA CORRECCIÓN", key=f"dev_{fila['id']}", width="stretch"):
                        if not motivo.strip():
                            st.error("Debes ingresar el motivo de devolución.")
                        else:
                            actualizar_estado_auditoria(fila["id"], "Requiere corrección", usuario_actual["nombre"], motivo)
                            st.warning("⚠️ Conciliación devuelta al preparador.")
                            st.rerun()


elif menu_seleccionado == "📊 Dashboard":
    st.title("📊 Dashboard de Conciliaciones")
    if empresa_activa_nombre != "Todas las empresas":
        st.caption(f"Filtrado por empresa: **{empresa_activa_nombre}**")
    
    historial = obtener_historial(empresa_activa_nombre)
    if len(historial) == 0:
        st.info("Aún no hay conciliaciones registradas para la empresa seleccionada.")
    else:
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Total Conciliaciones", len(historial))
        kpi2.metric("Aprobadas 🔒", len(historial[historial["workflow_status"] == "Aprobada"]))
        kpi3.metric("Pendientes 🕒", len(historial[historial["workflow_status"] == "Pendiente de revisión"]))
        st.divider()
        st.dataframe(historial[["id", "empresa", "mes", "banco", "resultado_final", "workflow_status"]].head(10), width="stretch", hide_index=True)


elif menu_seleccionado == "🏢 Empresas":
    st.title("🏢 Maestro de Empresas")

    empresa_edit_id = st.session_state.get("empresa_a_editar", None)
    if empresa_edit_id:
        emp_obj = obtener_empresa_por_id(empresa_edit_id)
        if emp_obj:
            st.info(f"✏️ **Modo Edición:** Editando la empresa **{emp_obj['nombre']}**")
            with st.form("form_editar_empresa"):
                edit_nom = st.text_input("Nombre de la empresa", value=emp_obj["nombre"])
                edit_nit = st.text_input("NIT", value=emp_obj["nit"])
                edit_logo = st.file_uploader("Actualizar Logo (Opcional - PNG, JPG)", type=["png", "jpg", "jpeg"])
                
                c_guard, c_canc = st.columns(2)
                with c_guard:
                    if st.form_submit_button("💾 Guardar Cambios", type="primary", width="stretch"):
                        logo_b = edit_logo.getvalue() if edit_logo else None
                        actualizar_empresa_db(empresa_edit_id, edit_nom, edit_nit, logo_b)
                        st.session_state.empresa_a_editar = None
                        st.success("Empresa actualizada con éxito.")
                        st.rerun()
                with c_canc:
                    if st.form_submit_button("❌ Cancelar", width="stretch"):
                        st.session_state.empresa_a_editar = None
                        st.rerun()
            st.divider()

    if rol_actual == "Administrador" and not empresa_edit_id:
        with st.expander("➕ Registrar nueva empresa con Logo"):
            with st.form("form_empresa_maestro"):
                nom = st.text_input("Nombre de la empresa")
                nit = st.text_input("NIT")
                logo_file = st.file_uploader("Logo de la Empresa (PNG, JPG)", type=["png", "jpg", "jpeg"])
                if st.form_submit_button("Guardar Empresa", type="primary"):
                    if not nom.strip() or not nit.strip():
                        st.error("Por favor completa el nombre y el NIT.")
                    else:
                        logo_b = logo_file.getvalue() if logo_file else None
                        guardar_empresa(nom, nit, logo_b)
                        st.success("Empresa registrada con éxito.")
                        st.rerun()

    st.subheader("Empresas Registradas")
    empresas_list = obtener_empresas()

    if empresas_list.empty:
        st.info("No hay empresas registradas.")
    else:
        for _, emp in empresas_list.iterrows():
            with st.container():
                col_lg, col_dt, col_act1, col_act2 = st.columns([1, 4, 1.5, 1.5])
                
                with col_lg:
                    if emp["logo"]:
                        st.image(emp["logo"], width=70)
                    else:
                        st.caption("Sin logo")

                with col_dt:
                    st.write(f"🏢 **{emp['nombre']}**")
                    st.write(f"🆔 NIT: {emp['nit']}")

                with col_act1:
                    if rol_actual == "Administrador":
                        if st.button("✏️ Editar", key=f"btn_edit_emp_{emp['id']}", width="stretch"):
                            st.session_state.empresa_a_editar = emp['id']
                            st.rerun()

                with col_act2:
                    if rol_actual == "Administrador":
                        if st.button("🗑️ Eliminar", key=f"btn_del_emp_{emp['id']}", width="stretch"):
                            st.session_state[f"confirm_del_emp_{emp['id']}"] = True

                if st.session_state.get(f"confirm_del_emp_{emp['id']}", False):
                    st.warning(f"⚠️ ¿Estás seguro de eliminar la empresa '{emp['nombre']}'?")
                    col_si, col_no = st.columns(2)
                    with col_si:
                        if st.button("Sí, eliminar", key=f"confirm_yes_{emp['id']}", type="primary", width="stretch"):
                            eliminar_empresa_db(emp['id'])
                            st.session_state[f"confirm_del_emp_{emp['id']}"] = False
                            st.success("Empresa eliminada.")
                            st.rerun()
                    with col_no:
                        if st.button("Cancelar", key=f"confirm_no_{emp['id']}", width="stretch"):
                            st.session_state[f"confirm_del_emp_{emp['id']}"] = False
                            st.rerun()

            st.divider()


elif menu_seleccionado == "🏦 Bancos y Cuentas":
    st.title("🏦 Maestro de Bancos y Cuentas Bancarias")
    if rol_actual == "Administrador":
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
    st.dataframe(obtener_cuentas(empresa_activa_id), width="stretch", hide_index=True)


elif menu_seleccionado == "📝 Nueva Conciliación":
    st.title("📝 Captura / Edición de Conciliación Bancaria")

    id_edicion = st.session_state.get("conciliacion_a_editar", None)
    if id_edicion:
        c_edit = obtener_conciliacion_por_id(id_edicion)
        consecutivo_edit_str = f"CONC-{int(id_edicion):06d}"
        st.info(f"✏️ **Modo Edición Activado:** Editando Conciliación {consecutivo_edit_str} ({c_edit.get('empresa')} - {c_edit.get('mes')})")
        if st.button("❌ Cancelar Edición y Crear Nueva"):
            st.session_state.conciliacion_a_editar = None
            if "datos_cargados_edit" in st.session_state:
                del st.session_state.datos_cargados_edit
            st.rerun()

    if "tabla1" not in st.session_state or id_edicion:
        st.session_state.tabla1 = pd.DataFrame(columns=["Fecha", "Beneficiario", "Documento", "Valor"])
    if "tabla2" not in st.session_state or id_edicion:
        st.session_state.tabla2 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    if "tabla3" not in st.session_state or id_edicion:
        st.session_state.tabla3 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    if "tabla4" not in st.session_state or id_edicion:
        st.session_state.tabla4 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    if "tabla5" not in st.session_state or id_edicion:
        st.session_state.tabla5 = pd.DataFrame(columns=["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"])

    if id_edicion and "datos_cargados_edit" not in st.session_state:
        c_edit = obtener_conciliacion_por_id(id_edicion)
        if c_edit:
            try:
                d_js = json.loads(c_edit.get("datos_json", "{}"))
                if d_js.get("salidas_extracto"): st.session_state.tabla1 = pd.DataFrame(d_js["salidas_extracto"])
                if d_js.get("salidas_libros"): st.session_state.tabla2 = pd.DataFrame(d_js["salidas_libros"])
                if d_js.get("entradas_libros"): st.session_state.tabla3 = pd.DataFrame(d_js["entradas_libros"])
                if d_js.get("entradas_extracto"): st.session_state.tabla4 = pd.DataFrame(d_js["entradas_extracto"])
                if d_js.get("gastos_bancarios"): st.session_state.tabla5 = pd.DataFrame(d_js["gastos_bancarios"])
            except Exception:
                pass
            st.session_state.datos_cargados_edit = True

    col_info_emp, col_logo_emp = st.columns([3, 1])

    with col_info_emp:
        st.subheader("Información General")
        col1, col2 = st.columns(2)

        with col1:
            if id_edicion:
                c_edit = obtener_conciliacion_por_id(id_edicion)
                empresa = c_edit.get("empresa", "")
                nit = c_edit.get("nit", "")
                mes = c_edit.get("mes", "")
                st.text_input("Empresa", value=empresa, disabled=True)
                st.text_input("NIT", value=nit, disabled=True)
                st.text_input("Mes / Año", value=mes, disabled=True)
                fecha_elaboracion = st.date_input("Fecha de elaboración", key="form_fecha_elaboracion_edit")
            else:
                if empresa_activa_nombre != "Todas las empresas" and empresa_activa_nombre != "Sin asignar":
                    empresa = empresa_activa_nombre
                    nit = empresa_activa_nit
                    st.text_input("Empresa Seleccionada", value=empresa, disabled=True)
                    st.text_input("NIT", value=nit, disabled=True)
                elif not empresas_df.empty:
                    empresa_obj = st.selectbox("Empresa Registrada", empresas_df["nombre"].tolist())
                    empresa = empresa_obj
                    nit = empresas_df.loc[empresas_df["nombre"] == empresa_obj, "nit"].values[0]
                    empresa_activa_id = empresas_df.loc[empresas_df["nombre"] == empresa_obj, "id"].values[0]
                    st.text_input("NIT", value=nit, disabled=True)
                else:
                    empresa = st.text_input("Nombre de la Empresa", key="form_empresa")
                    nit = st.text_input("NIT", key="form_nit")

                mes_nombre = st.selectbox("Mes a Conciliar", [
                    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
                ], index=datetime.now().month - 1)
                
                mes_num = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"].index(mes_nombre) + 1
                anio = st.number_input("Año", value=datetime.now().year, step=1)
                mes = f"{mes_nombre.upper()} {anio}"
                fecha_elaboracion = st.date_input("Fecha de elaboración", key="form_fecha_elaboracion")

        with col2:
            if id_edicion:
                c_edit = obtener_conciliacion_por_id(id_edicion)
                banco = c_edit.get("banco", "")
                cuenta = c_edit.get("cuenta", "")
                tipo = c_edit.get("tipo", "Cuenta de ahorros")
                st.text_input("Banco", value=banco, disabled=True)
                st.text_input("Cuenta / Tarjeta", value=cuenta, disabled=True)
                st.text_input("Tipo", value=tipo, disabled=True)
            else:
                cuentas_asig_df = obtener_cuentas_rotadas_por_usuario(
                    empresa_activa_id, mes_num, usuario_actual["id"]
                )
                if not cuentas_asig_df.empty:
                    st.info("🔄 **Cuentas asignadas para tu perfil este mes:**")
                    cta_sel = st.selectbox(
                        "Cuenta / Tarjeta Registrada",
                        cuentas_asig_df["id"].tolist(),
                        format_func=lambda x: f"{cuentas_asig_df.loc[cuentas_asig_df['id']==x, 'banco'].values[0]} - {cuentas_asig_df.loc[cuentas_asig_df['id']==x, 'numero_cuenta'].values[0]} ({cuentas_asig_df.loc[cuentas_asig_df['id']==x, 'tipo_cuenta'].values[0]})"
                    )
                    banco = cuentas_asig_df.loc[cuentas_asig_df["id"] == cta_sel, "banco"].values[0]
                    cuenta = cuentas_asig_df.loc[cuentas_asig_df["id"] == cta_sel, "numero_cuenta"].values[0]
                    tipo = cuentas_asig_df.loc[cuentas_asig_df["id"] == cta_sel, "tipo_cuenta"].values[0]
                else:
                    st.warning("No hay cuentas asignadas o registradas para preparadores activos en esta empresa.")
                    banco = st.text_input("Nombre del Banco", key="form_banco")
                    cuenta = st.text_input("Número de Cuenta", key="form_cuenta")
                    tipo = st.selectbox("Tipo de Cuenta", ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito"], key="form_tipo")

    with col_logo_emp:
        logo_bytes = obtener_logo_empresa(empresa)
        if logo_bytes:
            st.image(logo_bytes, caption=f"Logo - {empresa}", width=150)

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
    val_ext = float(obtener_conciliacion_por_id(id_edicion).get("saldo_extracto", 0.0)) if id_edicion else 0.0
    val_lib = float(obtener_conciliacion_por_id(id_edicion).get("saldo_libros", 0.0)) if id_edicion else 0.0

    with c1:
        saldo_extracto = st.number_input("Saldo según Extracto", value=val_ext, format="%.2f", key="form_saldo_extracto")
    with c2:
        saldo_libros = st.number_input("Saldo según Libros", value=val_lib, format="%.2f", key="form_saldo_libros")

    diferencia_inicial = saldo_extracto - saldo_libros
    st.metric("Diferencia a Justificar", f"${diferencia_inicial:,.2f}")

    st.divider()
    col_act1, col_act2 = st.columns([3, 1])
    with col_act2:
        if st.button("🔄 Actualizar y Recalcular Conciliación", width="stretch", type="secondary"):
            st.rerun()

    st.divider()
    st.subheader(f"1. {nombres_titulos['t1']}")
    salidas_extracto = st.data_editor(st.session_state.tabla1, num_rows="dynamic", width="stretch", key="editor_tabla1")
    st.session_state.tabla1 = salidas_extracto

    cols_t1 = ["Fecha", "Beneficiario", "Documento", "Valor"]
    c_p1, c_b1, c_del1 = st.columns([3, 1, 1])
    with c_p1:
        txt_t1 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Beneficiario | Documento | Valor):", key="paste_t1")
    with c_b1:
        if st.button("➕ Cargar Ítems", key="btn_parse_t1"):
            df_p1 = parsear_texto_pegado(txt_t1, cols_t1)
            if df_p1 is not None:
                st.session_state.tabla1 = pd.concat([st.session_state.tabla1, df_p1], ignore_index=True)
                st.rerun()
    with c_del1:
        if st.button("🗑️ Limpiar Ítem 1", key="btn_del_t1"):
            st.session_state.tabla1 = pd.DataFrame(columns=cols_t1)
            st.rerun()

    st.divider()
    st.subheader(f"2. {nombres_titulos['t2']}")
    salidas_libros = st.data_editor(st.session_state.tabla2, num_rows="dynamic", width="stretch", key="editor_tabla2")
    st.session_state.tabla2 = salidas_libros

    cols_t2 = ["Fecha", "Concepto", "Valor"]
    c_p2, c_b2, c_del2 = st.columns([3, 1, 1])
    with c_p2:
        txt_t2 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Concepto | Valor):", key="paste_t2")
    with c_b2:
        if st.button("➕ Cargar Ítems", key="btn_parse_t2"):
            df_p2 = parsear_texto_pegado(txt_t2, cols_t2)
            if df_p2 is not None:
                st.session_state.tabla2 = pd.concat([st.session_state.tabla2, df_p2], ignore_index=True)
                st.rerun()
    with c_del2:
        if st.button("🗑️ Limpiar Ítem 2", key="btn_del_t2"):
            st.session_state.tabla2 = pd.DataFrame(columns=cols_t2)
            st.rerun()

    st.divider()
    st.subheader(f"3. {nombres_titulos['t3']}")
    entradas_libros = st.data_editor(st.session_state.tabla3, num_rows="dynamic", width="stretch", key="editor_tabla3")
    st.session_state.tabla3 = entradas_libros

    cols_t3 = ["Fecha", "Concepto", "Valor"]
    c_p3, c_b3, c_del3 = st.columns([3, 1, 1])
    with c_p3:
        txt_t3 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Concepto | Valor):", key="paste_t3")
    with c_b3:
        if st.button("➕ Cargar Ítems", key="btn_parse_t3"):
            df_p3 = parsear_texto_pegado(txt_t3, cols_t3)
            if df_p3 is not None:
                st.session_state.tabla3 = pd.concat([st.session_state.tabla3, df_p3], ignore_index=True)
                st.rerun()
    with c_del3:
        if st.button("🗑️ Limpiar Ítem 3", key="btn_del_t3"):
            st.session_state.tabla3 = pd.DataFrame(columns=cols_t3)
            st.rerun()

    st.divider()
    st.subheader(f"4. {nombres_titulos['t4']}")
    entradas_extracto = st.data_editor(st.session_state.tabla4, num_rows="dynamic", width="stretch", key="editor_tabla4")
    st.session_state.tabla4 = entradas_extracto

    cols_t4 = ["Fecha", "Concepto", "Valor"]
    c_p4, c_b4, c_del4 = st.columns([3, 1, 1])
    with c_p4:
        txt_t4 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Concepto | Valor):", key="paste_t4")
    with c_b4:
        if st.button("➕ Cargar Ítems", key="btn_parse_t4"):
            df_p4 = parsear_texto_pegado(txt_t4, cols_t4)
            if df_p4 is not None:
                st.session_state.tabla4 = pd.concat([st.session_state.tabla4, df_p4], ignore_index=True)
                st.rerun()
    with c_del4:
        if st.button("🗑️ Limpiar Ítem 4", key="btn_del_t4"):
            st.session_state.tabla4 = pd.DataFrame(columns=cols_t4)
            st.rerun()

    st.divider()
    st.subheader("Gastos Bancarios")
    gastos_bancarios = st.data_editor(st.session_state.tabla5, num_rows="dynamic", width="stretch", key="editor_tabla5")
    st.session_state.tabla5 = gastos_bancarios

    cols_t5 = ["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"]
    c_p5, c_b5, c_del5 = st.columns([3, 1, 1])
    with c_p5:
        txt_t5 = st.text_input("📋 Pegar ítems desde Excel para Gastos Bancarios:", key="paste_t5")
    with c_b5:
        if st.button("➕ Cargar Ítems", key="btn_parse_t5"):
            df_p5 = parsear_texto_pegado(txt_t5, cols_t5)
            if df_p5 is not None:
                st.session_state.tabla5 = pd.concat([st.session_state.tabla5, df_p5], ignore_index=True)
                st.rerun()
    with c_del5:
        if st.button("🗑️ Limpiar Gastos", key="btn_del_t5"):
            st.session_state.tabla5 = pd.DataFrame(columns=cols_t5)
            st.rerun()

    m1 = total_columna(salidas_extracto)
    m2 = total_columna(salidas_libros)
    m3 = total_columna(entradas_libros)
    m4 = total_columna(entradas_extracto)

    if es_tc:
        diferencia_conciliada = - m1 + m2 - m3 + m4
    else:
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
    btn_col_borr, btn_col_env = st.columns(2)

    with btn_col_borr:
        if st.button("💾 Guardar como Borrador", key="btn_guardar_borrador", width="stretch"):
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data,
                workflow_status="Borrador",
                id_edicion=id_edicion
            )
            if "conciliacion_a_editar" in st.session_state:
                del st.session_state.conciliacion_a_editar
            if "datos_cargados_edit" in st.session_state:
                del st.session_state.datos_cargados_edit
            st.success(f"📝 Conciliación CONC-{int(id_g):06d} guardada exitosamente como Borrador.")

    with btn_col_env:
        if st.button("📤 Enviar a Revisión", key="btn_enviar_revision", type="primary", width="stretch"):
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data,
                workflow_status="Pendiente de revisión",
                id_edicion=id_edicion
            )
            if "conciliacion_a_editar" in st.session_state:
                del st.session_state.conciliacion_a_editar
            if "datos_cargados_edit" in st.session_state:
                del st.session_state.datos_cargados_edit
            st.success(f"✅ Conciliación CONC-{int(id_g):06d} enviada correctamente a Revisión.")


elif menu_seleccionado == "📚 Historial":
    st.title("📚 Historial de Conciliaciones")
    if empresa_activa_nombre != "Todas las empresas":
        st.caption(f"Mostrando conciliaciones de: **{empresa_activa_nombre}**")
    
    historial = obtener_historial(empresa_activa_nombre)

    col_f1, col_f2 = st.columns([2, 2])
    with col_f1:
        filtro_estado = st.selectbox("Filtrar por Estado de Revisión:", ["Todos los Estados", "Borrador", "Pendiente de revisión", "Aprobada", "Requiere corrección"])
    
    if filtro_estado != "Todos los Estados":
        historial = historial[historial["workflow_status"] == filtro_estado]

    if historial.empty:
        st.info("No hay conciliaciones guardadas para el filtro seleccionado.")
    else:
        for idx, fila in historial.iterrows():
            cuenta_txt = fila.get("cuenta") or "N/A"
            banco_txt = fila.get("banco") or "N/A"
            res_fin = fila.get("resultado_final", 0.0)
            wf_status = fila.get("workflow_status", "Pendiente de revisión")
            
            # CONSECUTIVO ÚNICO DE GESTIÓN Y CONTROL
            consecutivo_str = f"CONC-{int(fila['id']):06d}"

            badge_status = {
                "Borrador": "📝 BORRADOR",
                "Pendiente de revisión": "🕒 PENDIENTE DE REVISIÓN",
                "Aprobada": "🔒 APROBADA",
                "Requiere corrección": "❌ REQUIERE CORRECCIÓN"
            }.get(wf_status, wf_status)

            with st.expander(f"📌 {consecutivo_str} | {fila['empresa']} - {banco_txt} ({cuenta_txt}) | Mes: {fila['mes']} | [{badge_status}] | Resultado: ${res_fin:,.2f}"):
                
                c_data = obtener_conciliacion_por_id(fila["id"])
                try:
                    datos = json.loads(c_data.get("datos_json", "{}"))
                except Exception:
                    datos = {}

                if wf_status == "Requiere corrección" and c_data.get("motivo_correccion"):
                    st.error(f"⚠️ **Observaciones del Revisor / Auditor ({c_data.get('revisado_por_usuario', 'N/A')}):** {c_data.get('motivo_correccion')}")

                tipo_cta = c_data.get("tipo") or "Cuenta de ahorros"
                es_tc = "tarjeta" in tipo_cta.lower() or "crédito" in tipo_cta.lower() or "credito" in tipo_cta.lower()

                if es_tc:
                    t1_nombre, t2_nombre, t3_nombre, t4_nombre = (
                        "COMPRAS NO EVIDENCIADAS EN EXTRACTOS",
                        "COMPRAS NO CONTABILIZADAS EN LIBROS",
                        "DÉBITOS BANCARIOS NO CONTABILIZADOS EN LIBROS",
                        "ABONOS NO REGISTRADOS EN EXTRACTO"
                    )
                else:
                    t1_nombre, t2_nombre, t3_nombre, t4_nombre = (
                        "SALIDAS NO REGISTRADAS EN EXTRACTO",
                        "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
                        "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
                        "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS"
                    )

                # BOTONES DE ACCIÓN (EDITAR, ELIMINAR BORRADOR, EXCEL, PDF)
                cols_acciones = st.columns(4 if wf_status == "Borrador" else 3)
                
                with cols_acciones[0]:
                    # BOTÓN DE EDICIÓN CON REDIRECCIÓN DIRECTA A "NUEVA CONCILIACIÓN"
                    if st.button(f"✏️ Editar Conciliación", key=f"btn_edit_{fila['id']}", width="stretch"):
                        st.session_state.conciliacion_a_editar = fila['id']
                        if "datos_cargados_edit" in st.session_state:
                            del st.session_state.datos_cargados_edit
                        st.session_state.menu_override = "📝 Nueva Conciliación"
                        st.rerun()

                if wf_status == "Borrador":
                    with cols_acciones[1]:
                        if st.button(f"🗑️ Eliminar Borrador", key=f"btn_del_borr_{fila['id']}", width="stretch"):
                            st.session_state[f"confirm_del_conc_{fila['id']}"] = True

                    idx_ex = 2
                    idx_pdf = 3
                else:
                    idx_ex = 1
                    idx_pdf = 2

                with cols_acciones[idx_ex]:
                    if c_data.get("excel"):
                        nombre_ex = f"CONCILIACION_{consecutivo_str}_{limpiar_nombre_archivo(c_data.get('empresa'))}_{limpiar_nombre_archivo(c_data.get('mes'))}.xlsx"
                        st.download_button("📥 Descargar Excel (.xlsx)", data=bytes(c_data["excel"]), file_name=nombre_ex, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"down_ex_{fila['id']}", width="stretch")

                with cols_acciones[idx_pdf]:
                    nombres_titulos = {"t1": t1_nombre, "t2": t2_nombre, "t3": t3_nombre, "t4": t4_nombre}
                    pdf_bytes = generar_pdf_conciliacion(c_data, datos, nombres_titulos)
                    nombre_pdf = f"CONCILIACION_{consecutivo_str}_{limpiar_nombre_archivo(c_data.get('empresa'))}_{limpiar_nombre_archivo(c_data.get('mes'))}.pdf"
                    st.download_button("📥 Descargar PDF (.pdf)", data=pdf_bytes, file_name=nombre_pdf, mime="application/pdf", key=f"down_pdf_{fila['id']}", width="stretch")

                # CONFIRMACIÓN DE ELIMINACIÓN DE BORRADOR
                if st.session_state.get(f"confirm_del_conc_{fila['id']}", False):
                    st.warning(f"⚠️ ¿Estás seguro de eliminar permanentemente el Borrador **{consecutivo_str}**?")
                    col_si, col_no = st.columns(2)
                    with col_si:
                        if st.button("Sí, eliminar borrador", key=f"confirm_yes_conc_{fila['id']}", type="primary", width="stretch"):
                            eliminar_conciliacion_db(fila['id'])
                            st.session_state[f"confirm_del_conc_{fila['id']}"] = False
                            st.success("Borrador eliminado correctamente.")
                            st.rerun()
                    with col_no:
                        if st.button("Cancelar", key=f"confirm_no_conc_{fila['id']}", width="stretch"):
                            st.session_state[f"confirm_del_conc_{fila['id']}"] = False
                            st.rerun()

                st.divider()

                col_tit_h, col_lg_h = st.columns([4, 1])
                with col_tit_h:
                    st.markdown(
                        f"""
                        <div style="background-color: #1F4E78; color: white; padding: 10px; text-align: center; border-radius: 5px; font-weight: bold; font-size: 16px;">
                            VISTA VIRTUAL FORMATO OFICIAL - CONCILIACIÓN {tipo_cta.upper()} ({consecutivo_str})
                        </div>
                        """, unsafe_allow_html=True
                    )
                with col_lg_h:
                    logo_h = obtener_logo_empresa(c_data.get("empresa"))
                    if logo_h:
                        st.image(logo_h, width=110)

                col_inf1, col_inf2 = st.columns(2)
                with col_inf1:
                    st.write(f"🏢 **Empresa:** {c_data.get('empresa', 'N/A')} | **NIT:** {c_data.get('nit', 'N/A')}")
                    st.write(f"📅 **Mes/Año:** {c_data.get('mes', 'N/A')} | **Elaboración:** {c_data.get('fecha_elaboracion', 'N/A')}")
                with col_inf2:
                    st.write(f"🏦 **Banco:** {c_data.get('banco', 'N/A')} | **Cuenta No:** {c_data.get('cuenta', 'N/A')}")
                    st.write(f"👤 **Preparado por:** {datos.get('preparado_por', 'N/A')} | **Revisado por:** {c_data.get('revisado_por_usuario', 'Pendiente')}")

                st.markdown("**💰 SALDOS Y CÁLCULO DE LA CONCILIACIÓN**")
                df_saldos = pd.DataFrame([
                    {"Concepto": "SALDO SEGÚN EXTRACTO BANCARIO", "Valor": f"${c_data.get('saldo_extracto', 0):,.2f}"},
                    {"Concepto": "SALDO SEGÚN LIBROS", "Valor": f"${c_data.get('saldo_libros', 0):,.2f}"},
                    {"Concepto": "DIFERENCIA A JUSTIFICAR", "Valor": f"${c_data.get('diferencia_inicial', 0):,.2f}"},
                    {"Concepto": "DIFERENCIA CONCILIADA", "Valor": f"${c_data.get('diferencia_conciliada', 0):,.2f}"},
                    {"Concepto": "RESULTADO FINAL", "Valor": f"${c_data.get('resultado_final', 0):,.2f}"},
                    {"Concepto": "ESTADO", "Valor": c_data.get('estado', 'N/A')}
                ])
                st.table(df_saldos)

                st.markdown(f"**1. {t1_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("salidas_extracto", [])), width="stretch", hide_index=True)

                st.markdown(f"**2. {t2_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("salidas_libros", [])), width="stretch", hide_index=True)

                st.markdown(f"**3. {t3_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("entradas_libros", [])), width="stretch", hide_index=True)

                st.markdown(f"**4. {t4_nombre}**")
                st.dataframe(pd.DataFrame(datos.get("entradas_extracto", [])), width="stretch", hide_index=True)

                st.markdown("**GASTOS BANCARIOS**")
                st.dataframe(pd.DataFrame(datos.get("gastos_bancarios", [])), width="stretch", hide_index=True)


elif menu_seleccionado == "📈 Reportes":
    st.title("📈 Reportes")
    historial = obtener_historial(empresa_activa_nombre)
    if not historial.empty:
        csv = historial.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Historial (CSV)", data=csv, file_name=f"historial_{limpiar_nombre_archivo(empresa_activa_nombre)}.csv", mime="text/csv", width="stretch")


elif menu_seleccionado == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios y Roles")

    usuarios_df = obtener_usuarios()
    st.subheader("Lista de Usuarios Registrados")
    st.dataframe(usuarios_df, width="stretch", hide_index=True)

    if rol_actual == "Administrador":
        st.divider()
        col_crear, col_link = st.columns(2)

        with col_crear:
            with st.expander("➕ Crear Nuevo Usuario Manualmente", expanded=True):
                with st.form("form_crear_usuario_manual"):
                    nuevo_nombre = st.text_input("Nombre completo")
                    nuevo_usuario = st.text_input("Usuario")
                    nueva_pass = st.text_input("Contraseña", type="password")
                    nuevo_rol = st.selectbox("Rol de Acceso", ["Preparador", "Revisor", "Administrador"])
                    
                    emp_id_crear = None
                    if not empresas_df.empty:
                        emp_id_crear = st.selectbox(
                            "Empresa Asignada", 
                            [None] + empresas_df["id"].tolist(), 
                            format_func=lambda x: "Sin asignar (Todas)" if x is None else empresas_df.loc[empresas_df["id"]==x, "nombre"].values[0]
                        )
                    
                    if st.form_submit_button("Crear Usuario", type="primary", width="stretch"):
                        if not nuevo_nombre.strip() or not nuevo_usuario.strip() or not nueva_pass:
                            st.error("Completa todos los campos obligatorios.")
                        elif len(nueva_pass) < 8:
                            st.error("La contraseña debe tener al menos 8 caracteres.")
                        else:
                            try:
                                crear_usuario(nuevo_usuario, nuevo_nombre, nueva_pass, nuevo_rol, emp_id_crear)
                                st.success(f"Usuario '{nuevo_usuario}' creado con éxito con rol {nuevo_rol}.")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Ese nombre de usuario ya está registrado.")

        with col_link:
            with st.expander("🔗 Generar Enlace de Autoregistro", expanded=True):
                st.caption("Crea un link para enviar a un colaborador para que cree su propia cuenta.")
                rol_link = st.selectbox("Rol para el nuevo usuario", ["Preparador", "Revisor", "Administrador"], key="link_rol")
                
                emp_id_link = None
                if not empresas_df.empty:
                    emp_id_link = st.selectbox(
                        "Empresa predeterminada", 
                        [None] + empresas_df["id"].tolist(), 
                        format_func=lambda x: "Sin asignar" if x is None else empresas_df.loc[empresas_df["id"]==x, "nombre"].values[0],
                        key="link_empresa"
                    )

                url_base = st.query_params.get("base_url", "https://conciliacionweb.streamlit.app/")
                link_generado = f"{url_base}?registro=true&rol={rol_link}"
                if emp_id_link:
                    link_generado += f"&empresa_id={emp_id_link}"

                st.code(link_generado, language="text")
                st.info("Copie este enlace y envíelo al usuario para su registro.")

        if not usuarios_df.empty:
            st.divider()
            c_rol, c_emp = st.columns(2)
            
            with c_rol:
                st.subheader("🔄 Cambiar Rol de Usuario")
                with st.form("form_cambiar_rol"):
                    usr_sel_r = st.selectbox("Usuario", usuarios_df["id"].tolist(), format_func=lambda x: f"{usuarios_df.loc[usuarios_df['id']==x, 'nombre'].values[0]} ({usuarios_df.loc[usuarios_df['id']==x, 'rol'].values[0]})", key="sel_rol")
                    nuevo_rol_sel = st.selectbox("Nuevo Rol", ["Preparador", "Revisor", "Administrador"])
                    if st.form_submit_button("Actualizar Rol", type="primary"):
                        actualizar_rol_usuario(usr_sel_r, nuevo_rol_sel)
                        st.success("Rol actualizado con éxito.")
                        st.rerun()

            with c_emp:
                if not empresas_df.empty:
                    st.subheader("🏢 Cambiar Empresa Asignada")
                    with st.form("form_asignar_empresa"):
                        usr_sel_e = st.selectbox("Usuario", usuarios_df["id"].tolist(), format_func=lambda x: usuarios_df.loc[usuarios_df["id"]==x, "nombre"].values[0], key="sel_emp")
                        emp_sel = st.selectbox("Empresa", empresas_df["id"].tolist(), format_func=lambda x: empresas_df.loc[empresas_df["id"]==x, "nombre"].values[0])
                        if st.form_submit_button("Asignar Empresa", type="primary"):
                            actualizar_empresa_usuario(usr_sel_e, emp_sel)
                            st.success("Empresa asignada correctamente.")
                            st.rerun()
