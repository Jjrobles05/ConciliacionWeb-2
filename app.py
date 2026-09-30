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
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Conciliación Bancaria",
    layout="wide"
)


# =========================================================
# BASE DE DATOS - ETAPA 2: HISTORIAL
# =========================================================

DB_FILE = "conciliaciones.db"


def conectar_db():
    """Abre la base SQLite del historial."""
    return sqlite3.connect(DB_FILE)


def inicializar_db():
    """Crea la tabla de historial si todavía no existe."""
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
        conn.commit()


def guardar_conciliacion_historial(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial,
    diferencia_conciliada, resultado_final, salidas_extracto,
    salidas_libros, entradas_libros, entradas_extracto,
    gastos_bancarios, preparado_por, revisado_por, excel_data
):
    """Guarda una conciliación completa en SQLite."""
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
    """Obtiene el historial más reciente primero."""
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
    """Obtiene una conciliación individual por ID."""
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
    """Elimina una conciliación del historial."""
    with conectar_db() as conn:
        conn.execute("DELETE FROM conciliaciones WHERE id = ?", (int(id_conciliacion),))
        conn.commit()


def _datos_conciliacion(empresa, nit, mes, preparado_por, revisado_por,
                        salidas_extracto, salidas_libros, entradas_libros,
                        entradas_extracto, gastos_bancarios):
    def dataframe_a_registros(df):
        limpio = limpiar_dataframe(df)
        if limpio is None or limpio.empty:
            return []
        return json.loads(limpio.to_json(orient="records", date_format="iso"))
    return {
        "salidas_extracto": dataframe_a_registros(salidas_extracto),
        "salidas_libros": dataframe_a_registros(salidas_libros),
        "entradas_libros": dataframe_a_registros(entradas_libros),
        "entradas_extracto": dataframe_a_registros(entradas_extracto),
        "gastos_bancarios": dataframe_a_registros(gastos_bancarios),
        "preparado_por": preparado_por,
        "revisado_por": revisado_por,
    }


def actualizar_conciliacion_historial(
    id_conciliacion, empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
    resultado_final, salidas_extracto, salidas_libros, entradas_libros,
    entradas_extracto, gastos_bancarios, preparado_por, revisado_por,
    excel_data, usuario_actual
):
    estado = (
        "CONCILIACIÓN BANCARIA CORRECTA"
        if abs(resultado_final) < 0.005
        else "CONCILIACIÓN CON DIFERENCIA"
    )
    datos = _datos_conciliacion(
        empresa, nit, mes, preparado_por, revisado_por,
        salidas_extracto, salidas_libros, entradas_libros,
        entradas_extracto, gastos_bancarios
    )
    fecha_elaboracion_texto = (
        fecha_elaboracion.isoformat()
        if hasattr(fecha_elaboracion, "isoformat") else str(fecha_elaboracion)
    )
    with conectar_db() as conn:
        conn.execute(
            """
            UPDATE conciliaciones SET
                fecha_guardado = ?, empresa = ?, nit = ?, mes = ?, fecha_elaboracion = ?,
                banco = ?, cuenta = ?, tipo = ?, saldo_extracto = ?, saldo_libros = ?,
                diferencia_inicial = ?, diferencia_conciliada = ?, resultado_final = ?,
                estado = ?, datos_json = ?, excel = ?,
                workflow_status = 'Pendiente de revisión',
                revisado_por_usuario = NULL, fecha_revision = NULL,
                motivo_correccion = NULL, usuario_ultima_accion = ?
            WHERE id = ?
            """,
            (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"), empresa, nit, mes,
                fecha_elaboracion_texto, banco, cuenta, tipo, float(saldo_extracto),
                float(saldo_libros), float(diferencia_inicial), float(diferencia_conciliada),
                float(resultado_final), estado, json.dumps(datos, ensure_ascii=False),
                sqlite3.Binary(excel_data), usuario_actual, int(id_conciliacion)
            )
        )
        conn.commit()


def solicitar_correccion(id_conciliacion, usuario_revisor, motivo):
    with conectar_db() as conn:
        conn.execute(
            """UPDATE conciliaciones
               SET workflow_status = 'Corrección habilitada',
                   revisado_por_usuario = ?, fecha_revision = ?,
                   motivo_correccion = ?, usuario_ultima_accion = ?
             WHERE id = ?""",
            (usuario_revisor, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             motivo.strip(), usuario_revisor, int(id_conciliacion))
        )
        conn.commit()


def aprobar_conciliacion(id_conciliacion, usuario_revisor):
    with conectar_db() as conn:
        conn.execute(
            """UPDATE conciliaciones
               SET workflow_status = 'Aprobada',
                   revisado_por_usuario = ?, fecha_revision = ?,
                   motivo_correccion = NULL, usuario_ultima_accion = ?
             WHERE id = ?""",
            (usuario_revisor, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             usuario_revisor, int(id_conciliacion))
        )
        conn.commit()


def obtener_datos_para_edicion(id_conciliacion):
    fila = obtener_conciliacion(id_conciliacion)
    if not fila:
        return None
    datos = json.loads(fila[15] or "{}")
    def df_desde_registros(registros, columnas):
        if registros:
            return pd.DataFrame(registros)
        return pd.DataFrame(columns=columnas)
    return {
        "id": fila[0], "fecha_guardado": fila[1], "empresa": fila[2] or "",
        "nit": fila[3] or "", "mes": fila[4] or "",
        "fecha_elaboracion": fila[5], "banco": fila[6] or "", "cuenta": fila[7] or "",
        "tipo": fila[8] or "Ahorros", "saldo_extracto": fila[9], "saldo_libros": fila[10],
        "salidas_extracto": df_desde_registros(datos.get("salidas_extracto", []), ["Fecha","Beneficiario","Documento","Valor"]),
        "salidas_libros": df_desde_registros(datos.get("salidas_libros", []), ["Fecha","Concepto","Valor"]),
        "entradas_libros": df_desde_registros(datos.get("entradas_libros", []), ["Fecha","Concepto","Valor"]),
        "entradas_extracto": df_desde_registros(datos.get("entradas_extracto", []), ["Fecha","Concepto","Valor"]),
        "gastos_bancarios": df_desde_registros(datos.get("gastos_bancarios", []), ["Fecha","4 x 1000","Cuota de manejo","IVA","Rte. fuente","Comisión","Ing. x intereses"]),
        "preparado_por": datos.get("preparado_por", "") or "",
        "revisado_por": datos.get("revisado_por", "") or "",
        "workflow_status": fila[17] or "Pendiente de revisión",
        "motivo_correccion": fila[20] or "",
    }


inicializar_db()


# =========================================================
# ETAPA 4 - AUTENTICACIÓN Y USUARIOS
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


def recuperar_usuario_por_nombre_y_clave(nombre, password):
    """Recupera el nombre de usuario sin exponerlo solo con el nombre.

    Se exige nombre completo + contraseña para evitar revelar usuarios
    a cualquier persona que simplemente conozca el nombre de otra persona.
    """
    with conectar_db() as conn:
        filas = conn.execute(
            """
            SELECT id, usuario, nombre, password_hash, salt, rol
            FROM usuarios
            WHERE activo = 1 AND lower(trim(nombre)) = lower(trim(?))
            ORDER BY id
            """,
            (nombre.strip(),)
        ).fetchall()

    coincidencias = []
    for fila in filas:
        if verificar_password(password, fila[4], fila[3]):
            coincidencias.append({
                "id": fila[0],
                "usuario": fila[1],
                "nombre": fila[2],
                "rol": fila[5],
            })

    if len(coincidencias) == 1:
        return coincidencias[0]
    return None


def obtener_usuarios():
    with conectar_db() as conn:
        return pd.read_sql_query(
            "SELECT id, usuario, nombre, rol, activo, fecha_creacion FROM usuarios ORDER BY id",
            conn
        )


def cambiar_estado_usuario(id_usuario, activo):
    with conectar_db() as conn:
        conn.execute("UPDATE usuarios SET activo = ? WHERE id = ?", (int(activo), int(id_usuario)))
        conn.commit()


def cambiar_password_usuario(id_usuario, nueva_password):
    salt, password_hash = hash_password(nueva_password)
    with conectar_db() as conn:
        conn.execute(
            "UPDATE usuarios SET password_hash = ?, salt = ? WHERE id = ?",
            (password_hash, salt, int(id_usuario))
        )
        conn.commit()


def tiene_permiso(rol, permiso):
    permisos = {
        "Administrador": {"editar", "historial", "eliminar", "usuarios"},
        "Preparador": {"editar", "historial"},
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
                st.success(f"Administrador creado correctamente. Tu usuario es: {usuario.strip()}")
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
                st.error("Usuario o contraseña incorrectos, o usuario inactivo.")

        with st.expander("🔎 ¿Olvidaste tu nombre de usuario?"):
            st.caption("Para recuperar el usuario debes ingresar tu nombre completo y tu contraseña. La contraseña no se muestra ni se guarda en este formulario.")
            with st.form("form_recuperar_usuario"):
                nombre_recuperacion = st.text_input("Nombre completo", key="recuperacion_nombre")
                clave_recuperacion = st.text_input("Contraseña", type="password", key="recuperacion_clave")
                recuperar = st.form_submit_button("Mostrar mi usuario", width="stretch")
            if recuperar:
                if not nombre_recuperacion.strip() or not clave_recuperacion:
                    st.error("Completa el nombre completo y la contraseña.")
                else:
                    datos_recuperados = recuperar_usuario_por_nombre_y_clave(
                        nombre_recuperacion, clave_recuperacion
                    )
                    if datos_recuperados:
                        st.success(f"Tu nombre de usuario es: {datos_recuperados['usuario']}")
                        st.info(f"Rol: {datos_recuperados['rol']}")
                    else:
                        st.error("No se encontró una cuenta activa que coincida con esos datos.")
        st.stop()


iniciar_autenticacion()
usuario_actual = st.session_state.usuario_autenticado
rol_actual = usuario_actual["rol"]

st.title("📑 Conciliación Bancaria")
with st.sidebar:
    st.markdown("### 👤 Usuario")
    st.write(f"**{usuario_actual['nombre']}**")
    st.caption(f"Rol: {rol_actual}")
    if st.button("Cerrar sesión", width="stretch"):
        st.session_state.usuario_autenticado = None
        st.rerun()


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

        fila += 1

    if df.empty:
        fila += 1

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
    diferencia_conciliada,
    resultado_final,
    salidas_extracto,
    salidas_libros,
    entradas_libros,
    entradas_extracto,
    gastos_bancarios,
    preparado_por,
    revisado_por
):
    """
    Genera UN SOLO archivo Excel con UNA SOLA HOJA.

    La lógica de conciliación sigue el modelo del Excel de referencia:

    H15 = saldo extracto - saldo libros

    H18 = + salidas no registradas en extracto
    H19 = - salidas bancarias no contabilizadas en libros
    H20 = + entradas bancarias no contabilizadas en libros
    H21 = - entradas no evidenciadas en extractos

    H22 = suma de H18:H21
    H23 = H15 - H22
    """

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

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.40
    ws.page_margins.bottom = 0.40

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

    ws.merge_cells("A1:G1")
    ws["A1"] = "CONCILIACIÓN BANCARIA"
    estilo_titulo(ws["A1"], 16)
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
    # JUSTIFICACIÓN
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
            total_columna(salidas_extracto)
        ),
        (
            "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
            total_columna(salidas_libros)
        ),
        (
            "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
            total_columna(entradas_libros)
        ),
        (
            "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS",
            total_columna(entradas_extracto)
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

    # -----------------------------------------------------
    # RESUMEN DE CÁLCULO
    # -----------------------------------------------------

    fila += 1

    fila = escribir_seccion(
        ws,
        fila,
        "CÁLCULO DE LA CONCILIACIÓN",
        4
    )

    calculo = [
        (
            "DIFERENCIA A JUSTIFICAR",
            diferencia_inicial
        ),
        (
            "DIFERENCIA CONCILIADA",
            diferencia_conciliada
        ),
        (
            "RESULTADO FINAL",
            resultado_final
        ),
    ]

    for nombre, valor in calculo:

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
    ).value = "DIFERENCIA A JUSTIFICAR"

    ws.cell(
        row=fila,
        column=2
    ).value = diferencia_inicial
    ws.cell(
        row=fila,
        column=2
    ).number_format = '#,##0.00'

    fila += 1

    ws.cell(
        row=fila,
        column=1
    ).value = "DIFERENCIA CONCILIADA"

    ws.cell(
        row=fila,
        column=2
    ).value = diferencia_conciliada
    ws.cell(
        row=fila,
        column=2
    ).number_format = '#,##0.00'

    fila += 1

    ws.cell(
        row=fila,
        column=1
    ).value = "RESULTADO FINAL"

    ws.cell(
        row=fila,
        column=2
    ).value = resultado_final
    ws.cell(
        row=fila,
        column=2
    ).number_format = '#,##0.00'

    for r in range(fila - 2, fila + 1):
        ws.cell(
            row=r,
            column=1
        ).font = Font(bold=True)
        ws.cell(
            row=r,
            column=2
        ).font = Font(bold=True)

    fila += 1

    estado = (
        "CONCILIACIÓN BANCARIA CORRECTA"
        if abs(resultado_final) < 0.005
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
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    for fila_excel in ws.iter_rows():
        for celda in fila_excel:
            if celda.value is not None:
                celda.border = borde
                celda.alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

    for fila_excel in ws.iter_rows():
        for celda in fila_excel:
            if (
                celda.fill
                and celda.fill.fill_type == "solid"
                and celda.fill.fgColor.rgb
                and "1F4E78" in str(celda.fill.fgColor.rgb)
            ):
                celda.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

    ws.freeze_panes = "A3"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    nombre_archivo = (
        f"CONCILIACION_"
        f"{limpiar_nombre_archivo(empresa)}_"
        f"{limpiar_nombre_archivo(mes)}.xlsx"
    )

    return output.getvalue(), nombre_archivo


# =========================================================
# FLUJO DE EDICIÓN Y REVISIÓN
# =========================================================

if "edicion_conciliacion_id" not in st.session_state:
    st.session_state.edicion_conciliacion_id = None


def cargar_formulario_desde_historial(id_conciliacion):
    datos = obtener_datos_para_edicion(id_conciliacion)
    if not datos:
        return False

    st.session_state.form_empresa = datos["empresa"]
    st.session_state.form_nit = datos["nit"]
    st.session_state.form_mes = datos["mes"]
    try:
        st.session_state.form_fecha_elaboracion = datetime.fromisoformat(str(datos["fecha_elaboracion"])).date()
    except Exception:
        st.session_state.form_fecha_elaboracion = datetime.now().date()
    st.session_state.form_banco = datos["banco"]
    st.session_state.form_cuenta = datos["cuenta"]
    st.session_state.form_tipo = datos["tipo"] if datos["tipo"] in ["Ahorros", "Corriente"] else "Ahorros"
    st.session_state.form_saldo_extracto = float(datos["saldo_extracto"] or 0)
    st.session_state.form_saldo_libros = float(datos["saldo_libros"] or 0)
    st.session_state.tabla1 = datos["salidas_extracto"]
    st.session_state.tabla2 = datos["salidas_libros"]
    st.session_state.tabla3 = datos["entradas_libros"]
    st.session_state.tabla4 = datos["entradas_extracto"]
    st.session_state.tabla5 = datos["gastos_bancarios"]
    for key in ["tabla1", "tabla2", "tabla3", "tabla4", "tabla5"]:
        df = st.session_state[key].copy()
        if "Fecha" in df.columns and not df.empty:
            df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        st.session_state[key] = df
    st.session_state.form_preparado_por = datos["preparado_por"]
    st.session_state.form_revisado_por = datos["revisado_por"]
    st.session_state.edicion_conciliacion_id = int(id_conciliacion)
    st.session_state.edicion_cargada = int(id_conciliacion)
    return True


def limpiar_formulario_edicion():
    st.session_state.edicion_conciliacion_id = None
    st.session_state.edicion_cargada = None
    st.session_state.form_empresa = ""
    st.session_state.form_nit = ""
    st.session_state.form_mes = ""
    st.session_state.form_fecha_elaboracion = datetime.now().date()
    st.session_state.form_banco = ""
    st.session_state.form_cuenta = ""
    st.session_state.form_tipo = "Ahorros"
    st.session_state.form_saldo_extracto = 0.0
    st.session_state.form_saldo_libros = 0.0
    st.session_state.tabla1 = pd.DataFrame(columns=["Fecha", "Beneficiario", "Documento", "Valor"])
    st.session_state.tabla2 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    st.session_state.tabla3 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    st.session_state.tabla4 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
    st.session_state.tabla5 = pd.DataFrame(columns=["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"])
    st.session_state.form_preparado_por = ""
    st.session_state.form_revisado_por = ""


if "edicion_cargada" not in st.session_state:
    st.session_state.edicion_cargada = None

# Inicializar las tablas de movimientos antes de mostrarlas.
# Esto evita errores al abrir una conciliación nueva.
if "tabla1" not in st.session_state:
    st.session_state.tabla1 = pd.DataFrame(columns=["Fecha", "Beneficiario", "Documento", "Valor"])
if "tabla2" not in st.session_state:
    st.session_state.tabla2 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
if "tabla3" not in st.session_state:
    st.session_state.tabla3 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
if "tabla4" not in st.session_state:
    st.session_state.tabla4 = pd.DataFrame(columns=["Fecha", "Concepto", "Valor"])
if "tabla5" not in st.session_state:
    st.session_state.tabla5 = pd.DataFrame(
        columns=["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"]
    )

if st.session_state.edicion_conciliacion_id and st.session_state.edicion_cargada != st.session_state.edicion_conciliacion_id:
    if not cargar_formulario_desde_historial(st.session_state.edicion_conciliacion_id):
        limpiar_formulario_edicion()


# =========================================================
# CONCILIACIÓN - CAPTURA Y PREPARACIÓN
# =========================================================

if tiene_permiso(rol_actual, "editar"):

    if st.session_state.edicion_conciliacion_id:
        st.info(f"✏️ Modo corrección — estás editando la conciliación #{st.session_state.edicion_conciliacion_id}. Al guardar volverá a quedar pendiente de revisión.")
        if st.button("Cancelar corrección", key="cancelar_correccion", width="stretch"):
            limpiar_formulario_edicion()
            st.rerun()

    # =========================================================
    # INFORMACIÓN GENERAL
    # =========================================================
    
    st.subheader("Información General")
    
    col1, col2 = st.columns(2)
    
    with col1:
        empresa = st.text_input("Nombre de la Empresa", key="form_empresa")
        nit = st.text_input("NIT", key="form_nit")
        mes = st.text_input("Mes y Año", key="form_mes")
        fecha_elaboracion = st.date_input("Fecha de elaboración", key="form_fecha_elaboracion")
    
    with col2:
        banco = st.text_input("Nombre del Banco", key="form_banco")
        cuenta = st.text_input("Número de Cuenta", key="form_cuenta")
        tipo = st.selectbox(
            "Tipo de Cuenta",
            ["Ahorros", "Corriente"],
            key="form_tipo"
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
            format="%.2f",
            key="form_saldo_extracto"
        )
    
    with col2:
        saldo_libros = st.number_input(
            "Saldo según Libros",
            format="%.2f",
            key="form_saldo_libros"
        )
    
    diferencia_inicial = saldo_extracto - saldo_libros
    
    st.metric(
        "Diferencia a Justificar",
        f"${diferencia_inicial:,.2f}"
    )
    
    
    # =========================================================
    # MOVIMIENTOS
    # =========================================================
    
    st.divider()
    st.subheader("Salidas no Registradas en Extracto")
    
    salidas_extracto = st.data_editor(
        st.session_state.tabla1,
        num_rows="dynamic",
        width="stretch",
        key="tabla1",
        column_config={
            "Fecha": st.column_config.DateColumn(
                "Fecha",
                format="DD/MM/YYYY"
            ),
            "Beneficiario": st.column_config.TextColumn(
                "Beneficiario"
            ),
            "Documento": st.column_config.TextColumn(
                "Documento"
            ),
            "Valor": st.column_config.NumberColumn(
                "Valor",
                format="$ %.2f"
            )
        }
    )
    
    
    st.divider()
    st.subheader("Salidas Bancarias no Contabilizadas en Libros")
    
    salidas_libros = st.data_editor(
        st.session_state.tabla2,
        num_rows="dynamic",
        width="stretch",
        key="tabla2",
        column_config={
            "Fecha": st.column_config.DateColumn(
                "Fecha",
                format="DD/MM/YYYY"
            ),
            "Concepto": st.column_config.TextColumn(
                "Concepto"
            ),
            "Valor": st.column_config.NumberColumn(
                "Valor",
                format="$ %.2f"
            )
        }
    )
    
    
    st.divider()
    st.subheader("Entradas Bancarias no Contabilizadas en Libros")
    
    entradas_libros = st.data_editor(
        st.session_state.tabla3,
        num_rows="dynamic",
        width="stretch",
        key="tabla3",
        column_config={
            "Fecha": st.column_config.DateColumn(
                "Fecha",
                format="DD/MM/YYYY"
            ),
            "Concepto": st.column_config.TextColumn(
                "Concepto"
            ),
            "Valor": st.column_config.NumberColumn(
                "Valor",
                format="$ %.2f"
            )
        }
    )
    
    
    st.divider()
    st.subheader("Entradas no Evidenciadas en Extractos")
    
    entradas_extracto = st.data_editor(
        st.session_state.tabla4,
        num_rows="dynamic",
        width="stretch",
        key="tabla4",
        column_config={
            "Fecha": st.column_config.DateColumn(
                "Fecha",
                format="DD/MM/YYYY"
            ),
            "Concepto": st.column_config.TextColumn(
                "Concepto"
            ),
            "Valor": st.column_config.NumberColumn(
                "Valor",
                format="$ %.2f"
            )
        }
    )
    
    
    # =========================================================
    # GASTOS BANCARIOS
    # =========================================================
    
    st.divider()
    st.subheader("Gastos Bancarios")
    
    st.caption(
        "Esta sección se mantiene separada de las cuatro partidas "
        "de conciliación, como en el formato de referencia."
    )
    
    gastos_bancarios = st.data_editor(
        st.session_state.tabla5,
        num_rows="dynamic",
        width="stretch",
        key="tabla5",
        column_config={
            "Fecha": st.column_config.DateColumn(
                "Fecha",
                format="DD/MM/YYYY"
            ),
            "4 x 1000": st.column_config.NumberColumn(
                "4 x 1000",
                format="$ %.2f"
            ),
            "Cuota de manejo": st.column_config.NumberColumn(
                "Cuota de manejo",
                format="$ %.2f"
            ),
            "IVA": st.column_config.NumberColumn(
                "IVA",
                format="$ %.2f"
            ),
            "Rte. fuente": st.column_config.NumberColumn(
                "Rte. fuente",
                format="$ %.2f"
            ),
            "Comisión": st.column_config.NumberColumn(
                "Comisión",
                format="$ %.2f"
            ),
            "Ing. x intereses": st.column_config.NumberColumn(
                "Ing. x intereses",
                format="$ %.2f"
            )
        }
    )
    
    
    # =========================================================
    # TOTALES DE LAS CUATRO PARTIDAS
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
    
    
    # =========================================================
    # FÓRMULA SEGÚN EL EXCEL DE REFERENCIA
    # =========================================================
    
    # H15:
    # Diferencia a justificar = Extracto - Libros
    #
    # H18:
    # + Salidas no registradas en extracto
    #
    # H19:
    # - Salidas bancarias no contabilizadas en libros
    #
    # H20:
    # + Entradas bancarias no contabilizadas en libros
    #
    # H21:
    # - Entradas no evidenciadas en extractos
    #
    # H22:
    # Diferencia conciliada = H18 + H19 + H20 + H21
    #
    # H23:
    # Resultado final = H15 - H22
    
    diferencia_conciliada = (
        total_salidas_extracto
        - total_salidas_libros
        + total_entradas_libros
        - total_entradas_extracto
    )
    
    resultado_final = (
        diferencia_inicial
        - diferencia_conciliada
    )
    
    
    # =========================================================
    # RESULTADO
    # =========================================================
    
    st.divider()
    st.subheader("Resultado de la Conciliación")
    
    r1, r2, r3 = st.columns(3)
    
    with r1:
        st.metric(
            "Diferencia a Justificar",
            f"${diferencia_inicial:,.2f}"
        )
    
    with r2:
        st.metric(
            "Diferencia Conciliada",
            f"${diferencia_conciliada:,.2f}"
        )
    
    with r3:
        st.metric(
            "Resultado Final",
            f"${resultado_final:,.2f}"
        )
    
    conciliada = abs(resultado_final) < 0.005
    
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
    st.subheader("Responsables")
    
    col1, col2 = st.columns(2)
    
    with col1:
        preparado_por = st.text_input("Preparado por", key="form_preparado_por")
    
    with col2:
        revisado_por = st.text_input("Revisado por", key="form_revisado_por")
    
    
    # =========================================================
    # EXPORTACIÓN A EXCEL
    # =========================================================
    
    st.divider()
    st.subheader("Exportar Conciliación")
    
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
        diferencia_conciliada=diferencia_conciliada,
        resultado_final=resultado_final,
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
    
    # =========================================================
    # GUARDAR EN HISTORIAL
    # =========================================================
    
    st.divider()
    st.subheader("🗂️ Guardar en Historial")
    
    st.caption(
        "La conciliación se guarda en SQLite junto con sus movimientos "
        "y una copia del Excel generado."
    )
    
    texto_guardado = (
        "💾 Guardar corrección y enviar a revisión"
        if st.session_state.edicion_conciliacion_id
        else "💾 Guardar esta conciliación en el historial"
    )

    if st.button(texto_guardado, type="primary", width="stretch"):
        if st.session_state.edicion_conciliacion_id:
            id_guardado = st.session_state.edicion_conciliacion_id
            actualizar_conciliacion_historial(
                id_conciliacion=id_guardado, empresa=empresa, nit=nit, mes=mes,
                fecha_elaboracion=fecha_elaboracion, banco=banco, cuenta=cuenta, tipo=tipo,
                saldo_extracto=saldo_extracto, saldo_libros=saldo_libros,
                diferencia_inicial=diferencia_inicial, diferencia_conciliada=diferencia_conciliada,
                resultado_final=resultado_final, salidas_extracto=salidas_extracto,
                salidas_libros=salidas_libros, entradas_libros=entradas_libros,
                entradas_extracto=entradas_extracto, gastos_bancarios=gastos_bancarios,
                preparado_por=preparado_por, revisado_por=revisado_por,
                excel_data=excel_data, usuario_actual=usuario_actual["nombre"]
            )
            limpiar_formulario_edicion()
            st.success(f"✅ Conciliación #{id_guardado} corregida y enviada nuevamente a revisión.")
            st.rerun()
        else:
            id_guardado = guardar_conciliacion_historial(
                empresa=empresa, nit=nit, mes=mes, fecha_elaboracion=fecha_elaboracion,
                banco=banco, cuenta=cuenta, tipo=tipo, saldo_extracto=saldo_extracto,
                saldo_libros=saldo_libros, diferencia_inicial=diferencia_inicial,
                diferencia_conciliada=diferencia_conciliada, resultado_final=resultado_final,
                salidas_extracto=salidas_extracto, salidas_libros=salidas_libros,
                entradas_libros=entradas_libros, entradas_extracto=entradas_extracto,
                gastos_bancarios=gastos_bancarios, preparado_por=preparado_por,
                revisado_por=revisado_por, excel_data=excel_data
            )
            st.success(
                f"✅ Conciliación guardada correctamente. Número de historial: {id_guardado}. Estado: Pendiente de revisión."
            )
    
# =========================================================
# ETAPA 2 - HISTORIAL AVANZADO
# =========================================================

if rol_actual == "Revisor":
    st.info("Modo Revisor: puedes consultar y revisar el historial. La creación y eliminación están restringidas.")

st.divider()
st.subheader("📚 Historial de Conciliaciones")

historial = obtener_historial()

if historial.empty:
    st.info("Todavía no hay conciliaciones guardadas en el historial.")
else:
    st.caption(
        "Consulta las conciliaciones guardadas y utiliza los filtros para localizar "
        "rápidamente una empresa, período, banco o cuenta."
    )

    # -----------------------------------------------------
    # FILTROS
    # -----------------------------------------------------

    h1, h2, h3 = st.columns(3)

    with h1:
        filtro_empresa = st.text_input(
            "Empresa",
            placeholder="Ej. EMPRESA PRUEBA",
            key="hist_filtro_empresa"
        )

        filtro_nit = st.text_input(
            "NIT",
            placeholder="Número de identificación",
            key="hist_filtro_nit"
        )

    with h2:
        filtro_mes = st.text_input(
            "Mes / Año",
            placeholder="Ej. SEPTIEMBRE 2026",
            key="hist_filtro_mes"
        )

        filtro_banco = st.text_input(
            "Banco",
            placeholder="Ej. Bancolombia",
            key="hist_filtro_banco"
        )

    with h3:
        filtro_cuenta = st.text_input(
            "Cuenta",
            placeholder="Número de cuenta",
            key="hist_filtro_cuenta"
        )

        filtro_estado = st.selectbox(
            "Estado contable",
            [
                "Todos",
                "CONCILIACIÓN BANCARIA CORRECTA",
                "CONCILIACIÓN CON DIFERENCIA"
            ],
            key="hist_filtro_estado"
        )

        filtro_flujo = st.selectbox(
            "Estado de revisión",
            [
                "Todos",
                "Pendiente de revisión",
                "Corrección habilitada",
                "Aprobada"
            ],
            key="hist_filtro_flujo"
        )

    historial_filtrado = historial.copy()

    filtros_texto = [
        ("empresa", filtro_empresa),
        ("nit", filtro_nit),
        ("mes", filtro_mes),
        ("banco", filtro_banco),
        ("cuenta", filtro_cuenta),
    ]

    for columna, filtro in filtros_texto:
        if filtro.strip():
            historial_filtrado = historial_filtrado[
                historial_filtrado[columna]
                .fillna("")
                .astype(str)
                .str.contains(
                    filtro.strip(),
                    case=False,
                    na=False
                )
            ]

    if filtro_estado != "Todos":
        historial_filtrado = historial_filtrado[
            historial_filtrado["estado"] == filtro_estado
        ]

    if filtro_flujo != "Todos":
        historial_filtrado = historial_filtrado[
            historial_filtrado["workflow_status"] == filtro_flujo
        ]

    st.caption(
        f"Registros encontrados: {len(historial_filtrado)} de {len(historial)}"
    )

    # -----------------------------------------------------
    # TABLA DE HISTORIAL
    # -----------------------------------------------------

    mostrar_historial = historial_filtrado.rename(
        columns={
            "id": "ID",
            "fecha_guardado": "Guardado",
            "empresa": "Empresa",
            "nit": "NIT",
            "mes": "Mes/Año",
            "banco": "Banco",
            "cuenta": "Cuenta",
            "resultado_final": "Resultado final",
            "estado": "Estado",
            "workflow_status": "Estado de revisión"
        }
    )

    st.dataframe(
        mostrar_historial,
        width="stretch",
        hide_index=True,
        column_config={
            "ID": st.column_config.NumberColumn("ID", width="small"),
            "Resultado final": st.column_config.NumberColumn(
                "Resultado final",
                format="$ %.2f"
            ),
            "Estado": st.column_config.TextColumn(
                "Estado",
                width="large"
            )
        }
    )

    if historial_filtrado.empty:
        st.warning(
            "No se encontraron conciliaciones con los filtros seleccionados."
        )
    else:
        # -------------------------------------------------
        # SELECCIÓN DEL REGISTRO
        # -------------------------------------------------

        ids_disponibles = (
            historial_filtrado["id"]
            .astype(int)
            .tolist()
        )

        def texto_registro(id_registro):
            fila = historial.loc[
                historial["id"] == id_registro
            ].iloc[0]

            empresa_texto = fila["empresa"] or "Sin empresa"
            mes_texto = fila["mes"] or "Sin período"
            banco_texto = fila["banco"] or "Sin banco"

            return (
                f"#{id_registro} — {empresa_texto} — "
                f"{mes_texto} — {banco_texto}"
            )

        id_seleccionado = st.selectbox(
            "Selecciona una conciliación",
            ids_disponibles,
            format_func=texto_registro,
            key="hist_id_seleccionado"
        )

        registro = obtener_conciliacion(id_seleccionado)

        if registro:
            (
                rid,
                fecha_guardado,
                empresa_h,
                nit_h,
                mes_h,
                fecha_elaboracion_h,
                banco_h,
                cuenta_h,
                tipo_h,
                saldo_extracto_h,
                saldo_libros_h,
                diferencia_inicial_h,
                diferencia_conciliada_h,
                resultado_final_h,
                estado_h,
                datos_json_h,
                excel_h,
                workflow_status_h,
                revisado_por_usuario_h,
                fecha_revision_h,
                motivo_correccion_h,
                usuario_ultima_accion_h
            ) = registro

            # ---------------------------------------------
            # FICHA DE LA CONCILIACIÓN
            # ---------------------------------------------

            st.markdown(
                f"### Conciliación #{rid}"
            )

            info1, info2, info3 = st.columns(3)

            with info1:
                st.write(f"**Empresa:** {empresa_h or 'Sin empresa'}")
                st.write(f"**NIT:** {nit_h or 'Sin NIT'}")
                st.write(f"**Mes / Año:** {mes_h or 'Sin período'}")
                st.write(
                    f"**Fecha elaboración:** "
                    f"{fecha_elaboracion_h or 'Sin fecha'}"
                )

            with info2:
                st.write(f"**Banco:** {banco_h or 'Sin banco'}")
                st.write(f"**Cuenta:** {cuenta_h or 'Sin cuenta'}")
                st.write(f"**Tipo:** {tipo_h or 'Sin tipo'}")
                st.write(f"**Guardado:** {fecha_guardado}")

            with info3:
                st.metric(
                    "Saldo extracto",
                    f"${saldo_extracto_h:,.2f}"
                )
                st.metric(
                    "Saldo libros",
                    f"${saldo_libros_h:,.2f}"
                )

            st.divider()

            m1, m2, m3 = st.columns(3)

            with m1:
                st.metric(
                    "Diferencia a justificar",
                    f"${diferencia_inicial_h:,.2f}"
                )

            with m2:
                st.metric(
                    "Diferencia conciliada",
                    f"${diferencia_conciliada_h:,.2f}"
                )

            with m3:
                st.metric(
                    "Resultado final",
                    f"${resultado_final_h:,.2f}"
                )

            if estado_h == "CONCILIACIÓN BANCARIA CORRECTA":
                st.success(f"✅ {estado_h}")
            else:
                st.warning(f"⚠️ {estado_h}")

            if workflow_status_h == "Aprobada":
                st.success("🔒 Estado de revisión: APROBADA. La conciliación queda bloqueada para cambios.")
            elif workflow_status_h == "Corrección habilitada":
                st.warning("🔓 Estado de revisión: CORRECCIÓN HABILITADA. El Preparador puede corregirla y volver a enviarla a revisión.")
                if motivo_correccion_h:
                    st.info(f"Motivo de la corrección: {motivo_correccion_h}")
            else:
                st.info("🕒 Estado de revisión: PENDIENTE DE REVISIÓN")

            if revisado_por_usuario_h:
                st.caption(
                    f"Última revisión: {revisado_por_usuario_h} — {fecha_revision_h or 'sin fecha'}"
                )

            # ---------------------------------------------
            # DETALLE DE MOVIMIENTOS GUARDADOS
            # ---------------------------------------------

            try:
                datos_guardados = json.loads(datos_json_h or "{}")
            except (TypeError, json.JSONDecodeError):
                datos_guardados = {}

            with st.expander(
                "👁️ Ver detalle de movimientos guardados",
                expanded=False
            ):
                nombres_detalle = [
                    ("salidas_extracto", "Salidas no registradas en extracto"),
                    ("salidas_libros", "Salidas bancarias no contabilizadas en libros"),
                    ("entradas_libros", "Entradas bancarias no contabilizadas en libros"),
                    ("entradas_extracto", "Entradas no evidenciadas en extractos"),
                    ("gastos_bancarios", "Gastos bancarios"),
                ]

                for clave, titulo in nombres_detalle:
                    registros = datos_guardados.get(clave, [])
                    st.markdown(f"**{titulo}**")

                    if registros:
                        st.dataframe(
                            pd.DataFrame(registros),
                            width="stretch",
                            hide_index=True
                        )
                    else:
                        st.caption("Sin movimientos registrados.")

                st.write(
                    f"**Preparado por:** "
                    f"{datos_guardados.get('preparado_por') or 'Sin registrar'}"
                )
                st.write(
                    f"**Revisado por:** "
                    f"{datos_guardados.get('revisado_por') or 'Sin registrar'}"
                )

            # ---------------------------------------------
            # REVISIÓN Y AUTORIZACIÓN DE CORRECCIÓN
            # ---------------------------------------------

            if rol_actual in ["Revisor", "Administrador"] and workflow_status_h == "Pendiente de revisión":
                st.markdown("#### 🧐 Revisión de la conciliación")
                rv1, rv2 = st.columns(2)

                with rv1:
                    if st.button(
                        "✅ Aprobar conciliación",
                        key=f"aprobar_{rid}",
                        type="primary",
                        width="stretch"
                    ):
                        aprobar_conciliacion(rid, usuario_actual["nombre"])
                        st.success(f"Conciliación #{rid} aprobada correctamente.")
                        st.rerun()

                with rv2:
                    if st.button(
                        "🔓 Habilitar corrección",
                        key=f"habilitar_correccion_{rid}",
                        width="stretch"
                    ):
                        st.session_state[f"mostrar_motivo_{rid}"] = True

                if st.session_state.get(f"mostrar_motivo_{rid}"):
                    motivo = st.text_area(
                        "Motivo de la corrección",
                        placeholder="Indica qué debe corregir el Preparador.",
                        key=f"motivo_{rid}"
                    )
                    if st.button(
                        "Confirmar habilitación para corrección",
                        key=f"confirmar_correccion_{rid}",
                        type="primary",
                        width="stretch"
                    ):
                        if not motivo.strip():
                            st.error("Escribe el motivo de la corrección.")
                        else:
                            solicitar_correccion(rid, usuario_actual["nombre"], motivo)
                            st.session_state.pop(f"mostrar_motivo_{rid}", None)
                            st.success(f"Conciliación #{rid} habilitada para corrección.")
                            st.rerun()

            if rol_actual in ["Preparador", "Administrador"] and workflow_status_h == "Corrección habilitada":
                if st.button(
                    "✏️ Cargar esta conciliación para corregir",
                    key=f"editar_correccion_{rid}",
                    type="primary",
                    width="stretch"
                ):
                    cargar_formulario_desde_historial(rid)
                    st.rerun()

            # ---------------------------------------------
            # ACCIONES
            # ---------------------------------------------

            nombre_h = (
                f"CONCILIACION_"
                f"{limpiar_nombre_archivo(empresa_h)}_"
                f"{limpiar_nombre_archivo(mes_h)}.xlsx"
            )

            a1, a2 = st.columns(2)

            with a1:
                st.download_button(
                    "📥 Descargar Excel guardado",
                    data=excel_h,
                    file_name=nombre_h,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    key=f"descargar_historial_{rid}",
                    width="stretch"
                )

            with a2:
                if tiene_permiso(rol_actual, "eliminar"):
                    if st.button(
                        "🗑️ Eliminar esta conciliación",
                        key=f"eliminar_historial_{rid}",
                        width="stretch"
                    ):
                        st.session_state[
                            "confirmar_eliminacion_historial"
                        ] = rid
                else:
                    st.caption("Solo el Administrador puede eliminar conciliaciones.")

            # ---------------------------------------------
            # CONFIRMACIÓN DE ELIMINACIÓN
            # ---------------------------------------------

            if st.session_state.get(
                "confirmar_eliminacion_historial"
            ) == rid:
                st.warning(
                    f"Vas a eliminar la conciliación #{rid}. "
                    "Esta acción no se puede deshacer."
                )

                e1, e2 = st.columns(2)

                with e1:
                    if st.button(
                        "Sí, eliminar definitivamente",
                        key=f"confirmar_eliminar_{rid}",
                        type="primary",
                        width="stretch"
                    ):
                        eliminar_conciliacion(rid)
                        st.session_state.pop(
                            "confirmar_eliminacion_historial",
                            None
                        )
                        st.success(
                            f"Conciliación #{rid} eliminada correctamente."
                        )
                        st.rerun()

                with e2:
                    if st.button(
                        "Cancelar",
                        key=f"cancelar_eliminar_{rid}",
                        width="stretch"
                    ):
                        st.session_state.pop(
                            "confirmar_eliminacion_historial",
                            None
                        )
                        st.rerun()

# =========================================================
# ETAPA 4 - ADMINISTRACIÓN DE USUARIOS
# =========================================================

if tiene_permiso(rol_actual, "usuarios"):
    st.divider()
    st.subheader("👥 Administración de Usuarios")
    st.caption("Crea usuarios, asigna roles, activa o desactiva cuentas y restablece contraseñas.")

    with st.expander("➕ Crear nuevo usuario", expanded=False):
        with st.form("form_nuevo_usuario"):
            nuevo_usuario = st.text_input("Usuario", key="nuevo_usuario")
            nuevo_nombre = st.text_input("Nombre completo", key="nuevo_nombre")
            nuevo_password = st.text_input("Contraseña", type="password", key="nuevo_password")
            nuevo_confirmar = st.text_input("Confirmar contraseña", type="password", key="nuevo_confirmar")
            nuevo_rol = st.selectbox("Rol", ["Administrador", "Preparador", "Revisor"], key="nuevo_rol")
            crear_nuevo = st.form_submit_button("Crear usuario", type="primary")
        if crear_nuevo:
            if not nuevo_usuario.strip() or not nuevo_nombre.strip() or not nuevo_password:
                st.error("Completa todos los campos.")
            elif len(nuevo_password) < 8:
                st.error("La contraseña debe tener al menos 8 caracteres.")
            elif nuevo_password != nuevo_confirmar:
                st.error("Las contraseñas no coinciden.")
            else:
                try:
                    crear_usuario(nuevo_usuario, nuevo_nombre, nuevo_password, nuevo_rol)
                    st.success(f"Usuario {nuevo_usuario.strip()} creado correctamente.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Ese nombre de usuario ya existe.")

    usuarios_df = obtener_usuarios()
    mostrar_usuarios = usuarios_df.copy()
    mostrar_usuarios["Estado"] = mostrar_usuarios["activo"].map({1: "Activo", 0: "Inactivo"})
    mostrar_usuarios = mostrar_usuarios.rename(columns={"id":"ID","usuario":"Usuario","nombre":"Nombre","rol":"Rol","fecha_creacion":"Creado"})[["ID","Usuario","Nombre","Rol","Estado","Creado"]]
    st.dataframe(mostrar_usuarios, width="stretch", hide_index=True)

    if not usuarios_df.empty:
        st.markdown("#### Gestionar usuario")
        ids_usuarios = usuarios_df["id"].astype(int).tolist()
        def texto_usuario(id_usuario):
            fila = usuarios_df.loc[usuarios_df["id"] == id_usuario].iloc[0]
            estado = "Activo" if int(fila["activo"]) == 1 else "Inactivo"
            return f"#{id_usuario} — {fila['usuario']} — {fila['rol']} — {estado}"
        id_usuario_gestion = st.selectbox("Selecciona un usuario", ids_usuarios, format_func=texto_usuario, key="usuario_gestion")
        fila_usuario = usuarios_df.loc[usuarios_df["id"] == id_usuario_gestion].iloc[0]
        g1, g2 = st.columns(2)
        with g1:
            estado_actual = int(fila_usuario["activo"]) == 1
            texto_boton = "Desactivar usuario" if estado_actual else "Activar usuario"
            if st.button(texto_boton, key=f"estado_usuario_{id_usuario_gestion}", width="stretch"):
                cambiar_estado_usuario(id_usuario_gestion, not estado_actual)
                st.rerun()
        with g2:
            nueva_clave_admin = st.text_input("Nueva contraseña", type="password", key=f"nueva_clave_{id_usuario_gestion}")
            if st.button("Restablecer contraseña", key=f"reset_clave_{id_usuario_gestion}", width="stretch"):
                if len(nueva_clave_admin) < 8:
                    st.error("La contraseña debe tener al menos 8 caracteres.")
                else:
                    cambiar_password_usuario(id_usuario_gestion, nueva_clave_admin)
                    st.success("Contraseña actualizada correctamente.")

st.caption(
    "El Excel se genera en una sola hoja y utiliza la lógica de conciliación "
    "del formato de referencia. El historial se almacena en SQLite."
)
