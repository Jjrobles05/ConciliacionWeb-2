import io
import re
import json
import sqlite3
import hashlib
import hmac
import secrets
import random
import base64
import zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib.pagesizes import letter, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Sistema de Conciliación Bancaria",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. CONEXIÓN TURSO / SQLITE
# ==========================================
def conectar_db():
    """Conecta a Turso o SQLite local como respaldo."""
    if "TURSO_DATABASE_URL" in st.secrets and "TURSO_AUTH_TOKEN" in st.secrets:
        url = str(st.secrets["TURSO_DATABASE_URL"]).strip()
        token = str(st.secrets["TURSO_AUTH_TOKEN"]).strip()

        if url.startswith("libsql://"):
            url = "https://" + url[len("libsql://"):]

        try:
            import turso_serverless
            return turso_serverless.connect(url, auth_token=token)
        except Exception as e:
            st.error(f"⚠️ No fue posible conectar con Turso. {type(e).__name__}: {e}")
            st.stop()

    return sqlite3.connect("conciliaciones.db", check_same_thread=False)

# ==========================================
# 3. CREACIÓN Y ESTRUCTURA DE TABLAS
# ==========================================
def inicializar_db():
    """Crea únicamente las tablas si no existen usando el esquema Turso actual."""
    conn = conectar_db()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nit TEXT NOT NULL UNIQUE,
            razon_social TEXT NOT NULL
        )
    """)
    try:
        c.execute("ALTER TABLE empresas ADD COLUMN logo BLOB")
    except Exception:
        pass

    c.execute("""
        CREATE TABLE IF NOT EXISTS cuentas_bancarias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            banco TEXT NOT NULL,
            numero_cuenta TEXT NOT NULL,
            tipo_cuenta TEXT NOT NULL,
            empresa_id INTEGER,
            FOREIGN KEY(empresa_id) REFERENCES empresas(id)
        )
    """)
    try:
        c.execute("""
            CREATE TABLE IF NOT EXISTS cuentas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                empresa_id INTEGER NOT NULL,
                banco TEXT NOT NULL,
                numero_cuenta TEXT NOT NULL,
                tipo_cuenta TEXT NOT NULL,
                FOREIGN KEY (empresa_id) REFERENCES empresas(id)
            )
        """)
        c.execute("SELECT id, empresa_id, banco, numero_cuenta, tipo_cuenta FROM cuentas_bancarias ORDER BY id")
        cuentas_actuales = c.fetchall()

        for cuenta in cuentas_actuales:
            cuenta_id, empresa_id_cta, banco_cta, numero_cta, tipo_cta = cuenta
            c.execute("SELECT id FROM cuentas WHERE id=?", (int(cuenta_id),))
            if c.fetchone():
                c.execute("""
                    UPDATE cuentas
                       SET empresa_id=?, banco=?, numero_cuenta=?, tipo_cuenta=?
                     WHERE id=?
                """, (int(empresa_id_cta), str(banco_cta), str(numero_cta), str(tipo_cta), int(cuenta_id)))
            else:
                c.execute("""
                    INSERT INTO cuentas (id, empresa_id, banco, numero_cuenta, tipo_cuenta)
                    VALUES (?, ?, ?, ?, ?)
                """, (int(cuenta_id), int(empresa_id_cta), str(banco_cta), str(numero_cta), str(tipo_cta)))
        conn.commit()
    except Exception:
        pass

    c.execute("""
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
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS conciliaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER NOT NULL,
            cuenta_id INTEGER NOT NULL,
            periodo TEXT NOT NULL,
            saldo_libro REAL NOT NULL,
            saldo_banco REAL NOT NULL,
            estado TEXT NOT NULL,
            preparado_por TEXT,
            revisado_por TEXT,
            fecha_creacion TEXT,
            dictamen TEXT,
            observaciones TEXT,
            FOREIGN KEY(empresa_id) REFERENCES empresas(id),
            FOREIGN KEY(cuenta_id) REFERENCES cuentas_bancarias(id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS asignaciones_cuentas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER NOT NULL,
            cuenta_id INTEGER NOT NULL,
            usuario_id INTEGER NOT NULL,
            anio INTEGER NOT NULL,
            mes INTEGER NOT NULL,
            fecha_asignacion TEXT NOT NULL,
            FOREIGN KEY(empresa_id) REFERENCES empresas(id),
            FOREIGN KEY(cuenta_id) REFERENCES cuentas_bancarias(id),
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id),
            UNIQUE(cuenta_id, anio, mes)
        )
    """)
    conn.commit()
    conn.close()

inicializar_db()

# ==========================================
# 3B. COPIAS DE SEGURIDAD Y RESTAURACIÓN
# ==========================================
def _backup_encode_value(value):
    if isinstance(value, (bytes, bytearray, memoryview)):
        return {"__tipo__": "blob", "valor": base64.b64encode(bytes(value)).decode("ascii")}
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return {"__tipo__": "texto", "valor": str(value)}

def _backup_decode_value(value):
    if isinstance(value, dict) and value.get("__tipo__") == "blob":
        return base64.b64decode(value.get("valor", ""))
    if isinstance(value, dict) and value.get("__tipo__") == "texto":
        return value.get("valor", "")
    return value

def crear_copia_seguridad_bytes():
    conn = conectar_db()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT type, name, tbl_name, sql
            FROM sqlite_master
            WHERE name NOT LIKE 'sqlite_%'
              AND sql IS NOT NULL
              AND type IN ('table', 'index', 'trigger', 'view')
            ORDER BY CASE type WHEN 'table' THEN 1 WHEN 'index' THEN 2 WHEN 'trigger' THEN 3 ELSE 4 END, name
        """)
        objetos = c.fetchall()
        tablas = []
        esquemas = []

        for tipo, nombre, tbl_name, sql in objetos:
            esquemas.append({"type": tipo, "name": nombre, "tbl_name": tbl_name, "sql": sql})

        nombres_tablas = [str(nombre) for tipo, nombre, _tbl, _sql in objetos if tipo == "table"]
        for nombre in nombres_tablas:
            c.execute(f'SELECT * FROM "{nombre.replace(chr(34), chr(34)*2)}"')
            columnas = [d[0] for d in c.description or []]
            filas = [[_backup_encode_value(v) for v in fila] for fila in c.fetchall()]
            tablas.append({"name": nombre, "columns": columnas, "rows": filas})

        respaldo = {
            "formato": "ConciliacionWeb-Turso-Backup",
            "version": 1,
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "tablas": tablas,
            "esquemas": esquemas
        }
        contenido = json.dumps(respaldo, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        salida = io.BytesIO()
        with zipfile.ZipFile(salida, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("backup.json", contenido)
            z.writestr("README.txt", "COPIA DE SEGURIDAD - CONCILIACION WEB\n")
        return salida.getvalue()
    finally:
        conn.close()

def leer_copia_seguridad(archivo_bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(archivo_bytes), "r") as z:
            if "backup.json" not in z.namelist():
                raise ValueError("El ZIP no contiene backup.json.")
            data = json.loads(z.read("backup.json").decode("utf-8"))
        if data.get("formato") != "ConciliacionWeb-Turso-Backup":
            raise ValueError("El archivo no corresponde al formato de respaldo de Conciliación Web.")
        return data
    except zipfile.BadZipFile as e:
        raise ValueError("El archivo seleccionado no es un ZIP válido.") from e
    except json.JSONDecodeError as e:
        raise ValueError("El archivo de respaldo está dañado o incompleto.") from e

def restaurar_copia_seguridad(backup_data):
    conn = conectar_db()
    try:
        c = conn.cursor()
        try: c.execute("PRAGMA foreign_keys=OFF")
        except Exception: pass
        c.execute("BEGIN")

        esquemas = backup_data.get("esquemas", [])
        tablas_backup = [x for x in esquemas if x.get("type") == "table" and x.get("sql")]
        for obj in tablas_backup:
            sql = str(obj["sql"])
            c.execute(sql if "IF NOT EXISTS" in sql.upper() else _agregar_if_not_exists(sql))

        for tabla in backup_data.get("tablas", []):
            nombre = str(tabla.get("name", ""))
            columnas = [str(x) for x in tabla.get("columns", [])]
            if not nombre or not columnas:
                continue
            ident = nombre.replace('"', '""')
            c.execute(f'DELETE FROM "{ident}"')
            cols_sql = ", ".join(f'"{x.replace(chr(34), chr(34)*2)}"' for x in columnas)
            placeholders = ", ".join(["?"] * len(columnas))
            filas = [[_backup_decode_value(v) for v in fila] for fila in tabla.get("rows", [])]
            if filas:
                c.executemany(f'INSERT INTO "{ident}" ({cols_sql}) VALUES ({placeholders})', filas)

        for obj in esquemas:
            if obj.get("type") not in ("index", "trigger", "view") or not obj.get("sql"):
                continue
            sql = str(obj["sql"])
            try: c.execute(sql if "IF NOT EXISTS" in sql.upper() else _agregar_if_not_exists(sql))
            except Exception: pass

        conn.commit()
        try: c.execute("PRAGMA foreign_keys=ON")
        except Exception: pass
        st.cache_data.clear()
        return True, "Restauración completada correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible restaurar la copia: {type(e).__name__}: {e}"
    finally:
        conn.close()

def _agregar_if_not_exists(sql):
    texto = str(sql).strip()
    patron = re.compile(r"^(CREATE\s+(?:UNIQUE\s+)?(?:INDEX|TRIGGER|VIEW|TABLE)\s+)(?!IF\s+NOT\s+EXISTS)(.+)$", re.IGNORECASE | re.DOTALL)
    m = patron.match(texto)
    return m.group(1) + "IF NOT EXISTS " + m.group(2) if m else texto

# ==========================================
# 4. CONSULTAS OPTIMIZADAS (CON CACHÉ)
# ==========================================
@st.cache_data(ttl=600)
def obtener_empresas():
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, razon_social, nit, logo FROM empresas ORDER BY razon_social")
        rows = c.fetchall()
        conn.close()
        return pd.DataFrame(rows, columns=['id', 'nombre', 'nit', 'logo']) if rows else pd.DataFrame(columns=['id', 'nombre', 'nit', 'logo'])
    except Exception:
        return pd.DataFrame(columns=['id', 'nombre', 'nit', 'logo'])

def obtener_empresa_por_id(empresa_id):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, razon_social, nit, logo FROM empresas WHERE id=?", (int(empresa_id),))
        res = c.fetchone()
        conn.close()
        if res:
            return {"id": res[0], "nombre": res[1], "nit": res[2], "logo": res[3]}
    except Exception:
        pass
    return None

@st.cache_data(ttl=600)
def obtener_logo_empresa(nombre_empresa):
    if not nombre_empresa:
        return None
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT logo FROM empresas WHERE razon_social=? LIMIT 1", (str(nombre_empresa).strip(),))
        row = c.fetchone()
        conn.close()
        if row and row[0]:
            return bytes(row[0])
    except Exception:
        pass
    return None

def guardar_empresa(nombre, nit, logo_bytes=None):
    nombre = str(nombre or "").strip()
    nit = str(nit or "").strip()
    if not nombre or not nit:
        return False, "El nombre de la empresa y el NIT son obligatorios."
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT id, razon_social FROM empresas WHERE nit=?", (nit,))
        existente = c.fetchone()
        if existente:
            return False, f"El NIT {nit} ya está registrado para {existente[1]}."
        c.execute("INSERT INTO empresas (nit, razon_social, logo) VALUES (?, ?, ?)", (nit, nombre, logo_bytes))
        conn.commit()
        st.cache_data.clear()
        return True, "Empresa registrada con éxito."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible registrar la empresa: {type(e).__name__}: {e}"
    finally:
        conn.close()

def actualizar_empresa_db(empresa_id, nombre, nit, logo_bytes=None):
    nombre = str(nombre or "").strip()
    nit = str(nit or "").strip()
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT id, razon_social FROM empresas WHERE nit=? AND id<>?", (nit, int(empresa_id)))
        existente = c.fetchone()
        if existente:
            return False, f"El NIT {nit} ya está registrado para {existente[1]}."
        if logo_bytes is not None:
            c.execute("UPDATE empresas SET razon_social=?, nit=?, logo=? WHERE id=?", (nombre, nit, logo_bytes, int(empresa_id)))
        else:
            c.execute("UPDATE empresas SET razon_social=?, nit=? WHERE id=?", (nombre, nit, int(empresa_id)))
        conn.commit()
        st.cache_data.clear()
        return True, "Empresa actualizada con éxito."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible actualizar la empresa: {type(e).__name__}: {e}"
    finally:
        conn.close()

def eliminar_empresa_db(empresa_id):
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT COUNT(*) FROM usuarios WHERE empresa_id=?", (int(empresa_id),))
        usuarios = int(c.fetchone()[0] or 0)
        c.execute("SELECT COUNT(*) FROM cuentas_bancarias WHERE empresa_id=?", (int(empresa_id),))
        cuentas = int(c.fetchone()[0] or 0)
        c.execute("SELECT COUNT(*) FROM conciliaciones WHERE empresa_id=?", (int(empresa_id),))
        conciliaciones = int(c.fetchone()[0] or 0)
        if usuarios or cuentas or conciliaciones:
            partes = []
            if usuarios: partes.append(f"{usuarios} usuario(s)")
            if cuentas: partes.append(f"{cuentas} cuenta(s)")
            if conciliaciones: partes.append(f"{conciliaciones} conciliación(es)")
            return False, "No se puede eliminar: la empresa tiene " + ", ".join(partes) + "."
        c.execute("DELETE FROM empresas WHERE id=?", (int(empresa_id),))
        conn.commit()
        st.cache_data.clear()
        return True, "Empresa eliminada correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible eliminar la empresa: {type(e).__name__}: {e}"
    finally:
        conn.close()

def eliminar_conciliacion_db(conciliacion_id):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("DELETE FROM conciliaciones WHERE id=?", (int(conciliacion_id),))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def limpiar_datos_operativos():
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("DELETE FROM conciliaciones")
        c.execute("DELETE FROM cuentas_bancarias")
        try: c.execute("DELETE FROM cuentas")
        except Exception: pass

        for tabla in ["conciliaciones", "cuentas_bancarias", "cuentas"]:
            try: c.execute("DELETE FROM sqlite_sequence WHERE name=?", (tabla,))
            except Exception: pass

        conn.commit()
        st.cache_data.clear()
        return True, "Se eliminaron bancos, cuentas y conciliaciones. Empresas y usuarios se conservaron."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible limpiar los datos operativos: {type(e).__name__}: {e}"
    finally:
        conn.close()

def reiniciar_datos_aplicativo():
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("DELETE FROM conciliaciones")
        c.execute("DELETE FROM cuentas_bancarias")
        try: c.execute("DELETE FROM cuentas")
        except Exception: pass
        c.execute("DELETE FROM usuarios")
        c.execute("DELETE FROM empresas")

        for tabla in ["conciliaciones", "cuentas_bancarias", "cuentas", "usuarios", "empresas"]:
            try: c.execute("DELETE FROM sqlite_sequence WHERE name=?", (tabla,))
            except Exception: pass

        conn.commit()
        st.cache_data.clear()
        return True, "El aplicativo quedó completamente limpio. Ahora puedes crear el primer Administrador."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible reiniciar los datos: {type(e).__name__}: {e}"
    finally:
        conn.close()

@st.cache_data(ttl=300)
def obtener_cuentas(empresa_id=None):
    try:
        conn = conectar_db()
        c = conn.cursor()
        query = """
            SELECT c.id, c.banco, c.numero_cuenta, c.tipo_cuenta,
                   c.empresa_id, e.razon_social AS empresa_nombre
            FROM cuentas_bancarias c
            LEFT JOIN empresas e ON c.empresa_id = e.id
        """
        if empresa_id:
            query += " WHERE c.empresa_id = ? ORDER BY c.banco, c.numero_cuenta"
            c.execute(query, (int(empresa_id),))
        else:
            query += " ORDER BY c.banco, c.numero_cuenta"
            c.execute(query)
        rows = c.fetchall()
        conn.close()
        if rows:
            return pd.DataFrame(rows, columns=['id', 'banco', 'numero_cuenta', 'tipo_cuenta', 'empresa_id', 'empresa_nombre'])
    except Exception:
        pass
    return pd.DataFrame(columns=['id', 'banco', 'numero_cuenta', 'tipo_cuenta', 'empresa_id', 'empresa_nombre'])

def guardar_cuenta(banco, numero_cuenta, tipo_cuenta, empresa_id, datos_credito=None, usuario_id=None):
    banco = str(banco or "").strip()
    numero_cuenta = str(numero_cuenta or "").strip()
    tipo_cuenta = str(tipo_cuenta or "").strip()
    if not banco or not numero_cuenta or empresa_id is None:
        raise ValueError("Banco, número de cuenta y empresa son obligatorios.")

    tipos_validos = ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito", "Caja", "Crédito bancario"]
    if tipo_cuenta not in tipos_validos:
        raise ValueError("El tipo de cuenta seleccionado no es válido.")

    conn = conectar_db()
    c = conn.cursor()
    try:
        empresa_id = int(empresa_id)
        c.execute("SELECT id FROM empresas WHERE id=?", (empresa_id,))
        if not c.fetchone():
            raise ValueError("La empresa seleccionada no existe.")

        c.execute("SELECT id FROM cuentas_bancarias WHERE empresa_id=? AND banco=? AND numero_cuenta=? LIMIT 1",
                  (empresa_id, banco, numero_cuenta))
        existente = c.fetchone()

        if existente:
            cuenta_id = int(existente[0])
            c.execute("UPDATE cuentas_bancarias SET tipo_cuenta=? WHERE id=?", (tipo_cuenta, cuenta_id))
        else:
            c.execute("INSERT INTO cuentas_bancarias (banco, numero_cuenta, tipo_cuenta, empresa_id) VALUES (?, ?, ?, ?)",
                      (banco, numero_cuenta, tipo_cuenta, empresa_id))
            cuenta_id = int(c.lastrowid)

        conn.commit()
        st.cache_data.clear()
        return cuenta_id
    except Exception:
        try: conn.rollback()
        except Exception: pass
        raise
    finally:
        conn.close()

def actualizar_cuenta_db(cuenta_id, banco, numero_cuenta, tipo_cuenta, empresa_id):
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("UPDATE cuentas_bancarias SET banco=?, numero_cuenta=?, tipo_cuenta=?, empresa_id=? WHERE id=?",
                  (str(banco).strip(), str(numero_cuenta).strip(), str(tipo_cuenta), empresa_id, int(cuenta_id)))
        if c.rowcount == 0:
            return False, "No se encontró la cuenta para actualizar."
        conn.commit()
        st.cache_data.clear()
        return True, "Cuenta actualizada correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible actualizar la cuenta: {type(e).__name__}: {e}"
    finally:
        conn.close()

def eliminar_cuenta_db(cuenta_id):
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT banco, numero_cuenta FROM cuentas_bancarias WHERE id=?", (int(cuenta_id),))
        fila = c.fetchone()
        if not fila:
            return False, "La cuenta no existe."
        c.execute("SELECT COUNT(*) FROM conciliaciones WHERE cuenta_id=?", (int(cuenta_id),))
        conciliaciones = int(c.fetchone()[0] or 0)
        c.execute("SELECT COUNT(*) FROM asignaciones_cuentas WHERE cuenta_id=?", (int(cuenta_id),))
        asignaciones = int(c.fetchone()[0] or 0)
        if conciliaciones or asignaciones:
            partes = []
            if conciliaciones: partes.append(f"{conciliaciones} conciliación(es)")
            if asignaciones: partes.append(f"{asignaciones} asignación(es) mensuales")
            return False, "No se puede eliminar esta cuenta porque tiene " + " y ".join(partes) + "."
        c.execute("DELETE FROM cuentas_bancarias WHERE id=?", (int(cuenta_id),))
        conn.commit()
        st.cache_data.clear()
        return True, f"Cuenta {fila[0]} - {fila[1]} eliminada correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible eliminar la cuenta: {type(e).__name__}: {e}"
    finally:
        conn.close()

def preparar_excel_asignaciones(asignaciones_df, empresa, anio, mes_num):
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Asignaciones"
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    mes_nombre = meses[int(mes_num)-1]
    ws.merge_cells("A1:E1")
    ws["A1"] = f"ASIGNACIÓN MENSUAL DE CONCILIACIONES - {mes_nombre.upper()} {anio}"
    estilo_titulo(ws["A1"], 14)
    ws["A2"] = "Empresa"
    ws["B2"] = empresa
    ws["A2"].font = Font(bold=True)
    encabezados = ["Banco", "Cuenta / Tarjeta", "Tipo", "Preparador asignado", "Fecha de asignación"]
    fila = 4
    for col, encabezado in enumerate(encabezados, 1):
        celda = ws.cell(row=fila, column=col, value=encabezado)
        celda.font = Font(bold=True)
        celda.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        celda.alignment = Alignment(horizontal="center")
    fila += 1
    for _, r in asignaciones_df.iterrows():
        valores = [r.get("banco", ""), r.get("numero_cuenta", ""), r.get("tipo_cuenta", ""), r.get("preparador", ""), r.get("fecha_asignacion", "")]
        for col, valor in enumerate(valores, 1):
            ws.cell(row=fila, column=col, value=str(valor or ""))
        fila += 1
    for col, ancho in enumerate([22, 24, 22, 28, 22], 1):
        ws.column_dimensions[chr(64+col)].width = ancho

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    nombre = f"ASIGNACION_{limpiar_nombre_archivo(empresa)}_{mes_nombre.upper()}_{anio}.xlsx"
    return output.getvalue(), nombre

def generar_pdf_asignaciones(asignaciones_df, empresa, anio, mes_num):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(letter), rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    mes_nombre = meses[int(mes_num)-1]
    title_style = ParagraphStyle("AsignTitle", parent=styles["Heading1"], alignment=1, fontSize=13, spaceAfter=8)
    normal = ParagraphStyle("AsignNormal", parent=styles["Normal"], fontSize=8, leading=10)
    story.append(Paragraph("ASIGNACIÓN MENSUAL DE CONCILIACIONES", title_style))
    story.append(Paragraph(f"<b>Empresa:</b> {empresa}<br/><b>Periodo:</b> {mes_nombre} {anio}<br/><b>Estado:</b> Asignación fija", normal))
    story.append(Spacer(1, 10))
    datos = [["Banco", "Cuenta / Tarjeta", "Tipo", "Preparador"]]
    for _, r in asignaciones_df.iterrows():
        datos.append([str(r.get("banco", "")), str(r.get("numero_cuenta", "")), str(r.get("tipo_cuenta", "")), str(r.get("preparador", ""))])
    tabla = Table(datos, colWidths=[125, 125, 125, 150], repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F7F7")]),
    ]))
    story.append(tabla)
    doc.build(story)
    buffer.seek(0)
    nombre = f"ASIGNACION_{limpiar_nombre_archivo(empresa)}_{mes_nombre.upper()}_{anio}.pdf"
    return buffer.getvalue(), nombre

def obtener_asignaciones_mes(empresa_id, anio, mes_num):
    if not empresa_id:
        return pd.DataFrame()
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("""
            SELECT a.id, a.cuenta_id, c.banco, c.numero_cuenta, c.tipo_cuenta,
                   a.usuario_id, u.nombre AS preparador, u.usuario, a.anio, a.mes, a.fecha_asignacion
            FROM asignaciones_cuentas a
            JOIN cuentas_bancarias c ON c.id = a.cuenta_id
            JOIN usuarios u ON u.id = a.usuario_id
            WHERE a.empresa_id=? AND a.anio=? AND a.mes=?
            ORDER BY c.banco, c.numero_cuenta
        """, (int(empresa_id), int(anio), int(mes_num)))
        rows = c.fetchall()
        return pd.DataFrame(rows, columns=['id','cuenta_id','banco','numero_cuenta','tipo_cuenta','usuario_id','preparador','usuario','anio','mes','fecha_asignacion'])
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()

def contar_asignaciones_mes(empresa_id, anio, mes_num):
    if not empresa_id:
        return 0
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT COUNT(*) FROM asignaciones_cuentas WHERE empresa_id=? AND anio=? AND mes=?", (int(empresa_id), int(anio), int(mes_num)))
        return int(c.fetchone()[0] or 0)
    except Exception:
        return 0
    finally:
        conn.close()

def generar_asignacion_mensual(empresa_id, anio, mes_num):
    if not empresa_id:
        return False, "Selecciona una empresa."
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT COUNT(*) FROM asignaciones_cuentas WHERE empresa_id=? AND anio=? AND mes=?", (int(empresa_id), int(anio), int(mes_num)))
        if int(c.fetchone()[0] or 0):
            return False, "Este mes ya tiene una asignación generada."

        c.execute("SELECT id, nombre FROM usuarios WHERE empresa_id=? AND activo=1 AND rol='Preparador' ORDER BY id", (int(empresa_id),))
        preparadores = c.fetchall()
        c.execute("SELECT id FROM cuentas_bancarias WHERE empresa_id=? ORDER BY id", (int(empresa_id),))
        cuentas = [int(r[0]) for r in c.fetchall()]
        if not preparadores: return False, "No hay Preparadores activos asignados."
        if not cuentas: return False, "No hay cuentas bancarias registradas."

        usuarios = [int(r[0]) for r in preparadores]
        nombres = {int(r[0]): r[1] for r in preparadores}
        random.SystemRandom().shuffle(usuarios)
        random.SystemRandom().shuffle(cuentas)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for idx, cuenta_id in enumerate(cuentas):
            usuario_id = usuarios[idx % len(usuarios)]
            c.execute("""INSERT INTO asignaciones_cuentas (empresa_id, cuenta_id, usuario_id, anio, mes, fecha_asignacion)
                         VALUES (?,?,?,?,?,?)""", (int(empresa_id), cuenta_id, usuario_id, int(anio), int(mes_num), fecha))
        conn.commit()
        st.cache_data.clear()
        return True, f"Asignación generada correctamente para {mes_num:02d}/{anio}."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible generar la asignación: {type(e).__name__}: {e}"
    finally:
        conn.close()

def asignar_cuentas_nuevas_mes(empresa_id, anio, mes_num):
    if not empresa_id:
        return False, "Selecciona una empresa."

    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT id, nombre FROM usuarios WHERE empresa_id=? AND activo=1 AND rol='Preparador' ORDER BY id", (int(empresa_id),))
        preparadores = c.fetchall()
        if not preparadores: return False, "No hay Preparadores activos."

        c.execute("SELECT id FROM cuentas_bancarias WHERE empresa_id=? ORDER BY id", (int(empresa_id),))
        cuentas = [int(r[0]) for r in c.fetchall()]

        c.execute("SELECT cuenta_id, usuario_id FROM asignaciones_cuentas WHERE empresa_id=? AND anio=? AND mes=?", (int(empresa_id), int(anio), int(mes_num)))
        asignadas = c.fetchall()
        asignadas_ids = {int(r[0]) for r in asignadas}
        cuentas_nuevas = [cid for cid in cuentas if cid not in asignadas_ids]

        if not cuentas_nuevas:
            return False, "No hay cuentas nuevas pendientes de asignar."

        usuarios = [int(r[0]) for r in preparadores]
        cargas = {uid: 0 for uid in usuarios}
        for _, usuario_id in asignadas:
            if int(usuario_id) in cargas:
                cargas[int(usuario_id)] += 1

        rng = random.SystemRandom()
        rng.shuffle(cuentas_nuevas)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for cuenta_id in cuentas_nuevas:
            min_carga = min(cargas.values())
            candidatos = [uid for uid in usuarios if cargas[uid] == min_carga]
            usuario_id = rng.choice(candidatos)
            c.execute("""INSERT INTO asignaciones_cuentas (empresa_id, cuenta_id, usuario_id, anio, mes, fecha_asignacion)
                         VALUES (?,?,?,?,?,?)""", (int(empresa_id), cuenta_id, usuario_id, int(anio), int(mes_num), fecha))
            cargas[usuario_id] += 1

        conn.commit()
        st.cache_data.clear()
        return True, f"Se asignaron {len(cuentas_nuevas)} cuenta(s) nueva(s) para {mes_num:02d}/{anio}."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible asignar las cuentas nuevas: {type(e).__name__}: {e}"
    finally:
        conn.close()

def contar_cuentas_sin_asignar_mes(empresa_id, anio, mes_num):
    if not empresa_id:
        return 0
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("""
            SELECT COUNT(*)
            FROM cuentas_bancarias c
            WHERE c.empresa_id=?
              AND NOT EXISTS (
                  SELECT 1 FROM asignaciones_cuentas a
                  WHERE a.cuenta_id=c.id AND a.empresa_id=? AND a.anio=? AND a.mes=?
              )
        """, (int(empresa_id), int(empresa_id), int(anio), int(mes_num)))
        return int(c.fetchone()[0] or 0)
    except Exception:
        return 0
    finally:
        conn.close()

def obtener_cuentas_rotadas_por_usuario(empresa_id, anio, mes_num, usuario_id):
    if not empresa_id:
        return pd.DataFrame()
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT rol FROM usuarios WHERE id=?", (int(usuario_id),))
        fila = c.fetchone()
        rol_usuario = fila[0] if fila else None
        if rol_usuario != 'Preparador':
            c.execute("SELECT id, banco, numero_cuenta, tipo_cuenta FROM cuentas_bancarias WHERE empresa_id=? ORDER BY id", (int(empresa_id),))
        else:
            c.execute("""SELECT c.id, c.banco, c.numero_cuenta, c.tipo_cuenta
                         FROM asignaciones_cuentas a
                         JOIN cuentas_bancarias c ON c.id=a.cuenta_id
                         WHERE a.empresa_id=? AND a.anio=? AND a.mes=? AND a.usuario_id=?
                         ORDER BY c.banco, c.numero_cuenta""", (int(empresa_id), int(anio), int(mes_num), int(usuario_id)))
        rows = c.fetchall()
        return pd.DataFrame(rows, columns=['id','banco','numero_cuenta','tipo_cuenta'])
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()

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
                try: row_dict[col] = float(val_limpio)
                except ValueError: row_dict[col] = 0.0
            else:
                row_dict[col] = val
        rows.append(row_dict)
    return pd.DataFrame(rows)

def _json_datos_conciliacion(empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                              saldo_extracto, saldo_libros, diferencia_inicial,
                              diferencia_conciliada, resultado_final, salidas_extracto,
                              salidas_libros, entradas_libros, entradas_extracto,
                              gastos_bancarios, preparado_por, revisado_por, workflow_status):
    def dataframe_a_registros(df):
        limpio = limpiar_dataframe(df)
        if limpio is None or limpio.empty:
            return []
        return json.loads(limpio.to_json(orient="records", date_format="iso"))
    fecha_elaboracion_texto = fecha_elaboracion.isoformat() if hasattr(fecha_elaboracion, "isoformat") else str(fecha_elaboracion)
    return {
        "empresa": empresa, "nit": nit, "mes": mes,
        "fecha_elaboracion": fecha_elaboracion_texto, "banco": banco,
        "cuenta": cuenta, "tipo": tipo,
        "saldo_extracto": float(saldo_extracto), "saldo_libros": float(saldo_libros),
        "diferencia_inicial": float(diferencia_inicial),
        "diferencia_conciliada": float(diferencia_conciliada),
        "resultado_final": float(resultado_final),
        "estado": "CONCILIACIÓN BANCARIA CORRECTA" if abs(float(resultado_final)) < 0.005 else "CONCILIACIÓN CON DIFERENCIA",
        "salidas_extracto": dataframe_a_registros(salidas_extracto),
        "salidas_libros": dataframe_a_registros(salidas_libros),
        "entradas_libros": dataframe_a_registros(entradas_libros),
        "entradas_extracto": dataframe_a_registros(entradas_extracto),
        "gastos_bancarios": dataframe_a_registros(gastos_bancarios),
        "preparado_por": preparado_por, "revisado_por": revisado_por,
        "workflow_status": workflow_status,
    }

def _buscar_empresa_id(c, empresa, nit):
    c.execute("SELECT id FROM empresas WHERE nit=?", (str(nit).strip(),))
    row = c.fetchone()
    if row: return int(row[0])
    c.execute("SELECT id FROM empresas WHERE razon_social=?", (str(empresa).strip(),))
    row = c.fetchone()
    if row: return int(row[0])
    c.execute("INSERT INTO empresas (nit, razon_social) VALUES (?, ?)", (str(nit).strip(), str(empresa).strip()))
    return int(c.lastrowid)

def _buscar_cuenta_id(c, empresa_id, banco, cuenta, tipo):
    c.execute("SELECT id FROM cuentas_bancarias WHERE empresa_id=? AND banco=? AND numero_cuenta=? ORDER BY id LIMIT 1",
              (int(empresa_id), str(banco).strip(), str(cuenta).strip()))
    row = c.fetchone()
    if row: return int(row[0])
    c.execute("INSERT INTO cuentas_bancarias (banco, numero_cuenta, tipo_cuenta, empresa_id) VALUES (?, ?, ?, ?)",
              (str(banco).strip(), str(cuenta).strip(), str(tipo).strip(), int(empresa_id)))
    return int(c.lastrowid)

def construir_firmas_conciliacion(preparado_por="", revisado_por="", aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion=""):
    return {
        "preparado_por_usuario": preparado_por or "",
        "fecha_preparacion": fecha_preparacion or "",
        "revisado_por_usuario": revisado_por or "",
        "fecha_revision": fecha_revision or "",
        "aprobado_por_usuario": aprobado_por or "",
        "fecha_aprobacion": fecha_aprobacion or ""
    }

def renderizar_bloque_firmas_ui(preparado_por, fecha_preparacion="", revisado_por="", revisado_fecha="", aprobado_por="", aprobado_fecha="", estado="Pendiente de revisión"):
    st.divider()
    st.subheader("✍️ Firmas y trazabilidad")
    st.caption("Las firmas se generan automáticamente con el usuario autenticado y la fecha/hora de cada acción.")
    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown("**PREPARÓ**")
        st.write(f"👤 {preparado_por or 'Pendiente'}")
        st.caption(f"Fecha: {fecha_preparacion or 'Se registra al guardar'}")
        st.markdown("Firma: ____________________")
    with f2:
        st.markdown("**REVISÓ**")
        st.write(f"👤 {revisado_por or 'Pendiente de revisión'}")
        st.caption(f"Fecha: {revisado_fecha or 'Pendiente'}")
        st.markdown("Firma: ____________________")
    with f3:
        st.markdown("**APROBÓ**")
        st.write(f"👤 {aprobado_por or 'Pendiente de aprobación'}")
        st.caption(f"Fecha: {aprobado_fecha or 'Pendiente'}")
        st.markdown("Firma: ____________________")
    st.info(f"Estado actual: **{estado}**")

def guardar_conciliacion_historial(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial,
    diferencia_conciliada, resultado_final, salidas_extracto,
    salidas_libros, entradas_libros, entradas_extracto,
    gastos_bancarios, preparado_por, revisado_por, excel_data,
    workflow_status="Pendiente de revisión", id_edicion=None
):
    conn = conectar_db()
    c = conn.cursor()
    try:
        empresa_id = _buscar_empresa_id(c, empresa, nit)
        cuenta_id = _buscar_cuenta_id(c, empresa_id, banco, cuenta, tipo)
        datos = _json_datos_conciliacion(
            empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
            saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
            resultado_final, salidas_extracto, salidas_libros, entradas_libros,
            entradas_extracto, gastos_bancarios, preparado_por, revisado_por, workflow_status
        )
        ahora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        datos.update(construir_firmas_conciliacion(
            preparado_por=preparado_por,
            fecha_preparacion=datos.get('fecha_preparacion') or ahora,
            revisado_por=datos.get('revisado_por_usuario') or revisado_por,
            fecha_revision=datos.get('fecha_revision') or '',
            aprobado_por=datos.get('aprobado_por_usuario') or '',
            fecha_aprobacion=datos.get('fecha_aprobacion') or ''
        ))
        estado = datos["estado"]
        fecha_creacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        observaciones = json.dumps(datos, ensure_ascii=False)
        if id_edicion:
            c.execute("""
                UPDATE conciliaciones SET empresa_id=?, cuenta_id=?, periodo=?, saldo_libro=?, saldo_banco=?,
                    estado=?, preparado_por=?, revisado_por=?, fecha_creacion=?, dictamen=?, observaciones=?
                WHERE id=?
            """, (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), estado,
                     preparado_por or None, revisado_por or None, fecha_creacion, workflow_status, observaciones, int(id_edicion)))
            last_id = int(id_edicion)
        else:
            c.execute("""
                INSERT INTO conciliaciones (
                    empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, estado,
                    preparado_por, revisado_por, fecha_creacion, dictamen, observaciones
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), estado,
                     preparado_por or None, revisado_por or None, fecha_creacion, workflow_status, observaciones))
            last_id = int(c.lastrowid)
        conn.commit()
        st.cache_data.clear()
        return last_id
    finally:
        conn.close()

def _decodificar_observaciones(texto):
    if not texto: return {}
    try:
        data = json.loads(texto)
        return data if isinstance(data, dict) else {}
    except Exception: return {}

@st.cache_data(ttl=300)
def obtener_historial(empresa_nombre=None):
    conn = None
    try:
        conn = conectar_db()
        c = conn.cursor()
        query = """
            SELECT co.id, co.fecha_creacion, e.razon_social, e.nit, co.periodo,
                   cb.banco, cb.numero_cuenta, cb.tipo_cuenta, co.saldo_banco,
                   co.saldo_libro, co.estado, co.preparado_por, co.revisado_por,
                   co.dictamen, co.observaciones
            FROM conciliaciones co
            LEFT JOIN empresas e ON co.empresa_id = e.id
            LEFT JOIN cuentas_bancarias cb ON co.cuenta_id = cb.id
        """
        params = []
        if empresa_nombre and str(empresa_nombre).strip() not in ("Todas las empresas", "Sin asignar"):
            query += " WHERE (e.razon_social = ? OR json_extract(co.observaciones, '$.empresa') = ?)"
            params = [str(empresa_nombre).strip(), str(empresa_nombre).strip()]
        query += " ORDER BY co.id DESC"
        c.execute(query, tuple(params))
        rows = c.fetchall()

        salida = []
        for row in rows:
            d = _decodificar_observaciones(row[14])
            workflow = row[13] or d.get('workflow_status') or 'Pendiente de revisión'
            salida.append({
                'id': int(row[0]),
                'fecha_guardado': row[1] or '',
                'empresa': row[2] or d.get('empresa', ''),
                'nit': row[3] or d.get('nit', ''),
                'mes': row[4] or d.get('mes', ''),
                'banco': row[5] or d.get('banco', ''),
                'cuenta': row[6] or d.get('cuenta', ''),
                'tipo': row[7] or d.get('tipo', ''),
                'saldo_extracto': float(row[8] or d.get('saldo_extracto', 0) or 0),
                'saldo_libros': float(row[9] or d.get('saldo_libros', 0) or 0),
                'diferencia_inicial': float(d.get('diferencia_inicial', float(row[8] or 0) - float(row[9] or 0)) or 0),
                'diferencia_conciliada': float(d.get('diferencia_conciliada', 0) or 0),
                'resultado_final': float(d.get('resultado_final', 0) or 0),
                'estado': row[10] or d.get('estado', ''),
                'workflow_status': workflow,
                'preparado_por': row[11] or d.get('preparado_por', ''),
                'revisado_por': row[12] or d.get('revisado_por', ''),
                'revisado_por_usuario': row[12] or d.get('revisado_por_usuario', ''),
                'aprobado_por_usuario': d.get('aprobado_por_usuario', ''),
                'fecha_preparacion': d.get('fecha_preparacion', ''),
                'fecha_revision': d.get('fecha_revision', ''),
                'fecha_aprobacion': d.get('fecha_aprobacion', ''),
                'motivo_correccion': d.get('motivo_correccion'),
                'tipo_hallazgo': d.get('tipo_hallazgo'),
                'checklist_json': d.get('checklist_json'),
                'datos_json': json.dumps(d, ensure_ascii=False),
                'fecha_elaboracion': d.get('fecha_elaboracion', row[1])
            })
        return pd.DataFrame(salida)
    except Exception as e:
        st.error(f"❌ No fue posible cargar el Historial: {type(e).__name__}: {e}")
        return pd.DataFrame()
    finally:
        if conn is not None:
            try: conn.close()
            except Exception: pass

@st.cache_data(ttl=300)
def obtener_conciliacion_por_id(id_conciliacion):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("""
            SELECT co.id, co.empresa_id, co.cuenta_id, co.periodo, co.saldo_libro, co.saldo_banco,
                   co.estado, co.preparado_por, co.revisado_por, co.fecha_creacion, co.dictamen,
                   co.observaciones, e.razon_social, e.nit, cb.banco, cb.numero_cuenta, cb.tipo_cuenta
            FROM conciliaciones co
            LEFT JOIN empresas e ON co.empresa_id=e.id
            LEFT JOIN cuentas_bancarias cb ON co.cuenta_id=cb.id
            WHERE co.id=?
        """, (int(id_conciliacion),))
        row = c.fetchone()
        conn.close()
        if not row:
            return None
        datos = _decodificar_observaciones(row[11])
        empresa, nit = row[12] or datos.get('empresa',''), row[13] or datos.get('nit','')
        banco, cuenta = row[14] or datos.get('banco',''), row[15] or datos.get('cuenta','')
        tipo = row[16] or datos.get('tipo','')
        saldo_extracto = float(row[5] or datos.get('saldo_extracto',0))
        saldo_libros = float(row[4] or datos.get('saldo_libros',0))
        datos.setdefault('empresa', empresa); datos.setdefault('nit', nit); datos.setdefault('mes', row[3])
        datos.setdefault('banco', banco); datos.setdefault('cuenta', cuenta); datos.setdefault('tipo', tipo)
        datos.setdefault('saldo_extracto', saldo_extracto); datos.setdefault('saldo_libros', saldo_libros)
        datos.setdefault('diferencia_inicial', saldo_extracto-saldo_libros); datos.setdefault('estado', row[6])
        datos.setdefault('preparado_por', row[7]); datos.setdefault('revisado_por', row[8])
        datos.setdefault('workflow_status', row[10] or 'Pendiente de revisión')
        return {
            'id': row[0], 'empresa': empresa, 'nit': nit, 'mes': datos.get('mes',row[3]),
            'fecha_elaboracion': datos.get('fecha_elaboracion',row[9]), 'banco': banco, 'cuenta': cuenta, 'tipo': tipo,
            'saldo_extracto': saldo_extracto, 'saldo_libros': saldo_libros,
            'diferencia_inicial': datos.get('diferencia_inicial',saldo_extracto-saldo_libros),
            'diferencia_conciliada': datos.get('diferencia_conciliada',0), 'resultado_final': datos.get('resultado_final',0),
            'estado': row[6], 'datos_json': json.dumps(datos,ensure_ascii=False),
            'workflow_status': row[10] or datos.get('workflow_status','Pendiente de revisión'),
            'revisado_por_usuario': row[8] or datos.get('revisado_por_usuario',''),
            'fecha_revision': datos.get('fecha_revision'), 'motivo_correccion': datos.get('motivo_correccion'),
            'tipo_hallazgo': datos.get('tipo_hallazgo'), 'checklist_json': datos.get('checklist_json')
        }
    except Exception:
        return None

def actualizar_estado_auditoria(id_conciliacion, nuevo_estado, revisado_por, motivo_correccion=None, tipo_hallazgo=None, checklist=None):
    fecha_rev = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = conectar_db(); c = conn.cursor()
    c.execute("SELECT observaciones FROM conciliaciones WHERE id=?", (int(id_conciliacion),))
    row = c.fetchone(); datos = _decodificar_observaciones(row[0] if row else None)
    datos.update({'workflow_status': nuevo_estado, 'revisado_por_usuario': revisado_por, 'fecha_revision': fecha_rev,
                  'motivo_correccion': motivo_correccion, 'tipo_hallazgo': tipo_hallazgo, 'checklist_json': checklist})
    if nuevo_estado == 'Aprobada':
        datos['aprobado_por_usuario'] = revisado_por
        datos['fecha_aprobacion'] = fecha_rev
    else:
        datos.setdefault('aprobado_por_usuario', '')
        datos.setdefault('fecha_aprobacion', '')
    c.execute("UPDATE conciliaciones SET revisado_por=?, dictamen=?, observaciones=? WHERE id=?",
              (revisado_por, nuevo_estado, json.dumps(datos,ensure_ascii=False), int(id_conciliacion)))
    conn.commit(); conn.close()
    st.cache_data.clear()

def guardar_conciliacion_especial(empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                                  saldo_libros, saldo_extracto, resultado_final,
                                  datos_extra, preparado_por, revisado_por,
                                  workflow_status="Pendiente de revisión", id_edicion=None):
    conn = conectar_db(); c = conn.cursor()
    try:
        empresa_id = _buscar_empresa_id(c, empresa, nit)
        cuenta_id = _buscar_cuenta_id(c, empresa_id, banco, cuenta, tipo)
        payload = dict(datos_extra or {})
        ahora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        payload.update({
            "empresa": empresa, "nit": nit, "mes": mes,
            "fecha_elaboracion": fecha_elaboracion.isoformat() if hasattr(fecha_elaboracion, "isoformat") else str(fecha_elaboracion),
            "banco": banco, "cuenta": cuenta, "tipo": tipo,
            "saldo_libros": float(saldo_libros), "saldo_extracto": float(saldo_extracto),
            "diferencia_inicial": float(saldo_extracto) - float(saldo_libros),
            "resultado_final": float(resultado_final),
            "estado": "CONCILIACIÓN CORRECTA" if abs(float(resultado_final)) < 0.005 else "CONCILIACIÓN CON DIFERENCIA",
            "preparado_por": preparado_por, "revisado_por": revisado_por,
            "workflow_status": workflow_status,
            **construir_firmas_conciliacion(preparado_por=preparado_por, fecha_preparacion=payload.get("fecha_preparacion") or ahora, revisado_por=payload.get("revisado_por_usuario") or revisado_por, fecha_revision=payload.get("fecha_revision") or "", aprobado_por=payload.get("aprobado_por_usuario") or "", fecha_aprobacion=payload.get("fecha_aprobacion") or "")
        })
        observaciones = json.dumps(payload, ensure_ascii=False)
        fecha_creacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not id_edicion:
            c.execute("""SELECT id FROM conciliaciones
                         WHERE empresa_id=? AND cuenta_id=?
                           AND LOWER(TRIM(periodo))=LOWER(TRIM(?))
                         ORDER BY id DESC LIMIT 1""", (int(empresa_id), int(cuenta_id), str(mes)))
            existente = c.fetchone()
            if existente:
                id_edicion = int(existente[0])

        if id_edicion:
            c.execute("""UPDATE conciliaciones SET empresa_id=?, cuenta_id=?, periodo=?, saldo_libro=?, saldo_banco=?,
                         estado=?, preparado_por=?, revisado_por=?, fecha_creacion=?, dictamen=?, observaciones=? WHERE id=?""",
                      (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), payload["estado"],
                       preparado_por or None, revisado_por or None, fecha_creacion, workflow_status, observaciones, int(id_edicion)))
            last_id = int(id_edicion)
        else:
            try:
                c.execute("""INSERT INTO conciliaciones (empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, estado,
                             preparado_por, revisado_por, fecha_creacion, dictamen, observaciones)
                             VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                          (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), payload["estado"],
                           preparado_por or None, revisado_por or None, fecha_creacion, workflow_status, observaciones))
                last_id = int(c.lastrowid)
            except Exception as exc:
                c.execute("""SELECT id FROM conciliaciones
                             WHERE empresa_id=? AND cuenta_id=?
                               AND LOWER(TRIM(periodo))=LOWER(TRIM(?))
                             ORDER BY id DESC LIMIT 1""", (int(empresa_id), int(cuenta_id), str(mes)))
                existente = c.fetchone()
                if existente:
                    last_id = int(existente[0])
                    c.execute("""UPDATE conciliaciones SET empresa_id=?, cuenta_id=?, periodo=?, saldo_libro=?, saldo_banco=?,
                                 estado=?, preparado_por=?, revisado_por=?, fecha_creacion=?, dictamen=?, observaciones=? WHERE id=?""",
                              (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), payload["estado"],
                               preparado_por or None, revisado_por or None, fecha_creacion, workflow_status, observaciones, last_id))
                else:
                    raise exc
        conn.commit()
        st.cache_data.clear()
        return last_id
    except Exception:
        try: conn.rollback()
        except Exception: pass
        raise
    finally:
        conn.close()

def _filas_caja_default():
    return pd.DataFrame([{"N°":i,"Fecha":"","Proveedor":"","Concepto / Compra":"","Factura / Soporte":"",
                          "Forma de pago":"","Valor compra":0.0,"IVA / otros":0.0,"Total pagado":0.0,"Observaciones":""} for i in range(1,21)])

def preparar_excel_caja(empresa, mes, fecha, responsable, saldo_inicial, fondo, compras_df, efectivo_df, observaciones, logo_bytes=None, preparado_por="", revisado_por="", aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion="", estado="Pendiente de revisión"):
    wb=Workbook(); ws=wb.active; ws.title="Arqueo de Caja"
    ws.merge_cells("A1:J1"); ws["A1"]="ARQUEO DE CAJA GENERAL – COMPRAS"; ws["A1"].font=Font(bold=True,size=14); ws["A1"].alignment=Alignment(horizontal="center")
    ws["A3"]="Empresa:"; ws["B3"]=empresa; ws["D3"]="Responsable:"; ws["E3"]=responsable; ws["G3"]="Período:"; ws["H3"]=mes
    ws["A4"]="Fecha:"; ws["B4"]=str(fecha); ws["D4"]="Fondo autorizado:"; ws["E4"]=float(fondo); ws["G4"]="Saldo inicial para compras:"; ws["H4"]=float(saldo_inicial)
    if logo_bytes:
        try:
            img=XLImage(io.BytesIO(logo_bytes)); img.width=110; img.height=60; img.anchor="J2"; ws.add_image(img)
        except Exception: pass
    headers=["N°","Fecha","Proveedor","Concepto / Compra","Factura / Soporte","Forma de pago","Valor compra","IVA / otros","Total pagado","Observaciones"]
    for col,h in enumerate(headers,1):
        cell=ws.cell(6,col,h); cell.font=Font(bold=True); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for r_idx,row in enumerate(compras_df.to_dict("records"),7):
        for col,h in enumerate(headers,1): ws.cell(r_idx,col,row.get(h,""))
    fin=6+len(compras_df)
    ws.cell(fin+1,6,"TOTAL COMPRAS:"); ws.cell(fin+1,9,float(total_columna(compras_df,"Total pagado")))
    ws.cell(fin+3,2,"DETALLE DEL SALDO EN CAJA");
    ws.cell(fin+4,2,"Denominación"); ws.cell(fin+4,3,"Cantidad"); ws.cell(fin+4,4,"Monto")
    rr=fin+5
    for row in efectivo_df.to_dict("records"):
        ws.cell(rr,2,row.get("Denominación","")); ws.cell(rr,3,row.get("Cantidad",0)); ws.cell(rr,4,row.get("Monto",0)); rr+=1
    total_ef=total_columna(efectivo_df,"Monto")
    total_compras=total_columna(compras_df,"Total pagado")
    teorico=float(saldo_inicial)-total_compras
    ws.cell(rr+1,8,"Saldo teórico final:"); ws.cell(rr+1,10,teorico)
    ws.cell(rr+2,8,"Efectivo físico final:"); ws.cell(rr+2,10,total_ef)
    ws.cell(rr+3,8,"Diferencia (sobrante/faltante):"); ws.cell(rr+3,10,total_ef-teorico)
    ws.cell(rr+4,8,"Resultado del arqueo:"); ws.cell(rr+4,10,"CUADRA" if abs(total_ef-teorico)<0.005 else ("SOBRANTE" if total_ef-teorico>0 else "FALTANTE"))
    ws.cell(rr+6,2,"Observaciones:"); ws.cell(rr+7,2,observaciones or "")
    for col,w in enumerate([7,13,22,28,24,16,16,14,16,34],1): ws.column_dimensions[chr(64+col)].width=w
    sig_row = rr + 9
    ws.merge_cells(start_row=sig_row, start_column=1, end_row=sig_row, end_column=10)
    ws.cell(sig_row,1).value = "FIRMAS Y TRAZABILIDAD"; ws.cell(sig_row,1).font = Font(bold=True, size=12)
    sig_headers=["PREPARÓ","REVISÓ","APROBÓ"]
    sig_values=[f"{preparado_por or 'Pendiente'}\nFecha: {fecha_preparacion or 'Pendiente'}\nFirma: ____________________", f"{revisado_por or 'Pendiente de revisión'}\nFecha: {fecha_revision or 'Pendiente'}\nFirma: ____________________", f"{aprobado_por or 'Pendiente de aprobación'}\nFecha: {fecha_aprobacion or 'Pendiente'}\nFirma: ____________________"]
    for i in range(3):
        c1=1+i*3; ws.merge_cells(start_row=sig_row+1,start_column=c1,end_row=sig_row+1,end_column=c1+2); ws.cell(sig_row+1,c1).value=sig_headers[i]; ws.cell(sig_row+1,c1).font=Font(bold=True)
        ws.merge_cells(start_row=sig_row+2,start_column=c1,end_row=sig_row+4,end_column=c1+2); ws.cell(sig_row+2,c1).value=sig_values[i]; ws.cell(sig_row+2,c1).alignment=Alignment(wrap_text=True,vertical="top")
    ws.merge_cells(start_row=sig_row+6,start_column=1,end_row=sig_row+6,end_column=10); ws.cell(sig_row+6,1).value=f"Estado: {estado}"
    for row in ws.iter_rows():
        for cell in row: cell.alignment=Alignment(vertical="top",wrap_text=True)
    out=io.BytesIO(); wb.save(out); return out.getvalue(), f"ARQUEO_CAJA_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.xlsx"

def generar_pdf_caja(empresa, mes, fecha, responsable, saldo_inicial, fondo, compras_df, efectivo_df, observaciones, logo_bytes=None, preparado_por="", revisado_por="", aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion="", estado="Pendiente de revisión"):
    out=io.BytesIO(); doc=SimpleDocTemplate(out,pagesize=portrait(letter),rightMargin=25,leftMargin=25,topMargin=25,bottomMargin=25); styles=getSampleStyleSheet(); story=[]
    if logo_bytes:
        try: story.append(Image(io.BytesIO(logo_bytes),width=75,height=45)); story.append(Spacer(1,4))
        except Exception: pass
    story += [Paragraph("<b>ARQUEO DE CAJA GENERAL – COMPRAS</b>",styles["Title"]), Paragraph(f"Empresa: {empresa} &nbsp;&nbsp; Responsable: {responsable}",styles["Normal"]), Paragraph(f"Fecha: {fecha} &nbsp;&nbsp; Período: {mes}",styles["Normal"]), Paragraph(f"Saldo inicial: {formatear_moneda(saldo_inicial)} &nbsp;&nbsp; Fondo autorizado: {formatear_moneda(fondo)}",styles["Normal"]), Spacer(1,8)]
    headers=["N°","Fecha","Proveedor","Concepto","Soporte","Pago","Compra","IVA","Total","Observaciones"]
    data=[headers]
    for _,r in compras_df.iterrows(): data.append([r.get("N°",""),r.get("Fecha",""),r.get("Proveedor",""),r.get("Concepto / Compra",""),r.get("Factura / Soporte",""),r.get("Forma de pago",""),formatear_moneda(r.get("Valor compra",0)),formatear_moneda(r.get("IVA / otros",0)),formatear_moneda(r.get("Total pagado",0)),r.get("Observaciones","")])
    t=Table(data,repeatRows=1,colWidths=[20,45,60,70,55,40,48,45,50,80]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"TOP")])); story += [t,Spacer(1,8)]
    total_compras=total_columna(compras_df,"Total pagado"); teorico=float(saldo_inicial)-total_compras; fisico=total_columna(efectivo_df,"Monto"); diff=fisico-teorico
    story += [Paragraph(f"<b>Total compras:</b> {formatear_moneda(total_compras)}",styles["Normal"]),Paragraph(f"<b>Saldo teórico final:</b> {formatear_moneda(teorico)}",styles["Normal"]),Paragraph(f"<b>Efectivo físico final:</b> {formatear_moneda(fisico)}",styles["Normal"]),Paragraph(f"<b>Diferencia:</b> {formatear_moneda(diff)} &nbsp;&nbsp; <b>{'CUADRA' if abs(diff)<0.005 else ('SOBRANTE' if diff>0 else 'FALTANTE')}</b>",styles["Normal"]),Spacer(1,8),Paragraph(f"<b>Observaciones:</b> {observaciones or ''}",styles["Normal"])]
    story += [Spacer(1,12), Paragraph("<b>FIRMAS Y TRAZABILIDAD</b>", styles["Heading3"])]
    sig_data=[[Paragraph("<b>PREPARÓ</b>",styles["Normal"]),Paragraph("<b>REVISÓ</b>",styles["Normal"]),Paragraph("<b>APROBÓ</b>",styles["Normal"])], [Paragraph(f"{preparado_por or 'Pendiente'}<br/>Fecha: {fecha_preparacion or 'Pendiente'}<br/><br/>Firma: ____________________",styles["Normal"]), Paragraph(f"{revisado_por or 'Pendiente de revisión'}<br/>Fecha: {fecha_revision or 'Pendiente'}<br/><br/>Firma: ____________________",styles["Normal"]), Paragraph(f"{aprobado_por or 'Pendiente de aprobación'}<br/>Fecha: {fecha_aprobacion or 'Pendiente'}<br/><br/>Firma: ____________________",styles["Normal"] )]]
    tf=Table(sig_data,colWidths=[170,170,170]); tf.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP")])); story += [tf, Paragraph(f"Estado: <b>{estado}</b>",styles["Normal"])]
    doc.build(story); return out.getvalue(), f"ARQUEO_CAJA_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.pdf"

def preparar_excel_credito_bancario(empresa, mes, fecha, entidad, numero, fecha_inicio, fecha_vencimiento, tasa, saldo_libros, saldo_extracto, diferencias_df, observaciones, logo_bytes=None, preparado_por="", revisado_por="", aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion="", estado="Pendiente de revisión"):
    wb=Workbook(); ws=wb.active; ws.title="Crédito Bancario"; ws.merge_cells("A1:F1"); ws["A1"]="CONCILIACIÓN DE CRÉDITO BANCARIO"; ws["A1"].font=Font(bold=True,size=14); ws["A1"].alignment=Alignment(horizontal="center")
    ws["A3"]="Empresa:"; ws["B3"]=empresa; ws["D3"]="Período:"; ws["E3"]=mes; ws["A4"]="Entidad financiera:"; ws["B4"]=entidad; ws["D4"]="Número de crédito:"; ws["E4"]=numero; ws["A5"]="Fecha inicio:"; ws["B5"]=str(fecha_inicio); ws["D5"]="Vencimiento:"; ws["E5"]=str(fecha_vencimiento); ws["A6"]="Tasa:"; ws["B6"]=tasa
    ws["A8"]="Saldo según libros:"; ws["B8"]=float(saldo_libros); ws["D8"]="Saldo según extracto:"; ws["E8"]=float(saldo_extracto); ws["A9"]="Diferencia:"; ws["B9"]=float(saldo_extracto)-float(saldo_libros)
    headers=["Concepto","Valor","Observación"]
    for c,h in enumerate(headers,1): ws.cell(11,c,h).font=Font(bold=True)
    for r_idx,row in enumerate(diferencias_df.to_dict("records"),12): ws.cell(r_idx,1,row.get("Concepto","")); ws.cell(r_idx,2,row.get("Valor",0)); ws.cell(r_idx,3,row.get("Observación",""))
    total_dif=total_columna(diferencias_df,"Valor"); rr=12+len(diferencias_df); ws.cell(rr+1,1,"Total diferencias detalladas:"); ws.cell(rr+1,2,total_dif); ws.cell(rr+3,1,"Resultado:"); ws.cell(rr+3,2,"CONCILIADO" if abs(float(saldo_extracto)-float(saldo_libros))<0.005 else "NO CONCILIADO"); ws.cell(rr+5,1,"Observaciones:"); ws.cell(rr+6,1,observaciones or "")
    for col,w in zip("ABCDEF",[28,18,40,24,22,18]): ws.column_dimensions[col].width=w
    if logo_bytes:
        try: img=XLImage(io.BytesIO(logo_bytes)); img.width=110; img.height=60; img.anchor="F2"; ws.add_image(img)
        except Exception: pass
    sig_row=rr+8; ws.merge_cells(start_row=sig_row,start_column=1,end_row=sig_row,end_column=6); ws.cell(sig_row,1).value="FIRMAS Y TRAZABILIDAD"; ws.cell(sig_row,1).font=Font(bold=True,size=12)
    vals=[f"PREPARÓ\n{preparado_por or 'Pendiente'}\nFecha: {fecha_preparacion or 'Pendiente'}\nFirma: ____________________",f"REVISÓ\n{revisado_por or 'Pendiente de revisión'}\nFecha: {fecha_revision or 'Pendiente'}\nFirma: ____________________",f"APROBÓ\n{aprobado_por or 'Pendiente de aprobación'}\nFecha: {fecha_aprobacion or 'Pendiente'}\nFirma: ____________________"]
    for i,v in enumerate(vals):
        col=1+i*2; ws.merge_cells(start_row=sig_row+1,start_column=col,end_row=sig_row+3,end_column=col+1); ws.cell(sig_row+1,col).value=v; ws.cell(sig_row+1,col).alignment=Alignment(wrap_text=True,vertical="top")
    ws.merge_cells(start_row=sig_row+5,start_column=1,end_row=sig_row+5,end_column=6); ws.cell(sig_row+5,1).value=f"Estado: {estado}"
    out=io.BytesIO(); wb.save(out); return out.getvalue(), f"CREDITO_BANCARIO_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.xlsx"

def generar_pdf_credito_bancario(empresa, mes, fecha, entidad, numero, fecha_inicio, fecha_vencimiento, tasa, saldo_libros, saldo_extracto, diferencias_df, observaciones, logo_bytes=None, preparado_por="", revisado_por="", aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion="", estado="Pendiente de revisión"):
    out=io.BytesIO(); doc=SimpleDocTemplate(out,pagesize=portrait(letter),rightMargin=35,leftMargin=35,topMargin=30,bottomMargin=30); styles=getSampleStyleSheet(); story=[]
    if logo_bytes:
        try: story.append(Image(io.BytesIO(logo_bytes),width=75,height=45)); story.append(Spacer(1,4))
        except Exception: pass
    story += [Paragraph("<b>CONCILIACIÓN DE CRÉDITO BANCARIO</b>",styles["Title"]),Paragraph(f"Empresa: {empresa} &nbsp;&nbsp; Período: {mes}",styles["Normal"]),Paragraph(f"Entidad: {entidad} &nbsp;&nbsp; Crédito: {numero}",styles["Normal"]),Paragraph(f"Inicio: {fecha_inicio} &nbsp;&nbsp; Vencimiento: {fecha_vencimiento} &nbsp;&nbsp; Tasa: {tasa}",styles["Normal"]),Spacer(1,8)]
    diff=float(saldo_extracto)-float(saldo_libros); data=[["Concepto","Valor"],["Saldo según libros",formatear_moneda(saldo_libros)],["Saldo según extracto",formatear_moneda(saldo_extracto)],["Diferencia",formatear_moneda(diff)],["Resultado","CONCILIADO" if abs(diff)<0.005 else "NO CONCILIADO"]]
    t=Table(data,colWidths=[220,180]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),0.5,colors.grey),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold")])); story += [t,Spacer(1,10),Paragraph("<b>Diferencias / partidas identificadas</b>",styles["Heading3"])]
    det=[["Concepto","Valor","Observación"]]+[[r.get("Concepto",""),formatear_moneda(r.get("Valor",0)),r.get("Observación","")] for _,r in diferencias_df.iterrows()]
    t2=Table(det,colWidths=[170,90,190],repeatRows=1); t2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),8)])); story += [t2,Spacer(1,10),Paragraph(f"<b>Observaciones:</b> {observaciones or ''}",styles["Normal"])]
    story += [Spacer(1,12),Paragraph("<b>FIRMAS Y TRAZABILIDAD</b>",styles["Heading3"])]
    sig_data=[[Paragraph("<b>PREPARÓ</b>",styles["Normal"]),Paragraph("<b>REVISÓ</b>",styles["Normal"]),Paragraph("<b>APROBÓ</b>",styles["Normal"])],[Paragraph(f"{preparado_por or 'Pendiente'}<br/>Fecha: {fecha_preparacion or 'Pendiente'}<br/><br/>Firma: ____________________",styles["Normal"]),Paragraph(f"{revisado_por or 'Pendiente de revisión'}<br/>Fecha: {fecha_revision or 'Pendiente'}<br/><br/>Firma: ____________________",styles["Normal"]),Paragraph(f"{aprobado_por or 'Pendiente de aprobación'}<br/>Fecha: {fecha_aprobacion or 'Pendiente'}<br/><br/>Firma: ____________________",styles["Normal"])]]
    tf=Table(sig_data,colWidths=[170,170,170]); tf.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP")])); story += [tf,Paragraph(f"Estado: <b>{estado}</b>",styles["Normal"])]
    doc.build(story); return out.getvalue(), f"CREDITO_BANCARIO_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.pdf"

# ==========================================
# 5. AUTENTICACIÓN FLEXIBLE Y SEGURA
# ==========================================
def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000).hex()
    return salt, digest

def verificar_password(password, salt, password_hash):
    _, digest = hash_password(password, salt)
    return hmac.compare_digest(digest, password_hash)

def contar_usuarios():
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM usuarios")
        res = c.fetchone()
        conn.close()
        return res[0] if res else 0
    except Exception:
        return 0

def crear_usuario(usuario, nombre, password, rol, empresa_id=None):
    salt, password_hash = hash_password(str(password).strip())
    conn = conectar_db()
    c = conn.cursor()
    c.execute('''
        INSERT INTO usuarios (usuario, nombre, password_hash, salt, rol, activo, fecha_creacion, empresa_id)
        VALUES (?, ?, ?, ?, ?, 1, ?, ?)
    ''', (str(usuario).strip().lower(), nombre.strip(), password_hash, salt, rol, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), empresa_id))
    conn.commit()
    last_id = c.lastrowid
    conn.close()
    st.cache_data.clear()
    return last_id

def actualizar_empresa_usuario(usuario_id, empresa_id):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("UPDATE usuarios SET empresa_id=? WHERE id=?", (empresa_id, int(usuario_id)))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def actualizar_rol_usuario(usuario_id, nuevo_rol):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("UPDATE usuarios SET rol=? WHERE id=?", (nuevo_rol, int(usuario_id)))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def autenticar_usuario(usuario, password):
    try:
        conn = conectar_db(); c = conn.cursor()
        query = """
        SELECT u.id, u.usuario, u.nombre, u.password_hash, u.salt, u.rol, u.empresa_id,
               e.razon_social AS empresa_nombre, e.nit AS empresa_nit
        FROM usuarios u
        LEFT JOIN empresas e ON u.empresa_id = e.id
        WHERE LOWER(u.usuario) = LOWER(?) AND u.activo = 1
        """
        c.execute(query, (str(usuario).strip(),)); fila = c.fetchone(); conn.close()
        if not fila: return None
        if fila[3] and fila[4] and verificar_password(str(password).strip(), fila[4], fila[3]):
            return {"id":fila[0],"usuario":fila[1],"nombre":fila[2],"rol":fila[5],"empresa_id":fila[6],"empresa_nombre":fila[7],"empresa_nit":fila[8]}
    except Exception as e:
        st.error("Error técnico durante la autenticación: " + f"{type(e).__name__}: {e}")
    return None

def cambiar_contrasena_usuario(usuario_id, contrasena_actual, nueva_contrasena):
    if not nueva_contrasena or len(str(nueva_contrasena)) < 8:
        return False, "La nueva contraseña debe tener al menos 8 caracteres."
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("SELECT salt, password_hash FROM usuarios WHERE id=? AND activo=1", (int(usuario_id),))
        fila = c.fetchone()
        if not fila:
            return False, "El usuario no existe o está inactivo."
        if not verificar_password(str(contrasena_actual or "").strip(), fila[0], fila[1]):
            return False, "La contraseña actual es incorrecta."
        nuevo_salt, nuevo_hash = hash_password(str(nueva_contrasena).strip())
        c.execute("UPDATE usuarios SET password_hash=?, salt=? WHERE id=?", (nuevo_hash, nuevo_salt, int(usuario_id)))
        conn.commit()
        return True, "Contraseña cambiada correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible cambiar la contraseña: {type(e).__name__}: {e}"
    finally:
        conn.close()

def actualizar_usuario_db(usuario_id, usuario, nombre, rol, activo, empresa_id=None, nueva_contrasena=None):
    usuario = str(usuario or "").strip().lower()
    nombre = str(nombre or "").strip()
    rol = str(rol or "").strip()
    if not usuario or not nombre or not rol:
        return False, "Usuario, nombre y rol son obligatorios."
    if nueva_contrasena is not None and str(nueva_contrasena).strip() and len(str(nueva_contrasena).strip()) < 8:
        return False, "La nueva contraseña debe tener al menos 8 caracteres."

    conn = conectar_db(); c = conn.cursor()
    try:
        c.execute("SELECT usuario FROM usuarios WHERE id=?", (int(usuario_id),))
        if not c.fetchone():
            return False, "No se encontró el usuario que deseas actualizar."

        c.execute("SELECT id FROM usuarios WHERE LOWER(TRIM(usuario))=LOWER(TRIM(?))", (usuario,))
        duplicado = c.fetchone()
        if duplicado and int(duplicado[0]) != int(usuario_id):
            return False, f"El nombre de usuario '{usuario}' ya pertenece a otro usuario."

        if nueva_contrasena is not None and str(nueva_contrasena).strip():
            nuevo_salt, nuevo_hash = hash_password(str(nueva_contrasena).strip())
            c.execute("""UPDATE usuarios SET usuario=?, nombre=?, password_hash=?, salt=?, rol=?, activo=?, empresa_id=? WHERE id=?""",
                      (usuario, nombre, nuevo_hash, nuevo_salt, rol, 1 if activo else 0, empresa_id, int(usuario_id)))
        else:
            c.execute("""UPDATE usuarios SET usuario=?, nombre=?, rol=?, activo=?, empresa_id=? WHERE id=?""",
                      (usuario, nombre, rol, 1 if activo else 0, empresa_id, int(usuario_id)))
        conn.commit()
        st.cache_data.clear()
        return True, "Usuario actualizado correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible actualizar el usuario: {type(e).__name__}: {e}"
    finally:
        conn.close()

def eliminar_usuario_db(usuario_id, usuario_actual_id=None):
    if usuario_actual_id is not None and int(usuario_id) == int(usuario_actual_id):
        return False, "No puedes eliminar el usuario con el que estás conectado."
    conn = conectar_db(); c = conn.cursor()
    try:
        c.execute("SELECT usuario, rol, activo FROM usuarios WHERE id=?", (int(usuario_id),))
        fila = c.fetchone()
        if not fila:
            return False, "El usuario no existe."
        if fila[1] == "Administrador" and int(fila[2] or 0) == 1:
            c.execute("SELECT COUNT(*) FROM usuarios WHERE rol='Administrador' AND activo=1")
            if int(c.fetchone()[0] or 0) <= 1:
                return False, "No se puede eliminar al único Administrador activo."
        c.execute("DELETE FROM usuarios WHERE id=?", (int(usuario_id),))
        conn.commit()
        st.cache_data.clear()
        return True, f"Usuario '{fila[0]}' eliminado correctamente."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible eliminar el usuario: {type(e).__name__}: {e}"
    finally:
        conn.close()

@st.cache_data(ttl=300)
def obtener_usuarios():
    try:
        conn = conectar_db(); c = conn.cursor()
        c.execute("""SELECT u.id, u.usuario, u.nombre, u.rol, u.activo, u.fecha_creacion, e.razon_social AS empresa_nombre
                     FROM usuarios u LEFT JOIN empresas e ON u.empresa_id=e.id ORDER BY u.id""")
        rows = c.fetchall(); conn.close()
        if rows: return pd.DataFrame(rows, columns=['id','usuario','nombre','rol','activo','fecha_creacion','empresa_nombre'])
    except Exception: pass
    return pd.DataFrame(columns=['id','usuario','nombre','rol','activo','fecha_creacion','empresa_nombre'])

def obtener_opciones_menu(rol):
    if rol == "Preparador":
        return ["📊 Dashboard", "📝 Nueva Conciliación", "📋 Historial"]
    elif rol == "Revisor":
        return ["🔍 Auditoría y Revisiones", "📊 Dashboard", "📋 Historial", "🏢 Empresas", "🏦 Bancos y Cuentas", "📄 Reportes"]
    elif rol == "Administrador":
        return ["📊 Dashboard", "🔍 Auditoría y Revisiones", "📝 Nueva Conciliación", "📋 Historial", "🏢 Empresas", "🏦 Bancos y Cuentas", "👥 Usuarios", "⚙️ Administración", "📄 Reportes"]
    return ["📊 Dashboard"]

# ==========================================
# 6. FLUJO DE LOGIN Y AUTENTICACIÓN
# ==========================================
def iniciar_autenticacion():
    if "usuario_autenticado" not in st.session_state:
        st.session_state.usuario_autenticado = None

    params = st.query_params
    if "registro" in params and params["registro"] == "true":
        st.title("👤 Registro de Nuevo Usuario")
        st.caption("Completa tus datos para crear tu cuenta de acceso.")
        rol_invitado = params.get("rol", "Preparador")
        empresa_id_invitada = params.get("empresa_id", None)
        if empresa_id_invitada:
            try: empresa_id_invitada = int(empresa_id_invitada)
            except ValueError: empresa_id_invitada = None

        with st.form("form_autoregistro"):
            usuario = st.text_input("Nombre de usuario")
            nombre = st.text_input("Nombre completo")
            password = st.text_input("Contraseña", type="password")
            confirmar = st.text_input("Confirmar contraseña", type="password")
            registro_btn = st.form_submit_button("Crear mi cuenta", type="primary")

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
                    except Exception:
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
            crear = st.form_submit_button("Crear administrador", type="primary")

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
        st.title("⚖️ Inicio de sesión")
        st.caption("Ingresa tus credenciales para acceder a Conciliación Bancaria.")
        with st.form("form_login"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Iniciar sesión", type="primary")

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

# ==========================================
# 7. BARRA LATERAL (SIDEBAR)
# ==========================================
empresas_df = obtener_empresas()
with st.sidebar:
    logo_sidebar = st.empty()
    st.title("Conciliación Web")
    st.markdown(f"👤 **{usuario_actual['nombre']}**")
    st.caption(f"Rol: **{rol_actual}**")
    st.divider()

    if rol_actual in ["Administrador", "Revisor"]:
        st.subheader("🏢 Selección de Empresa")
        opciones_emp = ["Todas las empresas"] + empresas_df["nombre"].tolist() if not empresas_df.empty else ["Todas las empresas"]
        empresa_activa_nombre = st.selectbox("Empresa a Auditar / Revisar", opciones_emp)
        if empresa_activa_nombre != "Todas las empresas" and not empresas_df.empty:
            fila_empresa = empresas_df[empresas_df["nombre"] == empresa_activa_nombre].iloc[0]
            empresa_activa_id = int(fila_empresa["id"])
            empresa_activa_nit = fila_empresa["nit"]
        else:
            empresa_activa_id = None
            empresa_activa_nit = ""
    else:
        empresa_activa_nombre = usuario_actual.get("empresa_nombre") or "Sin asignar"
        empresa_activa_nit = usuario_actual.get("empresa_nit") or ""
        empresa_activa_id = usuario_actual.get("empresa_id")
        st.info(f"🏢 **Empresa:** {empresa_activa_nombre}")

    logo_empresa_sidebar = None
    if empresa_activa_nombre and empresa_activa_nombre not in ["Todas las empresas", "Sin asignar"]:
        logo_empresa_sidebar = obtener_logo_empresa(empresa_activa_nombre)

    if logo_empresa_sidebar:
        logo_sidebar.image(logo_empresa_sidebar, width=90)
    else:
        logo_sidebar.image("https://img.icons8.com/color/96/bank-building.png", width=70)

    st.divider()
    opciones_menu = obtener_opciones_menu(rol_actual)
    if "menu_override" in st.session_state:
        menu_seleccionado = st.session_state.pop("menu_override")
    else:
        menu_seleccionado = st.radio("Navegación principal", opciones_menu)

    st.divider()
    if st.button("🚪 Cerrar sesión"):
        st.session_state.usuario_autenticado = None
        st.rerun()

# ==========================================
# 8. EXPORTADORES A EXCEL Y PDF
# ==========================================
def formatear_moneda(valor):
    try: numero = float(valor or 0)
    except (TypeError, ValueError): numero = 0.0
    texto = f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {texto}"

def formatear_columna_valor(df, columna="Valor"):
    if df is None or df.empty or columna not in df.columns:
        return df
    resultado = df.copy()
    resultado[columna] = pd.to_numeric(resultado[columna], errors="coerce").fillna(0.0)
    return resultado

def config_monetaria(columnas):
    config = {}
    for columna in columnas:
        config[columna] = st.column_config.NumberColumn(columna, format="$ %,.2f", step=0.01)
    return config

def total_columna(df, columna="Valor"):
    if df is None or df.empty or columna not in df.columns:
        return 0.0
    return float(pd.to_numeric(df[columna], errors="coerce").fillna(0).sum())

def limpiar_dataframe(df):
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()
    resultado = df.copy()
    for columna in resultado.columns:
        if resultado[columna].dtype == "object":
            resultado[columna] = resultado[columna].replace(r"^\s*\$", "", regex=True)
    return resultado.dropna(how="all").reset_index(drop=True)

def limpiar_nombre_archivo(texto):
    texto = str(texto).strip()
    return re.sub(r'[<>:"/\\|?*]', "-", texto) if texto else "Empresa"

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
    numero_columnas = max(4, len(df.columns)) if not df.empty else 4
    fila = escribir_seccion(ws, fila, titulo, numero_columnas)
    
    if not df.empty:
        for columna, nombre in enumerate(df.columns, start=1):
            celda = ws.cell(row=fila, column=columna)
            celda.value = nombre
            estilo_encabezado(celda)
        fila += 1
        for _, registro in df.iterrows():
            for columna, nombre in enumerate(df.columns, start=1):
                valor = registro[nombre]
                celda = ws.cell(row=fila, column=columna)
                celda.value = valor if not pd.isna(valor) else ""
                if nombre == "Valor":
                    celda.number_format = '$ #,##0.00'
            fila += 1
    else:
        fila += 1

    columna_valor = None
    if not df.empty:
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
    
    if not df.empty:
        for _, registro in df.iterrows():
            for columna, nombre in enumerate(columnas, start=1):
                valor = registro.get(nombre, "")
                celda = ws.cell(row=fila, column=columna)
                celda.value = valor if not pd.isna(valor) else ""
                if nombre != "Fecha":
                    celda.number_format = '$ #,##0.00'
            fila += 1
            
    fila_total = fila
    ws.cell(row=fila_total, column=1).value = "TOTAL"
    ws.cell(row=fila_total, column=1).font = Font(bold=True)
    for columna, nombre in enumerate(columnas[1:], start=2):
        total = float(pd.to_numeric(df.get(nombre, pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not df.empty else 0.0
        ws.cell(row=fila_total, column=columna).value = total
        ws.cell(row=fila_total, column=columna).number_format = '#,##0.00'
        ws.cell(row=fila_total, column=columna).font = Font(bold=True)
    return fila_total + 2

def preparar_excel(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
    resultado_final, salidas_extracto, salidas_libros, entradas_libros,
    entradas_extracto, gastos_bancarios, preparado_por, revisado_por,
    nombres_titulos, aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion="", estado="Pendiente de revisión"
):
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Conciliación Bancaria"
    
    anchos = {"A": 15, "B": 28, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    ws.merge_cells("A1:G1")
    ws["A1"] = f"CONCILIACIÓN BANCARIA - {tipo.upper()}"
    estilo_titulo(ws["A1"], 16)

    fila = escribir_seccion(ws, 3, "INFORMACIÓN GENERAL", 7)
    logo_bytes = obtener_logo_empresa(empresa)
    if logo_bytes:
        try:
            logo_stream = io.BytesIO(logo_bytes)
            logo_excel = XLImage(logo_stream)
            logo_excel.width = 120
            logo_excel.height = 65
            logo_excel.anchor = "E4"
            ws.add_image(logo_excel)
            ws.row_dimensions[4].height = max(ws.row_dimensions[4].height or 15, 50)
        except Exception:
            pass

    datos_generales = [
        ("Empresa", empresa), ("NIT", nit), ("Mes y Año", mes),
        ("Fecha de elaboración", fecha_elaboracion), ("Banco", banco),
        ("Cuenta No.", cuenta), ("Tipo", tipo),
    ]
    for nombre, valor in datos_generales:
        ws.cell(row=fila, column=1).value = nombre
        ws.cell(row=fila, column=1).font = Font(bold=True)
        ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=4)
        ws.cell(row=fila, column=2).value = str(valor)
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
        ws.cell(row=fila, column=2).number_format = '$ #,##0.00'
        fila += 1

    fila = escribir_tabla(ws, fila, nombres_titulos["t1"], salidas_extracto)
    fila = escribir_tabla(ws, fila, nombres_titulos["t2"], salidas_libros)
    fila = escribir_tabla(ws, fila, nombres_titulos["t3"], entradas_libros)
    fila = escribir_tabla(ws, fila, nombres_titulos["t4"], entradas_extracto)
    fila = escribir_gastos_bancarios(ws, fila, gastos_bancarios)

    fila += 1
    fila = escribir_seccion(ws, fila, "FIRMAS Y TRAZABILIDAD", 6)
    for idx, titulo in enumerate(["PREPARÓ", "REVISÓ", "APROBÓ"], start=0):
        col = idx * 2 + 1
        ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col + 1)
        celda = ws.cell(row=fila, column=col)
        celda.value = titulo
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center")
    fila += 1
    firmas_excel = [
        (preparado_por or "Pendiente", fecha_preparacion or "Se registra al guardar"),
        (revisado_por or "Pendiente de revisión", fecha_revision or "Pendiente"),
        (aprobado_por or "Pendiente de aprobación", fecha_aprobacion or "Pendiente"),
    ]
    for idx, (nombre_firma, fecha_firma) in enumerate(firmas_excel, start=0):
        col = idx * 2 + 1
        ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col + 1)
        celda = ws.cell(row=fila, column=col)
        celda.value = f"Usuario: {nombre_firma}\nFecha: {fecha_firma}\nFirma: ____________________"
        celda.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[fila].height = 58
    fila += 1
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=6)
    ws.cell(row=fila, column=1).value = f"Estado de aprobación: {estado or 'Pendiente de revisión'}"
    ws.cell(row=fila, column=1).font = Font(bold=True)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    nombre_archivo = f"CONCILIACION_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.xlsx"
    return output.getvalue(), nombre_archivo

def generar_pdf_conciliacion(c_data, datos, nombres_titulos):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(letter), rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1, textColor=colors.white, backColor=colors.HexColor("#1F4E78"), fontSize=12)
    sec_style = ParagraphStyle('SecStyle', parent=styles['Heading2'], textColor=colors.HexColor("#1F4E78"), fontSize=10, spaceBefore=6, spaceAfter=4)
    normal_style = ParagraphStyle('NormStyle', parent=styles['Normal'], fontSize=8, leading=10)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold")

    consecutivo_str = f"CONC-{int(c_data.get('id', 0)):06d}"
    tipo_cta = c_data.get("tipo") or "Cuenta"
    logo_bytes = obtener_logo_empresa(c_data.get("empresa"))
    
    p_header_text = Paragraph(f"<b>CONCILIACIÓN BANCARIA - {tipo_cta.upper()}</b><br/><font size=8>Consecutivo No: {consecutivo_str}</font>", title_style)
    if logo_bytes:
        try:
            img_stream = io.BytesIO(logo_bytes)
            img_logo = Image(img_stream, width=80, height=45)
            header_table = Table([[p_header_text, img_logo]], colWidths=[440, 100])
            header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (1,0), (1,0), 'RIGHT')]))
            story.append(header_table)
        except Exception:
            story.append(p_header_text)
    else:
        story.append(p_header_text)
        
    story.append(Spacer(1, 6))

    data_gen = [
        [Paragraph("<b>Empresa:</b>", normal_style), Paragraph(str(c_data.get("empresa", "")), normal_style), Paragraph("<b>NIT:</b>", normal_style), Paragraph(str(c_data.get("nit", "")), normal_style)],
        [Paragraph("<b>Mes/Año:</b>", normal_style), Paragraph(str(c_data.get("mes", "")), normal_style), Paragraph("<b>Elaboración:</b>", normal_style), Paragraph(str(c_data.get("fecha_elaboracion", "")), normal_style)],
        [Paragraph("<b>Banco:</b>", normal_style), Paragraph(str(c_data.get("banco", "")), normal_style), Paragraph("<b>Cuenta No:</b>", normal_style), Paragraph(str(c_data.get("cuenta", "")), normal_style)],
    ]
    t_gen = Table(data_gen, colWidths=[80, 190, 80, 190])
    t_gen.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F2F2F2")), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    story.append(t_gen)
    story.append(Spacer(1, 8))

    story.append(Paragraph("SALDOS Y CÁLCULO DE LA CONCILIACIÓN", sec_style))
    res_final_val = c_data.get('resultado_final', 0.0)
    data_sal = [
        [Paragraph("SALDO SEGÚN EXTRACTO BANCARIO", normal_style), formatear_moneda(c_data.get('saldo_extracto', 0))],
        [Paragraph("SALDO SEGÚN LIBROS", normal_style), formatear_moneda(c_data.get('saldo_libros', 0))],
        [Paragraph("DIFERENCIA A JUSTIFICAR", normal_style), formatear_moneda(c_data.get('diferencia_inicial', 0))],
        [Paragraph("DIFERENCIA CONCILIADA", normal_style), formatear_moneda(c_data.get('diferencia_conciliada', 0))],
        [Paragraph("RESULTADO FINAL", bold_style), Paragraph(formatear_moneda(res_final_val), bold_style)],
        [Paragraph("ESTADO", bold_style), Paragraph(str(c_data.get('estado', '')), bold_style)],
    ]
    t_sal = Table(data_sal, colWidths=[320, 220])
    t_sal.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")), ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#D9EAF7")), ('ALIGN', (1, 0), (1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    story.append(t_sal)
    story.append(Spacer(1, 8))

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
                    try: val = formatear_moneda(val)
                    except (ValueError, TypeError): pass
                r.append(Paragraph(str(val), normal_style))
            rows.append(r)
        t_m = Table(rows, colWidths=col_widths)
        t_m.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D9EAF7")), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        story.append(t_m)
        story.append(Spacer(1, 4))

    agregar_tabla_pdf(f"1. {nombres_titulos['t1']}", datos.get("salidas_extracto", []), ["Fecha", "Beneficiario", "Documento", "Valor"], ["Fecha", "Beneficiario", "Doc", "Valor"], [70, 240, 110, 120])
    agregar_tabla_pdf(f"2. {nombres_titulos['t2']}", datos.get("salidas_libros", []), ["Fecha", "Concepto", "Valor"], ["Fecha", "Concepto", "Valor"], [80, 340, 120])
    agregar_tabla_pdf(f"3. {nombres_titulos['t3']}", datos.get("entradas_libros", []), ["Fecha", "Concepto", "Valor"], ["Fecha", "Concepto", "Valor"], [80, 340, 120])
    agregar_tabla_pdf(f"4. {nombres_titulos['t4']}", datos.get("entradas_extracto", []), ["Fecha", "Concepto", "Valor"], ["Fecha", "Concepto", "Valor"], [80, 340, 120])
    agregar_tabla_pdf("GASTOS BANCARIOS", datos.get("gastos_bancarios", []), ["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"], ["Fecha", "4x1000", "Cuota", "IVA", "Rte", "Comisión", "Intereses"], [60, 80, 80, 80, 80, 80, 80])

    story.append(Spacer(1, 12))
    story.append(Paragraph("FIRMAS Y TRAZABILIDAD", sec_style))
    firmas = [[
        Paragraph("<b>PREPARÓ</b>", normal_style),
        Paragraph("<b>REVISÓ</b>", normal_style),
        Paragraph("<b>APROBÓ</b>", normal_style)
    ],[
        Paragraph(f"{datos.get('preparado_por_usuario') or datos.get('preparado_por') or 'Pendiente'}<br/>Fecha: {datos.get('fecha_preparacion') or 'Pendiente'}<br/><br/>Firma: ____________________", normal_style),
        Paragraph(f"{datos.get('revisado_por_usuario') or c_data.get('revisado_por_usuario') or 'Pendiente de revisión'}<br/>Fecha: {datos.get('fecha_revision') or 'Pendiente'}<br/><br/>Firma: ____________________", normal_style),
        Paragraph(f"{datos.get('aprobado_por_usuario') or 'Pendiente de aprobación'}<br/>Fecha: {datos.get('fecha_aprobacion') or 'Pendiente'}<br/><br/>Firma: ____________________", normal_style)
    ]]
    t_firmas = Table(firmas, colWidths=[180, 180, 180])
    t_firmas.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")), ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#D9EAF7")), ('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(t_firmas)
    story.append(Paragraph(f"Estado de aprobación: <b>{c_data.get('workflow_status') or datos.get('workflow_status') or 'Pendiente de revisión'}</b>", normal_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 9. VISTAS Y NAVEGACIÓN
# ==========================================
if menu_seleccionado == "🔍 Auditoría y Revisiones":
    st.title("🔍 BANDEJA Y MÓDULO DE AUDITORÍA Y CONTROL INTERNO")
    st.caption("Panel de control exclusivo para la revisión analítica, verificación y dictamen de Conciliaciones Bancarias.")

    historial = obtener_historial(empresa_activa_nombre)
    pendientes = historial[historial["workflow_status"] == "Pendiente de revisión"] if not historial.empty else pd.DataFrame()

    k1, k2, k3 = st.columns(3)
    k1.metric("Pendientes de Revisar", len(pendientes))
    k2.metric("Aprobadas", len(historial[historial["workflow_status"] == "Aprobada"]) if not historial.empty else 0)
    k3.metric("Devueltas ❌", len(historial[historial["workflow_status"] == "Requiere corrección"]) if not historial.empty else 0)
    st.divider()

    if pendientes.empty:
        st.success("🎉 ¡Excelente! No hay conciliaciones pendientes por auditar en este momento.")
    else:
        st.subheader("📋 Conciliaciones Asignadas para Auditoría")
        for idx, fila in pendientes.iterrows():
            cuenta_txt = fila.get("cuenta") or "N/A"
            banco_txt = fila.get("banco") or "N/A"
            consecutivo_str = f"CONC-{int(fila['id']):06d}"
            wf_status = fila.get("workflow_status", "Pendiente de revisión")

            with st.expander(f"📌 {consecutivo_str} | {fila['empresa']} - {banco_txt} ({cuenta_txt}) | Mes: {fila['mes']}"):
                c_data = obtener_conciliacion_por_id(fila["id"])
                try: datos = json.loads(c_data.get("datos_json", "{}"))
                except Exception: datos = {}

                tipo_cta = c_data.get("tipo") or "Cuenta de ahorros"
                tipo_cta_lower = tipo_cta.lower()
                es_caja_hist = "caja" in tipo_cta_lower
                es_credito_hist = "crédito bancario" in tipo_cta_lower or "credito bancario" in tipo_cta_lower
                es_tc = "tarjeta" in tipo_cta_lower

                if es_caja_hist or es_credito_hist:
                    st.info(f"📌 Tipo de conciliación: **{tipo_cta}**")
                    datos_especiales = datos
                    wf_actual = c_data.get("workflow_status", "Pendiente de revisión")

                    compras_h = pd.DataFrame(datos_especiales.get("compras", [])) if es_caja_hist else pd.DataFrame()
                    efectivo_h = pd.DataFrame(datos_especiales.get("efectivo", [])) if es_caja_hist else pd.DataFrame()
                    dif_h = pd.DataFrame(datos_especiales.get("diferencias", [])) if es_credito_hist else pd.DataFrame()

                    a1, a2, a3, a4 = st.tabs(["📋 Resumen y Evidencia", "✅ Checklist de Control", "🚩 Hallazgos y Riesgo", "⚡ Dictamen y Decisión"])

                    with a1:
                        st.subheader("📋 Resumen de la conciliación")
                        r1, r2 = st.columns(2)
                        with r1:
                            st.write(f"🏢 **Empresa:** {c_data.get('empresa', 'N/A')} | **NIT:** {c_data.get('nit', 'N/A')}")
                            st.write(f"📅 **Período:** {c_data.get('mes', 'N/A')} | **Fecha:** {datos_especiales.get('fecha_elaboracion', 'N/A')}")
                            st.write(f"🏦 **Entidad/Banco:** {datos_especiales.get('entidad_financiera', c_data.get('banco', 'N/A'))}")
                        with r2:
                            st.write(f"💳 **Cuenta/Crédito:** {datos_especiales.get('numero_credito', c_data.get('cuenta', 'N/A'))}")
                            st.write(f"👤 **Preparado por:** {datos_especiales.get('preparado_por_usuario', c_data.get('preparado_por', 'N/A'))}")
                            st.write(f"🔎 **Estado:** {wf_actual}")

                        cexp1, cexp2, cexp3 = st.columns(3)
                        with cexp1:
                            if st.button("✏️ Editar Conciliación", key=f"btn_edit_{fila['id']}"):
                                st.session_state.conciliacion_a_editar = fila['id']
                                st.session_state.pop("datos_cargados_edit", None)
                                st.session_state.menu_override = "📝 Nueva Conciliación"
                                st.rerun()
                        with cexp2:
                            if st.button("📊 Descargar Excel", key=f"dlx_{fila['id']}"):
                                if es_caja_hist:
                                    excel_h, nombre_h = preparar_excel_caja(c_data.get("empresa"), c_data.get("mes"), datos_especiales.get("fecha_elaboracion", ""), datos_especiales.get("responsable", c_data.get("preparado_por", "")), datos_especiales.get("saldo_inicial", 0), datos_especiales.get("fondo_autorizado", 0), compras_h, efectivo_h, datos_especiales.get("observaciones_caja", ""), obtener_logo_empresa(c_data.get("empresa")), datos_especiales.get("preparado_por_usuario", c_data.get("preparado_por", "")), datos_especiales.get("revisado_por_usuario", c_data.get("revisado_por", "")), datos_especiales.get("aprobado_por_usuario", ""), datos_especiales.get("fecha_preparacion", ""), datos_especiales.get("fecha_revision", ""), datos_especiales.get("fecha_aprobacion", ""), wf_actual)
                                else:
                                    excel_h, nombre_h = preparar_excel_credito_bancario(c_data.get("empresa"), c_data.get("mes"), datos_especiales.get("fecha_elaboracion", ""), datos_especiales.get("entidad_financiera", c_data.get("banco", "")), datos_especiales.get("numero_credito", c_data.get("cuenta", "")), datos_especiales.get("fecha_inicio", ""), datos_especiales.get("fecha_vencimiento", ""), datos_especiales.get("tasa", ""), c_data.get("saldo_libros", 0), c_data.get("saldo_extracto", 0), dif_h, datos_especiales.get("observaciones_credito", ""), obtener_logo_empresa(c_data.get("empresa")), datos_especiales.get("preparado_por_usuario", c_data.get("preparado_por", "")), datos_especiales.get("revisado_por_usuario", c_data.get("revisado_por", "")), datos_especiales.get("aprobado_por_usuario", ""), datos_especiales.get("fecha_preparacion", ""), datos_especiales.get("fecha_revision", ""), datos_especiales.get("fecha_aprobacion", ""), wf_actual)
                                st.download_button("Guardar Excel", data=excel_h, file_name=nombre_h, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        with cexp3:
                            if st.button("📄 Descargar PDF", key=f"dlp_{fila['id']}"):
                                if es_caja_hist:
                                    pdf_h, nombre_p = generar_pdf_caja(c_data.get("empresa"), c_data.get("mes"), datos_especiales.get("fecha_elaboracion", ""), datos_especiales.get("responsable", c_data.get("preparado_por", "")), datos_especiales.get("saldo_inicial", 0), datos_especiales.get("fondo_autorizado", 0), compras_h, efectivo_h, datos_especiales.get("observaciones_caja", ""), obtener_logo_empresa(c_data.get("empresa")), datos_especiales.get("preparado_por_usuario", c_data.get("preparado_por", "")), datos_especiales.get("revisado_por_usuario", c_data.get("revisado_por", "")), datos_especiales.get("aprobado_por_usuario", ""), datos_especiales.get("fecha_preparacion", ""), datos_especiales.get("fecha_revision", ""), datos_especiales.get("fecha_aprobacion", ""), wf_actual)
                                else:
                                    pdf_h, nombre_p = generar_pdf_credito_bancario(c_data.get("empresa"), c_data.get("mes"), datos_especiales.get("fecha_elaboracion", ""), datos_especiales.get("entidad_financiera", c_data.get("banco", "")), datos_especiales.get("numero_credito", c_data.get("cuenta", "")), datos_especiales.get("fecha_inicio", ""), datos_especiales.get("fecha_vencimiento", ""), datos_especiales.get("tasa", ""), c_data.get("saldo_libros", 0), c_data.get("saldo_extracto", 0), dif_h, datos_especiales.get("observaciones_credito", ""), obtener_logo_empresa(c_data.get("empresa")), datos_especiales.get("preparado_por_usuario", c_data.get("preparado_por", "")), datos_especiales.get("revisado_por_usuario", c_data.get("revisado_por", "")), datos_especiales.get("aprobado_por_usuario", ""), datos_especiales.get("fecha_preparacion", ""), datos_especiales.get("fecha_revision", ""), datos_especiales.get("fecha_aprobacion", ""), wf_actual)
                                st.download_button("Guardar PDF", data=pdf_h, file_name=nombre_p, mime="application/pdf")

                    with a2:
                        chk1 = st.checkbox("Soportes verificados.", key=f"esp_chk1_{fila['id']}")
                        chk2 = st.checkbox("Saldos validados.", key=f"esp_chk2_{fila['id']}")

                    with a3:
                        hallazgo_esp = st.selectbox("Categoría del hallazgo", ["Sin hallazgos", "Otro"], key=f"esp_hallazgo_{fila['id']}")
                        detalle_esp = st.text_area("Detalle / observaciones", key=f"esp_detalle_{fila['id']}")

                    with a4:
                        decision_esp = st.radio("Seleccione Acción:", ["✅ Aprobar Conciliación", "❌ Devolver a Corrección"], key=f"esp_dec_{fila['id']}")
                        if "Aprobar" in decision_esp:
                            if st.button("✅ Confirmar Aprobación", key=f"esp_aprobar_{fila['id']}", type="primary"):
                                actualizar_estado_auditoria(fila["id"], "Aprobada", usuario_actual["nombre"])
                                st.success("🎉 Conciliación aprobada.")
                                st.rerun()
                        else:
                            if st.button("❌ Confirmar Devolución", key=f"esp_devolver_{fila['id']}"):
                                actualizar_estado_auditoria(fila["id"], "Requiere corrección", usuario_actual["nombre"], motivo_correccion=detalle_esp, tipo_hallazgo=hallazgo_esp)
                                st.warning("⚠️ Conciliación devuelta.")
                                st.rerun()
                    continue

                t1_nombre, t2_nombre, t3_nombre, t4_nombre = ("COMPRAS NO EVIDENCIADAS EN EXTRACTOS", "COMPRAS NO CONTABILIZADAS EN LIBROS", "DÉBITOS BANCARIOS NO CONTABILIZADOS EN LIBROS", "ABONOS NO REGISTRADOS EN EXTRACTO") if es_tc else ("SALIDAS NO REGISTRADAS EN EXTRACTO", "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS", "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS", "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS")

                tab_aud1, tab_aud2, tab_aud3, tab_aud4 = st.tabs(["📊 Movimientos y Saldos", "✅ Checklist de Verificación", "🚩 Análisis de Partidas", "⚡ Dictamen y Decisión"])

                with tab_aud1:
                    st.write(f"🏢 **Empresa:** {c_data.get('empresa', 'N/A')} | **Banco:** {c_data.get('banco', 'N/A')}")
                    df_saldos = pd.DataFrame([
                        {"Concepto": "SALDO EXTRACTO", "Valor": c_data.get('saldo_extracto', 0)},
                        {"Concepto": "SALDO LIBROS", "Valor": c_data.get('saldo_libros', 0)},
                        {"Concepto": "RESULTADO FINAL", "Valor": c_data.get('resultado_final', 0)}
                    ])
                    st.dataframe(df_saldos, use_container_width=True, hide_index=True)

                with tab_aud2:
                    chk_extracto = st.checkbox("El saldo de extracto coincide.", key=f"chk_ext_{fila['id']}")
                    chk_libros = st.checkbox("El saldo en libros fue verificado.", key=f"chk_lib_{fila['id']}")

                with tab_aud3:
                    st.metric("Total Registros", len(datos.get("salidas_extracto", [])) + len(datos.get("salidas_libros", [])))

                with tab_aud4:
                    decision = st.radio("Acción:", ["✅ Aprobar Conciliación", "❌ Devolver a Corrección"], key=f"dec_{fila['id']}")
                    if "Aprobar" in decision:
                        if st.button("✅ Confirmar Aprobación", key=f"btn_aprob_final_{fila['id']}", type="primary"):
                            actualizar_estado_auditoria(fila["id"], "Aprobada", usuario_actual["nombre"])
                            st.success("🎉 Conciliación aprobada.")
                            st.rerun()
                    else:
                        motivo_det = st.text_area("Observaciones:", key=f"mot_det_{fila['id']}")
                        if st.button("❌ Devolver al Preparador", key=f"btn_dev_final_{fila['id']}"):
                            actualizar_estado_auditoria(fila["id"], "Requiere corrección", usuario_actual["nombre"], motivo_correccion=motivo_det)
                            st.warning("⚠️ Conciliación devuelta.")
                            st.rerun()

elif menu_seleccionado == "📊 Dashboard":
    st.title("📊 DASHBOARD Y CENTRO DE CONTROL FINANCIERO")
    historial = obtener_historial(empresa_activa_nombre)
    if historial.empty:
        st.info("ℹ️ Aún no hay datos registrados para generar el análisis analítico.")
    else:
        total_conciliaciones = len(historial)
        exitosas = len(historial[historial["resultado_final"].abs() < 0.005])
        tasa_exito = (exitosas / total_conciliaciones * 100) if total_conciliaciones > 0 else 0.0

        aprobadas_cnt = len(historial[historial["workflow_status"] == "Aprobada"])
        pendientes_cnt = len(historial[historial["workflow_status"] == "Pendiente de revisión"])
        borradores_cnt = len(historial[historial["workflow_status"] == "Borrador"])
        devueltas_cnt = len(historial[historial["workflow_status"] == "Requiere corrección"])

        st.subheader("📈 Indicadores Clave de Desempeño (KPIs)")
        kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
        kpi_col1.metric("Total Conciliaciones", total_conciliaciones)
        kpi_col2.metric("Tasa de Éxito", f"{tasa_exito:.1f}%")
        kpi_col3.metric("Pendientes por Auditar", pendientes_cnt)

        st.divider()
        df_pie = pd.DataFrame({
            "Estado": ["Aprobada", "Pendiente de revisión", "Borrador", "Requiere corrección"],
            "Cantidad": [aprobadas_cnt, pendientes_cnt, borradores_cnt, devueltas_cnt]
        })
        df_pie = df_pie[df_pie["Cantidad"] > 0]
        if not df_pie.empty:
            fig_pie = px.pie(df_pie, names="Estado", values="Cantidad", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

elif menu_seleccionado == "🏢 Empresas":
    st.title("🏢 Maestro de Empresas")
    empresa_edit_id = st.session_state.get("empresa_a_editar", None)
    if empresa_edit_id:
        emp_obj = obtener_empresa_por_id(empresa_edit_id)
        if emp_obj:
            with st.form("form_editar_empresa"):
                edit_nom = st.text_input("Nombre de la empresa", value=emp_obj["nombre"])
                edit_nit = st.text_input("NIT", value=emp_obj["nit"])
                edit_logo = st.file_uploader("Actualizar Logo", type=["png", "jpg", "jpeg"])
                if st.form_submit_button("Guardar Cambios", type="primary"):
                    logo_b = edit_logo.getvalue() if edit_logo else None
                    ok, mensaje = actualizar_empresa_db(empresa_edit_id, edit_nom, edit_nit, logo_b)
                    if ok:
                        st.session_state.empresa_a_editar = None
                        st.success(mensaje)
                        st.rerun()

    if rol_actual == "Administrador" and not empresa_edit_id:
        with st.expander("➕ Registrar nueva empresa con Logo"):
            with st.form("form_empresa_maestro"):
                nom = st.text_input("Nombre de la empresa")
                nit = st.text_input("NIT")
                logo_file = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
                if st.form_submit_button("Guardar Empresa", type="primary"):
                    ok, mensaje = guardar_empresa(nom, nit, logo_file.getvalue() if logo_file else None)
                    if ok: st.success(mensaje); st.rerun()
                    else: st.warning(mensaje)

    empresas_list = obtener_empresas()
    if not empresas_list.empty:
        for idx, emp in empresas_list.iterrows():
            st.write(f"🏢 **{emp['nombre']}** - NIT: {emp['nit']}")
            if rol_actual == "Administrador":
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✏️ Editar", key=f"btn_edit_emp_{emp['id']}"):
                        st.session_state.empresa_a_editar = emp['id']
                        st.rerun()
                with c2:
                    if st.button("🗑️ Eliminar", key=f"btn_del_emp_{emp['id']}"):
                        ok, mensaje = eliminar_empresa_db(emp['id'])
                        if ok: st.success(mensaje); st.rerun()

elif menu_seleccionado == "🏦 Bancos y Cuentas":
    st.title("🏦 Maestro de Bancos y Cuentas Bancarias")
    if rol_actual == "Administrador":
        with st.expander("➕ Registrar nueva cuenta bancaria"):
            with st.form("form_cuenta_maestro"):
                banco = st.text_input("Banco")
                num = st.text_input("Número de Cuenta")
                tipo = st.selectbox("Tipo", ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito", "Caja", "Crédito bancario"])
                emp_id = st.selectbox("Empresa", empresas_df["id"].tolist(), format_func=lambda x: empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0]) if not empresas_df.empty else None

                if st.form_submit_button("💾 Guardar Cuenta", type="primary"):
                    if banco and num and emp_id:
                        guardar_cuenta(banco, num, tipo, emp_id)
                        st.success("Cuenta guardada.")
                        st.rerun()

    cuentas_reg = obtener_cuentas(empresa_activa_id)
    if not cuentas_reg.empty:
        for _, cuenta_fila in cuentas_reg.iterrows():
            st.write(f"**{cuenta_fila['banco']}** - {cuenta_fila['numero_cuenta']} ({cuenta_fila['tipo_cuenta']})")
            if st.button("🗑️ Eliminar", key=f"eliminar_cuenta_{cuenta_fila['id']}"):
                ok, msj = eliminar_cuenta_db(int(cuenta_fila['id']))
                if ok: st.success(msj); st.rerun()

elif menu_seleccionado == "📝 Nueva Conciliación":
    st.title("📝 Captura de Conciliación Bancaria")
    id_edicion = st.session_state.get("conciliacion_a_editar", None)

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

    col1, col2 = st.columns(2)
    with col1:
        empresa = st.text_input("Empresa", value=empresa_activa_nombre)
        nit = st.text_input("NIT", value=empresa_activa_nit)
        mes = st.text_input("Mes/Año", value=datetime.now().strftime("%B %Y"))
        fecha_elaboracion = st.date_input("Fecha Elaboración")
    with col2:
        banco = st.text_input("Banco")
        cuenta = st.text_input("Cuenta")
        tipo = st.selectbox("Tipo de Cuenta", ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito", "Caja", "Crédito bancario"])

    saldo_extracto = st.number_input("Saldo según Extracto", format="%.2f")
    saldo_libros = st.number_input("Saldo según Libros", format="%.2f")
    diferencia_inicial = saldo_extracto - saldo_libros

    st.subheader("Movimientos")
    salidas_extracto = st.data_editor(st.session_state.tabla1, num_rows="dynamic", use_container_width=True, key="editor_t1")
    salidas_libros = st.data_editor(st.session_state.tabla2, num_rows="dynamic", use_container_width=True, key="editor_t2")
    entradas_libros = st.data_editor(st.session_state.tabla3, num_rows="dynamic", use_container_width=True, key="editor_t3")
    entradas_extracto = st.data_editor(st.session_state.tabla4, num_rows="dynamic", use_container_width=True, key="editor_t4")
    gastos_bancarios = st.data_editor(st.session_state.tabla5, num_rows="dynamic", use_container_width=True, key="editor_t5")

    diferencia_conciliada = total_columna(salidas_extracto) - total_columna(salidas_libros) + total_columna(entradas_libros) - total_columna(entradas_extracto)
    resultado_final = diferencia_inicial - diferencia_conciliada

    nombres_titulos = {"t1": "SALIDAS NO REGISTRADAS EN EXTRACTO", "t2": "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS", "t3": "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS", "t4": "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS"}

    btn_col_borr, btn_col_env = st.columns(2)
    with btn_col_borr:
        if st.button("💾 Guardar como Borrador", key="btn_guardar_borrador"):
            excel_data, nombre_archivo = preparar_excel(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, usuario_actual["nombre"], "", nombres_titulos
            )
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, usuario_actual["nombre"], "", excel_data,
                workflow_status="Borrador", id_edicion=id_edicion
            )
            st.session_state.pop("conciliacion_a_editar", None)
            st.session_state.menu_override = "📋 Historial"
            st.success(f"💾 Conciliación CONC-{int(id_g):06d} guardada como Borrador.")
            st.rerun()

    with btn_col_env:
        if st.button("🚀 Enviar a Revisión", key="btn_enviar_revision", type="primary"):
            excel_data, nombre_archivo = preparar_excel(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, usuario_actual["nombre"], "", nombres_titulos
            )
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, usuario_actual["nombre"], "", excel_data,
                workflow_status="Pendiente de revisión", id_edicion=id_edicion
            )
            st.session_state.pop("conciliacion_a_editar", None)
            st.session_state.menu_override = "📋 Historial"
            st.success(f"🚀 Conciliación CONC-{int(id_g):06d} enviada a Revisión.")
            st.rerun()

elif menu_seleccionado == "📋 Historial":
    st.title("📋 Historial de Conciliaciones")
    historial = obtener_historial(empresa_activa_nombre)

    if historial.empty:
        st.info("No hay conciliaciones guardadas.")
    else:
        for idx, fila in historial.iterrows():
            consecutivo_str = f"CONC-{int(fila['id']):06d}"
            with st.expander(f"📌 {consecutivo_str} | {fila['empresa']} - {fila['banco']} | Mes: {fila['mes']}"):
                st.write(f"Saldo Extracto: {formatear_moneda(fila['saldo_extracto'])} | Saldo Libros: {formatear_moneda(fila['saldo_libros'])}")
                if rol_actual == "Administrador":
                    if st.button("🗑️ Eliminar", key=f"btn_delete_conc_{fila['id']}"):
                        eliminar_conciliacion_db(int(fila['id']))
                        st.success("Eliminada correctamente.")
                        st.rerun()

elif menu_seleccionado == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios y Roles")
    usuarios_df = obtener_usuarios()
    if not usuarios_df.empty:
        st.dataframe(usuarios_df[["usuario", "nombre", "rol", "empresa_nombre"]], use_container_width=True)

elif menu_seleccionado == "⚙️ Administración":
    st.title("⚙️ Administración")
    if st.button("💾 CREAR COPIA DE SEGURIDAD", type="primary"):
        datos_backup = crear_copia_seguridad_bytes()
        st.download_button("📥 DESCARGAR COPIA DE SEGURIDAD", data=datos_backup, file_name="backup_conciliacion.zip", mime="application/zip")

elif menu_seleccionado == "📄 Reportes":
    st.title("📄 Reportes y Descargas")
    historial = obtener_historial(empresa_activa_nombre)
    if not historial.empty:
        csv = historial.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Historial Completo (CSV)", data=csv, file_name="historial.csv", mime="text/csv")
