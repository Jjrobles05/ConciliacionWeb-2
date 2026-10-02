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
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
    """
    Conecta directamente a Turso cuando existen los Secrets
    TURSO_DATABASE_URL y TURSO_AUTH_TOKEN.

    Si los Secrets no están configurados, utiliza SQLite local
    como respaldo para desarrollo.
    """
    if "TURSO_DATABASE_URL" in st.secrets and "TURSO_AUTH_TOKEN" in st.secrets:
        url = str(st.secrets["TURSO_DATABASE_URL"]).strip()
        token = str(st.secrets["TURSO_AUTH_TOKEN"]).strip()

        # Turso Serverless utiliza HTTPS para el acceso remoto.
        if url.startswith("libsql://"):
            url = "https://" + url[len("libsql://"): ]

        try:
            import turso_serverless

            return turso_serverless.connect(
                url,
                auth_token=token
            )
        except Exception as e:
            st.error(
                "⚠️ No fue posible conectar con Turso. "
                f"{type(e).__name__}: {e}"
            )
            st.stop()

    return sqlite3.connect(
        "conciliaciones.db",
        check_same_thread=False
    )

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
    conn.commit()
    conn.close()

inicializar_db()

# ==========================================
# 4. CONSULTAS A LA BASE DE DATOS
# ==========================================
def obtener_empresas():
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, razon_social, nit FROM empresas ORDER BY razon_social")
        rows = c.fetchall()
        conn.close()
        df = pd.DataFrame(rows, columns=['id', 'nombre', 'nit']) if rows else pd.DataFrame(columns=['id', 'nombre', 'nit'])
        df['logo'] = None
        return df
    except Exception:
        return pd.DataFrame(columns=['id', 'nombre', 'nit', 'logo'])


def obtener_empresa_por_id(empresa_id):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, razon_social, nit FROM empresas WHERE id=?", (int(empresa_id),))
        res = c.fetchone()
        conn.close()
        if res:
            return {"id": res[0], "nombre": res[1], "nit": res[2], "logo": None}
    except Exception:
        pass
    return None


def obtener_logo_empresa(nombre_empresa):
    return None


def guardar_empresa(nombre, nit, logo_bytes=None):
    """Registra una empresa y controla NIT duplicado sin mostrar errores técnicos."""
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
            return False, f"El NIT {nit} ya está registrado para la empresa {existente[1]}."
        c.execute("INSERT INTO empresas (nit, razon_social) VALUES (?, ?)", (nit, nombre))
        conn.commit()
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
    conn = conectar_db(); c = conn.cursor()
    try:
        c.execute("SELECT id, razon_social FROM empresas WHERE nit=? AND id<>?", (nit, int(empresa_id)))
        existente = c.fetchone()
        if existente:
            return False, f"El NIT {nit} ya está registrado para {existente[1]}."
        c.execute("UPDATE empresas SET razon_social=?, nit=? WHERE id=?", (nombre, nit, int(empresa_id)))
        conn.commit()
        return True, "Empresa actualizada con éxito."
    except Exception as e:
        try: conn.rollback()
        except Exception: pass
        return False, f"No fue posible actualizar la empresa: {type(e).__name__}: {e}"
    finally:
        conn.close()



def eliminar_empresa_db(empresa_id):
    conn = conectar_db(); c = conn.cursor()
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


def guardar_cuenta(banco, numero_cuenta, tipo_cuenta, empresa_id):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("INSERT INTO cuentas_bancarias (banco, numero_cuenta, tipo_cuenta, empresa_id) VALUES (?, ?, ?, ?)",
              (banco.strip(), numero_cuenta.strip(), tipo_cuenta, empresa_id))
    conn.commit()
    last_id = c.lastrowid
    conn.close()
    return last_id


def obtener_cuentas_rotadas_por_usuario(empresa_id, mes_num, usuario_id):
    if not empresa_id:
        return pd.DataFrame()
    conn = conectar_db()
    c = conn.cursor()
    c.execute("SELECT id, nombre FROM usuarios WHERE empresa_id = ? AND activo = 1 AND rol = 'Preparador' ORDER BY id", (int(empresa_id),))
    usr_rows = c.fetchall()
    c.execute("SELECT id, banco, numero_cuenta, tipo_cuenta FROM cuentas_bancarias WHERE empresa_id = ? ORDER BY id", (int(empresa_id),))
    cta_rows = c.fetchall()
    conn.close()
    if not cta_rows:
        return pd.DataFrame()
    cuentas = pd.DataFrame(cta_rows, columns=['id', 'banco', 'numero_cuenta', 'tipo_cuenta'])
    if not usr_rows or usuario_id not in [u[0] for u in usr_rows]:
        return cuentas
    num_usuarios = len(usr_rows)
    ids_usuarios = [u[0] for u in usr_rows]
    cuentas_asignadas = []
    for idx_cuenta, fila_cuenta in cuentas.iterrows():
        idx_usuario_asignado = (idx_cuenta + mes_num) % num_usuarios
        if ids_usuarios[idx_usuario_asignado] == usuario_id:
            cuentas_asignadas.append(fila_cuenta)
    return pd.DataFrame(cuentas_asignadas) if cuentas_asignadas else pd.DataFrame(columns=cuentas.columns)


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


def _json_datos_conciliacion(empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                              saldo_extracto, saldo_libros, diferencia_inicial,
                              diferencia_conciliada, resultado_final, salidas_extracto,
                              salidas_libros, entradas_libros, entradas_extracto,
                              gastos_bancarios, preparado_por, revisado_por,
                              workflow_status):
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
    if row:
        return int(row[0])
    c.execute("SELECT id FROM empresas WHERE razon_social=?", (str(empresa).strip(),))
    row = c.fetchone()
    if row:
        return int(row[0])
    c.execute("INSERT INTO empresas (nit, razon_social) VALUES (?, ?)", (str(nit).strip(), str(empresa).strip()))
    return int(c.lastrowid)


def _buscar_cuenta_id(c, empresa_id, banco, cuenta, tipo):
    c.execute("SELECT id FROM cuentas_bancarias WHERE empresa_id=? AND banco=? AND numero_cuenta=? ORDER BY id LIMIT 1",
              (int(empresa_id), str(banco).strip(), str(cuenta).strip()))
    row = c.fetchone()
    if row:
        return int(row[0])
    c.execute("INSERT INTO cuentas_bancarias (banco, numero_cuenta, tipo_cuenta, empresa_id) VALUES (?, ?, ?, ?)",
              (str(banco).strip(), str(cuenta).strip(), str(tipo).strip(), int(empresa_id)))
    return int(c.lastrowid)


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
            entradas_extracto, gastos_bancarios, preparado_por, revisado_por,
            workflow_status
        )

        # Cuando se corrige una devolución, conservamos las observaciones de
        # Auditoría para que no desaparezcan al guardar o reenviar la corrección.
        if id_edicion:
            c.execute("SELECT observaciones FROM conciliaciones WHERE id=?", (int(id_edicion),))
            anterior = c.fetchone()
            datos_anteriores = _decodificar_observaciones(anterior[0] if anterior else None)
            for campo in ("motivo_correccion", "tipo_hallazgo", "revisado_por_usuario", "fecha_revision"):
                if datos_anteriores.get(campo):
                    datos[campo] = datos_anteriores[campo]

        estado = datos["estado"]
        fecha_creacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        observaciones = json.dumps(datos, ensure_ascii=False)
        if id_edicion:
            c.execute("""
                UPDATE conciliaciones SET empresa_id=?, cuenta_id=?, periodo=?, saldo_libro=?, saldo_banco=?,
                    estado=?, preparado_por=?, revisado_por=?, fecha_creacion=?, dictamen=?, observaciones=?
                WHERE id=?
            """, (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), estado,
                     preparado_por or None, revisado_por or None, fecha_creacion, workflow_status,
                     observaciones, int(id_edicion)))
            last_id = int(id_edicion)
        else:
            c.execute("""
                INSERT INTO conciliaciones (
                    empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, estado,
                    preparado_por, revisado_por, fecha_creacion, dictamen, observaciones
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (empresa_id, cuenta_id, mes, float(saldo_libros), float(saldo_extracto), estado,
                     preparado_por or None, revisado_por or None, fecha_creacion, workflow_status,
                     observaciones))
            last_id = int(c.lastrowid)
        conn.commit()
        return last_id
    finally:
        conn.close()


def _decodificar_observaciones(texto):
    if not texto:
        return {}
    try:
        data = json.loads(texto)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def obtener_historial(empresa_nombre=None):
    try:
        conn = conectar_db()
        c = conn.cursor()
        query = """
            SELECT co.id, co.fecha_creacion, e.razon_social, e.nit, co.periodo,
                   cb.banco, cb.numero_cuenta, cb.tipo_cuenta, co.saldo_banco,
                   co.saldo_libro, co.estado, co.preparado_por, co.revisado_por,
                   co.dictamen, co.observaciones
            FROM conciliaciones co
            LEFT JOIN empresas e ON co.empresa_id=e.id
            LEFT JOIN cuentas_bancarias cb ON co.cuenta_id=cb.id
        """
        params = ()
        if empresa_nombre and empresa_nombre != "Todas las empresas":
            query += " WHERE e.razon_social=?"
            params = (empresa_nombre,)
        query += " ORDER BY co.id DESC"
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()
        if not rows:
            return pd.DataFrame()
        salida = []
        for row in rows:
            d = _decodificar_observaciones(row[14])
            salida.append({
                'id': row[0], 'fecha_guardado': row[1], 'empresa': row[2] or d.get('empresa', ''),
                'nit': row[3] or d.get('nit', ''), 'mes': row[4] or d.get('mes', ''),
                'banco': row[5] or d.get('banco', ''), 'cuenta': row[6] or d.get('cuenta', ''),
                'tipo': row[7] or d.get('tipo', ''), 'saldo_extracto': row[8], 'saldo_libros': row[9],
                'diferencia_inicial': d.get('diferencia_inicial', float(row[8] or 0)-float(row[9] or 0)),
                'diferencia_conciliada': d.get('diferencia_conciliada', 0.0),
                'resultado_final': d.get('resultado_final', 0.0), 'estado': row[10] or d.get('estado', ''),
                'workflow_status': row[13] or d.get('workflow_status', 'Pendiente de revisión'),
                'revisado_por_usuario': row[12] or d.get('revisado_por_usuario', ''),
                'fecha_revision': d.get('fecha_revision'), 'motivo_correccion': d.get('motivo_correccion'),
                'tipo_hallazgo': d.get('tipo_hallazgo'), 'checklist_json': d.get('checklist_json'),
                'datos_json': json.dumps(d, ensure_ascii=False), 'fecha_elaboracion': d.get('fecha_elaboracion', row[1])
            })
        return pd.DataFrame(salida)
    except Exception:
        return pd.DataFrame()


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
        excel_bytes = None
        try:
            from datetime import date
            fecha_elab = datos.get('fecha_elaboracion', datetime.now().date().isoformat())
            try: fecha_elab = date.fromisoformat(str(fecha_elab)[:10])
            except Exception: pass
            excel_bytes, _ = preparar_excel(
                empresa, nit, datos.get('mes', row[3]), fecha_elab, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, datos.get('diferencia_inicial', saldo_extracto-saldo_libros),
                datos.get('diferencia_conciliada',0), datos.get('resultado_final',0),
                pd.DataFrame(datos.get('salidas_extracto',[])), pd.DataFrame(datos.get('salidas_libros',[])),
                pd.DataFrame(datos.get('entradas_libros',[])), pd.DataFrame(datos.get('entradas_extracto',[])),
                pd.DataFrame(datos.get('gastos_bancarios',[])), datos.get('preparado_por',row[7] or ''),
                datos.get('revisado_por',row[8] or ''),
                {'t1':'SALIDAS NO REGISTRADAS EN EXTRACTO','t2':'SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS',
                 't3':'ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS','t4':'ENTRADAS NO EVIDENCIADAS EN EXTRACTOS'}
            )
        except Exception:
            pass
        return {
            'id': row[0], 'empresa': empresa, 'nit': nit, 'mes': datos.get('mes',row[3]),
            'fecha_elaboracion': datos.get('fecha_elaboracion',row[9]), 'banco': banco, 'cuenta': cuenta, 'tipo': tipo,
            'saldo_extracto': saldo_extracto, 'saldo_libros': saldo_libros,
            'diferencia_inicial': datos.get('diferencia_inicial',saldo_extracto-saldo_libros),
            'diferencia_conciliada': datos.get('diferencia_conciliada',0), 'resultado_final': datos.get('resultado_final',0),
            'estado': row[6], 'datos_json': json.dumps(datos,ensure_ascii=False), 'excel': excel_bytes,
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
    c.execute("UPDATE conciliaciones SET revisado_por=?, dictamen=?, observaciones=? WHERE id=?",
              (revisado_por, nuevo_estado, json.dumps(datos,ensure_ascii=False), int(id_conciliacion)))
    conn.commit(); conn.close()

# ==========================================
# 5. AUTENTICACIÓN FLEXIBLE Y SEGURA
# ==========================================
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
    return last_id

def actualizar_empresa_usuario(usuario_id, empresa_id):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("UPDATE usuarios SET empresa_id=? WHERE id=?", (empresa_id, int(usuario_id)))
    conn.commit()
    conn.close()

def actualizar_rol_usuario(usuario_id, nuevo_rol):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("UPDATE usuarios SET rol=? WHERE id=?", (nuevo_rol, int(usuario_id)))
    conn.commit()
    conn.close()

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
        return ["📊 Dashboard", "🔍 Auditoría y Revisiones", "📝 Nueva Conciliación", "📋 Historial", "🏢 Empresas", "🏦 Bancos y Cuentas", "👥 Usuarios", "📄 Reportes"]
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
            try:
                empresa_id_invitada = int(empresa_id_invitada)
            except ValueError:
                empresa_id_invitada = None

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
        empresa_activa_nombre = usuario_actual.get("empresa_nombre") or "Sin asignar"
        empresa_activa_nit = usuario_actual.get("empresa_nit") or ""
        empresa_activa_id = usuario_actual.get("empresa_id")
        st.info(f"🏢 **Empresa:** {empresa_activa_nombre}")

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
                    celda.number_format = '#,##0.00'
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
                    celda.number_format = '#,##0.00'
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
    nombres_titulos
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
        ws.cell(row=fila, column=2).number_format = '#,##0.00'
        fila += 1

    fila = escribir_tabla(ws, fila, nombres_titulos["t1"], salidas_extracto)
    fila = escribir_tabla(ws, fila, nombres_titulos["t2"], salidas_libros)
    fila = escribir_tabla(ws, fila, nombres_titulos["t3"], entradas_libros)
    fila = escribir_tabla(ws, fila, nombres_titulos["t4"], entradas_extracto)
    fila = escribir_gastos_bancarios(ws, fila, gastos_bancarios)

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
        [Paragraph("SALDO SEGÚN EXTRACTO BANCARIO", normal_style), f"$ {c_data.get('saldo_extracto', 0):,.2f}"],
        [Paragraph("SALDO SEGÚN LIBROS", normal_style), f"$ {c_data.get('saldo_libros', 0):,.2f}"],
        [Paragraph("DIFERENCIA A JUSTIFICAR", normal_style), f"$ {c_data.get('diferencia_inicial', 0):,.2f}"],
        [Paragraph("DIFERENCIA CONCILIADA", normal_style), f"$ {c_data.get('diferencia_conciliada', 0):,.2f}"],
        [Paragraph("RESULTADO FINAL", bold_style), Paragraph(f"$ {res_final_val:,.2f}", bold_style)],
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
                    try:
                        val = f"$ {float(val):,.2f}"
                    except (ValueError, TypeError):
                        pass
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
    firmas = [
        [Paragraph(f"<b>Preparado por:</b> {datos.get('preparado_por', 'N/A')}", normal_style), Paragraph(f"<b>Revisado por:</b> {c_data.get('revisado_por_usuario', 'N/A')}", normal_style)]
    ]
    t_firmas = Table(firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([('LINEABOVE', (0, 0), (-1, -1), 1, colors.HexColor("#1F4E78"))]))
    story.append(t_firmas)

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
        st.info(
            "ℹ️ No hay conciliaciones con estado **Pendiente de revisión** "
            "para la empresa seleccionada."
        )
        st.caption(
            "Para que una conciliación aparezca aquí debe haberse enviado desde "
            "Nueva Conciliación o desde Historial usando **Enviar a Revisión**."
        )
    else:
        st.subheader("📋 Conciliaciones Asignadas para Auditoría")
        for idx, fila in pendientes.iterrows():
            cuenta_txt = fila.get("cuenta") or "N/A"
            banco_txt = fila.get("banco") or "N/A"
            consecutivo_str = f"CONC-{int(fila['id']):06d}"

            with st.expander(f"📌 {consecutivo_str} | {fila['empresa']} - {banco_txt} ({cuenta_txt}) | Mes: {fila['mes']}"):
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

                tab_aud1, tab_aud2, tab_aud3, tab_aud4 = st.tabs([
                    "📊 Movimientos y Saldos",
                    "✅ Checklist de Verificación",
                    "🚩 Análisis de Partidas y Antigüedad",
                    "⚡ Dictamen y Decisión"
                ])

                with tab_aud1:
                    col_tit, col_lg = st.columns([4, 1])
                    with col_tit:
                        st.markdown(f"""
                        <div style="background-color: #1F4E78; color: white; padding: 10px; text-align: center; border-radius: 5px;">
                            CONCILIACIÓN - {tipo_cta.upper()} ({consecutivo_str})
                        </div>
                        """, unsafe_allow_html=True)
                    with col_lg:
                        logo_aud = obtener_logo_empresa(c_data.get("empresa"))
                        if logo_aud:
                            st.image(logo_aud, width=100)

                    col_inf1, col_inf2 = st.columns(2)
                    with col_inf1:
                        st.write(f"🏢 **Empresa:** {c_data.get('empresa', 'N/A')} | **NIT:** {c_data.get('nit', 'N/A')}")
                        st.write(f"📅 **Mes/Año:** {c_data.get('mes', 'N/A')} | **Elaboración:** {c_data.get('fecha_elaboracion', 'N/A')}")
                    with col_inf2:
                        st.write(f"🏦 **Banco:** {c_data.get('banco', 'N/A')} | **Cuenta No:** {c_data.get('cuenta', 'N/A')}")
                        st.write(f"👤 **Preparado por:** {datos.get('preparado_por', 'N/A')}")

                    st.markdown("**SALDOS Y CÁLCULO DE LA CONCILIACIÓN**")
                    df_saldos = pd.DataFrame([
                        {"Concepto": "SALDO SEGÚN EXTRACTO BANCARIO", "Valor": f"$ {c_data.get('saldo_extracto', 0):,.2f}"},
                        {"Concepto": "SALDO SEGÚN LIBROS", "Valor": f"$ {c_data.get('saldo_libros', 0):,.2f}"},
                        {"Concepto": "DIFERENCIA A JUSTIFICAR", "Valor": f"$ {c_data.get('diferencia_inicial', 0):,.2f}"},
                        {"Concepto": "DIFERENCIA CONCILIADA", "Valor": f"$ {c_data.get('diferencia_conciliada', 0):,.2f}"},
                        {"Concepto": "RESULTADO FINAL", "Valor": f"$ {c_data.get('resultado_final', 0):,.2f}"},
                        {"Concepto": "ESTADO", "Valor": c_data.get('estado', 'N/A')}
                    ])
                    st.table(df_saldos)

                    st.markdown(f"**1. {t1_nombre}**")
                    st.dataframe(pd.DataFrame(datos.get("salidas_extracto", [])), use_container_width=True, hide_index=True)
                    st.markdown(f"**2. {t2_nombre}**")
                    st.dataframe(pd.DataFrame(datos.get("salidas_libros", [])), use_container_width=True, hide_index=True)
                    st.markdown(f"**3. {t3_nombre}**")
                    st.dataframe(pd.DataFrame(datos.get("entradas_libros", [])), use_container_width=True, hide_index=True)
                    st.markdown(f"**4. {t4_nombre}**")
                    st.dataframe(pd.DataFrame(datos.get("entradas_extracto", [])), use_container_width=True, hide_index=True)
                    st.markdown("**GASTOS BANCARIOS**")
                    st.dataframe(pd.DataFrame(datos.get("gastos_bancarios", [])), use_container_width=True, hide_index=True)

                with tab_aud2:
                    st.subheader("📋 Lista de Verificación de Control Interno")
                    st.caption("Marca las verificaciones ejecutadas antes de autorizar el cierre:")
                    chk_extracto = st.checkbox("El saldo de extracto bancario coincide exactamente con el PDF oficial.", key=f"chk_ext_{fila['id']}")
                    chk_libros = st.checkbox("El saldo en libros fue verificado contra el auxiliar contable.", key=f"chk_lib_{fila['id']}")
                    chk_gastos = st.checkbox("Los gastos bancarios (4x1000, comisiones, IVA) fueron contabilizados.", key=f"chk_gas_{fila['id']}")
                    chk_soportes = st.checkbox("Las partidas conciliadas poseen soportes documentales válidos.", key=f"chk_sop_{fila['id']}")
                    chk_antiguedad = st.checkbox("No existen partidas pendientes sin justificar mayores a 60 días.", key=f"chk_ant_{fila['id']}")

                with tab_aud3:
                    st.subheader("🚩 Análisis de Partidas Pendientes y Riesgo")
                    df_t1 = pd.DataFrame(datos.get("salidas_extracto", []))
                    df_t2 = pd.DataFrame(datos.get("salidas_libros", []))
                    df_t3 = pd.DataFrame(datos.get("entradas_libros", []))
                    df_t4 = pd.DataFrame(datos.get("entradas_extracto", []))
                    tot_partidas = len(df_t1) + len(df_t2) + len(df_t3) + len(df_t4)
                    st.metric("Total Partidas Conciliadas Registradas", tot_partidas)
                    if tot_partidas > 0:
                        st.markdown("🟢 **Riesgo Bajo:** Partidas del mes en curso.")
                        st.markdown("🟡 **Riesgo Medio:** Partidas con más de 30 días de antigüedad.")
                        st.markdown("🔴 **Riesgo Crítico:** Partidas con más de 60 días de antigüedad.")
                    else:
                        st.caption("No existen partidas pendientes registradas en esta conciliación.")

                with tab_aud4:
                    st.subheader("⚡ Emisión del Dictamen")
                    decision = st.radio("Seleccione Acción:", ["✅ Aprobar Conciliación", "❌ Devolver a Corrección"], key=f"dec_{fila['id']}")
                    if "Aprobar" in decision:
                        if st.button("✅ Confirmar y Emitir Aprobación", key=f"btn_aprob_final_{fila['id']}", type="primary"):
                            checklist_guardar = {
                                "chk_extracto": chk_extracto, "chk_libros": chk_libros,
                                "chk_gastos": chk_gastos, "chk_soportes": chk_soportes, "chk_antiguedad": chk_antiguedad
                            }
                            actualizar_estado_auditoria(fila["id"], "Aprobada", usuario_actual["nombre"], checklist=checklist_guardar)
                            st.success("🎉 Conciliación aprobada con éxito.")
                            st.rerun()
                    else:
                        tipo_hallazgo_sel = st.selectbox(
                            "Categoría Principal del Hallazgo:",
                            [
                                "Diferencia en saldos no justificada",
                                "Falta soporte documental",
                                "Error en clasificación de transacción",
                                "Partidas duplicadas o con fecha errónea",
                                "Gastos bancarios no contabilizados",
                                "Otro hallazgo de auditoría"
                            ],
                            key=f"cat_hallazgo_{fila['id']}"
                        )
                        motivo_det = st.text_area("Detalle de las Observaciones para el Preparador:", key=f"mot_det_{fila['id']}")
                        if st.button("❌ Confirmar y Devolver al Preparador", key=f"btn_dev_final_{fila['id']}"):
                            if not motivo_det.strip():
                                st.error("Debes ingresar el detalle de las observaciones.")
                            else:
                                actualizar_estado_auditoria(fila["id"], "Requiere corrección", usuario_actual["nombre"], motivo_correccion=motivo_det, tipo_hallazgo=tipo_hallazgo_sel)
                                st.warning("⚠️ Conciliación devuelta al preparador.")
                                st.rerun()

elif menu_seleccionado == "📊 Dashboard":
    st.title("📊 DASHBOARD Y CENTRO DE CONTROL FINANCIERO")
    if empresa_activa_nombre != "Todas las empresas":
        st.caption(f"Panel analítico filtrado por empresa: **{empresa_activa_nombre}**")
    else:
        st.caption("Panel analítico consolidado para **Todas las Empresas**")

    historial = obtener_historial(empresa_activa_nombre)
    if historial.empty:
        st.info("ℹ️ Aún no hay datos registrados para generar el análisis analítico.")
    else:
        total_conciliaciones = len(historial)
        exitosas = len(historial[historial["resultado_final"].abs() < 0.005])
        tasa_exito = (exitosas / total_conciliaciones * 100) if total_conciliaciones > 0 else 0.0
        monto_diferencias = historial["resultado_final"].abs().sum()

        aprobadas_cnt = len(historial[historial["workflow_status"] == "Aprobada"])
        pendientes_cnt = len(historial[historial["workflow_status"] == "Pendiente de revisión"])
        borradores_cnt = len(historial[historial["workflow_status"] == "Borrador"])
        devueltas_cnt = len(historial[historial["workflow_status"] == "Requiere corrección"])

        st.subheader("📈 Indicadores Clave de Desempeño (KPIs)")
        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        kpi_col1.metric("Total Conciliaciones", total_conciliaciones)
        kpi_col2.metric("Tasa de Éxito", f"{tasa_exito:.1f}%")
        kpi_col3.metric("Diferencia Total Pendiente", f"$ {monto_diferencias:,.2f}")
        kpi_col4.metric("Pendientes por Auditar", pendientes_cnt)

        st.divider()
        col_wf, col_pie = st.columns([2, 2])
        with col_wf:
            st.markdown("**Estado del Flujo de Auditoría**")
            w1, w2 = st.columns(2)
            w1.metric("Aprobadas", aprobadas_cnt)
            w2.metric("Pendientes de Revisión", pendientes_cnt)
            w3, w4 = st.columns(2)
            w3.metric("Borradores", borradores_cnt)
            w4.metric("Devueltas / Corrección", devueltas_cnt)

        with col_pie:
            st.markdown("**Distribución por Estado de Revisión**")
            df_pie = pd.DataFrame({
                "Estado": ["Aprobada", "Pendiente de revisión", "Borrador", "Requiere corrección"],
                "Cantidad": [aprobadas_cnt, pendientes_cnt, borradores_cnt, devueltas_cnt]
            })
            df_pie = df_pie[df_pie["Cantidad"] > 0]
            if not df_pie.empty:
                fig_pie = px.pie(
                    df_pie, names="Estado", values="Cantidad", color="Estado",
                    color_discrete_map={
                        "Aprobada": "#2ECC71",
                        "Pendiente de revisión": "#F39C12",
                        "Borrador": "#3498DB",
                        "Requiere corrección": "#E74C3C"
                    }, hole=0.4
                )
                fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=200)
                st.plotly_chart(fig_pie, use_container_width=True)

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
                edit_logo = st.file_uploader("Actualizar Logo (Opcional PNG, JPG)", type=["png", "jpg", "jpeg"])
                c_guard, c_canc = st.columns(2)
                with c_guard:
                    if st.form_submit_button("Guardar Cambios", type="primary"):
                        logo_b = edit_logo.getvalue() if edit_logo else None
                        ok, mensaje = actualizar_empresa_db(empresa_edit_id, edit_nom, edit_nit, logo_b)
                        if ok:
                            st.session_state.empresa_a_editar = None
                            st.success(mensaje)
                            st.rerun()
                        else:
                            st.warning(mensaje)
                with c_canc:
                    if st.form_submit_button("❌ Cancelar"):
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
                        ok, mensaje = guardar_empresa(nom, nit, logo_b)
                        if ok:
                            st.success(mensaje)
                            st.rerun()
                        else:
                            st.warning(mensaje)

    st.subheader("Empresas Registradas")
    empresas_list = obtener_empresas()
    if empresas_list.empty:
        st.info("No hay empresas registradas.")
    else:
        for idx, emp in empresas_list.iterrows():
            with st.container():
                col_lg, col_dt, col_act1, col_act2 = st.columns([1, 4, 1.5, 1.5])
                with col_lg:
                    if emp["logo"]:
                        st.image(emp["logo"], width=70)
                    else:
                        st.caption("Sin logo")
                with col_dt:
                    st.write(f"🏢 **{emp['nombre']}**")
                    st.write(f"📄 NIT: {emp['nit']}")
                with col_act1:
                    if rol_actual == "Administrador":
                        if st.button("✏️ Editar", key=f"btn_edit_emp_{emp['id']}"):
                            st.session_state.empresa_a_editar = emp['id']
                            st.rerun()
                with col_act2:
                    if rol_actual == "Administrador":
                        if st.button("🗑️ Eliminar", key=f"btn_del_emp_{emp['id']}"):
                            ok, mensaje = eliminar_empresa_db(emp['id'])
                            if ok:
                                st.success(mensaje)
                                st.rerun()
                            else:
                                st.warning(mensaje)
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
                    emp_id = st.selectbox("Asociar a Empresa", empresas_df["id"].tolist(), format_func=lambda x: empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0])
                if st.form_submit_button("Guardar Cuenta", type="primary"):
                    guardar_cuenta(banco, num, tipo, emp_id)
                    st.success("Cuenta guardada exitosamente.")
                    st.rerun()

    st.dataframe(obtener_cuentas(empresa_activa_id), use_container_width=True, hide_index=True)

elif menu_seleccionado == "📝 Nueva Conciliación":
    id_edicion = st.session_state.get("conciliacion_a_editar", None)
    modo_correccion = bool(st.session_state.get("modo_correccion", False)) and bool(id_edicion)
    c_edit = obtener_conciliacion_por_id(id_edicion) if id_edicion else None

    if modo_correccion:
        st.title("🔧 Corrección de Conciliación Devuelta")
        consecutivo_edit_str = f"CONC-{int(id_edicion):06d}"
        st.error(
            f"🔴 **{consecutivo_edit_str} fue devuelta por Auditoría y requiere corrección.**"
        )
        if c_edit:
            st.warning(
                f"**Observaciones del Auditor:** {c_edit.get('motivo_correccion') or 'Sin detalle registrado.'}\n\n"
                f"**Categoría:** {c_edit.get('tipo_hallazgo') or 'General'}\n\n"
                f"**Revisado por:** {c_edit.get('revisado_por_usuario') or 'N/A'}"
            )
        st.caption("Corrige los datos necesarios y utiliza **Enviar Corrección a Revisión** cuando hayas terminado.")
    else:
        st.title("📝 Captura / Edición de Conciliación Bancaria")

    if id_edicion and not modo_correccion:
        consecutivo_edit_str = f"CONC-{int(id_edicion):06d}"
        st.info(f"✏️ **Modo Edición Activado:** Editando Conciliación {consecutivo_edit_str} ({c_edit.get('empresa')} - {c_edit.get('banco')})")

    if id_edicion:
        if st.button("❌ Cancelar y volver al Historial"):
            st.session_state.conciliacion_a_editar = None
            st.session_state.modo_correccion = False
            st.session_state.pop("datos_cargados_edit", None)
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
                cuentas_asig_df = obtener_cuentas_rotadas_por_usuario(empresa_activa_id, mes_num, usuario_actual["id"])
                if not cuentas_asig_df.empty:
                    st.info("ℹ️ **Cuentas asignadas para tu perfil este mes:**")
                    cta_sel = st.selectbox(
                        "Cuenta / Tarjeta Registrada",
                        cuentas_asig_df["id"].tolist(),
                        format_func=lambda x: f"{cuentas_asig_df.loc[cuentas_asig_df['id']==x, 'banco'].values[0]} - {cuentas_asig_df.loc[cuentas_asig_df['id']==x, 'numero_cuenta'].values[0]}"
                    )
                    banco = cuentas_asig_df.loc[cuentas_asig_df["id"] == cta_sel, "banco"].values[0]
                    cuenta = cuentas_asig_df.loc[cuentas_asig_df["id"] == cta_sel, "numero_cuenta"].values[0]
                    tipo = cuentas_asig_df.loc[cuentas_asig_df["id"] == cta_sel, "tipo_cuenta"].values[0]
                else:
                    st.warning("No hay cuentas asignadas o registradas para preparadores activos en esta empresa.")
                    banco = st.text_input("Nombre del Banco", key="form_banco")
                    cuenta = st.text_input("Número de Cuenta", key="form_cuenta")
                    tipo = st.selectbox("Tipo de Cuenta", ["Cuenta de ahorros", "Cuenta corriente", "Tarjeta de crédito"])

    with col_logo_emp:
        logo_bytes = obtener_logo_empresa(empresa)
        if logo_bytes:
            st.image(logo_bytes, caption=f"Logo {empresa}", width=150)

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
    st.metric("Diferencia a Justificar", f"$ {diferencia_inicial:,.2f}")

    st.divider()
    st.subheader(f"1. {nombres_titulos['t1']}")
    salidas_extracto = st.data_editor(st.session_state.tabla1, num_rows="dynamic", use_container_width=True, key="editor_tabla1")
    cols_t1 = ["Fecha", "Beneficiario", "Documento", "Valor"]
    c_p1, c_b1, c_del1 = st.columns([3, 1, 1])
    with c_p1:
        txt_t1 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Beneficiario | Documento | Valor):", key="paste_t1")
    with c_b1:
        if st.button("+ Cargar Ítems 1", key="btn_parse_t1"):
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
    salidas_libros = st.data_editor(st.session_state.tabla2, num_rows="dynamic", use_container_width=True, key="editor_tabla2")
    cols_t2 = ["Fecha", "Concepto", "Valor"]
    c_p2, c_b2, c_del2 = st.columns([3, 1, 1])
    with c_p2:
        txt_t2 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Concepto | Valor):", key="paste_t2")
    with c_b2:
        if st.button("+ Cargar Ítems 2", key="btn_parse_t2"):
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
    entradas_libros = st.data_editor(st.session_state.tabla3, num_rows="dynamic", use_container_width=True, key="editor_tabla3")
    cols_t3 = ["Fecha", "Concepto", "Valor"]
    c_p3, c_b3, c_del3 = st.columns([3, 1, 1])
    with c_p3:
        txt_t3 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Concepto | Valor):", key="paste_t3")
    with c_b3:
        if st.button("+ Cargar Ítems 3", key="btn_parse_t3"):
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
    entradas_extracto = st.data_editor(st.session_state.tabla4, num_rows="dynamic", use_container_width=True, key="editor_tabla4")
    cols_t4 = ["Fecha", "Concepto", "Valor"]
    c_p4, c_b4, c_del4 = st.columns([3, 1, 1])
    with c_p4:
        txt_t4 = st.text_input("📋 Pegar ítems desde Excel (Fecha | Concepto | Valor):", key="paste_t4")
    with c_b4:
        if st.button("+ Cargar Ítems 4", key="btn_parse_t4"):
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
    gastos_bancarios = st.data_editor(st.session_state.tabla5, num_rows="dynamic", use_container_width=True, key="editor_tabla5")
    cols_t5 = ["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"]
    c_p5, c_b5, c_del5 = st.columns([3, 1, 1])
    with c_p5:
        txt_t5 = st.text_input("📋 Pegar ítems desde Excel para Gastos Bancarios:", key="paste_t5")
    with c_b5:
        if st.button("+ Cargar Ítems 5", key="btn_parse_t5"):
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
        diferencia_conciliada = m1 + m2 - m3 + m4
    else:
        diferencia_conciliada = m1 - m2 + m3 - m4

    resultado_final = diferencia_inicial - diferencia_conciliada

    st.divider()
    st.subheader("Resultado de la Conciliación")
    r1, r2, r3 = st.columns(3)
    r1.metric("Diferencia Inicial", f"$ {diferencia_inicial:,.2f}")
    r2.metric("Diferencia Conciliada", f"$ {diferencia_conciliada:,.2f}")
    r3.metric("Resultado Final", f"$ {resultado_final:,.2f}")

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
        texto_guardar = "💾 Guardar Corrección" if modo_correccion else "💾 Guardar como Borrador"
        if st.button(texto_guardar, key="btn_guardar_borrador"):
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data,
                workflow_status="Borrador", id_edicion=id_edicion
            )
            st.session_state.pop("conciliacion_a_editar", None)
            st.session_state.pop("datos_cargados_edit", None)
            st.session_state.modo_correccion = False
            st.success(f"💾 Conciliación CONC-{int(id_g):06d} guardada como Borrador.")

    with btn_col_env:
        texto_enviar = "🚀 Enviar Corrección a Revisión" if modo_correccion else "🚀 Enviar a Revisión"
        if st.button(texto_enviar, key="btn_enviar_revision", type="primary"):
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data,
                workflow_status="Pendiente de revisión", id_edicion=id_edicion
            )
            st.session_state.pop("conciliacion_a_editar", None)
            st.session_state.pop("datos_cargados_edit", None)
            st.session_state.modo_correccion = False

            # Si el usuario puede auditar, lo llevamos directamente a la bandeja
            # para que vea la conciliación recién enviada sin tener que buscarla.
            if rol_actual in ["Administrador", "Revisor"]:
                st.session_state.menu_override = "🔍 Auditoría y Revisiones"
                st.rerun()
            else:
                st.success(
                    f"🚀 Conciliación CONC-{int(id_g):06d} enviada correctamente a Revisión. "
                    "Ahora queda disponible para un usuario con rol Revisor o Administrador."
                )

elif menu_seleccionado == "📋 Historial":
    st.title("📋 Historial de Conciliaciones")
    if empresa_activa_nombre != "Todas las empresas":
        st.caption(f"Mostrando conciliaciones de: **{empresa_activa_nombre}**")

    historial = obtener_historial(empresa_activa_nombre)
    filtro_estado = st.selectbox("Filtrar por Estado de Revisión:", ["Todos los Estados", "Borrador", "Pendiente de revisión", "Aprobada", "Requiere corrección"])

    if filtro_estado != "Todos los Estados" and not historial.empty:
        historial = historial[historial["workflow_status"] == filtro_estado]

    if historial.empty:
        st.info("No hay conciliaciones guardadas para el filtro seleccionado.")
    else:
        for idx, fila in historial.iterrows():
            cuenta_txt = fila.get("cuenta") or "N/A"
            banco_txt = fila.get("banco") or "N/A"
            wf_status = fila.get("workflow_status", "Pendiente de revisión")
            consecutivo_str = f"CONC-{int(fila['id']):06d}"

            with st.expander(f"📌 {consecutivo_str} | {fila['empresa']} - {banco_txt} ({cuenta_txt}) | Mes: {fila['mes']} - Estado: {wf_status}"):
                c_data = obtener_conciliacion_por_id(fila["id"])
                try:
                    datos = json.loads(c_data.get("datos_json", "{}"))
                except Exception:
                    datos = {}

                if wf_status == "Requiere corrección" and c_data.get("motivo_correccion"):
                    st.error(f"⚠️️ **Observaciones del Auditor ({c_data.get('revisado_por_usuario', 'N/A')}):**\n"
                             f"• **Categoría:** {c_data.get('tipo_hallazgo', 'General')}\n"
                             f"• **Detalle:** {c_data.get('motivo_correccion')}")

                tipo_cta = c_data.get("tipo") or "Cuenta de ahorros"
                es_tc = "tarjeta" in tipo_cta.lower() or "crédito" in tipo_cta.lower() or "credito" in tipo_cta.lower()

                if es_tc:
                    t1_nombre, t2_nombre, t3_nombre, t4_nombre = (
                        "COMPRAS NO EVIDENCIADAS EN EXTRACTOS", "COMPRAS NO CONTABILIZADAS EN LIBROS",
                        "DÉBITOS BANCARIOS NO CONTABILIZADOS EN LIBROS", "ABONOS NO REGISTRADOS EN EXTRACTO"
                    )
                else:
                    t1_nombre, t2_nombre, t3_nombre, t4_nombre = (
                        "SALIDAS NO REGISTRADAS EN EXTRACTO", "SALIDAS BANCARIAS NO CONTABILIZADAS EN LIBROS",
                        "ENTRADAS BANCARIAS NO CONTABILIZADAS EN LIBROS", "ENTRADAS NO EVIDENCIADAS EN EXTRACTOS"
                    )

                c_act1, c_act2, c_act3 = st.columns(3)
                with c_act1:
                    if wf_status == "Requiere corrección" and rol_actual in ["Preparador", "Administrador"]:
                        if st.button("🔧 Corregir Conciliación", key=f"btn_corregir_{fila['id']}", type="primary"):
                            st.session_state.conciliacion_a_editar = fila['id']
                            st.session_state.modo_correccion = True
                            st.session_state.pop("datos_cargados_edit", None)
                            st.session_state.menu_override = "📝 Nueva Conciliación"
                            st.rerun()
                    elif wf_status != "Aprobada":
                        if st.button("✏️ Editar Conciliación", key=f"btn_edit_{fila['id']}"):
                            st.session_state.conciliacion_a_editar = fila['id']
                            st.session_state.modo_correccion = False
                            st.session_state.pop("datos_cargados_edit", None)
                            st.session_state.menu_override = "📝 Nueva Conciliación"
                            st.rerun()

                # Permite enviar un Borrador directamente desde el Historial.
                if wf_status == "Borrador" and rol_actual in ["Preparador", "Administrador"]:
                    if st.button("🚀 Enviar a Revisión", key=f"btn_hist_enviar_{fila['id']}", type="primary"):
                        actualizar_estado_auditoria(
                            fila["id"],
                            "Pendiente de revisión",
                            usuario_actual["nombre"]
                        )
                        st.success(
                            f"🚀 {consecutivo_str} fue enviado correctamente a Revisión."
                        )
                        st.rerun()

                with c_act2:
                    if c_data.get("excel"):
                        nombre_ex = f"CONCILIACION_{consecutivo_str}_{limpiar_nombre_archivo(c_data.get('empresa'))}.xlsx"
                        st.download_button("📊 Descargar Excel", data=bytes(c_data["excel"]), file_name=nombre_ex)

                with c_act3:
                    nombres_titulos = {"t1": t1_nombre, "t2": t2_nombre, "t3": t3_nombre, "t4": t4_nombre}
                    pdf_bytes = generar_pdf_conciliacion(c_data, datos, nombres_titulos)
                    nombre_pdf = f"CONCILIACION_{consecutivo_str}_{limpiar_nombre_archivo(c_data.get('empresa'))}.pdf"
                    st.download_button("📄 Descargar PDF", data=pdf_bytes, file_name=nombre_pdf, mime="application/pdf")

elif menu_seleccionado == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios y Roles")
    usuarios_df = obtener_usuarios()
    st.subheader("Lista de Usuarios Registrados")
    st.dataframe(usuarios_df, use_container_width=True, hide_index=True)

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
                            format_func=lambda x: "Sin asignar (Todas)" if x is None else empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0]
                        )
                    if st.form_submit_button("Crear Usuario", type="primary"):
                        if not nuevo_nombre.strip() or not nuevo_usuario.strip() or not nueva_pass:
                            st.error("Completa todos los campos obligatorios.")
                        elif len(nueva_pass) < 8:
                            st.error("La contraseña debe tener al menos 8 caracteres.")
                        else:
                            try:
                                crear_usuario(nuevo_usuario, nuevo_nombre, nueva_pass, nuevo_rol, emp_id_crear)
                                st.success(f"Usuario '{nuevo_usuario}' creado con éxito con rol {nuevo_rol}.")
                                st.rerun()
                            except Exception:
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
                        format_func=lambda x: "Sin asignar" if x is None else empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0],
                        key="link_empresa"
                    )
                url_base = st.query_params.get("base_url", "https://conciliacionweb.streamlit.app/")
                link_generado = f"{url_base}?registro=true&rol={rol_link}"
                if emp_id_link:
                    link_generado += f"&empresa_id={emp_id_link}"
                st.code(link_generado, language="text")

        if not usuarios_df.empty:
            st.divider()
            c_rol, c_emp = st.columns(2)
            with c_rol:
                st.subheader("🔄 Cambiar Rol de Usuario")
                with st.form("form_cambiar_rol"):
                    usr_sel_r = st.selectbox("Usuario", usuarios_df["id"].tolist(), format_func=lambda x: f"{usuarios_df.loc[usuarios_df['id'] == x, 'nombre'].values[0]} ({usuarios_df.loc[usuarios_df['id'] == x, 'usuario'].values[0]})")
                    nuevo_rol_sel = st.selectbox("Nuevo Rol", ["Preparador", "Revisor", "Administrador"])
                    if st.form_submit_button("Actualizar Rol", type="primary"):
                        actualizar_rol_usuario(usr_sel_r, nuevo_rol_sel)
                        st.success("Rol actualizado con éxito.")
                        st.rerun()

            with c_emp:
                if not empresas_df.empty:
                    st.subheader("🏢 Cambiar Empresa Asignada")
                    with st.form("form_asignar_empresa"):
                        usr_sel_e = st.selectbox("Usuario", usuarios_df["id"].tolist(), format_func=lambda x: f"{usuarios_df.loc[usuarios_df['id'] == x, 'nombre'].values[0]} ({usuarios_df.loc[usuarios_df['id'] == x, 'usuario'].values[0]})")
                        emp_sel = st.selectbox("Empresa", empresas_df["id"].tolist(), format_func=lambda x: empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0])
                        if st.form_submit_button("Asignar Empresa", type="primary"):
                            actualizar_empresa_usuario(usr_sel_e, emp_sel)
                            st.success("Empresa asignada correctamente.")
                            st.rerun()

elif menu_seleccionado == "📄 Reportes":
    st.title("📄 Reportes y Descargas")
    historial = obtener_historial(empresa_activa_nombre)
    if not historial.empty:
        csv = historial.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Historial Completo (CSV)", data=csv, file_name=f"historial_{limpiar_nombre_archivo(empresa_activa_nombre)}.csv", mime="text/csv")
    else:
        st.info("No hay información registrada para generar reportes.")
