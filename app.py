import streamlit as st
import sqlite3
import pandas as pd
import datetime
import io
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Conciliación Bancaria y Auditoría",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# ADAPTADOR COMPATIBLE PARA TURSO / SQLITE
# ==========================================
class TursoCursorWrapper:
    """Adaptador para que libsql_client responda exactamente igual que un cursor de SQLite."""
    def __init__(self, client):
        self.client = client
        self.lastrowid = None

    def execute(self, query, params=()):
        if params and not isinstance(params, (list, tuple)):
            params = (params,)
        
        res = self.client.execute(query, list(params) if params else [])
        self._last_result = res
        try:
            if res.last_insert_rowid is not None:
                self.lastrowid = res.last_insert_rowid
        except Exception:
            pass
        return self

    def fetchone(self):
        try:
            rows = self._last_result.rows
            if rows:
                return rows[0]
        except Exception:
            pass
        return None

    def fetchall(self):
        try:
            return self._last_result.rows
        except Exception:
            return []

class TursoConnectionWrapper:
    """Adaptador de conexión para libsql_client."""
    def __init__(self, client):
        self.client = client

    def cursor(self):
        return TursoCursorWrapper(self.client)

    def commit(self):
        pass

    def close(self):
        try:
            self.client.close()
        except Exception:
            pass

def conectar_db():
    """
    Conecta a la base de datos persistente en Turso usando libsql-client,
    asegurando que la URL utilice el protocolo HTTPS correcto.
    """
    if "TURSO_DATABASE_URL" in st.secrets and "TURSO_AUTH_TOKEN" in st.secrets:
        url = st.secrets["TURSO_DATABASE_URL"]
        token = st.secrets["TURSO_AUTH_TOKEN"]
        
        if url.startswith("libsql://"):
            url = url.replace("libsql://", "https://")
        
        try:
            import libsql_client
            client = libsql_client.create_client_sync(url=url, auth_token=token)
            return TursoConnectionWrapper(client)
        except Exception as e:
            st.warning(f"⚠️ Error conectando a Turso: {e}. Usando respaldo local.")
            
    return sqlite3.connect("conciliaciones.db")

def inicializar_bd():
    try:
        conn = conectar_db()
        c = conn.cursor()
        
        # Tabla Usuarios
        c.execute('''
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                nombre TEXT NOT NULL,
                password TEXT NOT NULL,
                rol TEXT NOT NULL,
                empresa_id INTEGER
            )
        ''')
        
        # Tabla Empresas
        c.execute('''
            CREATE TABLE IF NOT EXISTS empresas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nit TEXT UNIQUE NOT NULL,
                razon_social TEXT NOT NULL
            )
        ''')
        
        # Tabla Cuentas Bancarias
        c.execute('''
            CREATE TABLE IF NOT EXISTS cuentas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                empresa_id INTEGER NOT NULL,
                banco TEXT NOT NULL,
                numero_cuenta TEXT NOT NULL,
                tipo_cuenta TEXT NOT NULL,
                FOREIGN KEY (empresa_id) REFERENCES empresas (id)
            )
        ''')
        
        # Tabla Conciliaciones
        c.execute('''
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
                FOREIGN KEY (empresa_id) REFERENCES empresas (id),
                FOREIGN KEY (cuenta_id) REFERENCES cuentas (id)
            )
        ''')
        
        # Tabla Partidas Conciliatorias
        c.execute('''
            CREATE TABLE IF NOT EXISTS partidas_conciliatorias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conciliacion_id INTEGER NOT NULL,
                tipo TEXT NOT NULL,
                fecha TEXT,
                concepto TEXT,
                monto REAL NOT NULL,
                clasificacion TEXT,
                antiguedad_dias INTEGER,
                FOREIGN KEY (conciliacion_id) REFERENCES conciliaciones (id)
            )
        ''')
        
        # Tabla Bitácora de Auditoría
        c.execute('''
            CREATE TABLE IF NOT EXISTS bitacora_auditoria (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conciliacion_id INTEGER NOT NULL,
                usuario TEXT NOT NULL,
                accion TEXT NOT NULL,
                fecha_hora TEXT NOT NULL,
                comentario TEXT,
                FOREIGN KEY (conciliacion_id) REFERENCES conciliaciones (id)
            )
        ''')
        
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Error al inicializar la base de datos: {e}")

# Inicializar tablas en la BD
inicializar_bd()

# ==========================================
# FUNCIONES DE CONSULTA
# ==========================================
def contar_usuarios():
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM usuarios")
        res = c.fetchone()
        total = res[0] if res else 0
        conn.close()
        return total
    except Exception:
        return 0

def verificar_credenciales(username, password):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, username, nombre, rol, empresa_id FROM usuarios WHERE username = ? AND password = ?", (username, password))
        user = c.fetchone()
        conn.close()
        return user
    except Exception:
        return None

def registrar_usuario(username, nombre, password, rol, empresa_id=None):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("INSERT INTO usuarios (username, nombre, password, rol, empresa_id) VALUES (?, ?, ?, ?, ?)",
                  (username, nombre, password, rol, empresa_id))
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False

def obtener_empresas():
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, nit, razon_social FROM empresas")
        rows = c.fetchall()
        conn.close()
        if rows:
            return pd.DataFrame(rows, columns=['id', 'nit', 'razon_social'])
    except Exception:
        pass
    return pd.DataFrame(columns=['id', 'nit', 'razon_social'])

def registrar_empresa(nit, razon_social):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("INSERT INTO empresas (nit, razon_social) VALUES (?, ?)", (nit, razon_social))
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False

def registrar_cuenta(empresa_id, banco, numero_cuenta, tipo_cuenta):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("INSERT INTO cuentas (empresa_id, banco, numero_cuenta, tipo_cuenta) VALUES (?, ?, ?, ?)",
                  (empresa_id, banco, numero_cuenta, tipo_cuenta))
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Error al registrar la cuenta: {e}")

def obtener_cuentas_empresa(empresa_id):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, empresa_id, banco, numero_cuenta, tipo_cuenta FROM cuentas WHERE empresa_id = ?", (empresa_id,))
        rows = c.fetchall()
        conn.close()
        if rows:
            return pd.DataFrame(rows, columns=['id', 'empresa_id', 'banco', 'numero_cuenta', 'tipo_cuenta'])
    except Exception:
        pass
    return pd.DataFrame(columns=['id', 'empresa_id', 'banco', 'numero_cuenta', 'tipo_cuenta'])

def registrar_conciliacion(empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, preparado_por, df_partidas):
    try:
        conn = conectar_db()
        c = conn.cursor()
        fecha_hoy = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        c.execute('''
            INSERT INTO conciliaciones (empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, estado, preparado_por, fecha_creacion)
            VALUES (?, ?, ?, ?, ?, 'Pendiente', ?, ?)
        ''', (empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, preparado_por, fecha_hoy))
        
        conciliacion_id = c.lastrowid
        
        for _, row in df_partidas.iterrows():
            c.execute('''
                INSERT INTO partidas_conciliatorias (conciliacion_id, tipo, fecha, concepto, monto, clasificacion, antiguedad_dias)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (conciliacion_id, row['tipo'], str(row['fecha']), row['concepto'], float(row['monto']), row.get('clasificacion', 'General'), int(row.get('antiguedad_dias', 0))))
            
        c.execute('''
            INSERT INTO bitacora_auditoria (conciliacion_id, usuario, accion, fecha_hora, comentario)
            VALUES (?, ?, 'Creada y Enviada a Revisión', ?, 'Conciliación registrada')
        ''', (conciliacion_id, preparado_por, fecha_hoy))
        
        conn.commit()
        conn.close()
        return conciliacion_id
    except Exception as e:
        st.error(f"Error al guardar conciliación: {e}")
        return None

def obtener_conciliaciones(empresa_id=None):
    try:
        conn = conectar_db()
        c = conn.cursor()
        query = '''
            SELECT c.id, e.razon_social as empresa, b.banco, b.numero_cuenta, c.periodo, 
                   c.saldo_libro, c.saldo_banco, c.estado, c.preparado_por, c.revisado_por, c.fecha_creacion,
                   c.dictamen, c.observaciones
            FROM conciliaciones c
            JOIN empresas e ON c.empresa_id = e.id
            JOIN cuentas b ON c.cuenta_id = b.id
        '''
        if empresa_id:
            query += f" WHERE c.empresa_id = {empresa_id}"
        query += " ORDER BY c.id DESC"
        c.execute(query)
        rows = c.fetchall()
        conn.close()
        if rows:
            return pd.DataFrame(rows, columns=['id', 'empresa', 'banco', 'numero_cuenta', 'periodo', 'saldo_libro', 'saldo_banco', 'estado', 'preparado_por', 'revisado_por', 'fecha_creacion', 'dictamen', 'observaciones'])
    except Exception:
        pass
    return pd.DataFrame()

def actualizar_dictamen_conciliacion(conciliacion_id, usuario, nuevo_estado, dictamen, observaciones):
    try:
        conn = conectar_db()
        c = conn.cursor()
        fecha_hoy = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        c.execute('''
            UPDATE conciliaciones
            SET estado = ?, revisado_por = ?, dictamen = ?, observaciones = ?
            WHERE id = ?
        ''', (nuevo_estado, usuario, dictamen, observaciones, conciliacion_id))
        
        c.execute('''
            INSERT INTO bitacora_auditoria (conciliacion_id, usuario, accion, fecha_hora, comentario)
            VALUES (?, ?, ?, ?, ?)
        ''', (conciliacion_id, usuario, f"Dictamen: {nuevo_estado}", fecha_hoy, f"{dictamen} - {observaciones}"))
        
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Error al actualizar dictamen: {e}")

def obtener_partidas(conciliacion_id):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, conciliacion_id, tipo, fecha, concepto, monto, clasificacion, antiguedad_dias FROM partidas_conciliatorias WHERE conciliacion_id = ?", (conciliacion_id,))
        rows = c.fetchall()
        conn.close()
        if rows:
            return pd.DataFrame(rows, columns=['id', 'conciliacion_id', 'tipo', 'fecha', 'concepto', 'monto', 'clasificacion', 'antiguedad_dias'])
    except Exception:
        pass
    return pd.DataFrame()

def obtener_bitacora(conciliacion_id):
    try:
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, conciliacion_id, usuario, accion, fecha_hora, comentario FROM bitacora_auditoria WHERE conciliacion_id = ? ORDER BY id DESC", (conciliacion_id,))
        rows = c.fetchall()
        conn.close()
        if rows:
            return pd.DataFrame(rows, columns=['id', 'conciliacion_id', 'usuario', 'accion', 'fecha_hora', 'comentario'])
    except Exception:
        pass
    return pd.DataFrame()

# ==========================================
# GENERADOR DE REPORTE PDF
# ==========================================
def generar_pdf_conciliacion(conciliacion_info, df_partidas, df_bitacora):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, alignment=1, textColor=colors.HexColor('#1E3A8A'))
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Heading2'], fontSize=12, leading=16, textColor=colors.HexColor('#1F2937'))
    
    story.append(Paragraph(f"INFORME DE CONCILIACIÓN BANCARIA Y AUDITORÍA", title_style))
    story.append(Paragraph(f"Empresa: {conciliacion_info['empresa']} | Periodo: {conciliacion_info['periodo']}", subtitle_style))
    story.append(Spacer(1, 15))
    
    data_resumen = [
        [Paragraph("<b>Banco:</b>", styles['Normal']), conciliacion_info['banco'], Paragraph("<b>Cuenta:</b>", styles['Normal']), conciliacion_info['numero_cuenta']],
        [Paragraph("<b>Saldo Libros:</b>", styles['Normal']), f"${conciliacion_info['saldo_libro']:,.2f}", Paragraph("<b>Saldo Banco:</b>", styles['Normal']), f"${conciliacion_info['saldo_banco']:,.2f}"],
        [Paragraph("<b>Estado:</b>", styles['Normal']), conciliacion_info['estado'], Paragraph("<b>Preparado por:</b>", styles['Normal']), conciliacion_info['preparado_por']],
        [Paragraph("<b>Revisado por:</b>", styles['Normal']), conciliacion_info['revisado_por'] or "Pendiente", Paragraph("<b>Dictamen:</b>", styles['Normal']), conciliacion_info['dictamen'] or "N/A"]
    ]
    t_resumen = Table(data_resumen, colWidths=[100, 150, 100, 150])
    t_resumen.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F3F4F6')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(t_resumen)
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("Detalle de Partidas Conciliatorias", subtitle_style))
    story.append(Spacer(1, 5))
    
    data_partidas = [["Tipo", "Fecha", "Concepto", "Monto", "Clasificación"]]
    if not df_partidas.empty:
        for _, r in df_partidas.iterrows():
            data_partidas.append([r['tipo'], r['fecha'], r['concepto'], f"${r['monto']:,.2f}", r.get('clasificacion', 'General')])
    else:
        data_partidas.append(["N/A", "N/A", "Sin partidas registradas", "$0.00", "N/A"])
        
    t_partidas = Table(data_partidas, colWidths=[90, 70, 180, 80, 80])
    t_partidas.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(t_partidas)
    story.append(Spacer(1, 15))
    
    if conciliacion_info['observaciones']:
        story.append(Paragraph("Observaciones del Auditor", subtitle_style))
        story.append(Paragraph(conciliacion_info['observaciones'], styles['Normal']))
        story.append(Spacer(1, 15))
        
    doc.build(story)
    buffer.seek(0)
    return buffer

# ==========================================
# GESTIÓN DE SESIÓN Y VISTAS
# ==========================================
if 'usuario' not in st.session_state:
    st.session_state.usuario = None

with st.sidebar.expander("🛠️ Acceso de Emergencia (Admin)"):
    if st.button("Crear Admin de Respaldo"):
        if registrar_usuario("admin_emergencia", "Administrador Principal", "123456", "Administrador"):
            st.sidebar.success("Usuario creado: `admin_emergencia` / `123456`")
        else:
            st.sidebar.info("El usuario `admin_emergencia` ya existe. Úsalo con clave `123456`.")

if contar_usuarios() == 0:
    st.title("🔐 Configuración Inicial")
    st.info("Crea el primer usuario con rol Administrador para comenzar.")
    
    with st.form("form_init"):
        u_user = st.text_input("Usuario")
        u_nombre = st.text_input("Nombre completo")
        u_pass = st.text_input("Contraseña", type="password")
        u_pass2 = st.text_input("Confirmar contraseña", type="password")
        btn_init = st.form_submit_button("Crear Administrador")
        
        if btn_init:
            if not u_user or not u_pass or not u_nombre:
                st.error("Por favor completa todos los campos.")
            elif u_pass != u_pass2:
                st.error("Las contraseñas no coinciden.")
            else:
                if registrar_usuario(u_user, u_nombre, u_pass, "Administrador"):
                    st.success("Administrador creado con éxito. Ya puedes iniciar sesión.")
                    st.rerun()
                else:
                    st.error("El usuario ya existe.")
    st.stop()

if not st.session_state.usuario:
    st.title("⚖️ Sistema de Conciliación Bancaria y Auditoría")
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Iniciar Sesión")
        with st.form("form_login"):
            l_user = st.text_input("Usuario")
            l_pass = st.text_input("Contraseña", type="password")
            btn_login = st.form_submit_button("Ingresar")
            
            if btn_login:
                user = verificar_credenciales(l_user, l_pass)
                if user:
                    st.session_state.usuario = {
                        "id": user[0],
                        "username": user[1],
                        "nombre": user[2],
                        "rol": user[3],
                        "empresa_id": user[4]
                    }
                    st.success(f"Bienvenido {user[2]}")
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
    st.stop()

user_curr = st.session_state.usuario
st.sidebar.title(f"👤 {user_curr['nombre']}")
st.sidebar.caption(f"Rol: {user_curr['rol']}")

if st.sidebar.button("Cerrar Sesión"):
    st.session_state.usuario = None
    st.rerun()

st.sidebar.divider()

opciones_menu = ["📊 Dashboard", "📝 Nueva Conciliación", "🔍 Auditoría y Revisiones"]
if user_curr['rol'] == "Administrador":
    opciones_menu.extend(["🏢 Gestión de Empresas", "👥 Usuarios"])

menu = st.sidebar.radio("Navegación", opciones_menu)

# ==========================================
# MÓDULOS DE LA APLICACIÓN
# ==========================================
if menu == "📊 Dashboard":
    st.title("📊 Dashboard de Control y Auditoría")
    df_conc = obtener_conciliaciones(user_curr['empresa_id'] if user_curr['rol'] != "Administrador" else None)
    
    if df_conc.empty:
        st.info("No hay conciliaciones registradas en el sistema.")
    else:
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Conciliaciones", len(df_conc))
        kpi2.metric("Aprobadas", len(df_conc[df_conc['estado'] == 'Aprobada']))
        kpi3.metric("Pendientes", len(df_conc[df_conc['estado'] == 'Pendiente']))
        kpi4.metric("Requieren Corrección", len(df_conc[df_conc['estado'] == 'Requiere Corrección']))
        
        st.divider()
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            fig_pie = px.pie(df_conc, names='estado', title='Distribución por Estado de Conciliación',
                             color='estado', color_discrete_map={
                                 'Aprobada': '#10B981',
                                 'Pendiente': '#F59E0B',
                                 'Requiere Corrección': '#EF4444'
                             })
            st.plotly_chart(fig_pie, use_container_width=True)
        with col_chart2:
            fig_bar = px.bar(df_conc, x='periodo', y=['saldo_libro', 'saldo_banco'],
                             barmode='group', title='Comparativa Saldo Libros vs Saldo Bancos')
            st.plotly_chart(fig_bar, use_container_width=True)

elif menu == "📝 Nueva Conciliación":
    st.title("📝 Registrar Nueva Conciliación Bancaria")
    df_emp = obtener_empresas()
    if df_emp.empty:
        st.warning("Debe registrar al menos una empresa antes de conciliar.")
        st.stop()
        
    emp_dict = dict(zip(df_emp['razon_social'], df_emp['id']))
    emp_sel = st.selectbox("Empresa", list(emp_dict.keys()))
    emp_id = emp_dict[emp_sel]
    
    df_cuentas = obtener_cuentas_empresa(emp_id)
    if df_cuentas.empty:
        st.warning("Esta empresa no tiene cuentas bancarias registradas.")
        st.stop()
        
    cta_dict = {f"{r['banco']} - {r['numero_cuenta']} ({r['tipo_cuenta']})": r['id'] for _, r in df_cuentas.iterrows()}
    cta_sel = st.selectbox("Cuenta Bancaria", list(cta_dict.keys()))
    cta_id = cta_dict[cta_sel]
    
    col_a, col_b, col_c = st.columns(3)
    periodo = col_a.text_input("Periodo (Ej: 2026-03)", datetime.datetime.now().strftime("%Y-%m"))
    s_libro = col_b.number_input("Saldo según Libros ($)", value=0.0, step=1000.0)
    s_banco = col_c.number_input("Saldo según Extracto ($)", value=0.0, step=1000.0)
    
    st.subheader("Partidas Conciliatorias (Diferencias)")
    st.caption("Agrega Notas Débito/Crédito y Cheques en Tránsito o Consignaciones pendientes.")
    
    if 'partidas_temp' not in st.session_state:
        st.session_state.partidas_temp = pd.DataFrame(columns=['tipo', 'fecha', 'concepto', 'monto', 'clasificacion', 'antiguedad_dias'])
        
    with st.form("form_partida"):
        c1, c2, c3, c4 = st.columns([2, 2, 3, 2])
        p_tipo = c1.selectbox("Tipo de Partida", [
            "ND_LIBROS (Nota Débito Libros no Banco)",
            "NC_LIBROS (Nota Crédito Libros no Banco)",
            "ND_BANCO (Debito Banco no Libros)",
            "NC_BANCO (Credito Banco no Libros)"
        ])
        p_fecha = c2.date_input("Fecha Transacción")
        p_concepto = c3.text_input("Concepto / Descripción")
        p_monto = c4.number_input("Monto ($)", min_value=0.0, step=100.0)
        p_clasif = st.selectbox("Clasificación de Riesgo", ["Normal", "Gasto Bancario", "Partida Antigua / Pendiente"])
        
        btn_add_p = st.form_submit_button("＋ Agregar Partida")
        if btn_add_p:
            if p_monto > 0 and p_concepto:
                nueva_p = {
                    'tipo': p_tipo.split(" ")[0],
                    'fecha': str(p_fecha),
                    'concepto': p_concepto,
                    'monto': p_monto,
                    'clasificacion': p_clasif,
                    'antiguedad_dias': (datetime.date.today() - p_fecha).days
                }
                st.session_state.partidas_temp = pd.concat([st.session_state.partidas_temp, pd.DataFrame([nueva_p])], ignore_index=True)
                st.rerun()
            else:
                st.error("Proporcione concepto y monto mayor a 0.")
                
    if not st.session_state.partidas_temp.empty:
        st.dataframe(st.session_state.partidas_temp, use_container_width=True)
        if st.button("Limpiar Partidas"):
            st.session_state.partidas_temp = pd.DataFrame(columns=['tipo', 'fecha', 'concepto', 'monto', 'clasificacion', 'antiguedad_dias'])
            st.rerun()
            
    st.divider()
    if st.button("🚀 Enviar Conciliación a Revisión", type="primary"):
        cid = registrar_conciliacion(emp_id, cta_id, periodo, s_libro, s_banco, user_curr['nombre'], st.session_state.partidas_temp)
        if cid:
            st.session_state.partidas_temp = pd.DataFrame(columns=['tipo', 'fecha', 'concepto', 'monto', 'clasificacion', 'antiguedad_dias'])
            st.success(f"Conciliación #{cid} enviada a revisión con éxito.")

elif menu == "🔍 Auditoría y Revisiones":
    st.title("🔍 Bandeja de Auditoría y Revisiones")
    df_conc = obtener_conciliaciones(user_curr['empresa_id'] if user_curr['rol'] != "Administrador" else None)
    
    if df_conc.empty:
        st.info("No hay conciliaciones para revisar.")
    else:
        for idx, r in df_conc.iterrows():
            with st.expander(f"Conciliación #{r['id']} - {r['empresa']} | {r['banco']} ({r['periodo']}) - Estado: {r['estado']}"):
                col1, col2, col3 = st.columns(3)
                col1.write(f"**Saldo Libros:** ${r['saldo_libro']:,.2f}")
                col2.write(f"**Saldo Banco:** ${r['saldo_banco']:,.2f}")
                col3.write(f"**Preparado por:** {r['preparado_por']}")
                
                df_p = obtener_partidas(r['id'])
                tab1, tab2, tab3 = st.tabs(["📊 Movimientos y Saldos", "📋 Bitácora", "⚡ Dictamen"])
                with tab1:
                    st.dataframe(df_p, use_container_width=True)
                with tab2:
                    st.dataframe(obtener_bitacora(r['id']), use_container_width=True)
                with tab3:
                    if user_curr['rol'] in ["Auditor", "Administrador"]:
                        with st.form(f"form_dict_{r['id']}"):
                            n_est = st.selectbox("Estado", ["Aprobada", "Requiere Corrección", "Pendiente"])
                            n_dict = st.selectbox("Dictamen Auditor", ["Sin Salvedades", "Con Salvedades", "Abstención", "Adverso"])
                            n_obs = st.text_area("Observaciones del Auditor", value=r['observaciones'] or "")
                            if st.form_submit_button("Guardar Dictamen"):
                                actualizar_dictamen_conciliacion(r['id'], user_curr['nombre'], n_est, n_dict, n_obs)
                                st.success("Dictamen guardado.")
                                st.rerun()
                    else:
                        st.info("Solo Auditores o Administradores pueden dictaminar.")
                        
                st.download_button("📄 Descargar Reporte PDF", data=generar_pdf_conciliacion(r, df_p, obtener_bitacora(r['id'])), file_name=f"Conciliacion_{r['empresa']}_{r['periodo']}.pdf", mime="application/pdf", key=f"pdf_{r['id']}")

elif menu == "🏢 Gestión de Empresas":
    st.title("🏢 Gestión de Empresas y Cuentas Bancarias")
    tab_e1, tab_e2 = st.tabs(["Registrar Empresa", "Registrar Cuenta Bancaria"])
    
    with tab_e1:
        with st.form("form_emp"):
            nit_e = st.text_input("NIT")
            raz_e = st.text_input("Razón Social")
            if st.form_submit_button("Guardar Empresa"):
                if nit_e and raz_e:
                    if registrar_empresa(nit_e, raz_e):
                        st.success("Empresa registrada con éxito.")
                        st.rerun()
                    else:
                        st.error("El NIT ya está registrado en Turso.")
                else:
                    st.error("Completa todos los campos.")
        st.subheader("Empresas Registradas")
        st.dataframe(obtener_empresas(), use_container_width=True)
        
    with tab_e2:
        df_e = obtener_empresas()
        if not df_e.empty:
            e_dict = dict(zip(df_e['razon_social'], df_e['id']))
            with st.form("form_cta"):
                e_sel = st.selectbox("Empresa", list(e_dict.keys()))
                banco = st.text_input("Banco")
                num_c = st.text_input("Número de Cuenta")
                t_c = st.selectbox("Tipo de Cuenta", ["Ahorros", "Corriente"])
                if st.form_submit_button("Guardar Cuenta"):
                    if banco and num_c:
                        registrar_cuenta(e_dict[e_sel], banco, num_c, t_c)
                        st.success("Cuenta registrada.")
                        st.rerun()
                    else:
                        st.error("Completa los datos de la cuenta.")

elif menu == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios del Sistema")
    df_emp = obtener_empresas()
    emp_opts = {"Ninguna": None}
    if not df_emp.empty:
        for _, r in df_emp.iterrows():
            emp_opts[r['razon_social']] = r['id']
            
    with st.form("form_usr"):
        u_n = st.text_input("Usuario")
        u_nm = st.text_input("Nombre Completo")
        u_p = st.text_input("Contraseña", type="password")
        u_r = st.selectbox("Rol", ["Auxiliar", "Auditor", "Administrador"])
        u_e = st.selectbox("Empresa Asignada", list(emp_opts.keys()))
        if st.form_submit_button("Registrar Usuario"):
            if u_n and u_nm and u_p:
                if registrar_usuario(u_n, u_nm, u_p, u_r, emp_opts[u_e]):
                    st.success("Usuario registrado con éxito.")
                    st.rerun()
                else:
                    st.error("El usuario ya existe.")
            else:
                st.error("Completa todos los datos requeridos.")
