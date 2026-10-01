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
# CONEXIÓN A BASE DE DATOS (TURSO / SQLITE)
# ==========================================
def conectar_db():
    """
    Intenta conectar a la base de datos persistente de Turso en la nube
    utilizando los Secrets de Streamlit. Si falla o se ejecuta localmente,
    hace respaldo a SQLite local.
    """
    if "TURSO_DATABASE_URL" in st.secrets and "TURSO_AUTH_TOKEN" in st.secrets:
        try:
            import libsql_experimental as libsql
            return libsql.connect(
                database=st.secrets["TURSO_DATABASE_URL"],
                auth_token=st.secrets["TURSO_AUTH_TOKEN"]
            )
        except Exception as e:
            st.warning(f"⚠️ No se pudo conectar a Turso: {e}. Usando SQLite local.")
    return sqlite3.connect("conciliaciones.db")

def inicializar_bd():
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
            tipo TEXT NOT NULL, -- 'ND_LIBROS', 'NC_LIBROS', 'ND_BANCO', 'NC_BANCO'
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

# Inicializar tablas si no existen
inicializar_bd()

# ==========================================
# FUNCIONES DE CONSULTA
# ==========================================
def contar_usuarios():
    conn = conectar_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM usuarios")
    total = c.fetchone()[0]
    conn.close()
    return total

def verificar_credenciales(username, password):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("SELECT id, username, nombre, rol, empresa_id FROM usuarios WHERE username = ? AND password = ?", (username, password))
    user = c.fetchone()
    conn.close()
    return user

def registrar_usuario(username, nombre, password, rol, empresa_id=None):
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("INSERT INTO usuarios (username, nombre, password, rol, empresa_id) VALUES (?, ?, ?, ?, ?)",
                  (username, nombre, password, rol, empresa_id))
        conn.commit()
        res = True
    except sqlite3.IntegrityError:
        res = False
    conn.close()
    return res

def obtener_empresas():
    conn = conectar_db()
    df = pd.read_sql_query("SELECT * FROM empresas", conn)
    conn.close()
    return df

def registrar_empresa(nit, razon_social):
    conn = conectar_db()
    c = conn.cursor()
    try:
        c.execute("INSERT INTO empresas (nit, razon_social) VALUES (?, ?)", (nit, razon_social))
        conn.commit()
        res = True
    except sqlite3.IntegrityError:
        res = False
    conn.close()
    return res

def registrar_cuenta(empresa_id, banco, numero_cuenta, tipo_cuenta):
    conn = conectar_db()
    c = conn.cursor()
    c.execute("INSERT INTO cuentas (empresa_id, banco, numero_cuenta, tipo_cuenta) VALUES (?, ?, ?, ?)",
              (empresa_id, banco, numero_cuenta, tipo_cuenta))
    conn.commit()
    conn.close()

def obtener_cuentas_empresa(empresa_id):
    conn = conectar_db()
    df = pd.read_sql_query("SELECT * FROM cuentas WHERE empresa_id = ?", conn, params=(empresa_id,))
    conn.close()
    return df

def registrar_conciliacion(empresa_id, cuenta_id, periodo, saldo_libro, saldo_banco, preparado_por, df_partidas):
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

def obtener_conciliaciones(empresa_id=None):
    conn = conectar_db()
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
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def actualizar_dictamen_conciliacion(conciliacion_id, usuario, nuevo_estado, dictamen, observaciones):
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

def obtener_partidas(conciliacion_id):
    conn = conectar_db()
    df = pd.read_sql_query("SELECT * FROM partidas_conciliatorias WHERE conciliacion_id = ?", conn, params=(conciliacion_id,))
    conn.close()
    return df

def obtener_bitacora(conciliacion_id):
    conn = conectar_db()
    df = pd.read_sql_query("SELECT * FROM bitacora_auditoria WHERE conciliacion_id = ? ORDER BY id DESC", conn, params=(conciliacion_id,))
    conn.close()
    return df

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
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontName='Helvetica-Bold')
    
    # Encabezado
    story.append(Paragraph(f"INFORME DE CONCILIACIÓN BANCARIA Y AUDITORÍA", title_style))
    story.append(Paragraph(f"Empresa: {conciliacion_info['empresa']} | Periodo: {conciliacion_info['periodo']}", subtitle_style))
    story.append(Spacer(1, 15))
    
    # Datos Principales
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
    
    # Tabla de Partidas
    story.append(Paragraph("Detalle de Partidas Conciliatorias", subtitle_style))
    story.append(Spacer(1, 5))
    
    data_partidas = [["Tipo", "Fecha", "Concepto", "Monto", "Clasificación"]]
    for _, r in df_partidas.iterrows():
        data_partidas.append([r['tipo'], r['fecha'], r['concepto'], f"${r['monto']:,.2f}", r.get('clasificacion', 'General')])
        
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
    
    # Observaciones
    if conciliacion_info['observaciones']:
        story.append(Paragraph("Observaciones del Auditor", subtitle_style))
        story.append(Paragraph(conciliacion_info['observaciones'], styles['Normal']))
        story.append(Spacer(1, 15))
        
    doc.build(story)
    buffer.seek(0)
    return buffer

# ==========================================
# GESTIÓN DE SESIÓN
# ==========================================
if 'usuario' not in st.session_state:
    st.session_state.usuario = None

# Pantalla de Configuración Inicial si no existen usuarios
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

# Pantalla de Login
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

# ==========================================
# MENÚ LATERAL Y NAVEGACIÓN
# ==========================================
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
# MÓDULO: DASHBOARD
# ==========================================
if menu == "📊 Dashboard":
    st.title("📊 Dashboard de Control y Auditoría")
    
    df_conc = obtener_conciliaciones(user_curr['empresa_id'] if user_curr['rol'] != "Administrador" else None)
    
    if df_conc.empty:
        st.info("No hay conciliaciones registradas en el sistema.")
    else:
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        total_conc = len(df_conc)
        aprobadas = len(df_conc[df_conc['estado'] == 'Aprobada'])
        pendientes = len(df_conc[df_conc['estado'] == 'Pendiente'])
        corregir = len(df_conc[df_conc['estado'] == 'Requiere Corrección'])
        
        kpi1.metric("Total Conciliaciones", total_conc)
        kpi2.metric("Aprobadas", aprobadas)
        kpi3.metric("Pendientes", pendientes)
        kpi4.metric("Requieren Corrección", corregir)
        
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
                             barmode='group', title='Comparativa Saldo Libros vs Saldo Bancos',
                             labels={'value': 'Monto ($)', 'variable': 'Tipo Saldo'})
            st.plotly_chart(fig_bar, use_container_width=True)

# ==========================================
# MÓDULO: NUEVA CONCILIACIÓN
# ==========================================
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
                st.success("Partida agregada.")
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
        st.session_state.partidas_temp = pd.DataFrame(columns=['tipo', 'fecha', 'concepto', 'monto', 'clasificacion', 'antiguedad_dias'])
        st.success(f"Conciliación #{cid} enviada a revisión con éxito.")

# ==========================================
# MÓDULO: AUDITORÍA Y REVISIONES
# ==========================================
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
                    df_b = obtener_bitacora(r['id'])
                    st.dataframe(df_b, use_container_width=True)
                    
                with tab3:
                    if user_curr['rol'] in ["Auditor", "Administrador"]:
                        with st.form(f"form_dictamen_{r['id']}"):
                            nuevo_estado = st.selectbox("Estado", ["Aprobada", "Requiere Corrección", "Pendiente"])
                            dictamen_sel = st.selectbox("Dictamen Auditor", ["Sin Salvedades", "Con Salvedades", "Abstención", "Adverso"])
                            obs_text = st.text_area("Observaciones del Auditor", value=r['observaciones'] or "")
                            btn_dictamen = st.form_submit_button("Guardar Dictamen")
                            
                            if btn_dictamen:
                                actualizar_dictamen_conciliacion(r['id'], user_curr['nombre'], nuevo_estado, dictamen_sel, obs_text)
                                st.success("Dictamen guardado.")
                                st.rerun()
                    else:
                        st.info("Solo Auditores o Administradores pueden emitir dictamen.")
                        
                # Botón Descargar PDF
                pdf_bytes = generar_pdf_conciliacion(r, df_p, obtener_bitacora(r['id']))
                st.download_button(
                    label="📄 Descargar Reporte PDF",
                    data=pdf_bytes,
                    file_name=f"Conciliacion_{r['empresa']}_{r['periodo']}.pdf",
                    mime="application/pdf",
                    key=f"pdf_btn_{r['id']}"
                )

# ==========================================
# MÓDULO: GESTIÓN DE EMPRESAS
# ==========================================
elif menu == "🏢 Gestión de Empresas":
    st.title("🏢 Gestión de Empresas y Cuentas Bancarias")
    
    tab_e1, tab_e2 = st.tabs(["Registrar Empresa", "Registrar Cuenta Bancaria"])
    
    with tab_e1:
        with st.form("form_empresa"):
            nit_emp = st.text_input("NIT")
            razon_emp = st.text_input("Razón Social")
            btn_emp = st.form_submit_button("Guardar Empresa")
            if btn_emp:
                if nit_emp and razon_emp:
                    if registrar_empresa(nit_emp, razon_emp):
                        st.success("Empresa registrada con éxito.")
                        st.rerun()
                    else:
                        st.error("El NIT ya está registrado.")
                else:
                    st.error("Completa todos los campos.")
                    
        st.subheader("Empresas Registradas")
        st.dataframe(obtener_empresas(), use_container_width=True)
        
    with tab_e2:
        df_e = obtener_empresas()
        if not df_e.empty:
            e_dict = dict(zip(df_e['razon_social'], df_e['id']))
            with st.form("form_cuenta"):
                e_sel = st.selectbox("Empresa", list(e_dict.keys()))
                banco = st.text_input("Banco")
                num_cuenta = st.text_input("Número de Cuenta")
                tipo_cuenta = st.selectbox("Tipo de Cuenta", ["Ahorros", "Corriente"])
                btn_cta = st.form_submit_button("Guardar Cuenta")
                
                if btn_cta:
                    if banco and num_cuenta:
                        registrar_cuenta(e_dict[e_sel], banco, num_cuenta, tipo_cuenta)
                        st.success("Cuenta registrada.")
                        st.rerun()
                    else:
                        st.error("Completa los datos de la cuenta.")

# ==========================================
# MÓDULO: GESTIÓN DE USUARIOS
# ==========================================
elif menu == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios del Sistema")
    
    df_emp = obtener_empresas()
    emp_opts = {"Ninguna": None}
    for _, r in df_emp.iterrows():
        emp_opts[r['razon_social']] = r['id']
        
    with st.form("form_reg_usr"):
        st.subheader("Crear Nuevo Usuario")
        u_name = st.text_input("Usuario")
        u_nom = st.text_input("Nombre Completo")
        u_pwd = st.text_input("Contraseña", type="password")
        u_rol = st.selectbox("Rol", ["Auxiliar", "Auditor", "Administrador"])
        u_emp = st.selectbox("Empresa Asignada", list(emp_opts.keys()))
        
        btn_usr = st.form_submit_button("Registrar Usuario")
        if btn_usr:
            if u_name and u_nom and u_pwd:
                if registrar_usuario(u_name, u_nom, u_pwd, u_rol, emp_opts[u_emp]):
                    st.success("Usuario registrado exitosamente.")
                    st.rerun()
                else:
                    st.error("El nombre de usuario ya existe.")
            else:
                st.error("Completa todos los datos requeridos.")
