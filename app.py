import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import datetime

# ==========================================
# CONFIGURACIÓN INICIAL DE LA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Sistema de Conciliación Bancaria",
    page_icon="🏦",
    layout="wide"
)

# ==========================================
# FUNCIONES DE BASE DE DATOS Y BACKEND
# ==========================================
def inicializar_base_datos():
    conn = sqlite3.connect("conciliaciones.db")
    cursor = conn.cursor()
    
    # Tabla de Empresas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            nit TEXT NOT NULL
        )
    """)
    
    # Tabla de Usuarios
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL,
            password TEXT NOT NULL,
            rol TEXT NOT NULL,
            empresa_id INTEGER,
            FOREIGN KEY (empresa_id) REFERENCES empresas (id)
        )
    """)
    
    # Tabla de Historial de Conciliaciones
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historial_conciliaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_guardado TEXT,
            empresa TEXT,
            nit TEXT,
            mes TEXT,
            fecha_elaboracion TEXT,
            banco TEXT,
            cuenta TEXT,
            tipo TEXT,
            saldo_extracto REAL,
            saldo_libros REAL,
            diferencia_inicial REAL,
            diferencia_conciliada REAL,
            resultado_final REAL,
            estado TEXT,
            workflow_status TEXT,
            revisado_por_usuario TEXT,
            excel BLOB
        )
    """)
    
    conn.commit()
    
    # Crear usuario administrador por defecto si no existe
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO usuarios (username, nombre, password, rol, empresa_id)
            VALUES (?, ?, ?, ?, ?)
        """, ("admin", "Administrador Principal", "admin123", "Administrador", None))
        conn.commit()
        
    conn.close()

inicializar_base_datos()

def obtener_usuarios():
    conn = sqlite3.connect("conciliaciones.db")
    df = pd.read_sql_query("SELECT * FROM usuarios", conn)
    conn.close()
    return df

def crear_usuario(username, nombre, password, rol, empresa_id):
    conn = sqlite3.connect("conciliaciones.db")
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO usuarios (username, nombre, password, rol, empresa_id)
        VALUES (?, ?, ?, ?, ?)
    """, (username, nombre, password, rol, empresa_id))
    conn.commit()
    conn.close()

def actualizar_usuario(user_id, username, nombre, password, rol, empresa_id):
    conn = sqlite3.connect("conciliaciones.db")
    cursor = conn.cursor()
    if password.strip() != "":
        cursor.execute("""
            UPDATE usuarios 
            SET username = ?, nombre = ?, password = ?, rol = ?, empresa_id = ?
            WHERE id = ?
        """, (username, nombre, password, rol, empresa_id, user_id))
    else:
        cursor.execute("""
            UPDATE usuarios 
            SET username = ?, nombre = ?, rol = ?, empresa_id = ?
            WHERE id = ?
        """, (username, nombre, rol, empresa_id, user_id))
    conn.commit()
    conn.close()

def obtener_empresas():
    conn = sqlite3.connect("conciliaciones.db")
    df = pd.read_sql_query("SELECT * FROM empresas", conn)
    conn.close()
    return df

def guardar_conciliacion_historial(
    empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
    saldo_extracto, saldo_libros, diferencia_inicial,
    diferencia_conciliada, resultado_final, t1, t2, t3, t4, t5,
    preparado_por, revisado_por, excel_bytes, workflow_status="Borrador", id_edicion=None
):
    conn = sqlite3.connect("conciliaciones.db")
    cursor = conn.cursor()
    estado_str = "CONCILIACIÓN BANCARIA CORRECTA" if abs(resultado_final) < 0.005 else "CONCILIACIÓN CON DIFERENCIA"
    fecha_guardado = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if id_edicion:
        cursor.execute("""
            UPDATE historial_conciliaciones SET
                fecha_guardado = ?, empresa = ?, nit = ?, mes = ?, fecha_elaboracion = ?,
                banco = ?, cuenta = ?, tipo = ?, saldo_extracto = ?, saldo_libros = ?,
                diferencia_inicial = ?, diferencia_conciliada = ?, resultado_final = ?,
                estado = ?, workflow_status = ?, revisado_por_usuario = ?, excel = ?
            WHERE id = ?
        """, (
            fecha_guardado, empresa, nit, mes, str(fecha_elaboracion), banco, cuenta, tipo,
            saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
            resultado_final, estado_str, workflow_status, revisado_por, excel_bytes, id_edicion
        ))
    else:
        cursor.execute("""
            INSERT INTO historial_conciliaciones (
                fecha_guardado, empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, estado, workflow_status, revisado_por_usuario, excel
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            fecha_guardado, empresa, nit, mes, str(fecha_elaboracion), banco, cuenta, tipo,
            saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
            resultado_final, estado_str, workflow_status, revisado_por, excel_bytes
        ))
    conn.commit()
    conn.close()

def obtener_historial(empresa_nombre=None):
    conn = sqlite3.connect("conciliaciones.db")
    if empresa_nombre:
        df = pd.read_sql_query("SELECT * FROM historial_conciliaciones WHERE empresa = ?", conn, params=(empresa_nombre,))
    else:
        df = pd.read_sql_query("SELECT * FROM historial_conciliaciones", conn)
    conn.close()
    return df

def obtener_conciliacion_por_id(id_cons):
    conn = sqlite3.connect("conciliaciones.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM historial_conciliaciones WHERE id = ?", (id_cons,))
    row = cursor.fetchone()
    conn.close()
    if row:
        cols = [description[0] for description in cursor.description] if 'cursor' in locals() else [
            "id", "fecha_guardado", "empresa", "nit", "mes", "fecha_elaboracion", "banco",
            "cuenta", "tipo", "saldo_extracto", "saldo_libros", "diferencia_inicial",
            "diferencia_conciliada", "resultado_final", "estado", "workflow_status",
            "revisado_por_usuario", "excel"
        ]
        return dict(zip(cols, row))
    return None

def limpiar_datos_operativos():
    try:
        conn = sqlite3.connect("conciliaciones.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM historial_conciliaciones")
        conn.commit()
        conn.close()
        return True, "Datos operativos eliminados correctamente."
    except Exception as e:
        return False, str(e)

def reiniciar_datos_aplicativo():
    try:
        conn = sqlite3.connect("conciliaciones.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM historial_conciliaciones")
        cursor.execute("DELETE FROM empresas")
        cursor.execute("DELETE FROM usuarios")
        cursor.execute("""
            INSERT INTO usuarios (username, nombre, password, rol, empresa_id)
            VALUES (?, ?, ?, ?, ?)
        """, ("admin", "Administrador Principal", "admin123", "Administrador", None))
        conn.commit()
        conn.close()
        return True, "Aplicativo reiniciado por completo."
    except Exception as e:
        return False, str(e)

# ==========================================
# UTILIDADES DE FORMATO Y CÁLCULO
# ==========================================
def formatear_moneda(valor):
    try:
        return f"${float(valor):,.2f}"
    except:
        return "$0.00"

def total_columna(df, columna):
    if df is None or df.empty or columna not in df.columns:
        return 0.0
    return pd.to_numeric(df[columna], errors="coerce").fillna(0.0).sum()

def formatear_columna_valor(df):
    if df is not None and not df.empty:
        for col in df.columns:
            if "Valor" in col or col in ["4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return df

def config_monetaria(columnas):
    return {col: st.column_config.NumberColumn(format="$#,##0.00") for col in columnas}

def limpiar_nombre_archivo(texto):
    return "".join(c for c in str(texto) if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')

def preparar_excel(*args):
    # Generador simulado del reporte binario de Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_dummy = pd.DataFrame({"Reporte": ["Conciliación Bancaria Generada Exitosamente"]})
        df_dummy.to_excel(writer, sheet_name="Resumen", index=False)
    return output.getvalue(), "conciliacion_bancaria.xlsx"

def generar_pdf_conciliacion(*args):
    # Generador simulado del reporte binario de PDF
    return b"%PDF-1.4 Mock PDF Stream"

# ==========================================
# GESTIÓN DE SESIÓN Y AUTENTICACIÓN
# ==========================================
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = None

if "menu_override" not in st.session_state:
    st.session_state.menu_override = None

if "conciliacion_a_editar" not in st.session_state:
    st.session_state.conciliacion_a_editar = None

if st.session_state.usuario_actual is None:
    st.sidebar.title("🔐 Inicio de Sesión")
    usuarios_df = obtener_usuarios()
    with st.sidebar.form("form_login"):
        user_input = st.text_input("Usuario")
        pass_input = st.text_input("Contraseña", type="password")
        btn_login = st.form_submit_button("Ingresar", type="primary")
        
        if btn_login:
            user_match = usuarios_df[usuarios_df["username"] == user_input]
            if not user_match.empty and user_match.iloc[0]["password"] == pass_input:
                st.session_state.usuario_actual = user_match.iloc[0].to_dict()
                st.success("¡Bienvenido!")
                st.rerun()
            else:
                st.error("Credenciales incorrectas.")
    st.stop()

# ==========================================
# BARRA LATERAL Y NAVEGACIÓN
# ==========================================
usuario_actual = st.session_state.usuario_actual
rol_actual = usuario_actual["rol"]

st.sidebar.title("🏦 Menú Principal")
st.sidebar.write(f"Usuario: **{usuario_actual['nombre']}** (`{rol_actual}`)")

if st.sidebar.button("Cerrar Sesión"):
    st.session_state.usuario_actual = None
    st.session_state.menu_override = None
    st.rerun()

st.sidebar.divider()

opciones_menu = ["📝 Nueva Conciliación", "📋 Historial", "📄 Reportes", "⚙️️ Administración", "👥 Usuarios"]
menu_seleccionado = st.sidebar.selectbox("Ir a", opciones_menu, index=0 if st.session_state.menu_override is None else opciones_menu.index(st.session_state.menu_override))
st.session_state.menu_override = None

empresas_df = obtener_empresas()
empresa_activa_nombre = "Empresa General"
if not empresas_df.empty:
    empresa_activa_nombre = st.sidebar.selectbox("Empresa Activa", empresas_df["nombre"].tolist())

# ==========================================
# ESTRUCTURA DE PANTALLAS (VISTAS)
# ==========================================
if menu_seleccionado == "📝 Nueva Conciliación":
    st.title("📝 Módulo de Conciliación Bancaria")
    
    id_edicion = st.session_state.conciliacion_a_editar
    
    col1, col2, col3 = st.columns(3)
    empresa = col1.text_input("Empresa", value=empresa_activa_nombre)
    nit = col2.text_input("NIT", value="900.000.000-1")
    mes = col3.text_input("Mes Conciliación", value="Octubre 2026")
    
    col4, col5, col6 = st.columns(3)
    fecha_elaboracion = col4.date_input("Fecha de Elaboración", value=datetime.now())
    banco = col5.text_input("Banco", value="Bancolombia")
    cuenta = col6.text_input("Número de Cuenta", value="123-456789-00")
    
    tipo = st.selectbox("Tipo de Cuenta", ["Corriente", "Ahorros"])
    
    col_s1, col_s2 = st.columns(2)
    saldo_extracto = col_s1.number_input("Saldo según Extracto Bancario", value=0.0, format="%.2f")
    saldo_libros = col_s2.number_input("Saldo según Libros Contables", value=0.0, format="%.2f")
    
    diferencia_inicial = saldo_extracto - saldo_libros
    st.info(f"Diferencia Inicial: {formatear_moneda(diferencia_inicial)}")
    
    # Inicialización de tablas secundarias en session state si no existen
    nombres_titulos = {
        "t1": "Salidas Extracto (No registradas en Libros)",
        "t2": "Salidas Libros (No registradas en Banco)",
        "t3": "Entradas Libros (No registradas en Banco)",
        "t4": "Entradas Extracto (No registradas en Libros)"
    }
    
    cols_t1 = ["Fecha", "Documento", "Tercero", "Concepto", "Valor"]
    cols_t2 = ["Fecha", "Documento", "Tercero", "Concepto", "Valor"]
    cols_t3 = ["Fecha", "Documento", "Tercero", "Concepto", "Valor"]
    cols_t4 = ["Fecha", "Documento", "Tercero", "Concepto", "Valor"]
    
    for key, cols in [("tabla1", cols_t1), ("tabla2", cols_t2), ("tabla3", cols_t3), ("tabla4", cols_t4)]:
        if key not in st.session_state:
            st.session_state[key] = pd.DataFrame(columns=cols)
            
    if "tabla5" not in st.session_state:
        st.session_state.tabla5 = pd.DataFrame(columns=["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"])

    st.divider()
    st.subheader(nombres_titulos["t1"])
    salidas_extracto = formatear_columna_valor(st.data_editor(st.session_state.tabla1, num_rows="dynamic", use_container_width=True, key="editor_t1", column_config=config_monetaria(["Valor"])))
    if st.button("🗑️ Limpiar Ítem 1", key="btn_del_t1"):
        st.session_state.tabla1 = pd.DataFrame(columns=cols_t1)
        st.rerun()

    st.divider()
    st.subheader(nombres_titulos["t2"])
    salidas_libros = formatear_columna_valor(st.data_editor(st.session_state.tabla2, num_rows="dynamic", use_container_width=True, key="editor_t2", column_config=config_monetaria(["Valor"])))
    if st.button("🗑️️ Limpiar Ítem 2", key="btn_del_t2"):
        st.session_state.tabla2 = pd.DataFrame(columns=cols_t2)
        st.rerun()

    st.divider()
    st.subheader(nombres_titulos["t3"])
    entradas_libros = formatear_columna_valor(st.data_editor(st.session_state.tabla3, num_rows="dynamic", use_container_width=True, key="editor_t3", column_config=config_monetaria(["Valor"])))
    if st.button("🗑️ Limpiar Ítem 3", key="btn_del_t3"):
        st.session_state.tabla3 = pd.DataFrame(columns=cols_t3)
        st.rerun()

    st.divider()
    st.subheader(nombres_titulos["t4"])
    c_ed4, c_del4 = st.columns([6, 1])
    with c_ed4:
        entradas_extracto = formatear_columna_valor(st.data_editor(st.session_state.tabla4, num_rows="dynamic", use_container_width=True, key="editor_t4", column_config=config_monetaria(["Valor"])))
    with c_del4:
        if st.button("🗑️ Limpiar Ítem 4", key="btn_del_t4"):
            st.session_state.tabla4 = pd.DataFrame(columns=cols_t4)
            st.rerun()

    st.divider()
    st.subheader("GASTOS BANCARIOS")
    gastos_bancarios_df = formatear_columna_valor(st.session_state.tabla5)
    gastos_bancarios = st.data_editor(gastos_bancarios_df, num_rows="dynamic", use_container_width=True, key="editor_tabla5", column_config=config_monetaria(["4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"]))
    
    if st.button("🗑️ Limpiar Gastos Bancarios", key="btn_del_t5"):
        st.session_state.tabla5 = pd.DataFrame(columns=["Fecha", "4 x 1000", "Cuota de manejo", "IVA", "Rte. fuente", "Comisión", "Ing. x intereses"])
        st.rerun()

    st.divider()
    st.subheader("📊 Resumen y Resultado de la Conciliación")
    
    tot_t1 = total_columna(salidas_extracto, "Valor")
    tot_t2 = total_columna(salidas_libros, "Valor")
    tot_t3 = total_columna(entradas_libros, "Valor")
    tot_t4 = total_columna(entradas_extracto, "Valor")
    
    diferencia_conciliada = diferencia_inicial - tot_t1 - tot_t2 + tot_t3 + tot_t4
    resultado_final = diferencia_conciliada 
    
    res1, res2, res3 = st.columns(3)
    res1.metric("Total Partidas 1", formatear_moneda(tot_t1))
    res2.metric("Total Partidas 2", formatear_moneda(tot_t2))
    res3.metric("Total Partidas 3", formatear_moneda(tot_t3))
    
    res4, res5 = st.columns(2)
    res4.metric("Total Partidas 4", formatear_moneda(tot_t4))
    res5.metric("Resultado Final (Diferencia)", formatear_moneda(resultado_final), delta="Cuadra" if abs(resultado_final) < 0.005 else "Con diferencia", delta_color="normal" if abs(resultado_final) < 0.005 else "inverse")

    st.divider()
    c_guardar, c_excel, c_pdf = st.columns(3)
    
    preparado_por = st.text_input("Preparado por (Nombre del Analista)", value=usuario_actual["nombre"])
    revisado_por_val = st.text_input("Revisado por (Opcional)", value="")

    excel_bytes, excel_nombre = preparar_excel(
        empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
        saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
        resultado_final, salidas_extracto, salidas_libros, entradas_libros,
        entradas_extracto, gastos_bancarios, preparado_por, revisado_por_val,
        nombres_titulos
    )

    with c_guardar:
        if st.button("💾 Guardar Conciliación", type="primary", use_container_width=True):
            workflow_status = "Borrador" if abs(resultado_final) >= 0.005 else "Pendiente de revisión"
            guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial,
                diferencia_conciliada, resultado_final, salidas_extracto,
                salidas_libros, entradas_libros, entradas_extracto,
                gastos_bancarios, preparado_por, revisado_por_val, excel_bytes,
                workflow_status=workflow_status, id_edicion=id_edicion
            )
            st.success("¡Conciliación guardada y registrada en el historial con éxito!")
            if id_edicion:
                st.session_state.conciliacion_a_editar = None
            st.rerun()

    with c_excel:
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_bytes,
            file_name=excel_nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with c_pdf:
        try:
            pdf_bytes = generar_pdf_conciliacion(
                {
                    "id": id_edicion or 0, "empresa": empresa, "nit": nit, "mes": mes,
                    "fecha_elaboracion": str(fecha_elaboracion), "banco": banco, "cuenta": cuenta,
                    "tipo": tipo, "saldo_extracto": saldo_extracto, "saldo_libros": saldo_libros,
                    "diferencia_inicial": diferencia_inicial, "diferencia_conciliada": diferencia_conciliada,
                    "resultado_final": resultado_final, "estado": "CONCILIACIÓN BANCARIA CORRECTA" if abs(resultado_final) < 0.005 else "CONCILIACIÓN CON DIFERENCIA",
                    "revisado_por_usuario": revisado_por_val
                },
                {
                    "salidas_extracto": salidas_extracto.to_dict(orient="records"),
                    "salidas_libros": salidas_libros.to_dict(orient="records"),
                    "entradas_libros": entradas_libros.to_dict(orient="records"),
                    "entradas_extracto": entradas_extracto.to_dict(orient="records"),
                    "gastos_bancarios": gastos_bancarios.to_dict(orient="records")
                },
                nombres_titulos
            )
            st.download_button(
                label="📄 Descargar PDF",
                data=pdf_bytes,
                file_name=f"CONCILIACION_{limpiar_nombre_archivo(empresa)}_{limpiar_nombre_archivo(mes)}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"No fue posible generar el PDF: {e}")

elif menu_seleccionado == "📋 Historial":
    st.title("📋 Historial de Conciliaciones Bancarias")
    historial = obtener_historial(empresa_activa_nombre)
    if historial.empty:
        st.info("No hay conciliaciones guardadas en el historial.")
    else:
        st.dataframe(historial[["id", "fecha_guardado", "empresa", "mes", "banco", "cuenta", "estado", "workflow_status"]], use_container_width=True, hide_index=True)
        
        st.divider()
        st.subheader("🔍 Consultar / Descargar Conciliación por ID")
        id_cons = st.number_input("Ingrese ID de Conciliación", min_value=1, step=1)
        if st.button("Consultar Detalle"):
            c_det = obtener_conciliacion_por_id(id_cons)
            if c_det:
                st.success(f"Conciliación encontrada: {c_det['empresa']} - {c_det['mes']} ({c_det['banco']})")
                if c_det.get("excel"):
                    st.download_button("📥 Descargar Excel Guardado", data=c_det["excel"], file_name=f"conciliacion_{id_cons}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                if rol_actual in ["Administrador", "Revisor"]:
                    if st.button("✏️ Editar esta Conciliación"):
                        st.session_state.conciliacion_a_editar = id_cons
                        st.session_state.menu_override = "📝 Nueva Conciliación"
                        st.rerun()
            else:
                st.error("No se encontró ninguna conciliación con ese ID.")

elif menu_seleccionado == "📄 Reportes":
    st.title("📄 Módulo de Reportes Consolidados")
    st.caption("Genera reportes gerenciales y consolidados de las conciliaciones bancarias.")
    historial = obtener_historial(empresa_activa_nombre)
    if historial.empty:
        st.info("No hay datos suficientes para generar reportes.")
    else:
        st.dataframe(historial[["fecha_guardado", "empresa", "mes", "banco", "cuenta", "saldo_extracto", "saldo_libros", "estado", "workflow_status"]], use_container_width=True, hide_index=True)

elif menu_seleccionado == "⚙️ Administración":
    st.title("⚙️ Administración General del Sistema")
    st.caption("Herramientas avanzadas de mantenimiento, limpieza y control de datos operativos.")
    
    col_adm1, col_adm2 = st.columns(2)
    with col_adm1:
        with st.container(border=True):
            st.subheader("🧹 Limpiar Datos Operativos")
            st.write("Elimina bancos, cuentas y conciliaciones, conservando empresas y usuarios.")
            if st.button("Ejecutar Limpieza Operativa", type="secondary"):
                ok, msg = limpiar_datos_operativos()
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
                    
    with col_adm2:
        with st.container(border=True):
            st.subheader("⚠️ Reiniciar Todo el Aplicativo")
            st.write("Elimina todos los datos (empresas, cuentas, conciliaciones y usuarios) dejando el sistema como nueva instalación.")
            if st.button("🚨 Reiniciar Sistema Completo", type="primary"):
                ok, msg = reiniciar_datos_aplicativo()
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

elif menu_seleccionado == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios y Roles")
    if rol_actual != "Administrador":
        st.error("Acceso restringido a Administradores.")
    else:
        tab_crear, tab_editar = st.tabs(["➕ Crear Usuario", "✏️ Editar Usuario"])
        
        with tab_crear:
            with st.form("form_crear_usuario_adm"):
                nuevo_usr = st.text_input("Nombre de usuario")
                nuevo_nom = st.text_input("Nombre completo")
                nuevo_pas = st.text_input("Contraseña", type="password")
                nuevo_rol = st.selectbox("Rol", ["Preparador", "Revisor", "Administrador"])
                emp_asig = None
                if not empresas_df.empty:
                    emp_asig = st.selectbox(
                        "Asignar Empresa",
                        empresas_df["id"].tolist(),
                        format_func=lambda x: empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0],
                        key="emp_crear"
                    )
                if st.form_submit_button("Crear Usuario", type="primary"):
                    if not nuevo_usr.strip() or not nuevo_nom.strip() or not nuevo_pas:
                        st.error("Todos los campos obligatorios deben completarse.")
                    else:
                        try:
                            crear_usuario(nuevo_usr, nuevo_nom, nuevo_pas, nuevo_rol, emp_asig)
                            st.success("Usuario creado correctamente.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al crear usuario: {e}")
                            
        with tab_editar:
            df_usr = obtener_usuarios()
            if df_usr.empty:
                st.info("No hay usuarios registrados para editar.")
            else:
                usuario_a_editar_id = st.selectbox(
                    "Seleccione el usuario que desea editar",
                    df_usr["id"].tolist(),
                    format_func=lambda x: f"{df_usr.loc[df_usr['id'] == x, 'username'].values[0]} - {df_usr.loc[df_usr['id'] == x, 'nombre'].values[0]}"
                )
                
                user_data = df_usr[df_usr["id"] == usuario_a_editar_id].iloc[0]
                
                with st.form("form_editar_usuario_adm"):
                    edit_usr = st.text_input("Nombre de usuario", value=user_data["username"])
                    edit_nom = st.text_input("Nombre completo", value=user_data["nombre"])
                    edit_pas = st.text_input("Nueva contraseña (déjalo en blanco si no deseas cambiarla)", type="password", value="")
                    
                    roles_disponibles = ["Preparador", "Revisor", "Administrador"]
                    rol_actual_idx = roles_disponibles.index(user_data["rol"]) if user_data["rol"] in roles_disponibles else 0
                    edit_rol = st.selectbox("Rol", roles_disponibles, index=rol_actual_idx)
                    
                    emp_asig_edit = None
                    if not empresas_df.empty:
                        emp_asig_actual = user_data.get("empresa_id")
                        emp_ids = empresas_df["id"].tolist()
                        idx_emp = emp_ids.index(emp_asig_actual) if emp_asig_actual in emp_ids else 0
                        emp_asig_edit = st.selectbox(
                            "Asignar Empresa",
                            emp_ids,
                            index=idx_emp,
                            format_func=lambda x: empresas_df.loc[empresas_df["id"] == x, "nombre"].values[0],
                            key="emp_editar"
                        )
                        
                    if st.form_submit_button("💾 Guardar Cambios", type="primary"):
                        try:
                            actualizar_usuario(usuario_a_editar_id, edit_usr, edit_nom, edit_pas, edit_rol, emp_asig_edit)
                            st.success("¡Usuario actualizado correctamente!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al actualizar el usuario: {e}")

        st.divider()
        st.subheader("Usuarios Registrados en el Sistema")
        df_usr_show = obtener_usuarios()
        if not df_usr_show.empty:
            if "password" in df_usr_show.columns:
                df_usr_show = df_usr_show.drop(columns=["password"])
            st.dataframe(df_usr_show, use_container_width=True, hide_index=True)
