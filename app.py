    r3.metric("Resultado Final", formatear_moneda(resultado_final))

    preparado_por = usuario_actual["nombre"]
    revisado_por = ""
    renderizar_bloque_firmas_ui(preparado_por, "Se registra al guardar", "", "Pendiente", "", "Pendiente", "Borrador / Pendiente de revisión")

    excel_data, nombre_archivo = preparar_excel(
        empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
        saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
        resultado_final, salidas_extracto, salidas_libros, entradas_libros,
        entradas_extracto, gastos_bancarios, preparado_por, revisado_por,
        nombres_titulos, aprobado_por="", fecha_preparacion="", fecha_revision="", fecha_aprobacion="",
        estado="Borrador / Pendiente de revisión"
    )

    st.divider()
    btn_col_borr, btn_col_env = st.columns(2)
    with btn_col_borr:
        if st.button("💾 Guardar como Borrador", key="btn_guardar_borrador"):
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data,
                workflow_status="Borrador", id_edicion=id_edicion
            )
            st.session_state.pop("conciliacion_a_editar", None)
            st.session_state.pop("datos_cargados_edit", None)
            st.session_state["historial_id_destacado"] = int(id_g)
            st.session_state.menu_override = "📋 Historial"
            st.success(f"💾 Conciliación CONC-{int(id_g):06d} guardada como Borrador.")

    with btn_col_env:
        if st.button("🚀 Enviar a Revisión", key="btn_enviar_revision", type="primary"):
            id_g = guardar_conciliacion_historial(
                empresa, nit, mes, fecha_elaboracion, banco, cuenta, tipo,
                saldo_extracto, saldo_libros, diferencia_inicial, diferencia_conciliada,
                resultado_final, salidas_extracto, salidas_libros, entradas_libros,
                entradas_extracto, gastos_bancarios, preparado_por, revisado_por, excel_data,
                workflow_status="Pendiente de revisión", id_edicion=id_edicion
            )
            st.session_state.pop("conciliacion_a_editar", None)
            st.session_state.pop("datos_cargados_edit", None)
            st.session_state["historial_id_destacado"] = int(id_g)
            st.session_state.menu_override = "📋 Historial"
            st.success(f"🚀 Conciliación CONC-{int(id_g):06d} enviada correctamente a Revisión.")

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
            es_destacada = int(fila["id"]) == int(st.session_state.get("historial_id_destacado", -1))
            if es_destacada:
                st.success(f"✅ Esta es la conciliación que acabas de guardar: **{consecutivo_str}** — Estado: **{wf_status}**")
                st.session_state.pop("historial_id_destacado", None)

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

                c_act1, c_act2, c_act3, c_act4 = st.columns(4)
                with c_act1:
                    if st.button("✏️ Editar Conciliación", key=f"btn_edit_{fila['id']}"):
                        st.session_state.conciliacion_a_editar = fila['id']
                        st.session_state.pop("datos_cargados_edit", None)
                        st.session_state.menu_override = "📝 Nueva Conciliación"
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

                with c_act4:
                    if rol_actual == "Administrador":
                        if st.button("🗑️ Eliminar", key=f"btn_delete_conc_{fila['id']}", use_container_width=True):
                            st.session_state.conciliacion_a_eliminar = int(fila['id'])
                            st.rerun()

                if rol_actual == "Administrador" and int(st.session_state.get("conciliacion_a_eliminar", -1)) == int(fila['id']):
                    st.warning(f"⚠️ ¿Está seguro de eliminar la conciliación **{consecutivo_str}**? Esta acción no se puede deshacer.")
                    d1, d2 = st.columns(2)
                    with d1:
                        if st.button("🗑️ Sí, eliminar conciliación", key=f"confirm_delete_conc_{fila['id']}", type="primary", use_container_width=True):
                            eliminar_conciliacion_db(int(fila['id']))
                            st.session_state.pop("conciliacion_a_eliminar", None)
                            st.success(f"✅ {consecutivo_str} eliminada correctamente.")
                            try:
                                st.cache_data.clear()
                            except Exception:
                                pass
                            st.rerun()
                    with d2:
                        if st.button("❌ Cancelar", key=f"cancel_delete_conc_{fila['id']}", use_container_width=True):
                            st.session_state.pop("conciliacion_a_eliminar", None)
                            st.rerun()

elif menu_seleccionado == "👥 Usuarios":
    st.title("👥 Gestión de Usuarios y Roles")
    if st.session_state.get("mensaje_usuario_editado"):
        st.success(st.session_state.pop("mensaje_usuario_editado"))
    usuarios_df = obtener_usuarios()

    if rol_actual == "Administrador":
        st.subheader("👥 Usuarios registrados")
        if usuarios_df.empty:
            st.info("No hay usuarios registrados.")
        else:
            for _, usr in usuarios_df.iterrows():
                uid = int(usr["id"])
                estado_txt = "🟢 Activo" if int(usr["activo"] or 0) == 1 else "🔴 Inactivo"
                with st.container(border=True):
                    c_info, c_edit, c_del = st.columns([5, 1.3, 1.3])
                    with c_info:
                        st.markdown(f"**{usr['nombre']}**  ·  `@{usr['usuario']}`")
                        st.caption(f"Rol: **{usr['rol']}** · Empresa: **{usr['empresa_nombre'] or 'Sin asignar'}** · {estado_txt}")
                    with c_edit:
                        if st.button("✏️ Editar", key=f"editar_usuario_{uid}", use_container_width=True):
                            st.session_state.usuario_a_editar = uid
                            st.rerun()
                    with c_del:
                        if st.button("🗑️ Eliminar", key=f"eliminar_usuario_{uid}", use_container_width=True, disabled=(uid == int(usuario_actual["id"]))):
                            st.session_state.usuario_a_eliminar = uid
                            st.rerun()

            edit_id = st.session_state.get("usuario_a_editar")
            if edit_id:
                fila_edit = usuarios_df[usuarios_df["id"] == int(edit_id)]
                if not fila_edit.empty:
                    usr_edit = fila_edit.iloc[0]
                    st.divider()
                    st.subheader(f"✏️ Editar usuario: {usr_edit['nombre']}")
                    empresas_ids = [None] + (empresas_df["id"].tolist() if not empresas_df.empty else [])
                    empresa_actual = None if pd.isna(usr_edit.get("empresa_nombre")) else None
                    if not empresas_df.empty:
                        m = empresas_df[empresas_df["nombre"] == usr_edit.get("empresa_nombre")]
                        empresa_actual = int(m.iloc[0]["id"]) if not m.empty else None
                    try:
                        idx_emp = empresas_ids.index(empresa_actual)
                    except ValueError:
                        idx_emp = 0
                    with st.form("form_editar_usuario"):
                        e1, e2 = st.columns(2)
                        with e1:
                            edit_nombre = st.text_input("Nombre completo", value=str(usr_edit["nombre"]))
                            edit_usuario = st.text_input("Usuario", value=str(usr_edit["usuario"]))
                            edit_rol = st.selectbox("Rol", ["Preparador", "Revisor", "Administrador"], index=["Preparador", "Revisor", "Administrador"].index(str(usr_edit["rol"])) if str(usr_edit["rol"]) in ["Preparador", "Revisor", "Administrador"] else 0)
                        with e2:
                            edit_activo = st.checkbox("Usuario activo", value=bool(int(usr_edit["activo"] or 0)))
