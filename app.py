import datetime
import os
import sqlite3
import zipfile
import pandas as pd
import streamlit as st

import analisis_financiero
import mod_reportes
import reporte_iva
import reporte_renta
from sync_sheets import (
    forzar_refresco_desde_sheets,
    inicializar_y_descargar_bd,
    subir_bd_a_sheets,
)

# ==========================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Sistema Contable NIIF - Costa Rica",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def aplicar_estilos_corporativos():
    st.markdown(
        """
        <style>
        .main-header {
            background: linear-gradient(90deg, #1B365D 0%, #2C5282 100%);
            padding: 1.1rem 1.5rem;
            border-radius: 10px;
            color: white;
            margin-bottom: 1.2rem;
            box-shadow: 0 2px 6px rgba(0,0,0,0.12);
        }
        .main-header h2 {
            margin: 0;
            font-size: 1.45rem;
            color: #FFFFFF;
        }
        .main-header p {
            margin: 0.25rem 0 0 0;
            font-size: 0.9rem;
            color: #E2E8F0;
        }
        div[data-testid="stMetric"] {
            background-color: #F8FAFC;
            border: 1px solid #E2E8F0;
            padding: 0.8rem 1rem;
            border-radius: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def renderizar_encabezado_corporativo(empresa_info, subtitulo=""):
    razon = empresa_info.get("razon_social", "Empresa")
    cedula = empresa_info.get("cedula", "")
    tipo = empresa_info.get("tipo_contribuyente", "")
    st.markdown(
        f"""
        <div class="main-header">
            <h2>🏢 {razon}</h2>
            <p>Cédula: <b>{cedula}</b> | Régimen: <b>{tipo}</b> | {subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


aplicar_estilos_corporativos()

DB_PATH = "sistema_contable.db"


@st.cache_resource
def obtener_conexion():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


conn = obtener_conexion()
inicializar_y_descargar_bd(conn)


def guardar_y_sincronizar(conn):
    conn.commit()
    subir_bd_a_sheets(conn)


def obtener_empresas(conn):
    return pd.read_sql_query(
        "SELECT * FROM empresas WHERE activa = 1 ORDER BY razon_social", conn
    )


def obtener_catalogo(conn, empresa_id):
    return pd.read_sql_query(
        """
        SELECT id, codigo, nombre, tipo, acepta_movimiento
        FROM catalogo_cuentas
        WHERE empresa_id = ?
        ORDER BY codigo
        """,
        conn,
        params=(int(empresa_id),),
    )


def ejecutar_backup_automatico(db_path):
    os.makedirs("backups", exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_zip = os.path.join("backups", f"backup_contable_{ts}.zip")
    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.exists(db_path):
            zf.write(db_path, arcname=os.path.basename(db_path))
        for archivo in os.listdir("."):
            if archivo.endswith(".py"):
                zf.write(archivo, arcname=archivo)
    return ruta_zip


def restaurar_desde_backup(zip_file, db_path):
    try:
        with zipfile.ZipFile(zip_file, "r") as zf:
            nombres = zf.namelist()
            db_en_zip = [n for n in nombres if n.endswith(".db")]
            if not db_en_zip:
                return False, "El archivo ZIP no contiene ninguna base de datos (.db)."
            data_db = zf.read(db_en_zip[0])
            with open(db_path, "wb") as f_out:
                f_out.write(data_db)
        return True, "Base de datos restaurada correctamente desde el respaldo."
    except Exception as e:
        return False, f"Error al restaurar respaldo: {e}"


# ==========================================
# BARRA LATERAL (SELECTOR DE EMPRESA Y MENÚ)
# ==========================================
with st.sidebar:
    st.markdown("## 📊 Contabilidad NIIF CR")
    st.caption("Gestión Contable, Tributaria y Financiera")

    df_empresas = obtener_empresas(conn)
    if df_empresas.empty:
        st.warning("No hay empresas activas registradas. Cree una en 'Empresas y Catálogo'.")
        empresa_id = None
        empresa_info = {}
    else:
        opciones_emp = {
            f"{r['razon_social']} ({r['cedula']})": int(r["id"])
            for _, r in df_empresas.iterrows()
        }
        emp_label = st.selectbox("🏢 Empresa Activa:", list(opciones_emp.keys()))
        empresa_id = opciones_emp[emp_label]
        empresa_info = df_empresas[df_empresas["id"] == empresa_id].iloc[0].to_dict()

    st.divider()

    menu = st.radio(
        "📌 Módulos del Sistema:",
        [
            "🏢 Empresas y Catálogo",
            "💰 Saldos Iniciales",
            "📝 Asientos Contables",
            "📥 Importar desde Excel",
            "🏦 Conciliación Bancaria",
            "📑 Reportes Financieros NIIF",
            "📊 Análisis Financiero Integral",
            "🧾 Liquidación IVA (D-104)",
            "🏛️ Impuesto Renta (D-101)",
            "🛡️ Respaldos y Nube",
        ],
    )

    st.divider()
    col_sync1, col_sync2 = st.columns(2)
    with col_sync1:
        if st.button("⬇️ Nube", use_container_width=True, help="Descargar última versión desde Google Sheets"):
            forzar_refresco_desde_sheets(conn)
            st.toast("Base de datos sincronizada desde Google Sheets.", icon="✅")
            st.rerun()
    with col_sync2:
        if st.button("⬆️ Subir", use_container_width=True, help="Respaldar BD actual en Google Sheets"):
            subir_bd_a_sheets(conn)
            st.toast("Datos respaldados en Google Sheets.", icon="☁️")


# ==========================================
# ENCABEZADO CORPORATIVO
# ==========================================
if empresa_id:
    renderizar_encabezado_corporativo(
        empresa_info,
        subtitulo=f"Módulo Activo: {menu}",
    )
else:
    st.title("📊 Sistema Contable NIIF - Costa Rica")


# ==========================================
# 1. MÓDULO: EMPRESAS Y CATÁLOGO DE CUENTAS
# ==========================================
if menu == "🏢 Empresas y Catálogo":
    tab_emp, tab_cat, tab_plantilla = st.tabs(
        ["🏢 Gestión de Empresas", "📒 Catálogo de Cuentas", "📋 Plantilla Maestra"]
    )

    with tab_emp:
        st.subheader("Registrar Nueva Empresa / Contribuyente")
        with st.form("form_nueva_empresa", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            tipo_contrib = c1.selectbox(
                "Tipo de Contribuyente",
                ["Persona Jurídica", "Persona Física", "Asociación / ASADA", "Régimen Simplificado"],
            )
            cedula = c2.text_input("Cédula Jurídica / Física *", placeholder="Ej: 3-101-123456")
            razon_social = c3.text_input("Razón Social / Nombre *", placeholder="Ej: Inversiones del Valle S.A.")

            c4, c5, c6 = st.columns(3)
            nombre_comercial = c4.text_input("Nombre Comercial")
            representante = c5.text_input("Representante Legal")
            fecha_inicio_act = c6.date_input("Fecha Inicio Actividades", value=datetime.date(2025, 1, 1))

            actividad = st.text_input("Actividad Económica (CIIU)")
            clonar_plantilla = st.checkbox("Copiar automáticamente la Plantilla Maestra de Cuentas", value=True)

            if st.form_submit_button("💾 Guardar Empresa", type="primary"):
                if not cedula.strip() or not razon_social.strip():
                    st.error("La cédula y la razón social son obligatorias.")
                else:
                    try:
                        cur = conn.cursor()
                        cur.execute(
                            """
                            INSERT INTO empresas (
                                tipo_contribuyente, cedula, razon_social, nombre_comercial,
                                representante_legal, actividad_economica, fecha_inicio_actividades, activa
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, 1)
                            """,
                            (
                                tipo_contrib,
                                cedula.strip(),
                                razon_social.strip(),
                                nombre_comercial.strip(),
                                representante.strip(),
                                actividad.strip(),
                                str(fecha_inicio_act),
                            ),
                        )
                        nueva_emp_id = cur.lastrowid
                        if clonar_plantilla:
                            cur.execute(
                                """
                                INSERT INTO catalogo_cuentas (empresa_id, codigo, nombre, tipo, acepta_movimiento)
                                SELECT ?, codigo, nombre, tipo, acepta_movimiento
                                FROM plantilla_catalogo_cuentas
                                """,
                                (nueva_emp_id,),
                            )
                        guardar_y_sincronizar(conn)
                        st.success(f"Empresa '{razon_social}' creada correctamente.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Ya existe una empresa registrada con esa cédula.")

        if not df_empresas.empty:
            st.divider()
            st.subheader("Empresas Registradas")
            st.dataframe(
                df_empresas[
                    ["id", "cedula", "razon_social", "tipo_contribuyente", "representante_legal", "fecha_inicio_actividades"]
                ],
                use_container_width=True,
                hide_index=True,
            )

    with tab_cat:
        if not empresa_id:
            st.info("Seleccione o cree una empresa primero.")
        else:
            st.subheader(f"Catálogo de Cuentas — {empresa_info['razon_social']}")
            with st.form("form_nueva_cuenta", clear_on_submit=True):
                c1, c2, c3, c4 = st.columns([2, 4, 2, 2])
                cod_cta = c1.text_input("Código *", placeholder="Ej: 1.1.01.05")
                nom_cta = c2.text_input("Nombre de la Cuenta *", placeholder="Ej: Banco Nacional Colones")
                tipo_cta = c3.selectbox("Tipo *", ["Activo", "Pasivo", "Patrimonio", "Ingreso", "Costo", "Gasto"])
                acepta_mov = c4.selectbox("¿Acepta Movimiento?", ["Sí (Detalle)", "No (Título/Mayor)"])

                if st.form_submit_button("➕ Agregar Cuenta", type="primary"):
                    if not cod_cta.strip() or not nom_cta.strip():
                        st.error("El código y el nombre son obligatorios.")
                    else:
                        existe = pd.read_sql_query(
                            "SELECT id FROM catalogo_cuentas WHERE empresa_id = ? AND codigo = ?",
                            conn,
                            params=(empresa_id, cod_cta.strip()),
                        )
                        if not existe.empty:
                            st.error(f"El código {cod_cta.strip()} ya existe en esta empresa.")
                        else:
                            conn.execute(
                                """
                                INSERT INTO catalogo_cuentas (empresa_id, codigo, nombre, tipo, acepta_movimiento)
                                VALUES (?, ?, ?, ?, ?)
                                """,
                                (
                                    empresa_id,
                                    cod_cta.strip(),
                                    nom_cta.strip(),
                                    tipo_cta,
                                    1 if acepta_mov.startswith("Sí") else 0,
                                ),
                            )
                            guardar_y_sincronizar(conn)
                            st.success("Cuenta agregada correctamente.")
                            st.rerun()

            df_cat = obtener_catalogo(conn, empresa_id)
            if not df_cat.empty:
                st.dataframe(
                    df_cat,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "id": "ID",
                        "codigo": "Código",
                        "nombre": "Nombre de Cuenta",
                        "tipo": "Elemento NIIF",
                        "acepta_movimiento": st.column_config.CheckboxColumn("Movimiento"),
                    },
                )

    with tab_plantilla:
        st.subheader("Importar / Actualizar Plantilla Maestra desde Excel")
        st.caption("Columnas requeridas: codigo, nombre, tipo, acepta_movimiento (1 o 0)")
        archivo_plant = st.file_uploader("Seleccionar archivo Excel (.xlsx)", type=["xlsx"], key="up_plantilla")
        if archivo_plant and st.button("📥 Cargar Plantilla Maestra"):
            try:
                df_p = pd.read_excel(archivo_plant)
                df_p.columns = [str(c).strip().lower() for c in df_p.columns]
                req_cols = {"codigo", "nombre", "tipo"}
                if not req_cols.issubset(set(df_p.columns)):
                    st.error(f"Faltan columnas requeridas: {req_cols}")
                else:
                    if "acepta_movimiento" not in df_p.columns:
                        df_p["acepta_movimiento"] = 1
                    cur = conn.cursor()
                    cur.execute("DELETE FROM plantilla_catalogo_cuentas")
                    for _, r in df_p.iterrows():
                        cur.execute(
                            """
                            INSERT OR REPLACE INTO plantilla_catalogo_cuentas (codigo, nombre, tipo, acepta_movimiento)
                            VALUES (?, ?, ?, ?)
                            """,
                            (
                                str(r["codigo"]).strip(),
                                str(r["nombre"]).strip(),
                                str(r["tipo"]).strip(),
                                int(r["acepta_movimiento"]) if pd.notna(r["acepta_movimiento"]) else 1,
                            ),
                        )
                    guardar_y_sincronizar(conn)
                    st.success(f"Plantilla actualizada con {len(df_p)} cuentas.")
            except Exception as e:
                st.error(f"Error procesando plantilla: {e}")


# ==========================================
# 2. MÓDULO: SALDOS INICIALES
# ==========================================
elif menu == "💰 Saldos Iniciales":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        st.subheader("Registro de Saldos Iniciales de Apertura")
        df_cat = obtener_catalogo(conn, empresa_id)
        df_mov = df_cat[df_cat["acepta_movimiento"] == 1].copy()

        if df_mov.empty:
            st.warning("Esta empresa no tiene cuentas de movimiento configuradas.")
        else:
            df_si = pd.read_sql_query(
                """
                SELECT si.id, c.codigo, c.nombre, c.tipo, si.fecha_inicio, si.debe, si.haber
                FROM saldos_iniciales si
                JOIN catalogo_cuentas c ON si.cuenta_id = c.id
                WHERE si.empresa_id = ?
                ORDER BY c.codigo
                """,
                conn,
                params=(empresa_id,),
            )

            with st.form("form_saldo_inicial", clear_on_submit=True):
                c1, c2, c3, c4 = st.columns([4, 2, 2, 2])
                opciones_cta = {
                    f"{r['codigo']} - {r['nombre']} ({r['tipo']})": int(r["id"])
                    for _, r in df_mov.iterrows()
                }
                cta_sel = c1.selectbox("Cuenta Contable", list(opciones_cta.keys()))
                fecha_si = c2.date_input("Fecha Apertura", value=datetime.date(2025, 1, 1))
                debe_si = c3.number_input("Saldo Debe (₡)", min_value=0.0, step=100.0, format="%.2f")
                haber_si = c4.number_input("Saldo Haber (₡)", min_value=0.0, step=100.0, format="%.2f")

                if st.form_submit_button("💾 Registrar / Actualizar Saldo Inicial", type="primary"):
                    if debe_si > 0 and haber_si > 0:
                        st.error("Ingrese el monto únicamente en Debe o únicamente en Haber, no en ambos.")
                    else:
                        cta_id_sel = opciones_cta[cta_sel]
                        conn.execute(
                            "DELETE FROM saldos_iniciales WHERE empresa_id = ? AND cuenta_id = ?",
                            (empresa_id, cta_id_sel),
                        )
                        if debe_si > 0 or haber_si > 0:
                            conn.execute(
                                """
                                INSERT INTO saldos_iniciales (empresa_id, cuenta_id, fecha_inicio, debe, haber)
                                VALUES (?, ?, ?, ?, ?)
                                """,
                                (empresa_id, cta_id_sel, str(fecha_si), float(debe_si), float(haber_si)),
                            )
                        guardar_y_sincronizar(conn)
                        st.success("Saldo inicial actualizado.")
                        st.rerun()

            if not df_si.empty:
                tot_debe = df_si["debe"].sum()
                tot_haber = df_si["haber"].sum()
                dif = round(tot_debe - tot_haber, 2)

                m1, m2, m3 = st.columns(3)
                m1.metric("Total Debe Apertura", f"₡ {tot_debe:,.2f}")
                m2.metric("Total Haber Apertura", f"₡ {tot_haber:,.2f}")
                m3.metric("Diferencia (Debe - Haber)", f"₡ {dif:,.2f}", delta="Cuadrado" if dif == 0 else "Descuadrado")

                st.dataframe(df_si, use_container_width=True, hide_index=True)


# ==========================================
# 3. MÓDULO: ASIENTOS CONTABLES (LIBRO DIARIO)
# ==========================================
elif menu == "📝 Asientos Contables":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        tab_reg, tab_cons = st.tabs(["✍️ Registrar Nuevo Asiento", "📖 Consultar Libro Diario"])

        with tab_reg:
            df_cat = obtener_catalogo(conn, empresa_id)
            df_mov = df_cat[df_cat["acepta_movimiento"] == 1].copy()

            if df_mov.empty:
                st.warning("Configure primero el catálogo de cuentas.")
            else:
                cur = conn.cursor()
                max_num = cur.execute(
                    "SELECT COALESCE(MAX(num_asiento), 0) FROM asientos_cabecera WHERE empresa_id = ?",
                    (empresa_id,),
                ).fetchone()[0]

                c1, c2, c3, c4 = st.columns(4)
                num_asiento = c1.number_input("N° Asiento", min_value=1, value=int(max_num) + 1, step=1)
                fecha_asiento = c2.date_input("Fecha Contable", value=datetime.date.today())
                tipo_asiento = c3.selectbox("Tipo de Asiento", ["Diario", "Ingresos", "Egresos", "Ajuste", "Cierre"])
                num_comp = c4.text_input("N° Comprobante / Factura", placeholder="Ej: FE-00125")

                concepto = st.text_input("Concepto / Descripción General *", placeholder="Ej: Registro de factura de servicios profesionales")

                c5, c6, c7 = st.columns(3)
                forma_pago = c5.selectbox("Forma de Pago", ["Transferencia / SINPE", "Efectivo", "Cheque", "Tarjeta", "Crédito / N/A"])
                num_cheque = c6.text_input("N° Cheque / Referencia Bancaria")
                usuario_crea = c7.text_input("Elaborado por", value="Lic. Roy F. Esquivel González")

                st.markdown("#### Líneas del Asiento (Partida Doble)")
                opciones_cuentas = [f"{r['codigo']} - {r['nombre']}" for _, r in df_mov.iterrows()]
                mapa_cuentas = {f"{r['codigo']} - {r['nombre']}": int(r["id"]) for _, r in df_mov.iterrows()}

                if "lineas_asiento_df" not in st.session_state:
                    st.session_state["lineas_asiento_df"] = pd.DataFrame(
                        {
                            "Cuenta": [opciones_cuentas[0], opciones_cuentas[0]],
                            "Debe": [0.0, 0.0],
                            "Haber": [0.0, 0.0],
                        }
                    )

                df_editado = st.data_editor(
                    st.session_state["lineas_asiento_df"],
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "Cuenta": st.column_config.SelectboxColumn(
                            "Cuenta Contable", options=opciones_cuentas, required=True, width="large"
                        ),
                        "Debe": st.column_config.NumberColumn("Debe (₡)", min_value=0.0, format="%.2f", default=0.0),
                        "Haber": st.column_config.NumberColumn("Haber (₡)", min_value=0.0, format="%.2f", default=0.0),
                    },
                    key="editor_asientos",
                )

                tot_d = float(df_editado["Debe"].fillna(0).sum())
                tot_h = float(df_editado["Haber"].fillna(0).sum())
                dif_asiento = round(tot_d - tot_h, 2)

                k1, k2, k3 = st.columns(3)
                k1.metric("Total Debe", f"₡ {tot_d:,.2f}")
                k2.metric("Total Haber", f"₡ {tot_h:,.2f}")
                k3.metric("Diferencia", f"₡ {dif_asiento:,.2f}")

                if st.button("💾 Contabilizar Asiento", type="primary"):
                    lineas_validas = df_editado[
                        (df_editado["Debe"].fillna(0) > 0) | (df_editado["Haber"].fillna(0) > 0)
                    ]
                    if not concepto.strip():
                        st.error("Debe indicar el concepto del asiento.")
                    elif len(lineas_validas) < 2:
                        st.error("Un asiento contable requiere al menos 2 líneas con movimiento.")
                    elif abs(dif_asiento) >= 0.01:
                        st.error(f"El asiento está descuadrado por ₡ {dif_asiento:,.2f}. Debe = Haber.")
                    else:
                        periodo_str = fecha_asiento.strftime("%Y-%m")
                        ahora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        cur.execute(
                            """
                            INSERT INTO asientos_cabecera (
                                empresa_id, num_asiento, fecha, periodo, concepto, tipo_asiento,
                                fecha_registro, fecha_comprobante, forma_pago, num_comprobante,
                                num_cheque, usuario_creacion
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                empresa_id,
                                int(num_asiento),
                                str(fecha_asiento),
                                periodo_str,
                                concepto.strip(),
                                tipo_asiento,
                                ahora_str,
                                str(fecha_asiento),
                                forma_pago,
                                num_comp.strip(),
                                num_cheque.strip(),
                                usuario_crea.strip(),
                            ),
                        )
                        nuevo_asiento_id = cur.lastrowid
                        for _, linea in lineas_validas.iterrows():
                            cid = mapa_cuentas[linea["Cuenta"]]
                            d_val = float(linea["Debe"] or 0.0)
                            h_val = float(linea["Haber"] or 0.0)
                            cur.execute(
                                """
                                INSERT INTO asientos_detalle (asiento_id, cuenta_id, debe, haber, conciliado)
                                VALUES (?, ?, ?, ?, 0)
                                """,
                                (nuevo_asiento_id, cid, d_val, h_val),
                            )

                        guardar_y_sincronizar(conn)
                        st.session_state.pop("lineas_asiento_df", None)
                        st.success(f"Asiento #{num_asiento} contabilizado exitosamente.")
                        st.rerun()

        with tab_cons:
            st.subheader("Consulta de Asientos Registrados")
            f1, f2 = st.columns(2)
            f_desde = f1.date_input("Desde", value=datetime.date(datetime.date.today().year, 1, 1), key="ld_desde")
            f_hasta = f2.date_input("Hasta", value=datetime.date.today(), key="ld_hasta")

            df_asientos = pd.read_sql_query(
                """
                SELECT ac.id, ac.num_asiento AS "Asiento", ac.fecha AS "Fecha", ac.tipo_asiento AS "Tipo",
                       ac.num_comprobante AS "Comprobante", ac.concepto AS "Concepto",
                       COALESCE(SUM(ad.debe), 0) AS "Total Debe",
                       COALESCE(SUM(ad.haber), 0) AS "Total Haber"
                FROM asientos_cabecera ac
                LEFT JOIN asientos_detalle ad ON ac.id = ad.asiento_id
                WHERE ac.empresa_id = ? AND ac.fecha BETWEEN ? AND ?
                GROUP BY ac.id
                ORDER BY ac.fecha DESC, ac.num_asiento DESC
                """,
                conn,
                params=(empresa_id, str(f_desde), str(f_hasta)),
            )

            if df_asientos.empty:
                st.info("No hay asientos en el rango seleccionado.")
            else:
                st.dataframe(df_asientos.drop(columns=["id"]), use_container_width=True, hide_index=True)

                asiento_ver = st.selectbox(
                    "Ver detalle o eliminar asiento:",
                    df_asientos["id"].tolist(),
                    format_func=lambda x: (
                        f"Asiento #{df_asientos[df_asientos['id']==x]['Asiento'].values[0]} | "
                        f"{df_asientos[df_asientos['id']==x]['Fecha'].values[0]} - "
                        f"{df_asientos[df_asientos['id']==x]['Concepto'].values[0]}"
                    ),
                )
                if asiento_ver:
                    df_det = pd.read_sql_query(
                        """
                        SELECT c.codigo AS "Código", c.nombre AS "Cuenta", ad.debe AS "Debe", ad.haber AS "Haber"
                        FROM asientos_detalle ad
                        JOIN catalogo_cuentas c ON ad.cuenta_id = c.id
                        WHERE ad.asiento_id = ?
                        """,
                        conn,
                        params=(int(asiento_ver),),
                    )
                    st.table(df_det)
                    if st.button("🗑️ Eliminar este Asiento", type="secondary"):
                        conn.execute("DELETE FROM asientos_detalle WHERE asiento_id = ?", (int(asiento_ver),))
                        conn.execute("DELETE FROM asientos_cabecera WHERE id = ?", (int(asiento_ver),))
                        guardar_y_sincronizar(conn)
                        st.warning("Asiento eliminado.")
                        st.rerun()


# ==========================================
# 4. MÓDULO: IMPORTACIÓN MASIVA DESDE EXCEL
# ==========================================
elif menu == "📥 Importar desde Excel":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        st.subheader("Importación Masiva de Asientos desde Excel")
        st.info(
            "El archivo Excel debe contener las columnas: "
            "`num_asiento`, `fecha`, `concepto`, `codigo_cuenta`, `debe`, `haber` "
            "(opcionales: `tipo_asiento`, `num_comprobante`, `forma_pago`)."
        )
        archivo_excel = st.file_uploader("Subir archivo Excel (.xlsx)", type=["xlsx"], key="up_asientos")
        if archivo_excel and st.button("🚀 Procesar e Importar Asientos", type="primary"):
            try:
                df_imp = pd.read_excel(archivo_excel)
                df_imp.columns = [str(c).strip().lower() for c in df_imp.columns]
                req = {"num_asiento", "fecha", "concepto", "codigo_cuenta", "debe", "haber"}
                if not req.issubset(set(df_imp.columns)):
                    st.error(f"Faltan columnas obligatorias en el Excel: {req - set(df_imp.columns)}")
                else:
                    df_cat = obtener_catalogo(conn, empresa_id)
                    mapa_cod = {str(r["codigo"]).strip(): int(r["id"]) for _, r in df_cat.iterrows()}
                    df_imp["codigo_cuenta"] = df_imp["codigo_cuenta"].astype(str).str.strip()
                    df_imp["debe"] = pd.to_numeric(df_imp["debe"], errors="coerce").fillna(0.0)
                    df_imp["haber"] = pd.to_numeric(df_imp["haber"], errors="coerce").fillna(0.0)

                    faltantes = set(df_imp["codigo_cuenta"]) - set(mapa_cod.keys())
                    if faltantes:
                        st.error(f"Códigos de cuenta inexistentes en el catálogo: {', '.join(sorted(faltantes))}")
                    else:
                        cur = conn.cursor()
                        insertados = 0
                        for num_a, grupo in df_imp.groupby("num_asiento"):
                            dif_g = round(grupo["debe"].sum() - grupo["haber"].sum(), 2)
                            if abs(dif_g) >= 0.01:
                                st.error(f"El asiento #{num_a} está descuadrado por ₡ {dif_g:,.2f}. Importación cancelada.")
                                conn.rollback()
                                st.stop()
                            prim = grupo.iloc[0]
                            fecha_str = pd.to_datetime(prim["fecha"]).strftime("%Y-%m-%d")
                            periodo_str = fecha_str[:7]
                            cur.execute(
                                """
                                INSERT INTO asientos_cabecera (
                                    empresa_id, num_asiento, fecha, periodo, concepto, tipo_asiento,
                                    fecha_registro, forma_pago, num_comprobante, usuario_creacion
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    empresa_id,
                                    int(num_a),
                                    fecha_str,
                                    periodo_str,
                                    str(prim["concepto"]).strip(),
                                    str(prim.get("tipo_asiento", "Diario")),
                                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    str(prim.get("forma_pago", "Transferencia / SINPE")),
                                    str(prim.get("num_comprobante", "")),
                                    "Importación Excel",
                                ),
                            )
                            aid = cur.lastrowid
                            for _, r in grupo.iterrows():
                                cur.execute(
                                    """
                                    INSERT INTO asientos_detalle (asiento_id, cuenta_id, debe, haber, conciliado)
                                    VALUES (?, ?, ?, ?, 0)
                                    """,
                                    (aid, mapa_cod[r["codigo_cuenta"]], float(r["debe"]), float(r["haber"])),
                                )
                            insertados += 1

                        guardar_y_sincronizar(conn)
                        st.success(f"Se importaron {insertados} asientos contables correctamente.")
            except Exception as e:
                conn.rollback()
                st.error(f"Error durante la importación: {e}")


# ==========================================
# 5. MÓDULO: CONCILIACIÓN BANCARIA
# ==========================================
elif menu == "🏦 Conciliación Bancaria":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        st.subheader("Conciliación Bancaria Mensual")
        df_cat = obtener_catalogo(conn, empresa_id)
        df_bancos = df_cat[
            (df_cat["acepta_movimiento"] == 1)
            & (
                df_cat["codigo"].str.startswith("1.1.01")
                | df_cat["nombre"].str.upper().str.contains("BANCO|CUENTA CORRIENTE|AHORRO|BNCR|BCR|BAC")
            )
        ]
        if df_bancos.empty:
            df_bancos = df_cat[df_cat["acepta_movimiento"] == 1]

        if df_bancos.empty:
            st.warning("No hay cuentas bancarias disponibles en el catálogo.")
        else:
            c1, c2, c3, c4 = st.columns([4, 2, 2, 2])
            opciones_bco = {f"{r['codigo']} - {r['nombre']}": int(r["id"]) for _, r in df_bancos.iterrows()}
            bco_sel = c1.selectbox("Cuenta Bancaria", list(opciones_bco.keys()))
            cta_bco_id = opciones_bco[bco_sel]
            anio_conc = c2.number_input("Año", min_value=2020, max_value=2035, value=datetime.date.today().year)
            mes_conc = c3.number_input("Mes", min_value=1, max_value=12, value=datetime.date.today().month)
            saldo_estado_bco = c4.number_input("Saldo según Estado de Cuenta Banco (₡)", value=0.0, format="%.2f")

            fecha_ini_mes = datetime.date(int(anio_conc), int(mes_conc), 1)
            if int(mes_conc) == 12:
                fecha_fin_mes = datetime.date(int(anio_conc), 12, 31)
            else:
                fecha_fin_mes = datetime.date(int(anio_conc), int(mes_conc) + 1, 1) - datetime.timedelta(days=1)

            si_row = pd.read_sql_query(
                "SELECT COALESCE(SUM(debe - haber), 0) AS s FROM saldos_iniciales WHERE empresa_id = ? AND cuenta_id = ?",
                conn,
                params=(empresa_id, cta_bco_id),
            ).iloc[0]["s"]

            mov_ant = pd.read_sql_query(
                """
                SELECT COALESCE(SUM(ad.debe - ad.haber), 0) AS s
                FROM asientos_detalle ad
                JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
                WHERE ac.empresa_id = ? AND ad.cuenta_id = ? AND ac.fecha < ?
                """,
                conn,
                params=(empresa_id, cta_bco_id, str(fecha_ini_mes)),
            ).iloc[0]["s"]

            saldo_libros_ini = float(si_row) + float(mov_ant)

            df_movs_mes = pd.read_sql_query(
                """
                SELECT ad.id AS detalle_id, ac.fecha AS Fecha, ac.num_asiento AS Asiento,
                       ac.num_comprobante AS Documento, ac.concepto AS Concepto,
                       COALESCE(ad.debe, 0.0) AS Depósitos, COALESCE(ad.haber, 0.0) AS Retiros,
                       CAST(COALESCE(ad.conciliado, 0) AS INTEGER) AS Conciliado
                FROM asientos_detalle ad
                JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
                WHERE ac.empresa_id = ? AND ad.cuenta_id = ? AND ac.fecha BETWEEN ? AND ?
                ORDER BY ac.fecha, ac.num_asiento
                """,
                conn,
                params=(empresa_id, cta_bco_id, str(fecha_ini_mes), str(fecha_fin_mes)),
            )

            if df_movs_mes.empty:
                df_movs_mes = pd.DataFrame(
                    columns=["detalle_id", "Fecha", "Asiento", "Documento", "Concepto", "Depósitos", "Retiros", "Conciliado"]
                )

            df_movs_mes["Depósitos"] = pd.to_numeric(df_movs_mes["Depósitos"], errors="coerce").fillna(0.0)
            df_movs_mes["Retiros"] = pd.to_numeric(df_movs_mes["Retiros"], errors="coerce").fillna(0.0)
            df_movs_mes["Conciliado"] = df_movs_mes["Conciliado"].astype(bool)

            saldo_libros_fin = saldo_libros_ini + float(
                df_movs_mes["Depósitos"].sum() - df_movs_mes["Retiros"].sum()
            )

            df_conc_edit = st.data_editor(
                df_movs_mes,
                use_container_width=True,
                hide_index=True,
                disabled=["detalle_id", "Fecha", "Asiento", "Documento", "Concepto", "Depósitos", "Retiros"],
                key="editor_conciliacion",
            )

            if df_conc_edit.empty:
                depositos_transito = 0.0
                cheques_transito = 0.0
            else:
                no_conc = df_conc_edit[~df_conc_edit["Conciliado"]]
                depositos_transito = float(pd.to_numeric(no_conc["Depósitos"], errors="coerce").fillna(0.0).sum())
                cheques_transito = float(pd.to_numeric(no_conc["Retiros"], errors="coerce").fillna(0.0).sum())

            saldo_conciliado = saldo_estado_bco + depositos_transito - cheques_transito
            diferencia_conc = round(saldo_conciliado - saldo_libros_fin, 2)

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Saldo Libros Final", f"₡ {saldo_libros_fin:,.2f}")
            k2.metric("Depósitos en Tránsito (+)", f"₡ {depositos_transito:,.2f}")
            k3.metric("Cheques/Pagos Tránsito (-)", f"₡ {cheques_transito:,.2f}")
            k4.metric("Diferencia Conciliación", f"₡ {diferencia_conc:,.2f}")

            if st.button("💾 Guardar Conciliación Bancaria", type="primary"):
                cur = conn.cursor()
                for _, row in df_conc_edit.iterrows():
                    cur.execute(
                        "UPDATE asientos_detalle SET conciliado = ? WHERE id = ?",
                        (1 if row["Conciliado"] else 0, int(row["detalle_id"])),
                    )
                cur.execute(
                    "DELETE FROM conciliaciones_bancarias WHERE empresa_id = ? AND cuenta_id = ? AND anio = ? AND mes = ?",
                    (empresa_id, cta_bco_id, int(anio_conc), int(mes_conc)),
                )
                cur.execute(
                    """
                    INSERT INTO conciliaciones_bancarias (
                        empresa_id, cuenta_id, anio, mes, fecha_desde, fecha_hasta,
                        saldo_banco, saldo_libros_ini, saldo_libros_fin, saldo_conciliado,
                        diferencia, fecha_actualizacion
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        empresa_id,
                        cta_bco_id,
                        int(anio_conc),
                        int(mes_conc),
                        str(fecha_ini_mes),
                        str(fecha_fin_mes),
                        float(saldo_estado_bco),
                        float(saldo_libros_ini),
                        float(saldo_libros_fin),
                        float(saldo_conciliado),
                        float(diferencia_conc),
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ),
                )
                guardar_y_sincronizar(conn)
                st.success("Conciliación bancaria guardada y sincronizada.")


# ==========================================
# 6. MÓDULO: REPORTES FINANCIEROS NIIF
# ==========================================
elif menu == "📑 Reportes Financieros NIIF":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        mod_reportes.renderizar_modulo_reportes(conn, empresa_id, empresa_info)


# ==========================================
# 7. MÓDULO: ANÁLISIS FINANCIERO INTEGRAL
# ==========================================
elif menu == "📊 Análisis Financiero Integral":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        analisis_financiero.renderizar_modulo_analisis_financiero(conn, empresa_id, empresa_info)


# ==========================================
# 8. MÓDULO: LIQUIDACIÓN IVA (D-104)
# ==========================================
elif menu == "🧾 Liquidación IVA (D-104)":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        reporte_iva.renderizar_modulo_iva(conn, empresa_id, empresa_info)


# ==========================================
# 9. MÓDULO: IMPUESTO SOBRE LA RENTA (D-101)
# ==========================================
elif menu == "🏛️ Impuesto Renta (D-101)":
    if not empresa_id:
        st.info("Seleccione una empresa en el panel izquierdo.")
    else:
        reporte_renta.renderizar_modulo_renta(conn, empresa_id, empresa_info)


# ==========================================
# 10. MÓDULO: RESPALDOS Y NUBE
# ==========================================
elif menu == "🛡️ Respaldos y Nube":
    st.subheader("Respaldo Integral y Sincronización")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### 📦 Respaldo Local (.ZIP)")
        st.caption("Genera una copia comprimida de la base de datos SQLite y del código fuente.")
        if st.button("Crear Respaldo Ahora", type="primary"):
            ruta_zip = ejecutar_backup_automatico(DB_PATH)
            if ruta_zip and os.path.exists(ruta_zip):
                with open(ruta_zip, "rb") as f:
                    st.download_button(
                        label="⬇️ Descargar Respaldo (.ZIP)",
                        data=f.read(),
                        file_name=os.path.basename(ruta_zip),
                        mime="application/zip",
                    )
                st.success(f"Respaldo generado: {ruta_zip}")

        st.divider()
        st.markdown("### ♻️ Restaurar desde Respaldo (.ZIP)")
        zip_subido = st.file_uploader("Seleccionar archivo de respaldo (.zip)", type=["zip"])
        if zip_subido and st.button("⚠️ Restaurar Base de Datos"):
            ok, msg = restaurar_desde_backup(zip_subido, DB_PATH)
            if ok:
                st.cache_resource.clear()
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    with c2:
        st.markdown("### ☁️ Sincronización con Google Sheets")
        st.caption("Hoja vinculada: `BD_Sistema_Contable`")
        if st.button("⬆️ Forzar Subida Completa a Google Sheets", use_container_width=True):
            st.session_state.pop("ultima_firma_sheets", None)
            subir_bd_a_sheets(conn)
            st.success("Todas las tablas fueron enviadas a Google Sheets.")

        if st.button("⬇️ Forzar Descarga desde Google Sheets", use_container_width=True):
            forzar_refresco_desde_sheets(conn)
            st.success("Base de datos local actualizada desde Google Sheets.")
            st.rerun()

        if os.path.exists(DB_PATH):
            with open(DB_PATH, "rb") as db_file:
                st.download_button(
                    label="💾 Descargar archivo SQLite directo (.db)",
                    data=db_file.read(),
                    file_name=f"sistema_contable_{datetime.date.today()}.db",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
