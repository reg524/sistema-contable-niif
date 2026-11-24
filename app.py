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
import tc_bccr
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
            <p>Cédula: <b>{cedula}</b> | Régimen: <b>{tipo}</b> | Moneda Base: <b>Colones (₡)</b> | {subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _parsear_fecha_segura(valor, por_defecto=datetime.date(2025, 1, 1)):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return por_defecto
    texto = str(valor).strip()
    if not texto or texto.lower() in ("none", "nan", "nat"):
        return por_defecto
    try:
        dt = pd.to_datetime(texto, errors="coerce")
        if pd.isna(dt):
            return por_defecto
        return dt.date()
    except Exception:
        return por_defecto


aplicar_estilos_corporativos()

DB_PATH = "sistema_contable.db"


@st.cache_resource
def obtener_conexion():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def asegurar_columnas_multimoneda(conn):
    """Agrega columnas multimoneda de forma segura si aún no existen."""
    cols_cab = {
        r[1] for r in conn.execute("PRAGMA table_info(asientos_cabecera)").fetchall()
    }
    if "moneda_dig" not in cols_cab:
        conn.execute("ALTER TABLE asientos_cabecera ADD COLUMN moneda_dig TEXT DEFAULT 'CRC'")
    if "tipo_cambio_bccr" not in cols_cab:
        conn.execute("ALTER TABLE asientos_cabecera ADD COLUMN tipo_cambio_bccr REAL DEFAULT 1.0")
    if "modalidad_tc" not in cols_cab:
        conn.execute("ALTER TABLE asientos_cabecera ADD COLUMN modalidad_tc TEXT DEFAULT 'N/A'")

    cols_det = {
        r[1] for r in conn.execute("PRAGMA table_info(asientos_detalle)").fetchall()
    }
    if "moneda_origen" not in cols_det:
        conn.execute("ALTER TABLE asientos_detalle ADD COLUMN moneda_origen TEXT DEFAULT 'CRC'")
    if "monto_usd_debe" not in cols_det:
        conn.execute("ALTER TABLE asientos_detalle ADD COLUMN monto_usd_debe REAL DEFAULT 0.0")
    if "monto_usd_haber" not in cols_det:
        conn.execute("ALTER TABLE asientos_detalle ADD COLUMN monto_usd_haber REAL DEFAULT 0.0")
    if "tipo_cambio_aplicado" not in cols_det:
        conn.execute("ALTER TABLE asientos_detalle ADD COLUMN tipo_cambio_aplicado REAL DEFAULT 1.0")
    conn.commit()


conn = obtener_conexion()
inicializar_y_descargar_bd(conn)
asegurar_columnas_multimoneda(conn)
tc_bccr.asegurar_tabla_tc(conn)


def guardar_y_sincronizar(conn):
    conn.commit()
    subir_bd_a_sheets(conn)


def obtener_empresas(conn):
    return pd.read_sql_query(
        "SELECT * FROM empresas WHERE activa = 1 ORDER BY razon_social", conn
    )


def obtener_todas_empresas(conn):
    return pd.read_sql_query(
        "SELECT * FROM empresas ORDER BY id DESC", conn
    )


def _tabla_existe(conn, nombre_tabla):
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (nombre_tabla,),
    ).fetchone()
    return row is not None


def eliminar_empresa_completa(conn, emp_id):
    """Elimina una empresa y sus registros asociados de forma segura."""
    emp_id = int(emp_id)
    cur = conn.cursor()

    if _tabla_existe(conn, "asientos_detalle") and _tabla_existe(conn, "asientos_cabecera"):
        cur.execute(
            """
            DELETE FROM asientos_detalle
            WHERE asiento_id IN (SELECT id FROM asientos_cabecera WHERE empresa_id = ?)
            """,
            (emp_id,),
        )
        cur.execute("DELETE FROM asientos_cabecera WHERE empresa_id = ?", (emp_id,))

    for tabla in [
        "saldos_iniciales",
        "conciliaciones_bancarias",
        "declaraciones_iva",
        "declaraciones_renta",
        "catalogo_cuentas",
    ]:
        if _tabla_existe(conn, tabla):
            cur.execute(f"DELETE FROM {tabla} WHERE empresa_id = ?", (emp_id,))

    cur.execute("DELETE FROM empresas WHERE id = ?", (emp_id,))
    guardar_y_sincronizar(conn)


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


def crear_cuenta_catalogo(conn, empresa_id, codigo, nombre, tipo, acepta_movimiento=1):
    """Crea una cuenta contable validando duplicados por código en la empresa."""
    cod = str(codigo or "").strip()
    nom = str(nombre or "").strip()
    if not cod or not nom:
        return False, "El código y el nombre de la cuenta son obligatorios."

    existe = pd.read_sql_query(
        "SELECT id FROM catalogo_cuentas WHERE empresa_id = ? AND codigo = ?",
        conn,
        params=(int(empresa_id), cod),
    )
    if not existe.empty:
        return False, f"El código '{cod}' ya existe en esta empresa."

    conn.execute(
        """
        INSERT INTO catalogo_cuentas (empresa_id, codigo, nombre, tipo, acepta_movimiento)
        VALUES (?, ?, ?, ?, ?)
        """,
        (int(empresa_id), cod, nom, str(tipo).strip(), int(acepta_movimiento)),
    )
    guardar_y_sincronizar(conn)
    return True, f"Cuenta '{cod} - {nom}' creada correctamente."


def actualizar_cuenta_catalogo(conn, empresa_id, cuenta_id, codigo, nombre, tipo, acepta_movimiento):
    """Modifica los datos de una cuenta existente en la empresa."""
    cod = str(codigo or "").strip()
    nom = str(nombre or "").strip()
    if not cod or not nom:
        return False, "El código y el nombre de la cuenta son obligatorios."

    duplicado = pd.read_sql_query(
        "SELECT id FROM catalogo_cuentas WHERE empresa_id = ? AND codigo = ? AND id != ?",
        conn,
        params=(int(empresa_id), cod, int(cuenta_id)),
    )
    if not duplicado.empty:
        return False, f"Ya existe otra cuenta con el código '{cod}' en esta empresa."

    conn.execute(
        """
        UPDATE catalogo_cuentas
        SET codigo = ?, nombre = ?, tipo = ?, acepta_movimiento = ?
        WHERE id = ? AND empresa_id = ?
        """,
        (cod, nom, str(tipo).strip(), int(acepta_movimiento), int(cuenta_id), int(empresa_id)),
    )
    guardar_y_sincronizar(conn)
    return True, f"Cuenta '{cod} - {nom}' actualizada correctamente."


def eliminar_cuenta_catalogo(conn, empresa_id, cuenta_id):
    """Elimina una cuenta solo si no tiene asientos ni saldos iniciales asociados."""
    cid = int(cuenta_id)
    eid = int(empresa_id)

    en_asientos = conn.execute(
        """
        SELECT COUNT(*)
        FROM asientos_detalle ad
        JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
        WHERE ac.empresa_id = ? AND ad.cuenta_id = ?
        """,
        (eid, cid),
    ).fetchone()[0]

    en_saldos = conn.execute(
        "SELECT COUNT(*) FROM saldos_iniciales WHERE empresa_id = ? AND cuenta_id = ?",
        (eid, cid),
    ).fetchone()[0]

    if en_asientos > 0 or en_saldos > 0:
        return (
            False,
            f"No se puede eliminar la cuenta porque tiene {en_asientos} línea(s) en asientos y {en_saldos} saldo(s) inicial(es).",
        )

    conn.execute(
        "DELETE FROM catalogo_cuentas WHERE id = ? AND empresa_id = ?",
        (cid, eid),
    )
    guardar_y_sincronizar(conn)
    return True, "Cuenta eliminada correctamente."


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
                        st.session_state["emp_editar_sel_id"] = int(nueva_emp_id)
                        st.success(f"Empresa '{razon_social}' creada correctamente.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Ya existe una empresa registrada con esa cédula.")

        df_todas_emp = obtener_todas_empresas(conn)

        if not df_todas_emp.empty:
            st.divider()
            st.subheader("Empresas Registradas")
            st.dataframe(
                df_todas_emp[
                    [
                        "id",
                        "cedula",
                        "razon_social",
                        "nombre_comercial",
                        "tipo_contribuyente",
                        "representante_legal",
                        "fecha_inicio_actividades",
                        "activa",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

            st.divider()
            st.subheader("✏️ Modificar o 🗑️ Eliminar Empresa")

            ids_disponibles = [int(x) for x in df_todas_emp["id"].tolist()]
            default_id = st.session_state.get("emp_editar_sel_id", empresa_id or ids_disponibles[0])
            if default_id not in ids_disponibles:
                default_id = ids_disponibles[0]

            emp_sel_id = st.selectbox(
                "Seleccione la empresa a modificar o eliminar:",
                ids_disponibles,
                index=ids_disponibles.index(default_id),
                format_func=lambda x: (
                    f"[{x}] {df_todas_emp[df_todas_emp['id']==x]['razon_social'].values[0]} "
                    f"({df_todas_emp[df_todas_emp['id']==x]['cedula'].values[0]})"
                ),
                key="selector_gestion_empresa",
            )
            st.session_state["emp_editar_sel_id"] = int(emp_sel_id)
            emp_row = df_todas_emp[df_todas_emp["id"] == emp_sel_id].iloc[0]

            tipos_posibles = [
                "Persona Jurídica",
                "Persona Física",
                "Asociación / ASADA",
                "Régimen Simplificado",
            ]
            tipo_actual = str(emp_row.get("tipo_contribuyente") or tipos_posibles[0])
            if tipo_actual not in tipos_posibles:
                tipos_posibles.append(tipo_actual)

            fecha_act_dt = _parsear_fecha_segura(emp_row.get("fecha_inicio_actividades"))

            with st.form(f"form_editar_empresa_{emp_sel_id}"):
                e1, e2, e3 = st.columns(3)
                ed_tipo = e1.selectbox(
                    "Tipo de Contribuyente",
                    tipos_posibles,
                    index=tipos_posibles.index(tipo_actual),
                )
                ed_cedula = e2.text_input("Cédula Jurídica / Física *", value=str(emp_row.get("cedula") or ""))
                ed_razon = e3.text_input("Razón Social / Nombre *", value=str(emp_row.get("razon_social") or ""))

                e4, e5, e6 = st.columns(3)
                ed_nom_com = e4.text_input("Nombre Comercial", value=str(emp_row.get("nombre_comercial") or ""))
                ed_rep = e5.text_input("Representante Legal", value=str(emp_row.get("representante_legal") or ""))
                ed_fecha = e6.date_input("Fecha Inicio Actividades", value=fecha_act_dt)

                ea1, ea2 = st.columns([4, 1])
                ed_act_econ = ea1.text_input("Actividad Económica (CIIU)", value=str(emp_row.get("actividad_economica") or ""))
                ed_activa = ea2.checkbox("Empresa Activa", value=bool(int(emp_row.get("activa") or 0)))

                if st.form_submit_button("💾 Guardar Cambios de Empresa", type="primary"):
                    if not ed_cedula.strip() or not ed_razon.strip():
                        st.error("La cédula y la razón social son obligatorias.")
                    else:
                        try:
                            conn.execute(
                                """
                                UPDATE empresas
                                SET tipo_contribuyente = ?,
                                    cedula = ?,
                                    razon_social = ?,
                                    nombre_comercial = ?,
                                    representante_legal = ?,
                                    actividad_economica = ?,
                                    fecha_inicio_actividades = ?,
                                    activa = ?
                                WHERE id = ?
                                """,
                                (
                                    ed_tipo,
                                    ed_cedula.strip(),
                                    ed_razon.strip(),
                                    ed_nom_com.strip(),
                                    ed_rep.strip(),
                                    ed_act_econ.strip(),
                                    str(ed_fecha),
                                    1 if ed_activa else 0,
                                    int(emp_sel_id),
                                ),
                            )
                            guardar_y_sincronizar(conn)
                            st.success(f"Empresa '{ed_razon.strip()}' actualizada correctamente.")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Ya existe otra empresa registrada con esa cédula.")

            with st.expander(f"🗑️ Eliminar Empresa: {emp_row['razon_social']} (ID {emp_sel_id})"):
                st.warning(
                    "Esta acción eliminará permanentemente la empresa seleccionada y todos sus datos "
                    "contables asociados (catálogo de cuentas, saldos iniciales, asientos, conciliaciones y declaraciones)."
                )
                confirmar_borrado = st.checkbox(
                    f"Confirmo que deseo eliminar permanentemente a '{emp_row['razon_social']}'",
                    key=f"conf_borrar_emp_{emp_sel_id}",
                )
                if st.button(
                    "🗑️ Eliminar Empresa Permanentemente",
                    type="secondary",
                    disabled=not confirmar_borrado,
                    key=f"btn_borrar_emp_{emp_sel_id}",
                ):
                    eliminar_empresa_completa(conn, emp_sel_id)
                    st.session_state.pop("emp_editar_sel_id", None)
                    st.success("Empresa eliminada correctamente.")
                    st.rerun()

    with tab_cat:
        if not empresa_id:
            st.info("Seleccione o cree una empresa primero.")
        else:
            st.subheader(f"Catálogo de Cuentas — {empresa_info['razon_social']}")

            # 1. Crear cuenta
            with st.form("form_nueva_cuenta", clear_on_submit=True):
                st.markdown("**➕ Crear Nueva Cuenta**")
                c1, c2, c3, c4 = st.columns([2, 4, 2, 2])
                cod_cta = c1.text_input("Código *", placeholder="Ej: 1.1.01.05")
                nom_cta = c2.text_input("Nombre de la Cuenta *", placeholder="Ej: Banco Nacional Colones")
                tipo_cta = c3.selectbox("Tipo *", ["Activo", "Pasivo", "Patrimonio", "Ingreso", "Costo", "Gasto"])
                acepta_mov = c4.selectbox("¿Acepta Movimiento?", ["Sí (Detalle)", "No (Título/Mayor)"])

                if st.form_submit_button("➕ Agregar Cuenta", type="primary"):
                    ok, msg = crear_cuenta_catalogo(
                        conn,
                        empresa_id,
                        cod_cta,
                        nom_cta,
                        tipo_cta,
                        1 if acepta_mov.startswith("Sí") else 0,
                    )
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

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

                # 2. Modificar o borrar cuenta existente
                st.divider()
                st.markdown("**✏️ Modificar o 🗑️ Borrar Cuenta Existente**")
                ids_cta = [int(x) for x in df_cat["id"].tolist()]
                cta_edit_id = st.selectbox(
                    "Seleccione la cuenta a modificar o eliminar:",
                    ids_cta,
                    format_func=lambda x: (
                        f"{df_cat[df_cat['id']==x]['codigo'].values[0]} - "
                        f"{df_cat[df_cat['id']==x]['nombre'].values[0]} "
                        f"({df_cat[df_cat['id']==x]['tipo'].values[0]})"
                    ),
                    key="sel_cta_gestion",
                )
                cta_row = df_cat[df_cat["id"] == cta_edit_id].iloc[0]
                tipos_cta_disp = ["Activo", "Pasivo", "Patrimonio", "Ingreso", "Costo", "Gasto"]
                tipo_cta_act = str(cta_row["tipo"] or "Activo")
                if tipo_cta_act not in tipos_cta_disp:
                    tipos_cta_disp.append(tipo_cta_act)

                with st.form(f"form_editar_cuenta_{cta_edit_id}"):
                    ec1, ec2, ec3, ec4 = st.columns([2, 4, 2, 2])
                    ed_cod = ec1.text_input("Código *", value=str(cta_row["codigo"] or ""))
                    ed_nom = ec2.text_input("Nombre de la Cuenta *", value=str(cta_row["nombre"] or ""))
                    ed_tipo = ec3.selectbox("Tipo *", tipos_cta_disp, index=tipos_cta_disp.index(tipo_cta_act))
                    ed_mov = ec4.selectbox(
                        "¿Acepta Movimiento?",
                        ["Sí (Detalle)", "No (Título/Mayor)"],
                        index=0 if int(cta_row["acepta_movimiento"] or 0) == 1 else 1,
                    )

                    if st.form_submit_button("💾 Guardar Cambios de Cuenta", type="primary"):
                        ok, msg = actualizar_cuenta_catalogo(
                            conn,
                            empresa_id,
                            cta_edit_id,
                            ed_cod,
                            ed_nom,
                            ed_tipo,
                            1 if ed_mov.startswith("Sí") else 0,
                        )
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

                with st.expander(f"🗑️ Borrar Cuenta: {cta_row['codigo']} - {cta_row['nombre']}"):
                    st.caption("Solo se pueden eliminar cuentas que no tengan asientos contables ni saldos iniciales.")
                    conf_cta = st.checkbox(
                        f"Confirmo borrar la cuenta '{cta_row['codigo']} - {cta_row['nombre']}'",
                        key=f"conf_cta_{cta_edit_id}",
                    )
                    if st.button(
                        "🗑️ Borrar Cuenta",
                        type="secondary",
                        disabled=not conf_cta,
                        key=f"btn_borrar_cta_{cta_edit_id}",
                    ):
                        ok, msg = eliminar_cuenta_catalogo(conn, empresa_id, cta_edit_id)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

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
        st.subheader("Registro de Saldos Iniciales de Apertura (en Colones ₡)")
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
            # Creación rápida de cuenta sin salir de Asientos Contables
            with st.expander("➕ Crear Cuenta Rápida (sin ir a Catálogo de Cuentas)", expanded=False):
                with st.form("form_cuenta_rapida_asiento", clear_on_submit=True):
                    rc1, rc2, rc3, rc4 = st.columns([2, 4, 2, 2])
                    r_cod = rc1.text_input("Código Cuenta *", placeholder="Ej: 5.1.01.10")
                    r_nom = rc2.text_input("Nombre Cuenta *", placeholder="Ej: Combustibles y Lubricantes")
                    r_tipo = rc3.selectbox("Tipo NIIF *", ["Gasto", "Costo", "Ingreso", "Activo", "Pasivo", "Patrimonio"])
                    r_mov = rc4.selectbox("Acepta Movimiento", ["Sí (Detalle)", "No (Título/Mayor)"])

                    if st.form_submit_button("⚡ Crear y Habilitar en Asiento", type="primary"):
                        ok, msg = crear_cuenta_catalogo(
                            conn,
                            empresa_id,
                            r_cod,
                            r_nom,
                            r_tipo,
                            1 if r_mov.startswith("Sí") else 0,
                        )
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

            df_cat = obtener_catalogo(conn, empresa_id)
            df_mov = df_cat[df_cat["acepta_movimiento"] == 1].copy()

            if df_mov.empty:
                st.warning("Configure primero al menos una cuenta de movimiento (puedes crearla arriba).")
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

                st.markdown("#### 💱 Moneda de Digitación y Tipo de Cambio BCCR")
                m_c1, m_c2, m_c3, m_c4 = st.columns([2, 2, 2, 2])
                moneda_dig = m_c1.selectbox(
                    "Moneda a Digitar",
                    ["Colones (CRC)", "Dólares (USD)"],
                    key="asiento_moneda_dig",
                )
                es_usd = moneda_dig.startswith("Dólares")

                tipo_tc_sel = m_c2.selectbox(
                    "Tipo de Cambio BCCR",
                    ["Compra BCCR (317)", "Venta BCCR (318)", "Manual"],
                    index=1 if tipo_asiento in ["Egresos"] else 0,
                    key="asiento_tipo_tc",
                )

                modo_tc = "venta" if "318" in tipo_tc_sel else ("compra" if "317" in tipo_tc_sel else "manual")
                tc_sugerido, origen_tc = tc_bccr.obtener_tc_bccr(
                    conn,
                    fecha_asiento,
                    tipo_tc=modo_tc,
                    tc_manual_fallback=st.session_state.get("tc_guardado_manual", 510.0),
                )

                col_tc_btn, _ = m_c3.columns([1, 2])
                if col_tc_btn.button("🔄 BCCR", help="Consultar o actualizar tipo de cambio BCCR para la fecha"):
                    ind_code = "318" if modo_tc == "venta" else "317"
                    val_fresh, det_fresh = tc_bccr.consultar_tc_bccr_webservice(fecha_asiento, indicador=ind_code)
                    if val_fresh:
                        tc_bccr.guardar_tc_en_cache(conn, fecha_asiento, ind_code, val_fresh, fuente="BCCR")
                        st.session_state["tc_valor_actual"] = float(val_fresh)
                        st.toast(f"TC BCCR ({ind_code}) actualizado: ₡ {val_fresh:,.4f}", icon="✅")
                    else:
                        st.warning(det_fresh)

                valor_tc_default = float(
                    st.session_state.get(
                        "tc_valor_actual",
                        tc_sugerido if (tc_sugerido and tc_sugerido > 0) else 510.0,
                    )
                )
                if not es_usd and modo_tc != "manual":
                    valor_tc_default = float(tc_sugerido) if (tc_sugerido and tc_sugerido > 0) else 1.0

                tc_aplicado = m_c3.number_input(
                    "Tipo de Cambio (₡ / USD)",
                    min_value=0.0001,
                    value=valor_tc_default if es_usd else (valor_tc_default if valor_tc_default > 1 else 1.0),
                    step=0.05,
                    format="%.4f",
                    key="asiento_tc_valor",
                )
                if modo_tc == "manual" and es_usd:
                    st.session_state["tc_guardado_manual"] = float(tc_aplicado)
                    tc_bccr.guardar_tc_en_cache(
                        conn,
                        fecha_asiento,
                        "318" if "318" in tipo_tc_sel else "317",
                        float(tc_aplicado),
                        fuente="Manual",
                    )

                m_c4.caption(f"Estado TC: {origen_tc}")
                m_c4.info("Los libros y estados se contabilizan en **Colones (₡)**.")

                concepto = st.text_input(
                    "Concepto / Descripción General *",
                    placeholder="Ej: Pago de factura en USD convertida a colones según TC BCCR",
                )

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
                else:
                    # Asegura que cualquier cuenta recién creada aparezca en las opciones del selector
                    df_prev = st.session_state["lineas_asiento_df"].copy()
                    df_prev["Cuenta"] = df_prev["Cuenta"].apply(
                        lambda x: x if x in opciones_cuentas else opciones_cuentas[0]
                    )
                    st.session_state["lineas_asiento_df"] = df_prev

                etiqueta_debe = "Debe (USD $)" if es_usd else "Debe (₡)"
                etiqueta_haber = "Haber (USD $)" if es_usd else "Haber (₡)"

                df_editado = st.data_editor(
                    st.session_state["lineas_asiento_df"],
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "Cuenta": st.column_config.SelectboxColumn(
                            "Cuenta Contable", options=opciones_cuentas, required=True, width="large"
                        ),
                        "Debe": st.column_config.NumberColumn(etiqueta_debe, min_value=0.0, format="%.2f", default=0.0),
                        "Haber": st.column_config.NumberColumn(etiqueta_haber, min_value=0.0, format="%.2f", default=0.0),
                    },
                    key="editor_asientos",
                )

                factor_conv = float(tc_aplicado) if es_usd else 1.0
                tot_d_dig = float(df_editado["Debe"].fillna(0).sum())
                tot_h_dig = float(df_editado["Haber"].fillna(0).sum())
                dif_dig = round(tot_d_dig - tot_h_dig, 2)

                tot_d_crc = round(tot_d_dig * factor_conv, 2)
                tot_h_crc = round(tot_h_dig * factor_conv, 2)
                dif_crc = round(tot_d_crc - tot_h_crc, 2)

                if es_usd:
                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric("Total Debe (USD)", f"$ {tot_d_dig:,.2f}")
                    k2.metric("Total Haber (USD)", f"$ {tot_h_dig:,.2f}")
                    k3.metric("Convertido Debe (₡)", f"₡ {tot_d_crc:,.2f}")
                    k4.metric("Convertido Haber (₡)", f"₡ {tot_h_crc:,.2f}")
                else:
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Total Debe (₡)", f"₡ {tot_d_crc:,.2f}")
                    k2.metric("Total Haber (₡)", f"₡ {tot_h_crc:,.2f}")
                    k3.metric("Diferencia (₡)", f"₡ {dif_crc:,.2f}")

                if st.button("💾 Contabilizar Asiento", type="primary"):
                    lineas_validas = df_editado[
                        (df_editado["Debe"].fillna(0) > 0) | (df_editado["Haber"].fillna(0) > 0)
                    ]
                    if not concepto.strip():
                        st.error("Debe indicar el concepto del asiento.")
                    elif len(lineas_validas) < 2:
                        st.error("Un asiento contable requiere al menos 2 líneas con movimiento.")
                    elif es_usd and factor_conv <= 0:
                        st.error("Debe ingresar un tipo de cambio válido mayor a cero para convertir USD a colones.")
                    elif abs(dif_dig) >= 0.01:
                        st.error(
                            f"El asiento está descuadrado por {dif_dig:,.2f} "
                            f"{'USD' if es_usd else '₡'}. Debe = Haber."
                        )
                    else:
                        periodo_str = fecha_asiento.strftime("%Y-%m")
                        ahora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        moneda_cod = "USD" if es_usd else "CRC"
                        concepto_final = concepto.strip()
                        if es_usd and "(USD " not in concepto_final:
                            concepto_final = (
                                f"{concepto_final} [USD $ {tot_d_dig:,.2f} @ TC ₡ {factor_conv:,.4f}]"
                            )

                        cur.execute(
                            """
                            INSERT INTO asientos_cabecera (
                                empresa_id, num_asiento, fecha, periodo, concepto, tipo_asiento,
                                fecha_registro, fecha_comprobante, forma_pago, num_comprobante,
                                num_cheque, usuario_creacion, moneda_dig, tipo_cambio_bccr, modalidad_tc
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                empresa_id,
                                int(num_asiento),
                                str(fecha_asiento),
                                periodo_str,
                                concepto_final,
                                tipo_asiento,
                                ahora_str,
                                str(fecha_asiento),
                                forma_pago,
                                num_comp.strip(),
                                num_cheque.strip(),
                                usuario_crea.strip(),
                                moneda_cod,
                                float(factor_conv),
                                tipo_tc_sel if es_usd else "N/A",
                            ),
                        )
                        nuevo_asiento_id = cur.lastrowid
                        for _, linea in lineas_validas.iterrows():
                            cid = mapa_cuentas[linea["Cuenta"]]
                            d_dig = float(linea["Debe"] or 0.0)
                            h_dig = float(linea["Haber"] or 0.0)
                            d_crc = round(d_dig * factor_conv, 2)
                            h_crc = round(h_dig * factor_conv, 2)
                            d_usd = round(d_dig, 2) if es_usd else round(d_crc / factor_conv, 2) if factor_conv > 0 else 0.0
                            h_usd = round(h_dig, 2) if es_usd else round(h_crc / factor_conv, 2) if factor_conv > 0 else 0.0

                            cur.execute(
                                """
                                INSERT INTO asientos_detalle (
                                    asiento_id, cuenta_id, debe, haber, conciliado,
                                    moneda_origen, monto_usd_debe, monto_usd_haber, tipo_cambio_aplicado
                                ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, ?)
                                """,
                                (
                                    nuevo_asiento_id,
                                    cid,
                                    d_crc,
                                    h_crc,
                                    moneda_cod,
                                    d_usd,
                                    h_usd,
                                    float(factor_conv),
                                ),
                            )

                        guardar_y_sincronizar(conn)
                        st.session_state.pop("lineas_asiento_df", None)
                        st.session_state.pop("tc_valor_actual", None)
                        st.success(
                            f"Asiento #{num_asiento} contabilizado en Colones (₡ {tot_d_crc:,.2f})"
                            + (f" desde USD $ {tot_d_dig:,.2f} (TC ₡ {factor_conv:,.4f})." if es_usd else ".")
                        )
                        st.rerun()

        with tab_cons:
            st.subheader("Consulta de Asientos Registrados (Expresados en Colones ₡)")
            f1, f2 = st.columns(2)
            f_desde = f1.date_input("Desde", value=datetime.date(datetime.date.today().year, 1, 1), key="ld_desde")
            f_hasta = f2.date_input("Hasta", value=datetime.date.today(), key="ld_hasta")

            df_asientos = pd.read_sql_query(
                """
                SELECT ac.id, ac.num_asiento AS "Asiento", ac.fecha AS "Fecha", ac.tipo_asiento AS "Tipo",
                       COALESCE(ac.moneda_dig, 'CRC') AS "Moneda Dig.",
                       COALESCE(ac.tipo_cambio_bccr, 1.0) AS "TC BCCR",
                       ac.num_comprobante AS "Comprobante", ac.concepto AS "Concepto",
                       COALESCE(SUM(ad.debe), 0) AS "Total Debe (₡)",
                       COALESCE(SUM(ad.haber), 0) AS "Total Haber (₡)"
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
                        SELECT c.codigo AS "Código", c.nombre AS "Cuenta",
                               COALESCE(ad.monto_usd_debe, 0) AS "Debe (USD)",
                               COALESCE(ad.monto_usd_haber, 0) AS "Haber (USD)",
                               COALESCE(ad.tipo_cambio_aplicado, 1.0) AS "TC",
                               ad.debe AS "Debe (₡)", ad.haber AS "Haber (₡)"
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
        st.subheader("Importación Masiva de Asientos desde Excel (Convertidos a Colones ₡)")
        st.info(
            "El archivo Excel debe contener las columnas: "
            "`num_asiento`, `fecha`, `concepto`, `codigo_cuenta`, `debe`, `haber` "
            "(opcionales: `moneda` con `'CRC'` o `'USD'`, `tipo_cambio`, `tipo_asiento`, `num_comprobante`, `forma_pago`)."
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
                                st.error(f"El asiento #{num_a} está descuadrado por {dif_g:,.2f}. Importación cancelada.")
                                conn.rollback()
                                st.stop()
                            prim = grupo.iloc[0]
                            fecha_dt = pd.to_datetime(prim["fecha"]).date()
                            fecha_str = fecha_dt.strftime("%Y-%m-%d")
                            periodo_str = fecha_str[:7]
                            mon_fila = str(prim.get("moneda", "CRC")).strip().upper()
                            es_fila_usd = mon_fila in {"USD", "DOLARES", "DÓLARES", "$"}

                            if es_fila_usd:
                                tc_fila = pd.to_numeric(prim.get("tipo_cambio", None), errors="coerce")
                                if pd.notna(tc_fila) and float(tc_fila) > 0:
                                    tc_usado = float(tc_fila)
                                else:
                                    tc_bccr_val, _ = tc_bccr.obtener_tc_bccr(conn, fecha_dt, tipo_tc="compra")
                                    tc_usado = float(tc_bccr_val) if tc_bccr_val else 510.0
                            else:
                                tc_usado = 1.0

                            cur.execute(
                                """
                                INSERT INTO asientos_cabecera (
                                    empresa_id, num_asiento, fecha, periodo, concepto, tipo_asiento,
                                    fecha_registro, forma_pago, num_comprobante, usuario_creacion,
                                    moneda_dig, tipo_cambio_bccr, modalidad_tc
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                                    "USD" if es_fila_usd else "CRC",
                                    tc_usado,
                                    "BCCR/Excel" if es_fila_usd else "N/A",
                                ),
                            )
                            aid = cur.lastrowid
                            for _, r in grupo.iterrows():
                                d_orig = float(r["debe"])
                                h_orig = float(r["haber"])
                                d_crc = round(d_orig * tc_usado, 2)
                                h_crc = round(h_orig * tc_usado, 2)
                                d_usd = round(d_orig, 2) if es_fila_usd else 0.0
                                h_usd = round(h_orig, 2) if es_fila_usd else 0.0
                                cur.execute(
                                    """
                                    INSERT INTO asientos_detalle (
                                        asiento_id, cuenta_id, debe, haber, conciliado,
                                        moneda_origen, monto_usd_debe, monto_usd_haber, tipo_cambio_aplicado
                                    ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, ?)
                                    """,
                                    (
                                        aid,
                                        mapa_cod[r["codigo_cuenta"]],
                                        d_crc,
                                        h_crc,
                                        "USD" if es_fila_usd else "CRC",
                                        d_usd,
                                        h_usd,
                                        tc_usado,
                                    ),
                                )
                            insertados += 1

                        guardar_y_sincronizar(conn)
                        st.success(f"Se importaron {insertados} asientos contables convertidos a Colones (₡).")
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
        st.subheader("Conciliación Bancaria Mensual (en Colones ₡)")
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
