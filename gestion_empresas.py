from datetime import datetime
import os
import shutil
import pandas as pd
import streamlit as st
from sync_sheets import subir_bd_a_sheets

CATALOGO_BASICO_NIIF = [
    # ACTIVO
    ("1", "ACTIVO", "ACTIVO", 0),
    ("1.1", "ACTIVO CORRIENTE", "ACTIVO", 0),
    ("1.1.01", "EFECTIVO Y EQUIVALENTES DE EFECTIVO", "ACTIVO", 0),
    ("1.1.01.01", "Caja General", "ACTIVO", 1),
    ("1.1.01.02", "Caja Chica", "ACTIVO", 1),
    ("1.1.01.03", "Bancos Cuenta Corriente Moneda Nacional", "ACTIVO", 1),
    ("1.1.01.04", "Bancos Cuenta Corriente Moneda Extranjera", "ACTIVO", 1),
    ("1.1.02", "CUENTAS Y DOCUMENTOS POR COBRAR", "ACTIVO", 0),
    ("1.1.02.01", "Clientes Nacionales", "ACTIVO", 1),
    ("1.1.02.02", "Cuentas por Cobrar Empleados y Funcionarios", "ACTIVO", 1),
    ("1.1.03", "IMPUESTOS POR ACREDITAR Y ANTICIPOS", "ACTIVO", 0),
    ("1.1.03.01", "IVA Soportado / Acreditable (Compras)", "ACTIVO", 1),
    ("1.1.03.02", "Retenciones de Renta 2% por Acreditar", "ACTIVO", 1),
    ("1.1.03.03", "Pagos Parciales de Renta", "ACTIVO", 1),
    ("1.1.04", "INVENTARIOS", "ACTIVO", 0),
    ("1.1.04.01", "Inventario de Mercaderías", "ACTIVO", 1),
    ("1.2", "ACTIVO NO CORRIENTE", "ACTIVO", 0),
    ("1.2.01", "PROPIEDAD, PLANTA Y EQUIPO", "ACTIVO", 0),
    ("1.2.01.01", "Mobiliario y Equipo de Oficina", "ACTIVO", 1),
    ("1.2.01.02", "Equipo de Cómputo", "ACTIVO", 1),
    ("1.2.01.03", "Depreciación Acumulada Propiedad, Planta y Equipo", "ACTIVO", 1),
    # PASIVO
    ("2", "PASIVO", "PASIVO", 0),
    ("2.1", "PASIVO CORRIENTE", "PASIVO", 0),
    ("2.1.01", "CUENTAS Y PROVEEDORES POR PAGAR", "PASIVO", 0),
    ("2.1.01.01", "Proveedores Nacionales", "PASIVO", 1),
    ("2.1.01.02", "Cuentas por Pagar Varias", "PASIVO", 1),
    ("2.1.02", "IMPUESTOS Y CARGAS SOCIALES POR PAGAR", "PASIVO", 0),
    ("2.1.02.01", "IVA Devengado / Por Pagar (Ventas)", "PASIVO", 1),
    ("2.1.02.02", "Impuesto sobre la Renta por Pagar", "PASIVO", 1),
    ("2.1.02.03", "Cargas Sociales CCSS por Pagar", "PASIVO", 1),
    # PATRIMONIO
    ("3", "PATRIMONIO", "PATRIMONIO", 0),
    ("3.1", "CAPITAL Y RESULTADOS", "PATRIMONIO", 0),
    ("3.1.01", "Capital Social / Aportes del Propietario", "PATRIMONIO", 1),
    ("3.1.02", "Utilidades Acumuladas de Períodos Anteriores", "PATRIMONIO", 1),
    ("3.1.03", "Pérdidas Acumuladas de Períodos Anteriores", "PATRIMONIO", 1),
    # INGRESOS
    ("4", "INGRESOS", "INGRESO", 0),
    ("4.1", "INGRESOS OPERATIVOS", "INGRESO", 0),
    ("4.1.01", "Ingresos por Servicios Profesionales / Comerciales", "INGRESO", 1),
    ("4.1.02", "Ventas de Mercadería Gravadas", "INGRESO", 1),
    ("4.1.03", "Ventas de Mercadería Exentas", "INGRESO", 1),
    ("4.2", "OTROS INGRESOS", "INGRESO", 0),
    ("4.2.01", "Ingresos Financieros e Intereses Ganados", "INGRESO", 1),
    # COSTOS
    ("5", "COSTOS DE VENTAS Y SERVICIOS", "COSTO", 0),
    ("5.1", "COSTO DE OPERACIÓN", "COSTO", 0),
    ("5.1.01", "Costo de Mercadería Vendida", "COSTO", 1),
    ("5.1.02", "Costo Directo de Servicios Prestados", "COSTO", 1),
    # GASTOS
    ("6", "GASTOS OPERATIVOS", "GASTO", 0),
    ("6.1", "GASTOS ADMINISTRATIVOS Y DE VENTAS", "GASTO", 0),
    ("6.1.01", "Sueldos, Salarios y Cargas Sociales", "GASTO", 1),
    ("6.1.02", "Alquileres de Local y Oficina", "GASTO", 1),
    ("6.1.03", "Servicios Públicos (Electricidad, Agua, Internet, Teléfono)", "GASTO", 1),
    ("6.1.04", "Honorarios Profesionales", "GASTO", 1),
    ("6.1.05", "Papelería, Útiles y Suministros de Oficina", "GASTO", 1),
    ("6.1.06", "Gastos Bancarios y Comisiones", "GASTO", 1),
    ("6.1.07", "Gasto por Depreciación", "GASTO", 1),
]


def _asegurar_columnas_empresas(conn):
  cur = conn.cursor()
  cols = [
      row[1] for row in cur.execute("PRAGMA table_info(empresas)").fetchall()
  ]
  if "activa" not in cols:
    cur.execute("ALTER TABLE empresas ADD COLUMN activa INTEGER DEFAULT 1")
    cur.execute("UPDATE empresas SET activa = 1 WHERE activa IS NULL")
  if "tipo_contribuyente" not in cols:
    cur.execute(
        "ALTER TABLE empresas ADD COLUMN tipo_contribuyente TEXT DEFAULT"
        " 'Persona Jurídica'"
    )
    cur.execute(
        "UPDATE empresas SET tipo_contribuyente = 'Persona Jurídica' WHERE"
        " tipo_contribuyente IS NULL OR tipo_contribuyente = ''"
    )
  conn.commit()


def _clonar_catalogo_inicial(cur, nuevo_empresa_id):
  # Evitar duplicar si ya tiene cuentas
  ya_tiene = cur.execute(
      "SELECT COUNT(*) FROM catalogo_cuentas WHERE empresa_id=?",
      (nuevo_empresa_id,),
  ).fetchone()[0]
  if ya_tiene > 0:
    return ya_tiene

  # 1. Intentar desde plantilla_catalogo_cuentas
  plantilla = cur.execute(
      "SELECT codigo, nombre, tipo, acepta_movimiento FROM"
      " plantilla_catalogo_cuentas ORDER BY codigo"
  ).fetchall()

  # 2. Si la plantilla está vacía, tomar el catálogo base de la primera empresa existente
  if not plantilla:
    primera_emp = cur.execute(
        "SELECT empresa_id FROM catalogo_cuentas LIMIT 1"
    ).fetchone()
    if primera_emp:
      plantilla = cur.execute(
          "SELECT codigo, nombre, tipo, acepta_movimiento FROM"
          " catalogo_cuentas WHERE empresa_id=? ORDER BY codigo",
          (primera_emp[0],),
      ).fetchall()

  # 3. Si no hay ninguna plantilla ni empresa previa, usar el catálogo NIIF básico integrado
  if not plantilla:
    plantilla = CATALOGO_BASICO_NIIF

  max_cta_id = cur.execute(
      "SELECT COALESCE(MAX(id), 0) FROM catalogo_cuentas"
  ).fetchone()[0]
  siguiente_id = int(max_cta_id) + 1

  filas_insertar = []
  for codigo, nombre, tipo, acepta in plantilla:
    filas_insertar.append(
        (siguiente_id, nuevo_empresa_id, codigo, nombre, tipo, acepta)
    )
    siguiente_id += 1

  cur.executemany(
      """
        INSERT INTO catalogo_cuentas (id, empresa_id, codigo, nombre, tipo, acepta_movimiento)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
      filas_insertar,
  )
  return len(filas_insertar)


def _respaldo_seguridad_previo(conn, db_path="sistema_contable.db"):
  try:
    conn.commit()
    if os.path.exists(db_path):
      os.makedirs("backups_seguridad", exist_ok=True)
      ts = datetime.now().strftime("%Y%m%d_%H%M%S")
      shutil.copy2(db_path, f"backups_seguridad/pre_borrado_{ts}.db")
  except Exception:
    pass


def render_gestion_empresas(conn):
  _asegurar_columnas_empresas(conn)
  st.subheader(
      "🏢 Administración de Contribuyentes (Personas Jurídicas y Físicas)"
  )

  tab_lista, tab_nueva, tab_editar, tab_borrar, tab_papelera = st.tabs([
      "📋 Contribuyentes Activos",
      "➕ Nuevo Registro",
      "✏️ Editar Registro",
      "🗑️ Eliminar / Archivar",
      "♻️ Papelera (Recuperar Borrados)",
  ])

  df_todas = pd.read_sql_query(
      """
        SELECT rowid AS fila_sqlite, id,
               COALESCE(NULLIF(tipo_contribuyente, ''), 'Persona Jurídica') AS tipo_contribuyente,
               cedula, razon_social, nombre_comercial,
               representante_legal, actividad_economica,
               COALESCE(activa, 1) AS activa
        FROM empresas
        ORDER BY id, rowid
        """,
      conn,
  )

  df_activas = df_todas[df_todas["activa"] == 1].copy()
  df_papelera = df_todas[df_todas["activa"] == 0].copy()

  # --- PESTAÑA 1: LISTA DE ACTIVOS ---
  with tab_lista:
    if df_activas.empty:
      st.info("No hay empresas ni personas físicas activas registradas.")
    else:
      df_mostrar = df_activas.drop(columns=["fila_sqlite", "activa"]).rename(
          columns={
              "id": "ID",
              "tipo_contribuyente": "Tipo",
              "cedula": "Cédula",
              "razon_social": "Razón Social / Nombre Completo",
              "nombre_comercial": "Nombre Comercial",
              "representante_legal": "Representante Legal",
              "actividad_economica": "Actividad Económica",
          }
      )
      st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

  # --- PESTAÑA 2: NUEVA EMPRESA O PERSONA FÍSICA ---
  with tab_nueva:
    tipo_nuevo = st.radio(
        "Tipo de Contribuyente:",
        ["Persona Jurídica", "Persona Física"],
        horizontal=True,
        key="radio_tipo_contribuyente_nuevo",
    )
    es_fisica = tipo_nuevo == "Persona Física"

    with st.form("form_nueva_empresa_segura", clear_on_submit=True):
      col1, col2 = st.columns(2)
      cedula = col1.text_input(
          "Cédula Física / DIMEX / NITE *"
          if es_fisica
          else "Cédula Jurídica *"
      ).strip()
      razon_social = col2.text_input(
          "Nombre Completo (Apellidos y Nombre) *"
          if es_fisica
          else "Razón Social *"
      ).strip()
      nombre_comercial = col1.text_input(
          "Nombre Comercial / Fantasía (Opcional)"
          if es_fisica
          else "Nombre Comercial"
      ).strip()
      representante = (
          "" if es_fisica else col2.text_input("Representante Legal").strip()
      )
      actividad = (col2 if es_fisica else st).text_input(
          "Actividad Económica / Profesional"
      ).strip()
      st.caption(
          "📚 Se creará automáticamente un catálogo básico de cuentas NIIF para"
          " este contribuyente."
      )

      submitted = st.form_submit_button("💾 Guardar Registro", type="primary")
      if submitted:
        if not cedula or not razon_social:
          st.error("La Cédula y el Nombre / Razón Social son obligatorios.")
        else:
          ced_norm = cedula.replace("-", "").replace(" ", "").lower()
          razon_norm = razon_social.lower()

          duplicada = False
          for _, r in df_todas.iterrows():
            c_ex = (
                str(r["cedula"] or "").replace("-", "").replace(" ", "").lower()
            )
            r_ex = str(r["razon_social"] or "").strip().lower()
            if (ced_norm and c_ex == ced_norm) or (r_ex == razon_norm):
              duplicada = True
              break

          if duplicada:
            st.error(
                f"⚠️ Ya existe un registro con la cédula '{cedula}' o nombre"
                f" '{razon_social}' (revisa también la pestaña Papelera)."
            )
          else:
            cur = conn.cursor()
            max_id = cur.execute(
                "SELECT COALESCE(MAX(id), 0) FROM empresas"
            ).fetchone()[0]
            nuevo_id = int(max_id) + 1
            cur.execute(
                """
                INSERT INTO empresas (
                    id, tipo_contribuyente, cedula, razon_social,
                    nombre_comercial, representante_legal, actividad_economica, activa
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 1)
                """,
                (
                    nuevo_id,
                    tipo_nuevo,
                    cedula,
                    razon_social,
                    nombre_comercial,
                    representante,
                    actividad,
                ),
            )
            cuentas_creadas = _clonar_catalogo_inicial(cur, nuevo_id)
            conn.commit()
            subir_bd_a_sheets(conn)
            st.success(
                f"✅ {tipo_nuevo} '{razon_social}' registrada con"
                f" {cuentas_creadas} cuentas contables básicas."
            )
            st.rerun()

  # --- PESTAÑA 3: EDITAR REGISTRO ---
  with tab_editar:
    if df_activas.empty:
      st.info("No hay registros activos para editar.")
    else:
      opciones_ed = {
          f"ID {r['id']} [{r['tipo_contribuyente']}] - {r['razon_social']} ({r['cedula']})": int(
              r["id"]
          )
          for _, r in df_activas.iterrows()
      }
      sel_ed = st.selectbox(
          "Selecciona el registro a editar:", list(opciones_ed.keys())
      )
      emp_id_ed = opciones_ed[sel_ed]
      fila_actual = df_activas[df_activas["id"] == emp_id_ed].iloc[0]

      with st.form("form_editar_empresa"):
        tipo_ed = st.radio(
            "Tipo de Contribuyente:",
            ["Persona Jurídica", "Persona Física"],
            index=1
            if fila_actual["tipo_contribuyente"] == "Persona Física"
            else 0,
            horizontal=True,
        )
        es_fisica_ed = tipo_ed == "Persona Física"

        c1, c2 = st.columns(2)
        n_ced = c1.text_input(
            "Cédula Física / DIMEX" if es_fisica_ed else "Cédula Jurídica",
            value=str(fila_actual["cedula"] or ""),
        )
        n_raz = c2.text_input(
            "Nombre Completo" if es_fisica_ed else "Razón Social",
            value=str(fila_actual["razon_social"] or ""),
        )
        n_com = c1.text_input(
            "Nombre Comercial", value=str(fila_actual["nombre_comercial"] or "")
        )
        n_rep = c2.text_input(
            "Representante Legal (Opcional en Persona Física)",
            value=str(fila_actual["representante_legal"] or ""),
        )
        n_act = st.text_input(
            "Actividad Económica",
            value=str(fila_actual["actividad_economica"] or ""),
        )

        if st.form_submit_button("💾 Actualizar Datos", type="primary"):
          cur = conn.cursor()
          cur.execute(
              """
              UPDATE empresas
              SET tipo_contribuyente=?, cedula=?, razon_social=?,
                  nombre_comercial=?, representante_legal=?, actividad_economica=?
              WHERE id=?
              """,
              (
                  tipo_ed,
                  n_ced.strip(),
                  n_raz.strip(),
                  n_com.strip(),
                  n_rep.strip(),
                  n_act.strip(),
                  emp_id_ed,
              ),
          )
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success("✅ Datos actualizados correctamente.")
          st.rerun()

      cur = conn.cursor()
      cant_ctas = cur.execute(
          "SELECT COUNT(*) FROM catalogo_cuentas WHERE empresa_id=?",
          (emp_id_ed,),
      ).fetchone()[0]
      if cant_ctas == 0:
        st.warning(
            "⚠️ Este contribuyente aún no tiene catálogo de cuentas cargado."
        )
        if st.button("📚 Construir Catálogo Básico de Cuentas Ahora"):
          n_ins = _clonar_catalogo_inicial(cur, emp_id_ed)
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success(f"✅ Se construyeron {n_ins} cuentas contables básicas.")
          st.rerun()

  # --- PESTAÑA 4: ELIMINAR DUPLICADO O ENVIAR A PAPELERA ---
  with tab_borrar:
    if df_activas.empty:
      st.info("No hay registros activos.")
    else:
      cur = conn.cursor()
      opciones_del = {}
      conteo_ids = df_activas["id"].value_counts().to_dict()

      for _, r in df_activas.iterrows():
        eid = int(r["id"])
        f_sql = int(r["fila_sqlite"])
        n_cuentas = cur.execute(
            "SELECT COUNT(*) FROM catalogo_cuentas WHERE empresa_id=?", (eid,)
        ).fetchone()[0]
        n_asientos = cur.execute(
            "SELECT COUNT(*) FROM asientos_cabecera WHERE empresa_id=?", (eid,)
        ).fetchone()[0]
        es_duplicado_id = conteo_ids.get(eid, 0) > 1
        etiqueta_dup = " [FILA DUPLICADA]" if es_duplicado_id else ""
        etiqueta = (
            f"ID {eid} [{r['tipo_contribuyente']}] - {r['razon_social']}"
            f" ({r['cedula']}) | {n_cuentas} cuentas, {n_asientos}"
            f" asientos{etiqueta_dup}"
        )
        opciones_del[etiqueta] = (f_sql, eid, es_duplicado_id, r["razon_social"])

      sel_del = st.selectbox(
          "Selecciona el registro a eliminar:", list(opciones_del.keys())
      )
      fila_sql_sel, emp_id_del, es_dup_id, nombre_sel = opciones_del[sel_del]

      if es_dup_id:
        st.info(
            "ℹ️ Este registro tiene un **ID duplicado**. Se eliminará"
            " **únicamente esta fila repetida** conservando el registro original."
        )
        if st.button("🧹 Eliminar Solo Esta Fila Duplicada", type="primary"):
          _respaldo_seguridad_previo(conn)
          cur.execute("DELETE FROM empresas WHERE rowid=?", (fila_sql_sel,))
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success("✅ Fila duplicada eliminada.")
          st.rerun()
      else:
        st.warning(
            f"⚠️ El registro **{nombre_sel}** se enviará a la **Papelera** (no"
            " se perderán sus cuentas ni asientos, y podrás recuperarlo desde la"
            " pestaña *Papelera*)."
        )
        conf_papelera = st.checkbox(
            f"Confirmo enviar '{nombre_sel}' a la Papelera."
        )
        if st.button(
            "🗑️ Enviar a Papelera", type="primary", disabled=not conf_papelera
        ):
          _respaldo_seguridad_previo(conn)
          cur.execute("UPDATE empresas SET activa = 0 WHERE id=?", (emp_id_del,))
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success(
              f"✅ '{nombre_sel}' fue enviado a la Papelera y respaldado."
          )
          st.rerun()

  # --- PESTAÑA 5: PAPELERA ---
  with tab_papelera:
    st.caption(
        "Aquí aparecen las empresas y personas físicas eliminadas. Puedes"
        " restaurarlas con todos sus asientos intactos."
    )
    if df_papelera.empty:
      st.info("La papelera está vacía.")
    else:
      st.dataframe(
          df_papelera.drop(columns=["fila_sqlite", "activa"]),
          use_container_width=True,
          hide_index=True,
      )
      opciones_rec = {
          f"ID {r['id']} [{r['tipo_contribuyente']}] - {r['razon_social']} ({r['cedula']})": int(
              r["id"]
          )
          for _, r in df_papelera.iterrows()
      }
      sel_rec = st.selectbox(
          "Selecciona un registro para recuperar:", list(opciones_rec.keys())
      )
      id_rec = opciones_rec[sel_rec]

      col_r1, col_r2 = st.columns(2)
      if col_r1.button("♻️ Restaurar al Sistema", type="primary"):
        cur = conn.cursor()
        cur.execute("UPDATE empresas SET activa = 1 WHERE id=?", (id_rec,))
        conn.commit()
        subir_bd_a_sheets(conn)
        st.success("✅ ¡Registro restaurado nuevamente al sistema!")
        st.rerun()

      with col_r2.expander("🔥 Borrado Definitivo (Permanente)"):
        st.error(
            "Esto eliminará permanentemente el contribuyente, sus cuentas y"
            " sus asientos."
        )
        conf_def = st.checkbox("Confirmo borrado permanente e irreversible.")
        if st.button("Eliminar Definitivamente", disabled=not conf_def):
          _respaldo_seguridad_previo(conn)
          cur = conn.cursor()
          cur.execute(
              "DELETE FROM asientos_detalle WHERE asiento_id IN (SELECT id FROM"
              " asientos_cabecera WHERE empresa_id=?)",
              (id_rec,),
          )
          cur.execute(
              "DELETE FROM asientos_cabecera WHERE empresa_id=?", (id_rec,)
          )
          cur.execute(
              "DELETE FROM saldos_iniciales WHERE empresa_id=?", (id_rec,)
          )
          cur.execute(
              "DELETE FROM catalogo_cuentas WHERE empresa_id=?", (id_rec,)
          )
          cur.execute(
              "DELETE FROM conciliaciones_bancarias WHERE empresa_id=?",
              (id_rec,),
          )
          cur.execute(
              "DELETE FROM declaraciones_iva WHERE empresa_id=?", (id_rec,)
          )
          cur.execute(
              "DELETE FROM declaraciones_renta WHERE empresa_id=?", (id_rec,)
          )
          cur.execute("DELETE FROM empresas WHERE id=?", (id_rec,))
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success("✅ Registro eliminado permanentemente.")
          st.rerun()
