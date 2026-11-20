from datetime import datetime
import os
import shutil
import pandas as pd
import streamlit as st
from sync_sheets import subir_bd_a_sheets


def _asegurar_columna_activa(conn):
  cur = conn.cursor()
  cols = [
      row[1] for row in cur.execute("PRAGMA table_info(empresas)").fetchall()
  ]
  if "activa" not in cols:
    cur.execute("ALTER TABLE empresas ADD COLUMN activa INTEGER DEFAULT 1")
    cur.execute("UPDATE empresas SET activa = 1 WHERE activa IS NULL")
    conn.commit()


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
  _asegurar_columna_activa(conn)
  st.subheader("🏢 Administración y Control de Empresas")

  tab_lista, tab_nueva, tab_editar, tab_borrar, tab_papelera = st.tabs([
      "📋 Empresas Activas",
      "➕ Nueva Empresa",
      "✏️ Editar Empresa",
      "🗑️ Eliminar / Archivar",
      "♻️ Papelera (Recuperar Borradas)",
  ])

  # Traer todas las filas incluyendo rowid físico de SQLite para distinguir duplicados exactos
  df_todas = pd.read_sql_query(
      """
        SELECT rowid AS fila_sqlite, id, cedula, razon_social, nombre_comercial,
               representante_legal, actividad_economica, COALESCE(activa, 1) AS activa
        FROM empresas
        ORDER BY id, rowid
        """,
      conn,
  )

  df_activas = df_todas[df_todas["activa"] == 1].copy()
  df_papelera = df_todas[df_todas["activa"] == 0].copy()

  # --- PESTAÑA 1: LISTA DE ACTIVAS ---
  with tab_lista:
    if df_activas.empty:
      st.info("No hay empresas activas registradas.")
    else:
      st.dataframe(
          df_activas.drop(columns=["fila_sqlite", "activa"]),
          use_container_width=True,
          hide_index=True,
      )

  # --- PESTAÑA 2: NUEVA EMPRESA (CON VALIDACIÓN ANTI-DUPLICADOS) ---
  with tab_nueva:
    with st.form("form_nueva_empresa_segura", clear_on_submit=True):
      col1, col2 = st.columns(2)
      cedula = col1.text_input("Cédula Jurídica / Física *").strip()
      razon_social = col2.text_input("Razón Social *").strip()
      nombre_comercial = col1.text_input("Nombre Comercial").strip()
      representante = col2.text_input("Representante Legal").strip()
      actividad = st.text_input("Actividad Económica").strip()

      submitted = st.form_submit_button("💾 Guardar Empresa", type="primary")
      if submitted:
        if not cedula or not razon_social:
          st.error("La Cédula y la Razón Social son obligatorias.")
        else:
          ced_norm = cedula.replace("-", "").replace(" ", "").lower()
          razon_norm = razon_social.lower()

          duplicada = False
          for _, r in df_todas.iterrows():
            c_ex = str(r["cedula"] or "").replace("-", "").replace(" ", "").lower()
            r_ex = str(r["razon_social"] or "").strip().lower()
            if (ced_norm and c_ex == ced_norm) or (r_ex == razon_norm):
              duplicada = True
              break

          if duplicada:
            st.error(
                f"⚠️ Ya existe una empresa con la cédula '{cedula}' o razón"
                f" social '{razon_social}' (revisa también la pestaña Papelera)."
            )
          else:
            cur = conn.cursor()
            max_id = cur.execute(
                "SELECT COALESCE(MAX(id), 0) FROM empresas"
            ).fetchone()[0]
            nuevo_id = int(max_id) + 1
            cur.execute(
                """
                INSERT INTO empresas (id, cedula, razon_social, nombre_comercial, representante_legal, actividad_economica, activa)
                VALUES (?, ?, ?, ?, ?, ?, 1)
                """,
                (
                    nuevo_id,
                    cedula,
                    razon_social,
                    nombre_comercial,
                    representante,
                    actividad,
                ),
            )
            conn.commit()
            subir_bd_a_sheets(conn)
            st.success(f"✅ Empresa '{razon_social}' creada correctamente.")
            st.rerun()

  # --- PESTAÑA 3: EDITAR EMPRESA ---
  with tab_editar:
    if df_activas.empty:
      st.info("No hay empresas activas para editar.")
    else:
      opciones_ed = {
          f"ID {r['id']} - {r['razon_social']} ({r['cedula']})": int(r["id"])
          for _, r in df_activas.iterrows()
      }
      sel_ed = st.selectbox(
          "Selecciona la empresa a editar:", list(opciones_ed.keys())
      )
      emp_id_ed = opciones_ed[sel_ed]
      fila_actual = df_activas[df_activas["id"] == emp_id_ed].iloc[0]

      with st.form("form_editar_empresa"):
        c1, c2 = st.columns(2)
        n_ced = c1.text_input("Cédula", value=str(fila_actual["cedula"] or ""))
        n_raz = c2.text_input(
            "Razón Social", value=str(fila_actual["razon_social"] or "")
        )
        n_com = c1.text_input(
            "Nombre Comercial", value=str(fila_actual["nombre_comercial"] or "")
        )
        n_rep = c2.text_input(
            "Representante Legal",
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
              SET cedula=?, razon_social=?, nombre_comercial=?, representante_legal=?, actividad_economica=?
              WHERE id=?
              """,
              (
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

  # --- PESTAÑA 4: ELIMINAR DUPLICADO EXACTO O MOVER A PAPELERA ---
  with tab_borrar:
    if df_activas.empty:
      st.info("No hay empresas activas.")
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
            f"ID {eid} (Reg #{f_sql}) - {r['razon_social']} ({r['cedula']}) |"
            f" {n_cuentas} cuentas, {n_asientos} asientos{etiqueta_dup}"
        )
        opciones_del[etiqueta] = (f_sql, eid, es_duplicado_id, r["razon_social"])

      sel_del = st.selectbox(
          "Selecciona el registro a eliminar:", list(opciones_del.keys())
      )
      fila_sql_sel, emp_id_del, es_dup_id, nombre_sel = opciones_del[sel_del]

      if es_dup_id:
        st.info(
            "ℹ️ Este registro tiene un **ID duplicado**. Al eliminarlo, se"
            " borrará **únicamente esta fila repetida** y se conservará intacta"
            " la empresa original con todas sus cuentas y asientos."
        )
        if st.button("🧹 Eliminar Solo Esta Fila Duplicada", type="primary"):
          _respaldo_seguridad_previo(conn)
          cur.execute("DELETE FROM empresas WHERE rowid=?", (fila_sql_sel,))
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success(
              "✅ Fila duplicada eliminada sin afectar la empresa original."
          )
          st.rerun()
      else:
        st.warning(
            f"⚠️ La empresa **{nombre_sel}** se enviará a la **Papelera** (no"
            " se perderán sus cuentas ni asientos, y podrás recuperarla cuando"
            " quieras desde la pestaña *Papelera* o en tus respaldos)."
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
              f"✅ '{nombre_sel}' fue enviada a la Papelera y respaldada."
          )
          st.rerun()

  # --- PESTAÑA 5: PAPELERA (RECUPERAR EMPRESAS BORRADAS POR ERROR) ---
  with tab_papelera:
    st.caption(
        "Aquí aparecen las empresas eliminadas. Puedes restaurarlas con todos"
        " sus asientos intactos o borrarlas definitivamente."
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
          f"ID {r['id']} - {r['razon_social']} ({r['cedula']})": int(r["id"])
          for _, r in df_papelera.iterrows()
      }
      sel_rec = st.selectbox(
          "Selecciona una empresa para recuperar:", list(opciones_rec.keys())
      )
      id_rec = opciones_rec[sel_rec]

      col_r1, col_r2 = st.columns(2)
      if col_r1.button("♻️ Restaurar Empresa al Sistema", type="primary"):
        cur = conn.cursor()
        cur.execute("UPDATE empresas SET activa = 1 WHERE id=?", (id_rec,))
        conn.commit()
        subir_bd_a_sheets(conn)
        st.success("✅ ¡Empresa restaurada nuevamente al sistema!")
        st.rerun()

      with col_r2.expander("🔥 Borrado Definitivo (Permanente)"):
        st.error(
            "Esto eliminará permanentemente la empresa, sus cuentas y sus"
            " asientos."
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
          st.success("✅ Empresa eliminada permanentemente.")
          st.rerun()
