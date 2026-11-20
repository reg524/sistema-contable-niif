from sync_sheets import subir_bd_a_sheets
import pandas as pd
import streamlit as st


def render_gestion_empresas(conn):
  st.subheader("🏢 Administración de Empresas (Crear, Editar y Eliminar)")
  cur = conn.cursor()
  df_emp = pd.read_sql_query("SELECT * FROM empresas ORDER BY id", conn)

  tab_crear, tab_editar, tab_eliminar = st.tabs(
      ["➕ Crear Empresa", "✏️ Editar Empresa", "🗑️ Eliminar Empresa"]
  )

  # --- PESTAÑA 1: CREAR CON VALIDACIÓN DE DUPLICADOS ---
  with tab_crear:
    with st.form("form_crear_empresa_validada"):
      c1, c2 = st.columns(2)
      cedula = c1.text_input("Cédula Jurídica / Física *").strip()
      razon_social = c2.text_input("Razón Social / Nombre *").strip()
      nombre_comercial = c1.text_input("Nombre Comercial").strip()
      representante = c2.text_input("Representante Legal").strip()
      actividad = st.text_input("Actividad Económica").strip()

      forzar_duplicado = st.checkbox(
          "⚠️ Acepto crear otra contabilidad aunque ya exista una empresa con el"
          " mismo nombre o cédula"
      )
      submitted = st.form_submit_button("Guardar Empresa", type="primary")

    if submitted:
      if not cedula or not razon_social:
        st.error("La cédula y la razón social son obligatorias.")
      else:
        coincidencias = df_emp[
            (
                df_emp["razon_social"].astype(str).str.strip().str.lower()
                == razon_social.lower()
            )
            | (
                df_emp["cedula"].astype(str).str.strip().str.lower()
                == cedula.lower()
            )
        ]
        if not coincidencias.empty and not forzar_duplicado:
          st.warning(
              "⚠️ **Ya existe una empresa registrada con ese nombre o cédula:**"
          )
          st.dataframe(
              coincidencias[[
                  "id",
                  "cedula",
                  "razon_social",
                  "nombre_comercial",
              ]],
              use_container_width=True,
              hide_index=True,
          )
          st.info(
              "Si realmente deseas llevar una segunda contabilidad separada con"
              " el mismo nombre, marca la casilla de confirmación arriba y"
              " vuelve a presionar **Guardar Empresa**."
          )
        else:
          nuevo_id = int(df_emp["id"].max() + 1) if not df_emp.empty else 1
          cur.execute(
              """
                        INSERT INTO empresas (id, cedula, razon_social, nombre_comercial, representante_legal, actividad_economica)
                        VALUES (?, ?, ?, ?, ?, ?)
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
          st.success(
              f"✅ Empresa '{razon_social}' creada correctamente (ID:"
              f" {nuevo_id})."
          )
          st.rerun()

  # --- PESTAÑA 2: EDITAR EMPRESA EXISTENTE ---
  with tab_editar:
    if df_emp.empty:
      st.info("No hay empresas registradas para editar.")
    else:
      opciones_ed = {
          f"ID {row['id']} - {row['razon_social']} ({row['cedula']})": int(
              row["id"]
          )
          for _, row in df_emp.iterrows()
      }
      sel_ed = st.selectbox(
          "Selecciona la empresa a modificar:",
          list(opciones_ed.keys()),
          key="sel_ed_emp",
      )
      emp_id_ed = opciones_ed[sel_ed]
      fila_ed = df_emp[df_emp["id"] == emp_id_ed].iloc[0]

      with st.form("form_editar_empresa"):
        c1, c2 = st.columns(2)
        ed_cedula = c1.text_input(
            "Cédula Jurídica / Física", value=str(fila_ed.get("cedula", ""))
        ).strip()
        ed_razon = c2.text_input(
            "Razón Social / Nombre", value=str(fila_ed.get("razon_social", ""))
        ).strip()
        ed_comercial = c1.text_input(
            "Nombre Comercial", value=str(fila_ed.get("nombre_comercial", ""))
        ).strip()
        ed_rep = c2.text_input(
            "Representante Legal",
            value=str(fila_ed.get("representante_legal", "")),
        ).strip()
        ed_act = st.text_input(
            "Actividad Económica",
            value=str(fila_ed.get("actividad_economica", "")),
        ).strip()

        if st.form_submit_button("💾 Guardar Cambios", type="primary"):
          cur.execute(
              """
                        UPDATE empresas
                        SET cedula=?, razon_social=?, nombre_comercial=?, representante_legal=?, actividad_economica=?
                        WHERE id=?
                    """,
              (
                  ed_cedula,
                  ed_razon,
                  ed_comercial,
                  ed_rep,
                  ed_act,
                  emp_id_ed,
              ),
          )
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success("✅ Datos de la empresa actualizados y sincronizados.")
          st.rerun()

  # --- PESTAÑA 3: ELIMINAR EMPRESA (PARA QUITAR DUPLICADOS) ---
  with tab_eliminar:
    if df_emp.empty:
      st.info("No hay empresas registradas.")
    else:
      st.dataframe(
          df_emp[[
              "id",
              "cedula",
              "razon_social",
              "nombre_comercial",
              "representante_legal",
          ]],
          use_container_width=True,
          hide_index=True,
      )
      opciones_del = {
          f"ID {row['id']} - {row['razon_social']} ({row['cedula']})": int(
              row["id"]
          )
          for _, row in df_emp.iterrows()
      }
      sel_del = st.selectbox(
          "Selecciona la empresa que deseas eliminar:",
          list(opciones_del.keys()),
          key="sel_del_emp",
      )
      emp_id_del = opciones_del[sel_del]

      # Mostrar cuántos registros tiene asociados para que el usuario sepa cuál es la copia vacía y cuál tiene datos
      n_cuentas = pd.read_sql_query(
          "SELECT COUNT(*) c FROM catalogo_cuentas WHERE empresa_id=?",
          conn,
          params=(emp_id_del,),
      )["c"].iloc[0]
      n_asientos = pd.read_sql_query(
          "SELECT COUNT(*) c FROM asientos_cabecera WHERE empresa_id=?",
          conn,
          params=(emp_id_del,),
      )["c"].iloc[0]
      st.info(
          f"Esta empresa (ID {emp_id_del}) tiene **{n_cuentas} cuentas** y"
          f" **{n_asientos} asientos** asociados."
      )

      confirmar_borrado = st.checkbox(
          "Confirmo que deseo eliminar permanentemente esta empresa y todos sus"
          " registros asociados."
      )
      if st.button("🗑️ Eliminar Empresa Seleccionada", type="primary"):
        if not confirmar_borrado:
          st.warning("Marca primero la casilla de confirmación.")
        else:
          asientos_ids = pd.read_sql_query(
              "SELECT id FROM asientos_cabecera WHERE empresa_id=?",
              conn,
              params=(emp_id_del,),
          )["id"].tolist()
          if asientos_ids:
            placeholders = ",".join("?" * len(asientos_ids))
            cur.execute(
                "DELETE FROM asientos_detalle WHERE asiento_id IN"
                f" ({placeholders})",
                asientos_ids,
            )
          for tabla in [
              "asientos_cabecera",
              "saldos_iniciales",
              "catalogo_cuentas",
              "conciliaciones_bancarias",
              "declaraciones_iva",
              "declaraciones_renta",
          ]:
            try:
              cur.execute(
                  f"DELETE FROM {tabla} WHERE empresa_id=?", (emp_id_del,)
              )
            except Exception:
              pass
          cur.execute("DELETE FROM empresas WHERE id=?", (emp_id_del,))
          conn.commit()
          subir_bd_a_sheets(conn)
          st.success(f"✅ Empresa ID {emp_id_del} eliminada correctamente.")
          st.rerun()
