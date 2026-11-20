from datetime import datetime
import io
import os
import shutil
import sqlite3
import tempfile
import zipfile
import pandas as pd
import streamlit as st
from sync_sheets import TABLAS_ESQUEMA, subir_bd_a_sheets


def render_gestion_backups(conn, db_path="sistema_contable.db"):
  st.subheader("💾 Respaldo y Restauración de Base de Datos")
  st.caption(
      "Descarga una copia de seguridad íntegra (incluyendo empresas activas y"
      " en papelera) o restaura un respaldo anterior."
  )

  tab_descargar, tab_restaurar = st.tabs(
      ["⬇️ Descargar Backup", "⬆️ Recuperar / Restaurar Backup"]
  )

  # --- PESTAÑA 1: DESCARGAR BACKUP ---
  with tab_descargar:
    st.info(
        "El respaldo incluye todas tus empresas (activas y eliminadas en"
        " papelera), catálogos de cuentas, saldos iniciales, asientos,"
        " conciliaciones y declaraciones."
    )
    conn.commit()
    fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")

    col1, col2 = st.columns(2)

    if os.path.exists(db_path):
      with open(db_path, "rb") as f:
        db_bytes = f.read()
      col1.download_button(
          label="📥 Descargar Respaldo Directo (.db)",
          data=db_bytes,
          file_name=f"respaldo_contable_{fecha_str}.db",
          mime="application/octet-stream",
          type="primary",
          use_container_width=True,
      )

      zip_buffer = io.BytesIO()
      with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"sistema_contable_{fecha_str}.db", db_bytes)
        for tabla in TABLAS_ESQUEMA.keys():
          try:
            df_t = pd.read_sql_query(f"SELECT * FROM {tabla}", conn)
            zf.writestr(
                f"tablas_csv/{tabla}.csv",
                df_t.to_csv(index=False).encode("utf-8-sig"),
            )
          except Exception:
            pass
      col2.download_button(
          label="📦 Descargar Paquete Completo (.zip)",
          data=zip_buffer.getvalue(),
          file_name=f"respaldo_completo_{fecha_str}.zip",
          mime="application/zip",
          use_container_width=True,
      )

  # --- PESTAÑA 2: RECUPERAR BACKUP CON ADVERTENCIA ---
  with tab_restaurar:
    archivo_subido = st.file_uploader(
        "Selecciona el archivo de respaldo (.db o .zip):",
        type=["db", "sqlite", "sqlite3", "zip"],
        key="uploader_backup_db",
    )

    if archivo_subido is not None:
      st.error(
          "⚠️ **ADVERTENCIA CRÍTICA:** ¿Realmente desea recuperar este respaldo?"
          " En caso de recuperarlo, **SE BORRARÁN TODOS LOS DATOS ACTUALES** de"
          " todas las empresas, asientos, conciliaciones y declaraciones, y"
          " serán reemplazados por los del archivo seleccionado (tanto aquí"
          " como en Google Sheets)."
      )

      confirmar = st.checkbox(
          "Sí, entiendo la advertencia y confirmo que deseo borrar los datos"
          " actuales para restaurar este backup.",
          key="chk_confirmar_restauracion",
      )

      if st.button(
          "🔄 Ejecutar Restauración Ahora",
          type="primary",
          disabled=not confirmar,
      ):
        try:
          with tempfile.TemporaryDirectory() as tmpdir:
            ruta_temp_db = os.path.join(tmpdir, "backup_validado.db")

            if archivo_subido.name.lower().endswith(".zip"):
              with zipfile.ZipFile(io.BytesIO(archivo_subido.getvalue())) as zf:
                archivos_db = [
                    n
                    for n in zf.namelist()
                    if n.lower().endswith((".db", ".sqlite", ".sqlite3"))
                ]
                if not archivos_db:
                  st.error(
                      "El archivo .zip seleccionado no contiene ninguna base de"
                      " datos (.db)."
                  )
                  return
                with open(ruta_temp_db, "wb") as f_out:
                  f_out.write(zf.read(archivos_db[0]))
            else:
              with open(ruta_temp_db, "wb") as f_out:
                f_out.write(archivo_subido.getvalue())

            test_conn = sqlite3.connect(ruta_temp_db)
            tablas_en_backup = pd.read_sql_query(
                "SELECT name FROM sqlite_master WHERE type='table'", test_conn
            )["name"].tolist()

            if "empresas" not in tablas_en_backup:
              test_conn.close()
              st.error(
                  "El archivo subido no corresponde a una base de datos válida"
                  " del Sistema Contable."
              )
              return

            # Asegurar que todas las empresas del backup tengan activa=1 si no existía la columna
            cols_emp = [
                r[1]
                for r in test_conn.execute(
                    "PRAGMA table_info(empresas)"
                ).fetchall()
            ]
            if "activa" not in cols_emp:
              test_conn.execute(
                  "ALTER TABLE empresas ADD COLUMN activa INTEGER DEFAULT 1"
              )
            test_conn.execute(
                "UPDATE empresas SET activa = 1 WHERE activa IS NULL"
            )
            test_conn.commit()
            test_conn.close()

            # Cerrar conexión activa, reemplazar BD y limpiar cachés
            conn.commit()
            conn.close()
            shutil.copy2(ruta_temp_db, db_path)

            st.cache_data.clear()
            st.session_state.pop("ultima_firma_sheets", None)
            st.session_state["sheets_descargado"] = True

            nueva_conn = sqlite3.connect(db_path, check_same_thread=False)
            subir_bd_a_sheets(nueva_conn)
            nueva_conn.close()

            st.success(
                "✅ ¡Respaldo restaurado exitosamente! Los datos anteriores"
                " fueron reemplazados y sincronizados con Google Sheets."
            )
            st.rerun()

        except Exception as e:
          st.error(f"Error al intentar restaurar el respaldo: {e}")
