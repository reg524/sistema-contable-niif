import sqlite3
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import streamlit as st

TABLAS_ESQUEMA = {
    "empresas": """
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cedula TEXT,
            razon_social TEXT,
            nombre_comercial TEXT,
            representante_legal TEXT,
            actividad_economica TEXT
        )
    """,
    "catalogo_cuentas": """
        CREATE TABLE IF NOT EXISTS catalogo_cuentas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER,
            codigo TEXT,
            nombre TEXT,
            tipo TEXT,
            acepta_movimiento INTEGER
        )
    """,
    "saldos_iniciales": """
        CREATE TABLE IF NOT EXISTS saldos_iniciales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER,
            cuenta_id INTEGER,
            fecha_inicio TEXT,
            debe REAL,
            haber REAL
        )
    """,
    "asientos_cabecera": """
        CREATE TABLE IF NOT EXISTS asientos_cabecera (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER,
            num_asiento INTEGER,
            fecha TEXT,
            periodo TEXT,
            concepto TEXT,
            num_comprobante TEXT,
            usuario_creacion TEXT
        )
    """,
    "asientos_detalle": """
        CREATE TABLE IF NOT EXISTS asientos_detalle (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asiento_id INTEGER,
            cuenta_id INTEGER,
            debe REAL,
            haber REAL
        )
    """,
    "plantilla_catalogo_cuentas": """
        CREATE TABLE IF NOT EXISTS plantilla_catalogo_cuentas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            nombre TEXT,
            tipo TEXT,
            acepta_movimiento INTEGER
        )
    """,
    "conciliaciones_bancarias": """
        CREATE TABLE IF NOT EXISTS conciliaciones_bancarias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER NOT NULL,
            cuenta_id INTEGER NOT NULL,
            anio INTEGER NOT NULL,
            mes INTEGER NOT NULL,
            fecha_desde TEXT NOT NULL,
            fecha_hasta TEXT NOT NULL,
            saldo_banco REAL DEFAULT 0,
            saldo_libros_ini REAL DEFAULT 0,
            saldo_libros_fin REAL DEFAULT 0,
            saldo_conciliado REAL DEFAULT 0,
            diferencia REAL DEFAULT 0,
            detalle_libros_json TEXT,
            partidas_banco_json TEXT,
            observaciones TEXT,
            fecha_actualizacion TEXT
        )
    """,
    "declaraciones_iva": """
        CREATE TABLE IF NOT EXISTS declaraciones_iva (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER NOT NULL,
            anio INTEGER NOT NULL,
            mes INTEGER NOT NULL,
            fecha_desde TEXT NOT NULL,
            fecha_hasta TEXT NOT NULL,
            ventas_netas_gravadas REAL DEFAULT 0,
            ventas_exentas REAL DEFAULT 0,
            iva_debito REAL DEFAULT 0,
            compras_netas_gravadas REAL DEFAULT 0,
            compras_exentas REAL DEFAULT 0,
            iva_credito REAL DEFAULT 0,
            saldo_favor_anterior REAL DEFAULT 0,
            retenciones_iva REAL DEFAULT 0,
            impuesto_por_pagar REAL DEFAULT 0,
            saldo_favor_siguiente REAL DEFAULT 0,
            detalle_tarifas_json TEXT,
            observaciones TEXT,
            fecha_actualizacion TEXT
        )
    """,
    "declaraciones_renta": """
        CREATE TABLE IF NOT EXISTS declaraciones_renta (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER NOT NULL,
            anio_fiscal INTEGER NOT NULL,
            fecha_desde TEXT NOT NULL,
            fecha_hasta TEXT NOT NULL,
            regimen TEXT NOT NULL,
            ingresos_brutos REAL DEFAULT 0,
            costos_gastos_contables REAL DEFAULT 0,
            utilidad_contable REAL DEFAULT 0,
            gastos_no_deducibles REAL DEFAULT 0,
            ingresos_no_gravables REAL DEFAULT 0,
            renta_neta_imponible REAL DEFAULT 0,
            impuesto_determinado REAL DEFAULT 0,
            pagos_parciales REAL DEFAULT 0,
            retenciones_2pct REAL DEFAULT 0,
            creditos_fiscales REAL DEFAULT 0,
            impuesto_por_pagar REAL DEFAULT 0,
            saldo_favor REAL DEFAULT 0,
            detalle_ajustes_json TEXT,
            observaciones TEXT,
            fecha_actualizacion TEXT
        )
    """,
}


def _conectar_sheets():
  if "gcp_service_account" not in st.secrets:
    return None
  creds = Credentials.from_service_account_info(
      st.secrets["gcp_service_account"],
      scopes=[
          "https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive",
      ],
  )
  client = gspread.authorize(creds)
  return client.open("BD_Sistema_Contable")


def inicializar_y_descargar_bd(conn):
  cur = conn.cursor()
  for ddl in TABLAS_ESQUEMA.values():
    cur.execute(ddl)
  conn.commit()

  if st.session_state.get("sheets_descargado", False):
    return

  try:
    sh = _conectar_sheets()
    if sh is None:
      return
    for tabla in TABLAS_ESQUEMA.keys():
      try:
        ws = sh.worksheet(tabla)
        datos = ws.get_all_records()
        if datos:
          df = pd.DataFrame(datos)
          cur.execute(f"DELETE FROM {tabla}")
          df.to_sql(tabla, conn, if_exists="append", index=False)
      except Exception:
        pass
    conn.commit()
    st.session_state["sheets_descargado"] = True
  except Exception as e:
    st.warning(f"Modo local activo (sin sincronizar con Sheets): {e}")


def subir_bd_a_sheets(conn):
  try:
    sh = _conectar_sheets()
    if sh is None:
      return
    for tabla in TABLAS_ESQUEMA.keys():
      try:
        df = pd.read_sql_query(f"SELECT * FROM {tabla}", conn).fillna("")
        ws = sh.worksheet(tabla)
        ws.clear()
        ws.update([df.columns.tolist()] + df.values.tolist())
      except Exception:
        pass
  except Exception:
    pass
