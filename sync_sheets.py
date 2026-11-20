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

COLS_NUMERICAS = {
    "id",
    "empresa_id",
    "cuenta_id",
    "asiento_id",
    "num_asiento",
    "acepta_movimiento",
    "anio",
    "mes",
    "anio_fiscal",
    "debe",
    "haber",
    "saldo_banco",
    "saldo_libros_ini",
    "saldo_libros_fin",
    "saldo_conciliado",
    "diferencia",
    "ventas_netas_gravadas",
    "ventas_exentas",
    "iva_debito",
    "compras_netas_gravadas",
    "compras_exentas",
    "iva_credito",
    "saldo_favor_anterior",
    "retenciones_iva",
    "impuesto_por_pagar",
    "saldo_favor_siguiente",
    "ingresos_brutos",
    "costos_gastos_contables",
    "utilidad_contable",
    "gastos_no_deducibles",
    "ingresos_no_gravables",
    "renta_neta_imponible",
    "impuesto_determinado",
    "pagos_parciales",
    "retenciones_2pct",
    "creditos_fiscales",
    "saldo_favor",
}


def _buscar_credenciales(d):
  if isinstance(d, dict):
    if "private_key" in d and "client_email" in d:
      return d
    for v in d.values():
      res = _buscar_credenciales(v)
      if res:
        return res
  return None


@st.cache_resource(show_spinner=False)
def _obtener_cliente_gspread():
  info_creds = _buscar_credenciales(st.secrets.to_dict())
  if not info_creds:
    return None
  creds = Credentials.from_service_account_info(
      info_creds,
      scopes=[
          "https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive",
      ],
  )
  return gspread.authorize(creds)


def _conectar_sheets():
  client = _obtener_cliente_gspread()
  if client is None:
    return None
  return client.open("BD_Sistema_Contable")


@st.cache_data(ttl=600, show_spinner=False)
def _leer_todas_las_hojas_cache():
  sh = _conectar_sheets()
  if sh is None:
    return None
  tablas = list(TABLAS_ESQUEMA.keys())
  rangos = [f"'{t}'!A:Z" for t in tablas]
  resp = sh.values_batch_get(rangos)
  resultado = {}
  for tabla, vr in zip(tablas, resp.get("valueRanges", [])):
    filas = vr.get("values", [])
    if len(filas) > 1:
      headers = [str(h).strip() for h in filas[0] if str(h).strip()]
      n_cols = len(headers)
      filas_norm = [
          (f + [""] * (n_cols - len(f)))[:n_cols] for f in filas[1:] if any(f)
      ]
      if filas_norm:
        resultado[tabla] = pd.DataFrame(filas_norm, columns=headers)
  return resultado


def inicializar_y_descargar_bd(conn):
  cur = conn.cursor()
  for ddl in TABLAS_ESQUEMA.values():
    cur.execute(ddl)
  conn.commit()

  if st.session_state.get("sheets_descargado", False):
    return

  try:
    datos_hojas = _leer_todas_las_hojas_cache()
    if datos_hojas is None:
      return
    for tabla, df in datos_hojas.items():
      try:
        df_copia = df.copy()
        for col in df_copia.columns:
          if col in COLS_NUMERICAS:
            df_copia[col] = pd.to_numeric(
                df_copia[col], errors="coerce"
            ).fillna(0)
          else:
            df_copia[col] = df_copia[col].astype(str)
        df_copia.to_sql(tabla, conn, if_exists="replace", index=False)
      except Exception as e:
        st.warning(f"Aviso descargando '{tabla}': {e}")
    conn.commit()
    st.session_state["sheets_descargado"] = True
  except Exception as e:
    st.warning(f"Modo local activo (sin sincronizar con Sheets): {e}")


def subir_bd_a_sheets(conn):
  try:
    sh = _conectar_sheets()
    if sh is None:
      return

    # Verificar si hubo cambios reales en la base de datos antes de escribir
    firma_actual = []
    datos_batch = []
    rangos_limpiar = []

    for tabla in TABLAS_ESQUEMA.keys():
      df = pd.read_sql_query(f"SELECT * FROM {tabla}", conn).fillna("")
      valores = [df.columns.tolist()] + df.values.tolist()
      firma_actual.append((tabla, len(valores), str(valores[-1]) if valores else ""))
      rangos_limpiar.append(f"'{tabla}'!A:Z")
      datos_batch.append({"range": f"'{tabla}'!A1", "values": valores})

    firma_str = str(hash(tuple(firma_actual)))
    if st.session_state.get("ultima_firma_sheets") == firma_str:
      return

    sh.values_batch_clear({"ranges": rangos_limpiar})
    sh.values_batch_update(
        {"valueInputOption": "USER_ENTERED", "data": datos_batch}
    )
    st.session_state["ultima_firma_sheets"] = firma_str
  except Exception:
    pass
