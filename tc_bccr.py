import datetime
import sqlite3
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import streamlit as st

BCCR_WS_URL = (
    "https://gee.bccr.fi.cr/Indicadores/Suscripciones/WS/wsindicadoreseconomicos.asmx/ObtenerIndicadoresEconomicos"
)


def asegurar_tabla_tc(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS tipos_cambio_bccr (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            indicador TEXT NOT NULL,
            valor REAL NOT NULL,
            fuente TEXT DEFAULT 'BCCR',
            fecha_consulta TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_tc_fecha_ind
        ON tipos_cambio_bccr (fecha, indicador)
        """
    )
    conn.commit()


def _obtener_credenciales_bccr():
    email = ""
    token = ""
    nombre = "SistemaContableCR"
    try:
        if "bccr" in st.secrets:
            email = str(st.secrets["bccr"].get("email", "")).strip()
            token = str(st.secrets["bccr"].get("token", "")).strip()
            nombre = str(st.secrets["bccr"].get("nombre", nombre)).strip()
    except Exception:
        pass

    import os

    email = email or os.getenv("BCCR_EMAIL", "").strip()
    token = token or os.getenv("BCCR_TOKEN", "").strip()
    nombre = os.getenv("BCCR_NOMBRE", nombre).strip()
    return email, token, nombre


def consultar_tc_bccr_webservice(fecha_obj, indicador="317"):
    """
    Consulta el servicio web del BCCR para una fecha y código de indicador:
    - '317': Tipo de cambio de compra del dólar
    - '318': Tipo de cambio de venta del dólar
    """
    email, token, nombre = _obtener_credenciales_bccr()
    if not email or not token:
        return None, "No se encontraron credenciales BCCR (bccr.email / bccr.token) en st.secrets."

    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    params = {
        "Indicador": str(indicador),
        "FechaInicio": fecha_str,
        "FechaFinal": fecha_str,
        "Nombre": nombre,
        "SubNiveles": "N",
        "CorreoElectronico": email,
        "Token": token,
    }
    url = f"{BCCR_WS_URL}?{urllib.parse.urlencode(params)}"

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=12) as resp:
            xml_data = resp.read().decode("utf-8", errors="replace")

        root = ET.fromstring(xml_data)
        for elem in root.iter():
            if elem.tag.endswith("NUM_VALOR") and elem.text:
                valor = float(elem.text.strip().replace(",", "."))
                return round(valor, 4), "BCCR Web Service"

        return None, f"El BCCR no devolvió valor numérico para {fecha_str} (indicador {indicador})."
    except Exception as e:
        return None, f"Error consultando Web Service BCCR: {e}"


def obtener_tc_bccr(conn, fecha_obj, tipo_tc="compra", tc_manual_fallback=None):
    """
    Obtiene el tipo de cambio del BCCR ('compra' -> 317, 'venta' -> 318):
    1. Revisa caché local en SQLite para esa fecha e indicador.
    2. Si no está, consulta el Web Service del BCCR y lo guarda.
    3. Si falla la consulta externa, usa tc_manual_fallback si está disponible.
    """
    asegurar_tabla_tc(conn)
    ind_code = "318" if str(tipo_tc).lower().startswith("ven") else "317"
    fecha_iso = fecha_obj.strftime("%Y-%m-%d")

    cur = conn.cursor()
    row = cur.execute(
        "SELECT valor, fuente FROM tipos_cambio_bccr WHERE fecha = ? AND indicador = ?",
        (fecha_iso, ind_code),
    ).fetchone()
    if row and row[0] is not None and float(row[0]) > 0:
        return float(row[0]), f"Caché ({row[1]})"

    valor_ws, detalle_ws = consultar_tc_bccr_webservice(fecha_obj, indicador=ind_code)
    if valor_ws is not None and valor_ws > 0:
        guardar_tc_en_cache(conn, fecha_obj, ind_code, valor_ws, fuente="BCCR")
        return float(valor_ws), detalle_ws

    if tc_manual_fallback is not None and float(tc_manual_fallback) > 0:
        return float(tc_manual_fallback), f"Manual / Respaldo ({detalle_ws})"

    return None, detalle_ws


def guardar_tc_en_cache(conn, fecha_obj, indicador, valor, fuente="BCCR"):
    asegurar_tabla_tc(conn)
    fecha_iso = fecha_obj.strftime("%Y-%m-%d")
    ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """
        INSERT INTO tipos_cambio_bccr (fecha, indicador, valor, fuente, fecha_consulta)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(fecha, indicador) DO UPDATE SET
            valor = excluded.valor,
            fuente = excluded.fuente,
            fecha_consulta = excluded.fecha_consulta
        """,
        (fecha_iso, str(indicador), float(valor), str(fuente), ahora),
    )
    conn.commit()
