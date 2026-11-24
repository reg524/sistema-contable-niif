import datetime
import io
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from fpdf import FPDF


def _fmt_crc(valor):
    """Formatea montos en estilo costarricense/profesional."""
    val = float(valor or 0.0)
    return f"{val:,.2f}"


def _limpiar_pdf(texto):
    """Convierte caracteres Unicode a Latin-1 seguro para fpdf2."""
    if texto is None:
        return ""
    mapa = {
        "₡": "CRC ",
        "¢": "CRC ",
        "—": "-",
        "–": "-",
        "“": '"',
        "”": '"',
        "’": "'",
        "‘": "'",
        "•": "-",
    }
    s = str(texto)
    for k, v in mapa.items():
        s = s.replace(k, v)
    return s.encode("latin-1", "replace").decode("latin-1")


def _obtener_saldos_a_fecha(conn, empresa_id, fecha_corte):
    """
    Calcula el saldo acumulado (Debe - Haber) en Colones (₡) por cuenta
    desde saldos_iniciales hasta fecha_corte inclusive.
    """
    fecha_str = str(fecha_corte)
    df_si = pd.read_sql_query(
        """
        SELECT cuenta_id, COALESCE(SUM(debe - haber), 0.0) AS saldo
        FROM saldos_iniciales
        WHERE empresa_id = ? AND fecha_inicio <= ?
        GROUP BY cuenta_id
        """,
        conn,
        params=(int(empresa_id), fecha_str),
    )

    df_mov = pd.read_sql_query(
        """
        SELECT ad.cuenta_id, COALESCE(SUM(ad.debe - ad.haber), 0.0) AS saldo
        FROM asientos_detalle ad
        JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
        WHERE ac.empresa_id = ? AND ac.fecha <= ?
        GROUP BY ad.cuenta_id
        """,
        conn,
        params=(int(empresa_id), fecha_str),
    )

    saldos = {}
    for _, r in df_si.iterrows():
        cid = int(r["cuenta_id"])
        saldos[cid] = saldos.get(cid, 0.0) + float(r["saldo"] or 0.0)
    for _, r in df_mov.iterrows():
        cid = int(r["cuenta_id"])
        saldos[cid] = saldos.get(cid, 0.0) + float(r["saldo"] or 0.0)
    return saldos


def _obtener_movimientos_periodo(conn, empresa_id, fecha_desde, fecha_hasta):
    """Obtiene suma de Debe y Haber por cuenta en un rango de fechas."""
    df = pd.read_sql_query(
        """
        SELECT ad.cuenta_id,
               COALESCE(SUM(ad.debe), 0.0) AS debe,
               COALESCE(SUM(ad.haber), 0.0) AS haber
        FROM asientos_detalle ad
        JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
        WHERE ac.empresa_id = ? AND ac.fecha BETWEEN ? AND ?
        GROUP BY ad.cuenta_id
        """,
        conn,
        params=(int(empresa_id), str(fecha_desde), str(fecha_hasta)),
    )
    res = {}
    for _, r in df.iterrows():
        res[int(r["cuenta_id"])] = {
            "debe": float(r["debe"] or 0.0),
            "haber": float(r["haber"] or 0.0),
        }
    return res


def _cargar_catalogo(conn, empresa_id):
    df = pd.read_sql_query(
        """
        SELECT id, codigo, nombre, tipo, acepta_movimiento
        FROM catalogo_cuentas
        WHERE empresa_id = ?
        ORDER BY codigo
        """,
        conn,
        params=(int(empresa_id),),
    )
    if not df.empty:
        df["cod_str"] = df["codigo"].astype(str).str.strip()
        df["tipo_up"] = df["tipo"].astype(str).str.upper().str.strip()
        df["nom_up"] = df["nombre"].astype(str).str.upper().str.strip()
    return df


# ==============================================================================
# 1. CONSTRUCCIÓN DE ESTADOS FINANCIEROS SEGÚN MACHOTES
# ==============================================================================
def construir_estado_situacion_comparativo(conn, empresa_id, fecha_ant, fecha_act):
    """Construye Estado de Situación Financiera Comparativo según machote."""
    cat = _cargar_catalogo(conn, empresa_id)
    if cat.empty:
        return pd.DataFrame(), {}

    s_ant = _obtener_saldos_a_fecha(conn, empresa_id, fecha_ant)
    s_act = _obtener_saldos_a_fecha(conn, empresa_id, fecha_act)

    activos = cat[cat["tipo_up"].str.contains("ACTIVO") | cat["cod_str"].str.startswith("1")]
    pasivos = cat[
        (cat["tipo_up"].str.contains("PASIVO") | cat["cod_str"].str.startswith("2"))
        & (~cat.index.isin(activos.index))
    ]
    patrimonio = cat[
        (cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL") | cat["cod_str"].str.startswith("3"))
        & (~cat.index.isin(activos.index))
        & (~cat.index.isin(pasivos.index))
    ]
    resultados = cat[
        (
            cat["tipo_up"].str.contains("INGRESO|GASTO|COSTO|EGRESO")
            | cat["cod_str"].str.startswith(("4", "5", "6"))
        )
        & (~cat.index.isin(activos.index))
        & (~cat.index.isin(pasivos.index))
        & (~cat.index.isin(patrimonio.index))
    ]

    act_mov = activos[activos["acepta_movimiento"] == 1].copy()
    pas_mov = pasivos[pasivos["acepta_movimiento"] == 1].copy()
    pat_mov = patrimonio[patrimonio["acepta_movimiento"] == 1].copy()

    mask_fijo = (
        act_mov["cod_str"].str.startswith(("1.2", "1.3"))
        | act_mov["nom_up"].str.contains(
            "FIJO|NO CORRIENTE|TERRENO|EDIFICIO|VEHICULO|TRANSPORTE|MOBILIARIO|EQUIPO|HERRAMIENTA|HIDRANTE|LINEA|CONDUCCION|DEPRECIACION|OBRA"
        )
    )
    act_circ = act_mov[~mask_fijo]
    act_fijo = act_mov[mask_fijo]

    mask_pas_lp = (
        pas_mov["cod_str"].str.startswith("2.2")
        | pas_mov["nom_up"].str.contains("LARGO PLAZO|NO CORRIENTE|HIPOTECA|PRESTAMO L.P")
    )
    pas_circ = pas_mov[~mask_pas_lp]
    pas_lp = pas_mov[mask_pas_lp]

    filas = []
    nota_idx = 1

    def _agregar_grupo(titulo_sec, subtitulo, df_sub, signo=1.0, asignar_nota=True):
        nonlocal nota_idx
        filas.append(
            {
                "Rubro": titulo_sec if titulo_sec else subtitulo,
                "Nota": "",
                "Anterior": None,
                "Actual": None,
                "TipoFila": "SECCION",
            }
        )
        if titulo_sec and subtitulo:
            filas.append(
                {
                    "Rubro": subtitulo,
                    "Nota": "",
                    "Anterior": None,
                    "Actual": None,
                    "TipoFila": "SUBSECCION",
                }
            )
        tot_a = 0.0
        tot_b = 0.0
        for _, r in df_sub.iterrows():
            cid = int(r["id"])
            va = round(signo * s_ant.get(cid, 0.0), 2)
            vb = round(signo * s_act.get(cid, 0.0), 2)
            if abs(va) >= 0.005 or abs(vb) >= 0.005:
                n_str = str(nota_idx) if (asignar_nota and va >= 0 and vb >= 0) else ""
                if n_str:
                    nota_idx += 1
                filas.append(
                    {
                        "Rubro": f"  {r['nombre']}",
                        "Nota": n_str,
                        "Anterior": va,
                        "Actual": vb,
                        "TipoFila": "DETALLE",
                    }
                )
                tot_a += va
                tot_b += vb
        return round(tot_a, 2), round(tot_b, 2)

    # ACTIVOS
    t_ac_ant, t_ac_act = _agregar_grupo("ACTIVOS", "CIRCULANTE", act_circ, signo=1.0, asignar_nota=True)
    filas.append(
        {"Rubro": "Total de Circulante", "Nota": "", "Anterior": t_ac_ant, "Actual": t_ac_act, "TipoFila": "SUBTOTAL"}
    )

    t_af_ant, t_af_act = _agregar_grupo("", "FIJOS (NO CORRIENTES)", act_fijo, signo=1.0, asignar_nota=False)
    filas.append(
        {"Rubro": "Total de Fijos", "Nota": "", "Anterior": t_af_ant, "Actual": t_af_act, "TipoFila": "SUBTOTAL"}
    )

    tot_act_ant = round(t_ac_ant + t_af_ant, 2)
    tot_act_act = round(t_ac_act + t_af_act, 2)
    filas.append(
        {"Rubro": "TOTAL DE ACTIVOS", "Nota": "", "Anterior": tot_act_ant, "Actual": tot_act_act, "TipoFila": "TOTAL"}
    )

    # PASIVOS
    t_pc_ant, t_pc_act = _agregar_grupo("PASIVOS", "CIRCULANTES", pas_circ, signo=-1.0, asignar_nota=True)
    filas.append(
        {"Rubro": "Total Pasivo Circulante", "Nota": "", "Anterior": t_pc_ant, "Actual": t_pc_act, "TipoFila": "SUBTOTAL"}
    )

    t_pl_ant, t_pl_act = 0.0, 0.0
    if not pas_lp.empty:
        t_pl_ant, t_pl_act = _agregar_grupo("", "NO CORRIENTES (LARGO PLAZO)", pas_lp, signo=-1.0, asignar_nota=True)
        if abs(t_pl_ant) >= 0.005 or abs(t_pl_act) >= 0.005:
            filas.append(
                {
                    "Rubro": "Total Pasivo No Circulante",
                    "Nota": "",
                    "Anterior": t_pl_ant,
                    "Actual": t_pl_act,
                    "TipoFila": "SUBTOTAL",
                }
            )

    tot_pas_ant = round(t_pc_ant + t_pl_ant, 2)
    tot_pas_act = round(t_pc_act + t_pl_act, 2)
    filas.append(
        {"Rubro": "TOTAL DE PASIVOS", "Nota": "", "Anterior": tot_pas_ant, "Actual": tot_pas_act, "TipoFila": "TOTAL"}
    )

    # PATRIMONIO
    t_pat_cta_ant, t_pat_cta_act = _agregar_grupo("PATRIMONIO", "", pat_mov, signo=-1.0, asignar_nota=True)

    # Resultado acumulado / del período
    res_ant = round(sum(-s_ant.get(int(cid), 0.0) for cid in resultados["id"]), 2)
    res_act = round(sum(-s_act.get(int(cid), 0.0) for cid in resultados["id"]), 2)
    if abs(res_ant) >= 0.005 or abs(res_act) >= 0.005:
        filas.append(
            {
                "Rubro": "  Excedentes / Resultados del Período",
                "Nota": str(nota_idx),
                "Anterior": res_ant,
                "Actual": res_act,
                "TipoFila": "DETALLE",
            }
        )

    tot_pat_ant = round(t_pat_cta_ant + res_ant, 2)
    tot_pat_act = round(t_pat_cta_act + res_act, 2)
    filas.append(
        {"Rubro": "Total de Patrimonio", "Nota": "", "Anterior": tot_pat_ant, "Actual": tot_pat_act, "TipoFila": "SUBTOTAL"}
    )

    tot_pp_ant = round(tot_pas_ant + tot_pat_ant, 2)
    tot_pp_act = round(tot_pas_act + tot_pat_act, 2)
    filas.append(
        {
            "Rubro": "TOTAL DE PASIVO Y PATRIMONIO",
            "Nota": "",
            "Anterior": tot_pp_ant,
            "Actual": tot_pp_act,
            "TipoFila": "TOTAL",
        }
    )

    df_esf = pd.DataFrame(filas)
    resumen = {
        "tot_act_ant": tot_act_ant,
        "tot_act_act": tot_act_act,
        "tot_pas_ant": tot_pas_ant,
        "tot_pas_act": tot_pas_act,
        "tot_pat_ant": tot_pat_ant,
        "tot_pat_act": tot_pat_act,
        "tot_pp_ant": tot_pp_ant,
        "tot_pp_act": tot_pp_act,
        "cajas_bancos_ant": round(
            sum(
                s_ant.get(int(r["id"]), 0.0)
                for _, r in act_circ.iterrows()
                if any(w in r["nom_up"] for w in ["CAJA", "BANCO", "EFECTIVO"])
            ),
            2,
        ),
        "cajas_bancos_act": round(
            sum(
                s_act.get(int(r["id"]), 0.0)
                for _, r in act_circ.iterrows()
                if any(w in r["nom_up"] for w in ["CAJA", "BANCO", "EFECTIVO"])
            ),
            2,
        ),
    }
    return df_esf, resumen


def construir_estado_resultados(conn, empresa_id, fecha_desde, fecha_hasta, fecha_ant_desde=None, fecha_ant_hasta=None):
    """Construye Estado de Resultados (comparativo si se pasan fechas anteriores)."""
    cat = _cargar_catalogo(conn, empresa_id)
    if cat.empty:
        return pd.DataFrame(), {"ingresos": 0.0, "gastos": 0.0, "utilidad": 0.0}

    mov_act = _obtener_movimientos_periodo(conn, empresa_id, fecha_desde, fecha_hasta)
    mov_ant = (
        _obtener_movimientos_periodo(conn, empresa_id, fecha_ant_desde, fecha_ant_hasta)
        if (fecha_ant_desde and fecha_ant_hasta)
        else {}
    )

    ing_ctas = cat[
        (cat["acepta_movimiento"] == 1)
        & (cat["tipo_up"].str.contains("INGRESO") | cat["cod_str"].str.startswith("4"))
    ]
    gast_ctas = cat[
        (cat["acepta_movimiento"] == 1)
        & (
            cat["tipo_up"].str.contains("GASTO|COSTO|EGRESO")
            | cat["cod_str"].str.startswith(("5", "6"))
        )
    ]

    filas = [
        {"Rubro": "INGRESOS", "Nota": "11", "Anterior": None, "Actual": None, "TipoFila": "SECCION"}
    ]
    tot_ing_ant = 0.0
    tot_ing_act = 0.0
    for _, r in ing_ctas.iterrows():
        cid = int(r["id"])
        va = round(mov_ant.get(cid, {}).get("haber", 0.0) - mov_ant.get(cid, {}).get("debe", 0.0), 2)
        vb = round(mov_act.get(cid, {}).get("haber", 0.0) - mov_act.get(cid, {}).get("debe", 0.0), 2)
        if abs(va) >= 0.005 or abs(vb) >= 0.005:
            filas.append(
                {"Rubro": f"  {r['nombre']}", "Nota": "", "Anterior": va, "Actual": vb, "TipoFila": "DETALLE"}
            )
            tot_ing_ant += va
            tot_ing_act += vb

    tot_ing_ant = round(tot_ing_ant, 2)
    tot_ing_act = round(tot_ing_act, 2)
    filas.append(
        {"Rubro": "TOTAL DE INGRESOS", "Nota": "", "Anterior": tot_ing_ant, "Actual": tot_ing_act, "TipoFila": "SUBTOTAL"}
    )

    filas.append(
        {"Rubro": "GASTOS DE OPERACIÓN", "Nota": "12", "Anterior": None, "Actual": None, "TipoFila": "SECCION"}
    )
    tot_gas_ant = 0.0
    tot_gas_act = 0.0
    for _, r in gast_ctas.iterrows():
        cid = int(r["id"])
        va = round(mov_ant.get(cid, {}).get("debe", 0.0) - mov_ant.get(cid, {}).get("haber", 0.0), 2)
        vb = round(mov_act.get(cid, {}).get("debe", 0.0) - mov_act.get(cid, {}).get("haber", 0.0), 2)
        if abs(va) >= 0.005 or abs(vb) >= 0.005:
            filas.append(
                {"Rubro": f"  {r['nombre']}", "Nota": "", "Anterior": va, "Actual": vb, "TipoFila": "DETALLE"}
            )
            tot_gas_ant += va
            tot_gas_act += vb

    tot_gas_ant = round(tot_gas_ant, 2)
    tot_gas_act = round(tot_gas_act, 2)
    filas.append(
        {"Rubro": "TOTAL DE GASTOS", "Nota": "", "Anterior": tot_gas_ant, "Actual": tot_gas_act, "TipoFila": "SUBTOTAL"}
    )

    util_ant = round(tot_ing_ant - tot_gas_ant, 2)
    util_act = round(tot_ing_act - tot_gas_act, 2)
    filas.append(
        {"Rubro": "UTILIDAD / EXCEDENTE NETO", "Nota": "", "Anterior": util_ant, "Actual": util_act, "TipoFila": "TOTAL"}
    )

    df_er = pd.DataFrame(filas)
    resumen = {
        "ingresos_ant": tot_ing_ant,
        "ingresos": tot_ing_act,
        "gastos_ant": tot_gas_ant,
        "gastos": tot_gas_act,
        "utilidad_ant": util_ant,
        "utilidad": util_act,
    }
    return df_er, resumen


def construir_cambios_patrimonio(conn, empresa_id, fecha_ant, fecha_act, utilidad_periodo_act, utilidad_periodo_ant):
    """Construye Estado de Cambios en el Patrimonio según machote matricial."""
    cat = _cargar_catalogo(conn, empresa_id)
    if cat.empty:
        return pd.DataFrame()

    s_ant = _obtener_saldos_a_fecha(conn, empresa_id, fecha_ant)
    s_act = _obtener_saldos_a_fecha(conn, empresa_id, fecha_act)

    pat = cat[
        (cat["acepta_movimiento"] == 1)
        & (cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL") | cat["cod_str"].str.startswith("3"))
    ]

    aportes_ant = 0.0
    aportes_act = 0.0
    retenidas_ant = 0.0
    retenidas_act = 0.0

    for _, r in pat.iterrows():
        cid = int(r["id"])
        va = round(-s_ant.get(cid, 0.0), 2)
        vb = round(-s_act.get(cid, 0.0), 2)
        if any(k in r["nom_up"] for k in ["CAPITAL", "APORTE", "DONACION", "ACCION", "RESERVA"]):
            aportes_ant += va
            aportes_act += vb
        else:
            retenidas_ant += va
            retenidas_act += vb

    var_aportes = round(aportes_act - aportes_ant, 2)
    var_retenidas = round(retenidas_act - retenidas_ant, 2)
    tot_ini = round(aportes_ant + retenidas_ant + utilidad_periodo_ant, 2)
    tot_fin = round(aportes_act + retenidas_act + utilidad_periodo_act, 2)

    filas = [
        {
            "Concepto": f"SALDOS AL {fecha_ant}",
            "Aportes / Capital": round(aportes_ant, 2),
            "Excedentes Retenidos": round(retenidas_ant, 2),
            "Excedentes del Período": round(utilidad_periodo_ant, 2),
            "Totales": tot_ini,
        },
        {
            "Concepto": "Aumento / Movimientos en Capital y Aportes",
            "Aportes / Capital": var_aportes,
            "Excedentes Retenidos": 0.0,
            "Excedentes del Período": 0.0,
            "Totales": var_aportes,
        },
        {
            "Concepto": "Excedente / Utilidad Neta del Período",
            "Aportes / Capital": 0.0,
            "Excedentes Retenidos": 0.0,
            "Excedentes del Período": round(utilidad_periodo_act, 2),
            "Totales": round(utilidad_periodo_act, 2),
        },
        {
            "Concepto": "Ajustes / Traslado Excedentes Períodos Anteriores",
            "Aportes / Capital": 0.0,
            "Excedentes Retenidos": var_retenidas,
            "Excedentes del Período": round(-utilidad_periodo_ant, 2),
            "Totales": round(var_retenidas - utilidad_periodo_ant, 2),
        },
        {
            "Concepto": f"SALDOS AL {fecha_act}",
            "Aportes / Capital": round(aportes_act, 2),
            "Excedentes Retenidos": round(retenidas_act, 2),
            "Excedentes del Período": round(utilidad_periodo_act, 2),
            "Totales": tot_fin,
        },
    ]
    return pd.DataFrame(filas)


def construir_flujo_efectivo(conn, empresa_id, fecha_ant, fecha_act, utilidad_act, resumen_esf):
    """Construye Estado de Flujos de Efectivo (Método Indirecto) según machote."""
    cat = _cargar_catalogo(conn, empresa_id)
    if cat.empty:
        return pd.DataFrame()

    s_ant = _obtener_saldos_a_fecha(conn, empresa_id, fecha_ant)
    s_act = _obtener_saldos_a_fecha(conn, empresa_id, fecha_act)

    act_mov = cat[(cat["acepta_movimiento"] == 1) & (cat["tipo_up"].str.contains("ACTIVO") | cat["cod_str"].str.startswith("1"))]
    pas_mov = cat[(cat["acepta_movimiento"] == 1) & (cat["tipo_up"].str.contains("PASIVO") | cat["cod_str"].str.startswith("2"))]
    pat_mov = cat[(cat["acepta_movimiento"] == 1) & (cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL") | cat["cod_str"].str.startswith("3"))]

    # Variaciones de capital de trabajo operativo (excluyendo caja/bancos)
    var_act_oper = []
    for _, r in act_mov.iterrows():
        if any(w in r["nom_up"] for w in ["CAJA", "BANCO", "EFECTIVO"]):
            continue
        is_fijo = r["cod_str"].startswith(("1.2", "1.3")) or any(
            w in r["nom_up"]
            for w in ["FIJO", "TERRENO", "EDIFICIO", "TRANSPORTE", "MOBILIARIO", "EQUIPO", "HERRAMIENTA", "HIDRANTE", "LINEA", "DEPRECIACION", "OBRA"]
        )
        if not is_fijo:
            cid = int(r["id"])
            delta = round(-(s_act.get(cid, 0.0) - s_ant.get(cid, 0.0)), 2)
            if abs(delta) >= 0.005:
                var_act_oper.append((r["nombre"], delta))

    var_pas_oper = []
    for _, r in pas_mov.iterrows():
        is_lp = r["cod_str"].startswith("2.2") or "LARGO PLAZO" in r["nom_up"]
        if not is_lp:
            cid = int(r["id"])
            delta = round((-s_act.get(cid, 0.0)) - (-s_ant.get(cid, 0.0)), 2)
            if abs(delta) >= 0.005:
                var_pas_oper.append((r["nombre"], delta))

    # Inversiones (Activos fijos y depreciaciones)
    var_inv = []
    for _, r in act_mov.iterrows():
        is_fijo = r["cod_str"].startswith(("1.2", "1.3")) or any(
            w in r["nom_up"]
            for w in ["FIJO", "TERRENO", "EDIFICIO", "TRANSPORTE", "MOBILIARIO", "EQUIPO", "HERRAMIENTA", "HIDRANTE", "LINEA", "DEPRECIACION", "OBRA"]
        )
        if is_fijo:
            cid = int(r["id"])
            delta = round(-(s_act.get(cid, 0.0) - s_ant.get(cid, 0.0)), 2)
            if abs(delta) >= 0.005:
                var_inv.append((r["nombre"], delta))

    # Financiación (Pasivo LP y Patrimonio)
    var_fin = []
    for _, r in pas_mov.iterrows():
        if r["cod_str"].startswith("2.2") or "LARGO PLAZO" in r["nom_up"]:
            cid = int(r["id"])
            delta = round((-s_act.get(cid, 0.0)) - (-s_ant.get(cid, 0.0)), 2)
            if abs(delta) >= 0.005:
                var_fin.append((r["nombre"], delta))
    for _, r in pat_mov.iterrows():
        cid = int(r["id"])
        delta = round((-s_act.get(cid, 0.0)) - (-s_ant.get(cid, 0.0)), 2)
        if abs(delta) >= 0.005:
            var_fin.append((r["nombre"], delta))

    flujo_oper = round(utilidad_act + sum(v for _, v in var_act_oper) + sum(v for _, v in var_pas_oper), 2)
    flujo_inv = round(sum(v for _, v in var_inv), 2)
    flujo_fin = round(sum(v for _, v in var_fin), 2)

    ef_ini = resumen_esf.get("cajas_bancos_ant", 0.0)
    ef_fin = resumen_esf.get("cajas_bancos_act", 0.0)
    aumento_neto = round(ef_fin - ef_ini, 2)

    # Ajuste de cuadre operativo si hubo movimientos directos sin partida de resultados
    dif_cuadre = round(aumento_neto - (flujo_oper + flujo_inv + flujo_fin), 2)
    if abs(dif_cuadre) >= 0.01:
        flujo_oper = round(flujo_oper + dif_cuadre, 2)

    filas = [
        {"Actividad / Rubro": "ACTIVIDADES DE LA OPERACIÓN", "Monto (₡)": None, "TipoFila": "SECCION"},
        {"Actividad / Rubro": "  Utilidad / Excedente Neto del Período", "Monto (₡)": round(utilidad_act, 2), "TipoFila": "DETALLE"},
    ]
    for nom, val in var_act_oper + var_pas_oper:
        filas.append({"Actividad / Rubro": f"  Variación en {nom}", "Monto (₡)": val, "TipoFila": "DETALLE"})
    if abs(dif_cuadre) >= 0.01:
        filas.append({"Actividad / Rubro": "  Otras variaciones operativas netas", "Monto (₡)": dif_cuadre, "TipoFila": "DETALLE"})
    filas.append(
        {"Actividad / Rubro": "Efectivo neto provisto por las actividades de operación", "Monto (₡)": flujo_oper, "TipoFila": "SUBTOTAL"}
    )

    filas.append({"Actividad / Rubro": "ACTIVIDADES DE INVERSIÓN", "Monto (₡)": None, "TipoFila": "SECCION"})
    if not var_inv:
        filas.append({"Actividad / Rubro": "  Sin movimientos en propiedad, planta y equipo", "Monto (₡)": 0.0, "TipoFila": "DETALLE"})
    else:
        for nom, val in var_inv:
            filas.append({"Actividad / Rubro": f"  {nom}", "Monto (₡)": val, "TipoFila": "DETALLE"})
    filas.append(
        {"Actividad / Rubro": "Efectivo neto provisto/usado por actividades de inversión", "Monto (₡)": flujo_inv, "TipoFila": "SUBTOTAL"}
    )

    filas.append({"Actividad / Rubro": "ACTIVIDADES DE FINANCIACIÓN", "Monto (₡)": None, "TipoFila": "SECCION"})
    if not var_fin:
        filas.append({"Actividad / Rubro": "  Sin movimientos de financiación externa o aportes", "Monto (₡)": 0.0, "TipoFila": "DETALLE"})
    else:
        for nom, val in var_fin:
            filas.append({"Actividad / Rubro": f"  {nom}", "Monto (₡)": val, "TipoFila": "DETALLE"})
    filas.append(
        {"Actividad / Rubro": "Efectivo neto provisto por actividades de financiación", "Monto (₡)": flujo_fin, "TipoFila": "SUBTOTAL"}
    )

    filas.extend(
        [
            {"Actividad / Rubro": "Aumento (Disminución) neto de efectivo para el período", "Monto (₡)": aumento_neto, "TipoFila": "SUBTOTAL"},
            {"Actividad / Rubro": f"Dinero en efectivo al inicio del período ({fecha_ant})", "Monto (₡)": ef_ini, "TipoFila": "DETALLE"},
            {"Actividad / Rubro": f"EFECTIVO AL FINAL DEL PERÍODO ({fecha_act})", "Monto (₡)": ef_fin, "TipoFila": "TOTAL"},
        ]
    )
    return pd.DataFrame(filas)


def generar_texto_notas_nic(empresa_info, fecha_corte):
    """Genera el texto estructurado de Notas a los Estados Financieros según NIC/NIIF."""
    razon = empresa_info.get("razon_social", "LA ENTIDAD")
    cedula = empresa_info.get("cedula", "")
    actividad = empresa_info.get("actividad_economica") or "la actividad económica autorizada en su objeto social"
    return [
        (
            "1. Declaración de Normas de Contabilidad",
            f"Los estados financieros de {razon} (Cédula {cedula}) al {fecha_corte} se han elaborado de conformidad "
            "con las Normas Internacionales de Contabilidad (NIC) y las Normas Internacionales de Información Financiera "
            "para Pequeñas y Medianas Entidades (NIIF para las PYMES) vigentes en Costa Rica. Se han aplicado los criterios "
            "relativos al registro, valuación, presentación y revelación de políticas contables exigidos por dichas normas.",
        ),
        (
            "2. Domicilio, Actividad y Moneda Funcional",
            f"La entidad opera en la República de Costa Rica, dedicándose principalmente a {actividad}. "
            "Todos sus registros contables y estados financieros se expresan en colones costarricenses (₡), "
            "moneda oficial de Costa Rica y moneda funcional de la entidad. Las transacciones en moneda extranjera "
            "se convierten utilizando el tipo de cambio de referencia del Banco Central de Costa Rica (BCCR).",
        ),
        (
            "3. Políticas Contables Significativas",
            "a) Caja y Bancos: Incluye el efectivo en caja y los depósitos a la vista en entidades del Sistema Bancario Nacional.\n"
            "b) Cuentas por Cobrar: Se registran al valor nominal de las facturas emitidas y se evalúa su recuperabilidad al cierre.\n"
            "c) Inventarios: Se valúan al costo de adquisición o valor neto realizable, el que sea menor.\n"
            "d) Propiedad, Planta y Equipo: Se registran al costo histórico menos la depreciación acumulada calculada por el método de línea recta.\n"
            "e) Cuentas por Pagar y Obligaciones Tributarias: Reconocen pasivos exigibles a corto plazo, incluyendo IVA por pagar y cargas sociales (CCSS).\n"
            "f) Reconocimiento de Ingresos y Gastos: Se contabilizan por el principio de devengo (acumulación) conforme se prestan los servicios o se incurren las erogaciones.",
        ),
        (
            "4. Cambios en Políticas Contables y Eventos Posteriores",
            "No se han realizado cambios significativos en las políticas contables durante el período reportado, "
            "ni se han producido eventos posteriores al cierre que requieran ajustes materiales en los estados financieros.",
        ),
    ]


# ==============================================================================
# 2. EXPORTADORES PROFESIONALES (EXCEL, WORD, PDF) CON FIRMAS A IZQ. Y DER.
# ==============================================================================
def generar_excel_estados_financieros(
    empresa_info,
    fecha_ant,
    fecha_act,
    df_esf,
    df_er,
    df_pat,
    df_fe,
    notas_lista,
    contador_nombre,
    contador_carne,
    representante_nombre,
    representante_cargo,
):
    """Genera cuaderno Excel profesional con los 5 reportes y bloque de firmas."""
    output = io.BytesIO()
    razon = empresa_info.get("razon_social", "Empresa")
    cedula = empresa_info.get("cedula", "")

    def _escribir_firmas_excel(ws, fila_inicio):
        r_linea = fila_inicio + 2
        ws.cell(row=r_linea, column=1, value="________________________________________")
        ws.cell(row=r_linea, column=3, value="________________________________________")
        ws.cell(row=r_linea + 1, column=1, value=contador_nombre)
        ws.cell(row=r_linea + 1, column=3, value=representante_nombre)
        ws.cell(row=r_linea + 2, column=1, value=f"Contador Privado Carné {contador_carne}")
        ws.cell(row=r_linea + 2, column=3, value=representante_cargo)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. Estado de Situación Financiera
        df_esf_exp = df_esf[["Rubro", "Nota", "Anterior", "Actual"]].rename(
            columns={"Anterior": str(fecha_ant), "Actual": str(fecha_act)}
        )
        df_esf_exp.to_excel(writer, sheet_name="Situacion_Financiera", index=False, startrow=4)
        ws1 = writer.sheets["Situacion_Financiera"]
        ws1["A1"] = "ESTADO DE LA SITUACIÓN FINANCIERA COMPARATIVO"
        ws1["A2"] = f"{razon} - Cédula: {cedula}"
        ws1["A3"] = "(Expresado en colones costarricenses ₡)"
        _escribir_firmas_excel(ws1, len(df_esf_exp) + 6)

        # 2. Estado de Resultados
        df_er_exp = df_er[["Rubro", "Nota", "Anterior", "Actual"]].rename(
            columns={"Anterior": f"Corte {fecha_ant}", "Actual": f"Período a {fecha_act}"}
        )
        df_er_exp.to_excel(writer, sheet_name="Estado_Resultados", index=False, startrow=4)
        ws2 = writer.sheets["Estado_Resultados"]
        ws2["A1"] = "ESTADO DE RESULTADOS COMPARATIVO"
        ws2["A2"] = f"{razon} - Cédula: {cedula}"
        ws2["A3"] = "(Expresado en colones costarricenses ₡)"
        _escribir_firmas_excel(ws2, len(df_er_exp) + 6)

        # 3. Cambios en el Patrimonio
        df_pat.to_excel(writer, sheet_name="Cambios_Patrimonio", index=False, startrow=4)
        ws3 = writer.sheets["Cambios_Patrimonio"]
        ws3["A1"] = "ESTADO DE CAMBIOS EN EL PATRIMONIO"
        ws3["A2"] = f"{razon} - Cédula: {cedula}"
        ws3["A3"] = "(Expresado en colones costarricenses ₡)"
        _escribir_firmas_excel(ws3, len(df_pat) + 6)

        # 4. Flujo de Efectivo
        df_fe_exp = df_fe[["Actividad / Rubro", "Monto (₡)"]]
        df_fe_exp.to_excel(writer, sheet_name="Flujo_Efectivo", index=False, startrow=4)
        ws4 = writer.sheets["Flujo_Efectivo"]
        ws4["A1"] = "ESTADO DE FLUJOS DE EFECTIVO"
        ws4["A2"] = f"{razon} - Cédula: {cedula}"
        ws4["A3"] = f"Corte al {fecha_act} (Expresado en colones costarricenses ₡)"
        _escribir_firmas_excel(ws4, len(df_fe_exp) + 6)

        # 5. Notas NIC/NIIF
        df_notas = pd.DataFrame(notas_lista, columns=["Nota", "Detalle"])
        df_notas.to_excel(writer, sheet_name="Notas_NIC_NIIF", index=False, startrow=4)
        ws5 = writer.sheets["Notas_NIC_NIIF"]
        ws5["A1"] = f"NOTAS A LOS ESTADOS FINANCIEROS AL {fecha_act}"
        ws5["A2"] = f"{razon} - Cédula: {cedula}"
        ws5["A3"] = "Normas Internacionales de Contabilidad (NIC / NIIF)"
        _escribir_firmas_excel(ws5, len(df_notas) + 6)

    return output.getvalue()


def _agregar_firmas_word(doc, contador_nombre, contador_carne, representante_nombre, representante_cargo):
    """Agrega tabla sin bordes con firma del Contador a la izquierda y Representante/Dueño a la derecha."""
    doc.add_paragraph("\n")
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(3.2)
    tbl.columns[1].width = Inches(3.2)

    c_izq = tbl.rows[0].cells[0]
    p_izq = c_izq.paragraphs[0]
    p_izq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_izq.add_run(
        f"_____________________________________\n"
        f"{contador_nombre}\n"
        f"Contador Privado Carné {contador_carne}"
    )
    r1.bold = True
    r1.font.size = Pt(9.5)

    c_der = tbl.rows[0].cells[1]
    p_der = c_der.paragraphs[0]
    p_der.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p_der.add_run(
        f"_____________________________________\n"
        f"{representante_nombre}\n"
        f"{representante_cargo}"
    )
    r2.bold = True
    r2.font.size = Pt(9.5)


def generar_word_estados_financieros(
    empresa_info,
    fecha_ant,
    fecha_act,
    df_esf,
    df_er,
    df_pat,
    df_fe,
    notas_lista,
    contador_nombre,
    contador_carne,
    representante_nombre,
    representante_cargo,
):
    """Genera paquete completo de Estados Financieros y Notas en Word (.docx) según machotes."""
    doc = Document()
    razon = empresa_info.get("razon_social", "EMPRESA").upper()
    cedula = empresa_info.get("cedula", "")

    def _encabezado_pagina(titulo_estado, subtitulo_extra=""):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rt = p.add_run(f"{titulo_estado}\n")
        rt.bold = True
        rt.font.size = Pt(13)
        rt.font.color.rgb = RGBColor(27, 54, 93)

        rs = p.add_run(
            f"{razon}\nCÉDULA JURÍDICA / FÍSICA: {cedula}\n"
            f"{subtitulo_extra}(Expresado en colones costarricenses)"
        )
        rs.font.size = Pt(9.5)
        rs.font.color.rgb = RGBColor(70, 70, 70)

    def _render_tabla_comparativa_word(df_data, col_ant_lbl, col_act_lbl):
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = tbl.rows[0].cells
        hdr[0].text = "Detalle / Cuenta"
        hdr[1].text = "N"
        hdr[2].text = col_ant_lbl
        hdr[3].text = col_act_lbl
        for c in hdr:
            c.paragraphs[0].runs[0].bold = True

        for _, r in df_data.iterrows():
            row = tbl.add_row().cells
            tf = r.get("TipoFila", "DETALLE")
            row[0].text = str(r["Rubro"])
            row[1].text = str(r.get("Nota", ""))
            row[2].text = _fmt_crc(r["Anterior"]) if pd.notna(r["Anterior"]) else ""
            row[3].text = _fmt_crc(r["Actual"]) if pd.notna(r["Actual"]) else ""
            row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if tf in ("SECCION", "SUBSECCION", "SUBTOTAL", "TOTAL"):
                for c in row:
                    if c.paragraphs[0].runs:
                        c.paragraphs[0].runs[0].bold = True

    # 1. Estado de la Situación Financiera
    _encabezado_pagina("ESTADO DE LA SITUACIÓN FINANCIERA", f"COMPARATIVO AL {fecha_ant} Y {fecha_act}\n")
    _render_tabla_comparativa_word(df_esf, str(fecha_ant), str(fecha_act))
    _agregar_firmas_word(doc, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # 2. Estado de Resultados
    doc.add_page_break()
    _encabezado_pagina("ESTADO DE RESULTADOS", f"COMPARATIVO {fecha_ant} - {fecha_act}\n")
    _render_tabla_comparativa_word(df_er, str(fecha_ant), str(fecha_act))
    _agregar_firmas_word(doc, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # 3. Estado de Cambios en el Patrimonio
    doc.add_page_break()
    _encabezado_pagina("ESTADO DE CAMBIOS EN EL PATRIMONIO", f"PERÍODO {fecha_ant} AL {fecha_act}\n")
    tbl_p = doc.add_table(rows=1, cols=5)
    tbl_p.style = "Table Grid"
    cols_p = list(df_pat.columns)
    for i, c_name in enumerate(cols_p):
        tbl_p.rows[0].cells[i].text = c_name
        tbl_p.rows[0].cells[i].paragraphs[0].runs[0].bold = True
    for idx_r, r in df_pat.iterrows():
        row = tbl_p.add_row().cells
        for i, c_name in enumerate(cols_p):
            val = r[c_name]
            if i == 0:
                row[i].text = str(val)
            else:
                row[i].text = _fmt_crc(val)
                row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if idx_r in (0, len(df_pat) - 1) and row[i].paragraphs[0].runs:
                row[i].paragraphs[0].runs[0].bold = True
    _agregar_firmas_word(doc, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # 4. Estado de Flujo de Efectivo
    doc.add_page_break()
    _encabezado_pagina("ESTADO DE FLUJOS DE EFECTIVO", f"AL {fecha_act}\n")
    tbl_f = doc.add_table(rows=1, cols=2)
    tbl_f.style = "Table Grid"
    tbl_f.rows[0].cells[0].text = "Actividades"
    tbl_f.rows[0].cells[1].text = f"Al {fecha_act} (₡)"
    tbl_f.rows[0].cells[0].paragraphs[0].runs[0].bold = True
    tbl_f.rows[0].cells[1].paragraphs[0].runs[0].bold = True
    for _, r in df_fe.iterrows():
        row = tbl_f.add_row().cells
        row[0].text = str(r["Actividad / Rubro"])
        row[1].text = _fmt_crc(r["Monto (₡)"]) if pd.notna(r["Monto (₡)"]) else ""
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if r.get("TipoFila") in ("SECCION", "SUBTOTAL", "TOTAL"):
            for c in row:
                if c.paragraphs[0].runs:
                    c.paragraphs[0].runs[0].bold = True
    _agregar_firmas_word(doc, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # 5. Notas a los Estados Financieros
    doc.add_page_break()
    _encabezado_pagina(f"NOTAS A LOS ESTADOS FINANCIEROS AL {fecha_act}", "SEGÚN NORMAS INTERNACIONALES (NIC / NIIF)\n")
    for tit_n, cuerpo_n in notas_lista:
        h = doc.add_heading(tit_n, level=2)
        h.runs[0].font.color.rgb = RGBColor(27, 54, 93)
        p_n = doc.add_paragraph(cuerpo_n)
        p_n.paragraph_format.space_after = Pt(6)
    _agregar_firmas_word(doc, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _dibujar_firmas_pdf(pdf, contador_nombre, contador_carne, representante_nombre, representante_cargo):
    """Dibuja firma del Contador a la izquierda y Representante/Dueño a la derecha en PDF."""
    if pdf.get_y() > 232:
        pdf.add_page()
    pdf.ln(12)
    y_firma = pdf.get_y()
    pdf.set_draw_color(60, 60, 60)
    pdf.line(18, y_firma, 94, y_firma)
    pdf.line(118, y_firma, 194, y_firma)

    pdf.set_xy(14, y_firma + 1.5)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(84, 4.5, _limpiar_pdf(contador_nombre), align="C")

    pdf.set_xy(114, y_firma + 1.5)
    pdf.cell(84, 4.5, _limpiar_pdf(representante_nombre), align="C", ln=True)

    pdf.set_x(14)
    pdf.set_font("Helvetica", "", 8)
    pdf.cell(84, 4.5, _limpiar_pdf(f"Contador Privado Carné {contador_carne}"), align="C")

    pdf.set_x(114)
    pdf.cell(84, 4.5, _limpiar_pdf(representante_cargo), align="C", ln=True)


def generar_pdf_estados_financieros(
    empresa_info,
    fecha_ant,
    fecha_act,
    df_esf,
    df_er,
    df_pat,
    df_fe,
    notas_lista,
    contador_nombre,
    contador_carne,
    representante_nombre,
    representante_cargo,
):
    """Genera paquete PDF profesional con los 5 reportes según machotes y firmas izquierda/derecha."""
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=16)

    razon = empresa_info.get("razon_social", "EMPRESA").upper()
    cedula = empresa_info.get("cedula", "")

    def _cabecera_pdf(titulo, subtitulo):
        pdf.add_page()
        pdf.set_fill_color(27, 54, 93)
        pdf.rect(10, 10, 196, 22, "F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 11.5)
        pdf.set_xy(12, 12)
        pdf.cell(192, 6, _limpiar_pdf(titulo), align="C", ln=True)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_x(12)
        pdf.cell(192, 5, _limpiar_pdf(f"{razon} | CÉDULA: {cedula}"), align="C", ln=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_x(12)
        pdf.cell(192, 5, _limpiar_pdf(f"{subtitulo} (Expresado en colones costarricenses)"), align="C", ln=True)
        pdf.ln(5)

    def _tabla_comp_pdf(df_data):
        pdf.set_fill_color(230, 236, 245)
        pdf.set_text_color(20, 35, 60)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(108, 6.5, "Detalle / Cuenta", border=1, fill=True)
        pdf.cell(12, 6.5, "N", border=1, fill=True, align="C")
        pdf.cell(38, 6.5, _limpiar_pdf(str(fecha_ant)), border=1, fill=True, align="R")
        pdf.cell(38, 6.5, _limpiar_pdf(str(fecha_act)), border=1, fill=True, align="R", ln=True)

        for _, r in df_data.iterrows():
            tf = r.get("TipoFila", "DETALLE")
            rubro = _limpiar_pdf(r["Rubro"])[:62]
            nota = _limpiar_pdf(r.get("Nota", ""))
            v_ant = _fmt_crc(r["Anterior"]) if pd.notna(r["Anterior"]) else ""
            v_act = _fmt_crc(r["Actual"]) if pd.notna(r["Actual"]) else ""

            if tf in ("SECCION", "SUBSECCION"):
                pdf.set_fill_color(242, 245, 250)
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(27, 54, 93)
                pdf.cell(196, 5.8, rubro, border=1, fill=True, ln=True)
            elif tf in ("SUBTOTAL", "TOTAL"):
                pdf.set_fill_color(235, 240, 248) if tf == "TOTAL" else pdf.set_fill_color(248, 250, 252)
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(15, 23, 42)
                pdf.cell(108, 6, rubro, border=1, fill=True)
                pdf.cell(12, 6, nota, border=1, fill=True, align="C")
                pdf.cell(38, 6, v_ant, border=1, fill=True, align="R")
                pdf.cell(38, 6, v_act, border=1, fill=True, align="R", ln=True)
            else:
                pdf.set_font("Helvetica", "", 7.8)
                pdf.set_text_color(30, 30, 30)
                pdf.cell(108, 5.4, rubro, border=1)
                pdf.cell(12, 5.4, nota, border=1, align="C")
                pdf.cell(38, 5.4, v_ant, border=1, align="R")
                pdf.cell(38, 5.4, v_act, border=1, align="R", ln=True)

    # Página 1: Estado de Situación Financiera
    _cabecera_pdf("ESTADO DE LA SITUACIÓN FINANCIERA", f"COMPARATIVO {fecha_ant} - {fecha_act}")
    _tabla_comp_pdf(df_esf)
    _dibujar_firmas_pdf(pdf, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # Página 2: Estado de Resultados
    _cabecera_pdf("ESTADO DE RESULTADOS", f"COMPARATIVO {fecha_ant} - {fecha_act}")
    _tabla_comp_pdf(df_er)
    _dibujar_firmas_pdf(pdf, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # Página 3: Cambios en el Patrimonio
    _cabecera_pdf("ESTADO DE CAMBIOS EN EL PATRIMONIO", f"PERÍODO {fecha_ant} AL {fecha_act}")
    pdf.set_fill_color(230, 236, 245)
    pdf.set_text_color(20, 35, 60)
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.cell(68, 6.5, "Concepto", border=1, fill=True)
    pdf.cell(32, 6.5, "Aportes / Capital", border=1, fill=True, align="R")
    pdf.cell(32, 6.5, "Exc. Retenidos", border=1, fill=True, align="R")
    pdf.cell(32, 6.5, "Exc. Período", border=1, fill=True, align="R")
    pdf.cell(32, 6.5, "Totales", border=1, fill=True, align="R", ln=True)

    for idx_r, r in df_pat.iterrows():
        is_bold = idx_r in (0, len(df_pat) - 1)
        pdf.set_font("Helvetica", "B" if is_bold else "", 7.5)
        pdf.set_text_color(20, 20, 20)
        pdf.cell(68, 6, _limpiar_pdf(r["Concepto"])[:40], border=1)
        pdf.cell(32, 6, _fmt_crc(r["Aportes / Capital"]), border=1, align="R")
        pdf.cell(32, 6, _fmt_crc(r["Excedentes Retenidos"]), border=1, align="R")
        pdf.cell(32, 6, _fmt_crc(r["Excedentes del Período"]), border=1, align="R")
        pdf.cell(32, 6, _fmt_crc(r["Totales"]), border=1, align="R", ln=True)
    _dibujar_firmas_pdf(pdf, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # Página 4: Estado de Flujos de Efectivo
    _cabecera_pdf("ESTADO DE FLUJOS DE EFECTIVO", f"AL {fecha_act}")
    pdf.set_fill_color(230, 236, 245)
    pdf.set_text_color(20, 35, 60)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(146, 6.5, "Actividades", border=1, fill=True)
    pdf.cell(50, 6.5, _limpiar_pdf(f"Al {fecha_act} (CRC)"), border=1, fill=True, align="R", ln=True)

    for _, r in df_fe.iterrows():
        tf = r.get("TipoFila", "DETALLE")
        act_txt = _limpiar_pdf(r["Actividad / Rubro"])[:78]
        m_txt = _fmt_crc(r["Monto (₡)"]) if pd.notna(r["Monto (₡)"]) else ""
        if tf == "SECCION":
            pdf.set_fill_color(242, 245, 250)
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_text_color(27, 54, 93)
            pdf.cell(196, 6, act_txt, border=1, fill=True, ln=True)
        elif tf in ("SUBTOTAL", "TOTAL"):
            pdf.set_fill_color(238, 243, 250)
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_text_color(15, 23, 42)
            pdf.cell(146, 6, act_txt, border=1, fill=True)
            pdf.cell(50, 6, m_txt, border=1, fill=True, align="R", ln=True)
        else:
            pdf.set_font("Helvetica", "", 7.8)
            pdf.set_text_color(30, 30, 30)
            pdf.cell(146, 5.5, act_txt, border=1)
            pdf.cell(50, 5.5, m_txt, border=1, align="R", ln=True)
    _dibujar_firmas_pdf(pdf, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    # Página 5: Notas a los Estados Financieros según NIC/NIIF
    _cabecera_pdf(f"NOTAS A LOS ESTADOS FINANCIEROS AL {fecha_act}", "SEGÚN NIC / NIIF")
    for tit_n, cuerpo_n in notas_lista:
        pdf.set_text_color(27, 54, 93)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(0, 6, _limpiar_pdf(tit_n), ln=True)
        pdf.set_text_color(30, 30, 30)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.multi_cell(0, 4.8, _limpiar_pdf(cuerpo_n))
        pdf.ln(2.5)
    _dibujar_firmas_pdf(pdf, contador_nombre, contador_carne, representante_nombre, representante_cargo)

    raw = pdf.output()
    return bytes(raw) if isinstance(raw, (bytes, bytearray)) else raw.encode("latin-1")


# ==============================================================================
# 3. BALANZA DE COMPROBACIÓN Y LIBRO MAYOR
# ==============================================================================
def construir_balanza_comprobacion(conn, empresa_id, fecha_desde, fecha_hasta):
    cat = _cargar_catalogo(conn, empresa_id)
    if cat.empty:
        return pd.DataFrame()

    dia_ant = fecha_desde - datetime.timedelta(days=1)
    s_ini = _obtener_saldos_a_fecha(conn, empresa_id, dia_ant)
    movs = _obtener_movimientos_periodo(conn, empresa_id, fecha_desde, fecha_hasta)

    filas = []
    for _, r in cat[cat["acepta_movimiento"] == 1].iterrows():
        cid = int(r["id"])
        si = round(s_ini.get(cid, 0.0), 2)
        d = round(movs.get(cid, {}).get("debe", 0.0), 2)
        h = round(movs.get(cid, {}).get("haber", 0.0), 2)
        sf = round(si + d - h, 2)
        if abs(si) >= 0.005 or abs(d) >= 0.005 or abs(h) >= 0.005 or abs(sf) >= 0.005:
            filas.append(
                {
                    "Código": r["cod_str"],
                    "Cuenta": r["nombre"],
                    "Tipo": r["tipo"],
                    "Saldo Inicial (₡)": si,
                    "Debe Período (₡)": d,
                    "Haber Período (₡)": h,
                    "Saldo Final (₡)": sf,
                }
            )
    return pd.DataFrame(filas)


# ==============================================================================
# 4. INTERFAZ STREAMLIT DEL MÓDULO DE REPORTES FINANCIEROS NIIF
# ==============================================================================
def renderizar_modulo_reportes(conn, empresa_id, empresa_info):
    st.subheader("📑 Estados Financieros NIIF (Formato Oficial con Firmas)")

    hoy = datetime.date.today()
    c1, c2, c3 = st.columns(3)
    fecha_ant = c1.date_input("Fecha Corte Comparativo Anterior", value=datetime.date(hoy.year - 1, 12, 31), key="rep_f_ant")
    fecha_act = c2.date_input("Fecha Corte Actual", value=hoy, key="rep_f_act")
    fecha_er_ini = c3.date_input("Inicio Período Resultados", value=datetime.date(hoy.year, 1, 1), key="rep_er_ini")

    # Datos de firma (Izquierda: Contador / Derecha: Representante Legal o Dueño)
    with st.expander("✍️ Configurar Firmas al Pie de los Estados Financieros", expanded=False):
        f_c1, f_c2, f_c3, f_c4 = st.columns(4)
        contador_nombre = f_c1.text_input(
            "Firma Izquierda (Contador):",
            value="Lic. Roy F. Esquivel González",
            key="rep_cont_nom",
        )
        contador_carne = f_c2.text_input(
            "Carné Contador:",
            value="8788",
            key="rep_cont_carne",
        )
        rep_default = str(empresa_info.get("representante_legal") or "").strip() or "Representante Legal / Propietario"
        representante_nombre = f_c3.text_input(
            "Firma Derecha (Representante / Dueño):",
            value=rep_default,
            key="rep_rep_nom",
        )
        representante_cargo = f_c4.text_input(
            "Cargo Firma Derecha:",
            value="Representante Legal / Presidente",
            key="rep_rep_cargo",
        )

    df_esf, res_esf = construir_estado_situacion_comparativo(conn, empresa_id, fecha_ant, fecha_act)
    if df_esf.empty:
        st.info("No hay cuentas ni movimientos registrados para esta empresa.")
        return

    fecha_er_ant_ini = datetime.date(fecha_ant.year, 1, 1)
    df_er, res_er = construir_estado_resultados(
        conn, empresa_id, fecha_er_ini, fecha_act, fecha_ant_desde=fecha_er_ant_ini, fecha_ant_hasta=fecha_ant
    )
    df_pat = construir_cambios_patrimonio(
        conn, empresa_id, fecha_ant, fecha_act, res_er["utilidad"], res_er["utilidad_ant"]
    )
    df_fe = construir_flujo_efectivo(conn, empresa_id, fecha_ant, fecha_act, res_er["utilidad"], res_esf)
    notas_lista = generar_texto_notas_nic(empresa_info, fecha_act)

    # Descarga bajo demanda (no genera archivos pesados hasta que se presiona el botón)
    st.markdown("### 📥 Descarga del Juego Completo de Estados Financieros (PDF, Word y Excel)")
    col_gen, col_d1, col_d2, col_d3 = st.columns([1.3, 1, 1, 1])

    firma_params = (
        int(empresa_id),
        str(fecha_ant),
        str(fecha_act),
        str(fecha_er_ini),
        contador_nombre,
        contador_carne,
        representante_nombre,
        representante_cargo,
    )

    if col_gen.button("⚙️ Preparar Archivos (PDF/Word/Excel)", type="primary", use_container_width=True):
        st.session_state["paquete_eeff_key"] = firma_params
        st.session_state["paquete_eeff_bytes"] = {
            "pdf": generar_pdf_estados_financieros(
                empresa_info, fecha_ant, fecha_act, df_esf, df_er, df_pat, df_fe, notas_lista,
                contador_nombre, contador_carne, representante_nombre, representante_cargo
            ),
            "docx": generar_word_estados_financieros(
                empresa_info, fecha_ant, fecha_act, df_esf, df_er, df_pat, df_fe, notas_lista,
                contador_nombre, contador_carne, representante_nombre, representante_cargo
            ),
            "xlsx": generar_excel_estados_financieros(
                empresa_info, fecha_ant, fecha_act, df_esf, df_er, df_pat, df_fe, notas_lista,
                contador_nombre, contador_carne, representante_nombre, representante_cargo
            ),
        }

    archivos_listos = (
        st.session_state.get("paquete_eeff_key") == firma_params
        and "paquete_eeff_bytes" in st.session_state
    )
    ced_file = str(empresa_info.get("cedula", "empresa")).replace(" ", "_")

    if archivos_listos:
        b_pack = st.session_state["paquete_eeff_bytes"]
        with col_d1:
            st.download_button(
                "📕 Descargar PDF",
                data=b_pack["pdf"],
                file_name=f"Estados_Financieros_{ced_file}_{fecha_act}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with col_d2:
            st.download_button(
                "📘 Descargar Word",
                data=b_pack["docx"],
                file_name=f"Estados_Financieros_{ced_file}_{fecha_act}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with col_d3:
            st.download_button(
                "📗 Descargar Excel",
                data=b_pack["xlsx"],
                file_name=f"Estados_Financieros_{ced_file}_{fecha_act}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        col_d1.caption("Presione **Preparar Archivos** para habilitar las descargas.")

    st.divider()

    tabs = st.tabs(
        [
            "🏛️ Situación Financiera",
            "📈 Estado de Resultados",
            "🔄 Cambios en Patrimonio",
            "💵 Flujo de Efectivo",
            "📝 Notas NIC/NIIF",
            "⚖️ Balanza y Mayor",
        ]
    )

    with tabs[0]:
        st.markdown(f"**ESTADO DE LA SITUACIÓN FINANCIERA COMPARATIVO ({fecha_ant} vs {fecha_act})**")
        df_show = df_esf[["Rubro", "Nota", "Anterior", "Actual"]].rename(
            columns={"Anterior": str(fecha_ant), "Actual": str(fecha_act)}
        )
        st.dataframe(df_show, use_container_width=True, hide_index=True)

    with tabs[1]:
        st.markdown(f"**ESTADO DE RESULTADOS COMPARATIVO ({fecha_ant} vs {fecha_act})**")
        df_er_show = df_er[["Rubro", "Nota", "Anterior", "Actual"]].rename(
            columns={"Anterior": str(fecha_ant), "Actual": str(fecha_act)}
        )
        st.dataframe(df_er_show, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.markdown(f"**ESTADO DE CAMBIOS EN EL PATRIMONIO ({fecha_ant} AL {fecha_act})**")
        st.dataframe(df_pat, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.markdown(f"**ESTADO DE FLUJOS DE EFECTIVO AL {fecha_act}**")
        st.dataframe(df_fe[["Actividad / Rubro", "Monto (₡)"]], use_container_width=True, hide_index=True)

    with tabs[4]:
        st.markdown(f"**NOTAS A LOS ESTADOS FINANCIEROS AL {fecha_act} (SEGÚN NIC / NIIF)**")
        for tit_n, cuerpo_n in notas_lista:
            st.markdown(f"**{tit_n}**")
            st.write(cuerpo_n)

    with tabs[5]:
        st.markdown("**BALANZA DE COMPROBACIÓN**")
        df_bal = construir_balanza_comprobacion(conn, empresa_id, fecha_er_ini, fecha_act)
        if df_bal.empty:
            st.info("Sin saldos en el rango seleccionado.")
        else:
            st.dataframe(df_bal, use_container_width=True, hide_index=True)
