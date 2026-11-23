import calendar
import datetime
import io
import json
import pandas as pd
import streamlit as st

from sync_sheets import subir_bd_a_sheets


def _asegurar_tabla_iva(conn):
    conn.execute(
        """
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
        """
    )
    conn.commit()


def _calcular_movimientos_iva(conn, empresa_id, fecha_desde, fecha_hasta):
    """Obtiene movimientos del período agrupados por cuenta contable."""
    df = pd.read_sql_query(
        """
        SELECT c.id AS cuenta_id,
               c.codigo,
               c.nombre,
               c.tipo,
               COALESCE(SUM(ad.debe), 0.0) AS debe,
               COALESCE(SUM(ad.haber), 0.0) AS haber
        FROM asientos_detalle ad
        JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
        JOIN catalogo_cuentas c ON ad.cuenta_id = c.id
        WHERE ac.empresa_id = ? AND ac.fecha BETWEEN ? AND ?
        GROUP BY c.id, c.codigo, c.nombre, c.tipo
        ORDER BY c.codigo
        """,
        conn,
        params=(int(empresa_id), str(fecha_desde), str(fecha_hasta)),
    )
    return df


def _exportar_iva_excel(resumen_dict, df_movs, empresa_info, periodo_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_res = pd.DataFrame(
            [{"Concepto": k, "Monto (₡)": v} for k, v in resumen_dict.items()]
        )
        df_res.to_excel(writer, sheet_name="Liquidacion_IVA", index=False, startrow=4)
        if df_movs is not None and not df_movs.empty:
            df_movs.to_excel(writer, sheet_name="Detalle_Cuentas", index=False)

        ws = writer.sheets["Liquidacion_IVA"]
        ws["A1"] = empresa_info.get("razon_social", "")
        ws["A2"] = f"Cédula: {empresa_info.get('cedula', '')}"
        ws["A3"] = f"Liquidación Mensual IVA (D-104) - Período: {periodo_str}"
    return output.getvalue()


def renderizar_modulo_iva(conn, empresa_id, empresa_info):
    _asegurar_tabla_iva(conn)
    st.subheader("🧾 Liquidación Mensual del Impuesto al Valor Agregado (IVA / D-104)")

    hoy = datetime.date.today()
    c1, c2, c3 = st.columns([2, 2, 3])
    anio = int(c1.number_input("Año Fiscal", min_value=2020, max_value=2035, value=hoy.year, key="iva_anio"))
    mes = int(c2.number_input("Mes", min_value=1, max_value=12, value=hoy.month, key="iva_mes"))

    ultimo_dia = calendar.monthrange(anio, mes)[1]
    fecha_desde = datetime.date(anio, mes, 1)
    fecha_hasta = datetime.date(anio, mes, ultimo_dia)
    periodo_str = f"{anio}-{mes:02d}"

    c3.info(f"Período analizado: **{fecha_desde}** al **{fecha_hasta}**")

    # Buscar declaración previa guardada para este período
    prev_df = pd.read_sql_query(
        """
        SELECT * FROM declaraciones_iva
        WHERE empresa_id = ? AND anio = ? AND mes = ?
        ORDER BY id DESC LIMIT 1
        """,
        conn,
        params=(int(empresa_id), anio, mes),
    )
    prev = prev_df.iloc[0].to_dict() if not prev_df.empty else {}

    # Detección automática preliminar desde movimientos contables del mes
    df_movs = _calcular_movimientos_iva(conn, empresa_id, fecha_desde, fecha_hasta)

    def _sumar_por_patron(df_sub, patron_cod, patron_nom, signo="haber_menos_debe"):
        if df_sub.empty:
            return 0.0
        mask = df_sub["codigo"].astype(str).str.startswith(patron_cod) | (
            df_sub["nombre"].astype(str).str.upper().str.contains(patron_nom, regex=True)
        )
        sub = df_sub[mask]
        if sub.empty:
            return 0.0
        if signo == "haber_menos_debe":
            return round(float((sub["haber"] - sub["debe"]).sum()), 2)
        return round(float((sub["debe"] - sub["haber"]).sum()), 2)

    ventas_auto = max(0.0, _sumar_por_patron(df_movs, "4", r"VENTA|INGRESO.*GRAV", "haber_menos_debe"))
    iva_deb_auto = max(0.0, _sumar_por_patron(df_movs, "2.1.04", r"IVA.*DEBIT|IVA.*VENTA|IMPUESTO.*VENTA", "haber_menos_debe"))
    compras_auto = max(0.0, _sumar_por_patron(df_movs, "5", r"COMPRA|COSTO|GASTO.*GRAV", "debe_menos_haber"))
    iva_cred_auto = max(0.0, _sumar_por_patron(df_movs, "1.1.06", r"IVA.*CREDIT|IVA.*COMPRA|CRÉDITO.*IVA", "debe_menos_haber"))

    st.markdown("### 1. Débito Fiscal (Ventas y Operaciones Gravadas)")
    d1, d2, d3 = st.columns(3)
    ventas_grav = d1.number_input(
        "Ventas Netas Gravadas (₡)",
        min_value=0.0,
        value=float(prev.get("ventas_netas_gravadas", ventas_auto) or 0.0),
        step=1000.0,
        format="%.2f",
    )
    ventas_exentas = d2.number_input(
        "Ventas Exentas / No Sujetas (₡)",
        min_value=0.0,
        value=float(prev.get("ventas_exentas", 0.0) or 0.0),
        step=1000.0,
        format="%.2f",
    )
    iva_debito = d3.number_input(
        "IVA Débito Fiscal Generado (₡)",
        min_value=0.0,
        value=float(prev.get("iva_debito", iva_deb_auto) or 0.0),
        step=1000.0,
        format="%.2f",
    )

    st.markdown("### 2. Crédito Fiscal, Saldo a Favor y Retenciones")
    c_1, c_2, c_3 = st.columns(3)
    compras_grav = c_1.number_input(
        "Compras y Gastos Gravados (₡)",
        min_value=0.0,
        value=float(prev.get("compras_netas_gravadas", compras_auto) or 0.0),
        step=1000.0,
        format="%.2f",
    )
    compras_exentas = c_2.number_input(
        "Compras Exentas / No Sujetas (₡)",
        min_value=0.0,
        value=float(prev.get("compras_exentas", 0.0) or 0.0),
        step=1000.0,
        format="%.2f",
    )
    iva_credito = c_3.number_input(
        "IVA Crédito Fiscal Acreditable (₡)",
        min_value=0.0,
        value=float(prev.get("iva_credito", iva_cred_auto) or 0.0),
        step=1000.0,
        format="%.2f",
    )

    r1, r2, r3 = st.columns(3)
    saldo_favor_ant = r1.number_input(
        "Saldo a Favor Período Anterior (₡)",
        min_value=0.0,
        value=float(prev.get("saldo_favor_anterior", 0.0) or 0.0),
        step=1000.0,
        format="%.2f",
    )
    retenciones_iva = r2.number_input(
        "Retenciones / Percepciones IVA (₡)",
        min_value=0.0,
        value=float(prev.get("retenciones_iva", 0.0) or 0.0),
        step=1000.0,
        format="%.2f",
    )
    obs_iva = r3.text_input(
        "Observaciones / Notas de Declaración",
        value=str(prev.get("observaciones", "") or ""),
    )

    # Cálculo de liquidación
    diferencia_neta = round(iva_debito - iva_credito - saldo_favor_ant - retenciones_iva, 2)
    impuesto_pagar = max(0.0, diferencia_neta)
    saldo_favor_sig = abs(min(0.0, diferencia_neta))

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("IVA Débito Fiscal", f"₡ {iva_debito:,.2f}")
    m2.metric("IVA Crédito Fiscal", f"₡ {iva_credito:,.2f}")
    m3.metric("Impuesto por Pagar", f"₡ {impuesto_pagar:,.2f}")
    m4.metric("Saldo a Favor Siguiente", f"₡ {saldo_favor_sig:,.2f}")

    resumen_dict = {
        "Período": periodo_str,
        "Ventas Netas Gravadas": round(ventas_grav, 2),
        "Ventas Exentas": round(ventas_exentas, 2),
        "IVA Débito Fiscal": round(iva_debito, 2),
        "Compras Netas Gravadas": round(compras_grav, 2),
        "Compras Exentas": round(compras_exentas, 2),
        "IVA Crédito Fiscal": round(iva_credito, 2),
        "Saldo a Favor Anterior": round(saldo_favor_ant, 2),
        "Retenciones IVA": round(retenciones_iva, 2),
        "Impuesto por Pagar": round(impuesto_pagar, 2),
        "Saldo a Favor Próximo Período": round(saldo_favor_sig, 2),
    }

    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾 Guardar Liquidación IVA (D-104)", type="primary", use_container_width=True):
            cur = conn.cursor()
            cur.execute(
                "DELETE FROM declaraciones_iva WHERE empresa_id = ? AND anio = ? AND mes = ?",
                (int(empresa_id), anio, mes),
            )
            cur.execute(
                """
                INSERT INTO declaraciones_iva (
                    empresa_id, anio, mes, fecha_desde, fecha_hasta,
                    ventas_netas_gravadas, ventas_exentas, iva_debito,
                    compras_netas_gravadas, compras_exentas, iva_credito,
                    saldo_favor_anterior, retenciones_iva,
                    impuesto_por_pagar, saldo_favor_siguiente,
                    detalle_tarifas_json, observaciones, fecha_actualizacion
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(empresa_id),
                    anio,
                    mes,
                    str(fecha_desde),
                    str(fecha_hasta),
                    float(ventas_grav),
                    float(ventas_exentas),
                    float(iva_debito),
                    float(compras_grav),
                    float(compras_exentas),
                    float(iva_credito),
                    float(saldo_favor_ant),
                    float(retenciones_iva),
                    float(impuesto_pagar),
                    float(saldo_favor_sig),
                    json.dumps(resumen_dict, ensure_ascii=False),
                    obs_iva.strip(),
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )
            conn.commit()
            subir_bd_a_sheets(conn)
            st.success(f"Liquidación de IVA para el período {periodo_str} guardada correctamente.")

    with b2:
        excel_bytes = _exportar_iva_excel(resumen_dict, df_movs, empresa_info, periodo_str)
        st.download_button(
            label="📥 Exportar Liquidación IVA (.xlsx)",
            data=excel_bytes,
            file_name=f"IVA_D104_{empresa_info.get('cedula', '')}_{periodo_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if not df_movs.empty:
        with st.expander("🔍 Ver movimientos contables detectados en el período"):
            st.dataframe(df_movs, use_container_width=True, hide_index=True)
