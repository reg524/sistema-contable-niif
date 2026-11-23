import datetime
import io
import json
import pandas as pd
import streamlit as st

from sync_sheets import subir_bd_a_sheets


def _asegurar_tabla_renta(conn):
    conn.execute(
        """
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
        """
    )
    conn.commit()


def _calcular_resultados_contables(conn, empresa_id, fecha_desde, fecha_hasta):
    """Calcula ingresos y costos/gastos contables del período fiscal."""
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
    if df.empty:
        return 0.0, 0.0, df

    df["cod_str"] = df["codigo"].astype(str).str.strip()
    df["tipo_up"] = df["tipo"].astype(str).str.upper().str.strip()

    ing_mask = (df["tipo_up"].str.contains("INGRESO") | df["cod_str"].str.startswith("4")) & (
        ~df["tipo_up"].str.contains("ACTIVO|PASIVO|PATRIMONIO")
    )
    gas_mask = (
        df["tipo_up"].str.contains("GASTO|COSTO|EGRESO") | df["cod_str"].str.startswith(("5", "6"))
    ) & (~df["tipo_up"].str.contains("ACTIVO|PASIVO|PATRIMONIO|INGRESO"))

    ingresos = round(float((df[ing_mask]["haber"] - df[ing_mask]["debe"]).sum()), 2)
    gastos = round(float((df[gas_mask]["debe"] - df[gas_mask]["haber"]).sum()), 2)
    return ingresos, gastos, df


def _calcular_impuesto_pj_cr(renta_neta):
    """Escala progresiva orientativa para Persona Jurídica en Costa Rica."""
    if renta_neta <= 0:
        return 0.0
    if renta_neta <= 5_870_000:
        return round(renta_neta * 0.05, 2)
    if renta_neta <= 8_810_000:
        return round(renta_neta * 0.10, 2)
    if renta_neta <= 11_760_000:
        return round(renta_neta * 0.15, 2)
    if renta_neta <= 117_850_000:
        return round(renta_neta * 0.20, 2)
    return round(renta_neta * 0.30, 2)


def _exportar_renta_excel(resumen_dict, df_movs, empresa_info, anio_fiscal):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_res = pd.DataFrame(
            [{"Concepto": k, "Monto (₡)": v} for k, v in resumen_dict.items()]
        )
        df_res.to_excel(writer, sheet_name="Conciliacion_Renta", index=False, startrow=4)
        if df_movs is not None and not df_movs.empty:
            df_movs.to_excel(writer, sheet_name="Movimientos_Periodo", index=False)

        ws = writer.sheets["Conciliacion_Renta"]
        ws["A1"] = empresa_info.get("razon_social", "")
        ws["A2"] = f"Cédula: {empresa_info.get('cedula', '')}"
        ws["A3"] = f"Conciliación y Cálculo Impuesto sobre la Renta (D-101) - Período {anio_fiscal}"
    return output.getvalue()


def renderizar_modulo_renta(conn, empresa_id, empresa_info):
    _asegurar_tabla_renta(conn)
    st.subheader("🏛️ Conciliación Contable-Fiscal e Impuesto sobre la Renta (D-101)")

    hoy = datetime.date.today()
    c1, c2, c3 = st.columns([2, 3, 3])
    anio_fiscal = int(
        c1.number_input("Año Fiscal", min_value=2020, max_value=2035, value=hoy.year, key="renta_anio")
    )
    regimen = c2.selectbox(
        "Régimen Tributario",
        [
            "Régimen Tradicional - Persona Jurídica",
            "Régimen Tradicional - Persona Física / Actividades Lucrativas",
            "Régimen Simplificado / Exento",
        ],
    )
    fecha_desde = datetime.date(anio_fiscal, 1, 1)
    fecha_hasta = datetime.date(anio_fiscal, 12, 31)
    c3.info(f"Período fiscal: **{fecha_desde}** al **{fecha_hasta}**")

    prev_df = pd.read_sql_query(
        """
        SELECT * FROM declaraciones_renta
        WHERE empresa_id = ? AND anio_fiscal = ?
        ORDER BY id DESC LIMIT 1
        """,
        conn,
        params=(int(empresa_id), anio_fiscal),
    )
    prev = prev_df.iloc[0].to_dict() if not prev_df.empty else {}

    ing_auto, gas_auto, df_movs = _calcular_resultados_contables(
        conn, empresa_id, fecha_desde, fecha_hasta
    )

    st.markdown("### 1. Determinación de la Utilidad Contable")
    u1, u2, u3 = st.columns(3)
    ingresos_brutos = u1.number_input(
        "Ingresos Brutos Contables (₡)",
        min_value=0.0,
        value=float(prev.get("ingresos_brutos", ing_auto) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    costos_gastos = u2.number_input(
        "Costos y Gastos Contables (₡)",
        min_value=0.0,
        value=float(prev.get("costos_gastos_contables", gas_auto) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    utilidad_contable = round(ingresos_brutos - costos_gastos, 2)
    u3.metric("Utilidad / Pérdida Contable", f"₡ {utilidad_contable:,.2f}")

    st.markdown("### 2. Conciliación Fiscal (Ajustes NIIF vs. Ley del Impuesto sobre la Renta)")
    a1, a2, a3 = st.columns(3)
    gastos_no_ded = a1.number_input(
        "(+) Gastos No Deducibles (₡)",
        min_value=0.0,
        value=float(prev.get("gastos_no_deducibles", 0.0) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    ingresos_no_grav = a2.number_input(
        "(-) Ingresos No Gravables / Exentos (₡)",
        min_value=0.0,
        value=float(prev.get("ingresos_no_gravables", 0.0) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    renta_neta = round(utilidad_contable + gastos_no_ded - ingresos_no_grav, 2)
    a3.metric("Renta Neta Imponible", f"₡ {max(0.0, renta_neta):,.2f}")

    st.markdown("### 3. Impuesto Determinado, Pagos Parciales y Retenciones")
    p1, p2, p3, p4 = st.columns(4)
    imp_calculado_sug = (
        _calcular_impuesto_pj_cr(max(0.0, renta_neta))
        if "Jurídica" in regimen
        else round(max(0.0, renta_neta) * 0.15, 2)
    )
    impuesto_det = p1.number_input(
        "Impuesto Determinado (₡)",
        min_value=0.0,
        value=float(prev.get("impuesto_determinado", imp_calculado_sug) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    pagos_parciales = p2.number_input(
        "(-) Pagos Parciales a Cuenta (₡)",
        min_value=0.0,
        value=float(prev.get("pagos_parciales", 0.0) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    retenciones_2pct = p3.number_input(
        "(-) Retenciones 2% y Otras (₡)",
        min_value=0.0,
        value=float(prev.get("retenciones_2pct", 0.0) or 0.0),
        step=10000.0,
        format="%.2f",
    )
    creditos_fisc = p4.number_input(
        "(-) Créditos Fiscales / Saldo Favor Ant. (₡)",
        min_value=0.0,
        value=float(prev.get("creditos_fiscales", 0.0) or 0.0),
        step=10000.0,
        format="%.2f",
    )

    obs_renta = st.text_input(
        "Observaciones de Conciliación Fiscal",
        value=str(prev.get("observaciones", "") or ""),
    )

    saldo_neto = round(impuesto_det - pagos_parciales - retenciones_2pct - creditos_fisc, 2)
    impuesto_pagar = max(0.0, saldo_neto)
    saldo_favor = abs(min(0.0, saldo_neto))

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Utilidad Contable", f"₡ {utilidad_contable:,.2f}")
    m2.metric("Renta Neta Imponible", f"₡ {max(0.0, renta_neta):,.2f}")
    m3.metric("Impuesto por Pagar", f"₡ {impuesto_pagar:,.2f}")
    m4.metric("Saldo a Favor", f"₡ {saldo_favor:,.2f}")

    resumen_dict = {
        "Año Fiscal": anio_fiscal,
        "Régimen": regimen,
        "Ingresos Brutos Contables": round(ingresos_brutos, 2),
        "Costos y Gastos Contables": round(costos_gastos, 2),
        "Utilidad Contable": utilidad_contable,
        "Gastos No Deducibles": round(gastos_no_ded, 2),
        "Ingresos No Gravables": round(ingresos_no_grav, 2),
        "Renta Neta Imponible": round(max(0.0, renta_neta), 2),
        "Impuesto Determinado": round(impuesto_det, 2),
        "Pagos Parciales": round(pagos_parciales, 2),
        "Retenciones 2% y Otras": round(retenciones_2pct, 2),
        "Créditos Fiscales": round(creditos_fisc, 2),
        "Impuesto por Pagar": impuesto_pagar,
        "Saldo a Favor": saldo_favor,
    }

    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾 Guardar Declaración Renta (D-101)", type="primary", use_container_width=True):
            cur = conn.cursor()
            cur.execute(
                "DELETE FROM declaraciones_renta WHERE empresa_id = ? AND anio_fiscal = ?",
                (int(empresa_id), anio_fiscal),
            )
            cur.execute(
                """
                INSERT INTO declaraciones_renta (
                    empresa_id, anio_fiscal, fecha_desde, fecha_hasta, regimen,
                    ingresos_brutos, costos_gastos_contables, utilidad_contable,
                    gastos_no_deducibles, ingresos_no_gravables, renta_neta_imponible,
                    impuesto_determinado, pagos_parciales, retenciones_2pct,
                    creditos_fiscales, impuesto_por_pagar, saldo_favor,
                    detalle_ajustes_json, observaciones, fecha_actualizacion
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(empresa_id),
                    anio_fiscal,
                    str(fecha_desde),
                    str(fecha_hasta),
                    regimen,
                    float(ingresos_brutos),
                    float(costos_gastos),
                    float(utilidad_contable),
                    float(gastos_no_ded),
                    float(ingresos_no_grav),
                    float(max(0.0, renta_neta)),
                    float(impuesto_det),
                    float(pagos_parciales),
                    float(retenciones_2pct),
                    float(creditos_fiscales),
                    float(impuesto_pagar),
                    float(saldo_favor),
                    json.dumps(resumen_dict, ensure_ascii=False),
                    obs_renta.strip(),
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )
            conn.commit()
            subir_bd_a_sheets(conn)
            st.success(f"Declaración de Renta para el año fiscal {anio_fiscal} guardada correctamente.")

    with b2:
        excel_bytes = _exportar_renta_excel(resumen_dict, df_movs, empresa_info, anio_fiscal)
        st.download_button(
            label="📥 Exportar Conciliación Renta (.xlsx)",
            data=excel_bytes,
            file_name=f"Renta_D101_{empresa_info.get('cedula', '')}_{anio_fiscal}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if df_movs is not None and not df_movs.empty:
        with st.expander("🔍 Ver detalle de cuentas del período fiscal"):
            st.dataframe(df_movs, use_container_width=True, hide_index=True)
