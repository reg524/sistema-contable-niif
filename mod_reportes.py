import datetime
import io
import pandas as pd
import streamlit as st

import motor_contable as mc


def _obtener_saldos_a_fecha(conn, empresa_id, fecha_corte):
    """Calcula el saldo acumulado (Debe - Haber) por cuenta hasta fecha_corte."""
    df_si = pd.read_sql_query(
        """
        SELECT cuenta_id, COALESCE(SUM(debe - haber), 0.0) AS saldo
        FROM saldos_iniciales
        WHERE empresa_id = ?
        GROUP BY cuenta_id
        """,
        conn,
        params=(int(empresa_id),),
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
        params=(int(empresa_id), str(fecha_corte)),
    )

    saldos = {}
    for _, r in df_si.iterrows():
        saldos[int(r["cuenta_id"])] = float(r["saldo"])
    for _, r in df_mov.iterrows():
        cid = int(r["cuenta_id"])
        saldos[cid] = saldos.get(cid, 0.0) + float(r["saldo"])
    return saldos


def _obtener_movimientos_periodo(conn, empresa_id, fecha_desde, fecha_hasta):
    """Calcula los movimientos netos (Debe - Haber) dentro de un rango de fechas."""
    df_mov = pd.read_sql_query(
        """
        SELECT ad.cuenta_id,
               COALESCE(SUM(ad.debe), 0.0) AS debe,
               COALESCE(SUM(ad.haber), 0.0) AS haber,
               COALESCE(SUM(ad.debe - ad.haber), 0.0) AS neto
        FROM asientos_detalle ad
        JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
        WHERE ac.empresa_id = ? AND ac.fecha BETWEEN ? AND ?
        GROUP BY ad.cuenta_id
        """,
        conn,
        params=(int(empresa_id), str(fecha_desde), str(fecha_hasta)),
    )
    return df_mov


def construir_estado_situacion_comparativo(conn, empresa_id, fecha_act, fecha_ant):
    """Construye el Estado de Situación Financiera (NIIF) comparativo entre dos fechas."""
    cat = pd.read_sql_query(
        """
        SELECT id, codigo, nombre, tipo, acepta_movimiento
        FROM catalogo_cuentas
        WHERE empresa_id = ?
        ORDER BY codigo
        """,
        conn,
        params=(int(empresa_id),),
    )

    if cat.empty:
        return pd.DataFrame(), {}

    cat["cod_str"] = cat["codigo"].astype(str).str.strip()
    cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()

    saldos_act = _obtener_saldos_a_fecha(conn, empresa_id, fecha_act)
    saldos_ant = _obtener_saldos_a_fecha(conn, empresa_id, fecha_ant)

    activos = cat[
        cat["tipo_up"].str.contains("ACTIVO") | cat["cod_str"].str.startswith("1")
    ]
    pasivos = cat[
        (cat["tipo_up"].str.contains("PASIVO") | cat["cod_str"].str.startswith("2"))
        & (~cat.index.isin(activos.index))
    ]
    patrimonio = cat[
        (cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL") | cat["cod_str"].str.startswith("3"))
        & (~cat.index.isin(activos.index))
        & (~cat.index.isin(pasivos.index))
    ]

    # Cuentas de resultados (excluyendo explícitamente Activo, Pasivo y Patrimonio)
    ing_gast = cat[
        (
            cat["tipo_up"].str.contains("INGRESO|GASTO|COSTO|EGRESO")
            | cat["cod_str"].str.startswith(("4", "5", "6"))
        )
        & (~cat.index.isin(activos.index))
        & (~cat.index.isin(pasivos.index))
        & (~cat.index.isin(patrimonio.index))
    ]

    res_ej_ant = sum(-saldos_ant.get(int(cid), 0.0) for cid in ing_gast["id"])
    res_ej_act = sum(-saldos_act.get(int(cid), 0.0) for cid in ing_gast["id"])

    filas = []

    # --- ACTIVOS ---
    filas.append({"Sección": "ACTIVOS", "Código": "", "Cuenta": "ACTIVOS", "Saldo Anterior": None, "Saldo Actual": None, "Variación": None})
    tot_act_ant = 0.0
    tot_act_act = 0.0
    for _, r in activos[activos["acepta_movimiento"] == 1].iterrows():
        cid = int(r["id"])
        s_ant = saldos_ant.get(cid, 0.0)
        s_act = saldos_act.get(cid, 0.0)
        if abs(s_ant) >= 0.005 or abs(s_act) >= 0.005:
            tot_act_ant += s_ant
            tot_act_act += s_act
            filas.append(
                {
                    "Sección": "ACTIVOS",
                    "Código": r["cod_str"],
                    "Cuenta": r["nombre"],
                    "Saldo Anterior": round(s_ant, 2),
                    "Saldo Actual": round(s_act, 2),
                    "Variación": round(s_act - s_ant, 2),
                }
            )
    filas.append(
        {
            "Sección": "ACTIVOS",
            "Código": "",
            "Cuenta": "TOTAL ACTIVOS",
            "Saldo Anterior": round(tot_act_ant, 2),
            "Saldo Actual": round(tot_act_act, 2),
            "Variación": round(tot_act_act - tot_act_ant, 2),
        }
    )

    # --- PASIVOS ---
    filas.append({"Sección": "PASIVOS", "Código": "", "Cuenta": "PASIVOS", "Saldo Anterior": None, "Saldo Actual": None, "Variación": None})
    tot_pas_ant = 0.0
    tot_pas_act = 0.0
    for _, r in pasivos[pasivos["acepta_movimiento"] == 1].iterrows():
        cid = int(r["id"])
        s_ant = -saldos_ant.get(cid, 0.0)
        s_act = -saldos_act.get(cid, 0.0)
        if abs(s_ant) >= 0.005 or abs(s_act) >= 0.005:
            tot_pas_ant += s_ant
            tot_pas_act += s_act
            filas.append(
                {
                    "Sección": "PASIVOS",
                    "Código": r["cod_str"],
                    "Cuenta": r["nombre"],
                    "Saldo Anterior": round(s_ant, 2),
                    "Saldo Actual": round(s_act, 2),
                    "Variación": round(s_act - s_ant, 2),
                }
            )
    filas.append(
        {
            "Sección": "PASIVOS",
            "Código": "",
            "Cuenta": "TOTAL PASIVOS",
            "Saldo Anterior": round(tot_pas_ant, 2),
            "Saldo Actual": round(tot_pas_act, 2),
            "Variación": round(tot_pas_act - tot_pas_ant, 2),
        }
    )

    # --- PATRIMONIO ---
    filas.append({"Sección": "PATRIMONIO", "Código": "", "Cuenta": "PATRIMONIO", "Saldo Anterior": None, "Saldo Actual": None, "Variación": None})
    tot_pat_ant = 0.0
    tot_pat_act = 0.0
    for _, r in patrimonio[patrimonio["acepta_movimiento"] == 1].iterrows():
        cid = int(r["id"])
        s_ant = -saldos_ant.get(cid, 0.0)
        s_act = -saldos_act.get(cid, 0.0)
        if abs(s_ant) >= 0.005 or abs(s_act) >= 0.005:
            tot_pat_ant += s_ant
            tot_pat_act += s_act
            filas.append(
                {
                    "Sección": "PATRIMONIO",
                    "Código": r["cod_str"],
                    "Cuenta": r["nombre"],
                    "Saldo Anterior": round(s_ant, 2),
                    "Saldo Actual": round(s_act, 2),
                    "Variación": round(s_act - s_ant, 2),
                }
            )

    # Resultado acumulado / del período
    tot_pat_ant += res_ej_ant
    tot_pat_act += res_ej_act
    filas.append(
        {
            "Sección": "PATRIMONIO",
            "Código": "3.9.99",
            "Cuenta": "Resultado del Ejercicio / Acumulado",
            "Saldo Anterior": round(res_ej_ant, 2),
            "Saldo Actual": round(res_ej_act, 2),
            "Variación": round(res_ej_act - res_ej_ant, 2),
        }
    )
    filas.append(
        {
            "Sección": "PATRIMONIO",
            "Código": "",
            "Cuenta": "TOTAL PATRIMONIO",
            "Saldo Anterior": round(tot_pat_ant, 2),
            "Saldo Actual": round(tot_pat_act, 2),
            "Variación": round(tot_pat_act - tot_pat_ant, 2),
        }
    )
    filas.append(
        {
            "Sección": "TOTAL",
            "Código": "",
            "Cuenta": "TOTAL PASIVO Y PATRIMONIO",
            "Saldo Anterior": round(tot_pas_ant + tot_pat_ant, 2),
            "Saldo Actual": round(tot_pas_act + tot_pat_act, 2),
            "Variación": round((tot_pas_act + tot_pat_act) - (tot_pas_ant + tot_pat_ant), 2),
        }
    )

    resumen = {
        "total_activos_act": round(tot_act_act, 2),
        "total_pasivos_act": round(tot_pas_act, 2),
        "total_patrimonio_act": round(tot_pat_act, 2),
        "ecuacion_diferencia": round(tot_act_act - (tot_pas_act + tot_pat_act), 2),
    }
    return pd.DataFrame(filas), resumen


def construir_estado_resultados(conn, empresa_id, fecha_desde, fecha_hasta):
    """Construye el Estado de Resultados Integral (NIIF) para un rango de fechas."""
    cat = pd.read_sql_query(
        """
        SELECT id, codigo, nombre, tipo, acepta_movimiento
        FROM catalogo_cuentas
        WHERE empresa_id = ?
        ORDER BY codigo
        """,
        conn,
        params=(int(empresa_id),),
    )
    if cat.empty:
        return pd.DataFrame(), {}

    cat["cod_str"] = cat["codigo"].astype(str).str.strip()
    cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()

    df_mov = _obtener_movimientos_periodo(conn, empresa_id, fecha_desde, fecha_hasta)
    mapa_neto = {int(r["cuenta_id"]): float(r["neto"]) for _, r in df_mov.iterrows()}

    ingresos = cat[
        (cat["tipo_up"].str.contains("INGRESO") | cat["cod_str"].str.startswith("4"))
        & (~cat["tipo_up"].str.contains("ACTIVO|PASIVO|PATRIMONIO"))
    ]
    costos_gastos = cat[
        (
            cat["tipo_up"].str.contains("GASTO|COSTO|EGRESO")
            | cat["cod_str"].str.startswith(("5", "6"))
        )
        & (~cat["tipo_up"].str.contains("ACTIVO|PASIVO|PATRIMONIO|INGRESO"))
    ]

    filas = []
    tot_ing = 0.0
    for _, r in ingresos[ingresos["acepta_movimiento"] == 1].iterrows():
        cid = int(r["id"])
        monto = -mapa_neto.get(cid, 0.0)
        if abs(monto) >= 0.005:
            tot_ing += monto
            filas.append({"Rubro": "INGRESOS", "Código": r["cod_str"], "Cuenta": r["nombre"], "Monto": round(monto, 2)})

    filas.append({"Rubro": "INGRESOS", "Código": "", "Cuenta": "TOTAL INGRESOS", "Monto": round(tot_ing, 2)})

    tot_gas = 0.0
    for _, r in costos_gastos[costos_gastos["acepta_movimiento"] == 1].iterrows():
        cid = int(r["id"])
        monto = mapa_neto.get(cid, 0.0)
        if abs(monto) >= 0.005:
            tot_gas += monto
            filas.append({"Rubro": "COSTOS Y GASTOS", "Código": r["cod_str"], "Cuenta": r["nombre"], "Monto": round(monto, 2)})

    filas.append({"Rubro": "COSTOS Y GASTOS", "Código": "", "Cuenta": "TOTAL COSTOS Y GASTOS", "Monto": round(tot_gas, 2)})

    utilidad = tot_ing - tot_gas
    filas.append({"Rubro": "RESULTADO", "Código": "", "Cuenta": "UTILIDAD / EXCEDENTE DEL PERÍODO", "Monto": round(utilidad, 2)})

    resumen = {
        "ingresos": round(tot_ing, 2),
        "gastos": round(tot_gas, 2),
        "utilidad": round(utilidad, 2),
    }
    return pd.DataFrame(filas), resumen


def construir_balanza_comprobacion(conn, empresa_id, fecha_desde, fecha_hasta):
    """Construye la Balanza de Comprobación de Sumas y Saldos."""
    cat = pd.read_sql_query(
        """
        SELECT id, codigo, nombre, tipo
        FROM catalogo_cuentas
        WHERE empresa_id = ? AND acepta_movimiento = 1
        ORDER BY codigo
        """,
        conn,
        params=(int(empresa_id),),
    )
    if cat.empty:
        return pd.DataFrame()

    fecha_previa = (pd.to_datetime(fecha_desde) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    saldos_ini = _obtener_saldos_a_fecha(conn, empresa_id, fecha_previa)
    df_mov = _obtener_movimientos_periodo(conn, empresa_id, fecha_desde, fecha_hasta)
    mapa_debe = {int(r["cuenta_id"]): float(r["debe"]) for _, r in df_mov.iterrows()}
    mapa_haber = {int(r["cuenta_id"]): float(r["haber"]) for _, r in df_mov.iterrows()}

    filas = []
    for _, r in cat.iterrows():
        cid = int(r["id"])
        s_ini = saldos_ini.get(cid, 0.0)
        m_debe = mapa_debe.get(cid, 0.0)
        m_haber = mapa_haber.get(cid, 0.0)
        s_fin = s_ini + m_debe - m_haber

        if abs(s_ini) >= 0.005 or m_debe >= 0.005 or m_haber >= 0.005 or abs(s_fin) >= 0.005:
            filas.append(
                {
                    "Código": r["codigo"],
                    "Cuenta": r["nombre"],
                    "Saldo Inicial Debe": round(s_ini, 2) if s_ini > 0 else 0.0,
                    "Saldo Inicial Haber": round(-s_ini, 2) if s_ini < 0 else 0.0,
                    "Mov. Debe": round(m_debe, 2),
                    "Mov. Haber": round(m_haber, 2),
                    "Saldo Final Debe": round(s_fin, 2) if s_fin > 0 else 0.0,
                    "Saldo Final Haber": round(-s_fin, 2) if s_fin < 0 else 0.0,
                }
            )

    df_bal = pd.DataFrame(filas)
    if not df_bal.empty:
        totales = {
            "Código": "",
            "Cuenta": "SUMAS IGUALES",
            "Saldo Inicial Debe": round(df_bal["Saldo Inicial Debe"].sum(), 2),
            "Saldo Inicial Haber": round(df_bal["Saldo Inicial Haber"].sum(), 2),
            "Mov. Debe": round(df_bal["Mov. Debe"].sum(), 2),
            "Mov. Haber": round(df_bal["Mov. Haber"].sum(), 2),
            "Saldo Final Debe": round(df_bal["Saldo Final Debe"].sum(), 2),
            "Saldo Final Haber": round(df_bal["Saldo Final Haber"].sum(), 2),
        }
        df_bal = pd.concat([df_bal, pd.DataFrame([totales])], ignore_index=True)
    return df_bal


def _exportar_df_excel(df, nombre_hoja, titulo_reporte, empresa_info):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja[:31], index=False, startrow=4)
        ws = writer.sheets[nombre_hoja[:31]]
        ws["A1"] = empresa_info.get("razon_social", "")
        ws["A2"] = f"Cédula: {empresa_info.get('cedula', '')}"
        ws["A3"] = titulo_reporte
    return output.getvalue()


def renderizar_modulo_reportes(conn, empresa_id, empresa_info):
    st.subheader("📑 Estados Financieros NIIF y Auxiliares Contables")

    tab_esf, tab_er, tab_bal, tab_may = st.tabs(
        [
            "⚖️ Estado de Situación Financiera",
            "📈 Estado de Resultados",
            "🧮 Balanza de Comprobación",
            "📚 Libro Mayor Auxiliar",
        ]
    )

    hoy = datetime.date.today()

    # 1. ESTADO DE SITUACIÓN FINANCIERA
    with tab_esf:
        c1, c2 = st.columns(2)
        fecha_act = c1.date_input("Fecha de Corte Actual", value=hoy, key="esf_act")
        fecha_ant = c2.date_input(
            "Fecha Comparativa Anterior",
            value=datetime.date(hoy.year - 1, 12, 31),
            key="esf_ant",
        )

        df_esf, res_esf = construir_estado_situacion_comparativo(conn, empresa_id, fecha_act, fecha_ant)
        if df_esf.empty:
            st.info("No hay datos contables registrados para esta empresa.")
        else:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Activos", f"₡ {res_esf['total_activos_act']:,.2f}")
            m2.metric("Total Pasivos", f"₡ {res_esf['total_pasivos_act']:,.2f}")
            m3.metric("Total Patrimonio", f"₡ {res_esf['total_patrimonio_act']:,.2f}")
            m4.metric("Cuadre Ecuación", f"₡ {res_esf['ecuacion_diferencia']:,.2f}")

            st.dataframe(df_esf, use_container_width=True, hide_index=True)

            excel_esf = _exportar_df_excel(
                df_esf,
                "Situacion_Financiera",
                f"Estado de Situación Financiera al {fecha_act} vs {fecha_ant}",
                empresa_info,
            )
            st.download_button(
                "📥 Descargar Estado de Situación (.xlsx)",
                data=excel_esf,
                file_name=f"Estado_Situacion_{empresa_info.get('cedula', '')}_{fecha_act}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # 2. ESTADO DE RESULTADOS
    with tab_er:
        c1, c2 = st.columns(2)
        er_desde = c1.date_input("Desde", value=datetime.date(hoy.year, 1, 1), key="er_desde")
        er_hasta = c2.date_input("Hasta", value=hoy, key="er_hasta")

        df_er, res_er = construir_estado_resultados(conn, empresa_id, er_desde, er_hasta)
        if df_er.empty:
            st.info("No hay movimientos de resultados en el período seleccionado.")
        else:
            m1, m2, m3 = st.columns(3)
            m1.metric("Ingresos Totales", f"₡ {res_er['ingresos']:,.2f}")
            m2.metric("Costos y Gastos", f"₡ {res_er['gastos']:,.2f}")
            m3.metric("Utilidad / Excedente", f"₡ {res_er['utilidad']:,.2f}")

            st.dataframe(df_er, use_container_width=True, hide_index=True)

            excel_er = _exportar_df_excel(
                df_er,
                "Estado_Resultados",
                f"Estado de Resultados del {er_desde} al {er_hasta}",
                empresa_info,
            )
            st.download_button(
                "📥 Descargar Estado de Resultados (.xlsx)",
                data=excel_er,
                file_name=f"Estado_Resultados_{empresa_info.get('cedula', '')}_{er_hasta}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # 3. BALANZA DE COMPROBACIÓN
    with tab_bal:
        c1, c2 = st.columns(2)
        bal_desde = c1.date_input("Desde", value=datetime.date(hoy.year, 1, 1), key="bal_desde")
        bal_hasta = c2.date_input("Hasta", value=hoy, key="bal_hasta")

        df_bal = construir_balanza_comprobacion(conn, empresa_id, bal_desde, bal_hasta)
        if df_bal.empty:
            st.info("No hay saldos ni movimientos para mostrar en la balanza.")
        else:
            st.dataframe(df_bal, use_container_width=True, hide_index=True)
            excel_bal = _exportar_df_excel(
                df_bal,
                "Balanza_Comprobacion",
                f"Balanza de Comprobación del {bal_desde} al {bal_hasta}",
                empresa_info,
            )
            st.download_button(
                "📥 Descargar Balanza (.xlsx)",
                data=excel_bal,
                file_name=f"Balanza_{empresa_info.get('cedula', '')}_{bal_hasta}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # 4. LIBRO MAYOR AUXILIAR POR CUENTA
    with tab_may:
        cat_mov = pd.read_sql_query(
            """
            SELECT id, codigo, nombre
            FROM catalogo_cuentas
            WHERE empresa_id = ? AND acepta_movimiento = 1
            ORDER BY codigo
            """,
            conn,
            params=(int(empresa_id),),
        )
        if cat_mov.empty:
            st.info("No hay cuentas de movimiento registradas.")
        else:
            c1, c2, c3 = st.columns([4, 2, 2])
            opciones = {f"{r['codigo']} - {r['nombre']}": int(r["id"]) for _, r in cat_mov.iterrows()}
            cta_sel = c1.selectbox("Cuenta Contable", list(opciones.keys()), key="may_cta")
            may_desde = c2.date_input("Desde", value=datetime.date(hoy.year, 1, 1), key="may_desde")
            may_hasta = c3.date_input("Hasta", value=hoy, key="may_hasta")

            cid_sel = opciones[cta_sel]
            fecha_previa = (pd.to_datetime(may_desde) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
            saldo_ini_cta = _obtener_saldos_a_fecha(conn, empresa_id, fecha_previa).get(cid_sel, 0.0)

            df_aux = pd.read_sql_query(
                """
                SELECT ac.fecha AS Fecha, ac.num_asiento AS Asiento, ac.num_comprobante AS Documento,
                       ac.concepto AS Concepto, ad.debe AS Debe, ad.haber AS Haber
                FROM asientos_detalle ad
                JOIN asientos_cabecera ac ON ad.asiento_id = ac.id
                WHERE ac.empresa_id = ? AND ad.cuenta_id = ? AND ac.fecha BETWEEN ? AND ?
                ORDER BY ac.fecha, ac.num_asiento
                """,
                conn,
                params=(int(empresa_id), cid_sel, str(may_desde), str(may_hasta)),
            )

            saldo_corr = saldo_ini_cta
            saldos_acum = []
            for _, r in df_aux.iterrows():
                saldo_corr += float(r["Debe"]) - float(r["Haber"])
                saldos_acum.append(round(saldo_corr, 2))
            df_aux["Saldo Acumulado"] = saldos_acum

            st.metric("Saldo Inicial al " + str(may_desde), f"₡ {saldo_ini_cta:,.2f}")
            st.dataframe(df_aux, use_container_width=True, hide_index=True)
