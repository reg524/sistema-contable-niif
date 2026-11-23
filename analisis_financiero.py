import datetime
import io
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from fpdf import FPDF

import mod_reportes


def _limpiar_texto_pdf(texto):
    """Evita errores de codificación en FPDF estándar (Latin-1)."""
    if texto is None:
        return ""
    reemplazos = {
        "₡": "CRC ",
        "—": "-",
        "–": "-",
        "“": '"',
        "”": '"',
        "’": "'",
        "‘": "'",
        "•": "-",
    }
    res = str(texto)
    for k, v in reemplazos.items():
        res = res.replace(k, v)
    return res.encode("latin-1", "replace").decode("latin-1")


def calcular_analisis_integral(conn, empresa_id, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta):
    """Calcula análisis vertical, horizontal, razones financieras y diagnóstico ejecutivo."""
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
        return None

    cat["cod_str"] = cat["codigo"].astype(str).str.strip()
    cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()

    saldos_act = mod_reportes._obtener_saldos_a_fecha(conn, empresa_id, fecha_act)
    saldos_ant = mod_reportes._obtener_saldos_a_fecha(conn, empresa_id, fecha_ant)

    # Clasificación de cuentas de balance
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
    ing_gast = cat[
        (
            cat["tipo_up"].str.contains("INGRESO|GASTO|COSTO|EGRESO")
            | cat["cod_str"].str.startswith(("4", "5", "6"))
        )
        & (~cat.index.isin(activos.index))
        & (~cat.index.isin(pasivos.index))
        & (~cat.index.isin(patrimonio.index))
    ]

    # Subclasificaciones corrientes vs no corrientes (por código 1.1 / 2.1 o nombre)
    act_mov = activos[activos["acepta_movimiento"] == 1]
    pas_mov = pasivos[pasivos["acepta_movimiento"] == 1]
    pat_mov = patrimonio[patrimonio["acepta_movimiento"] == 1]

    act_corr_mask = act_mov["cod_str"].str.startswith("1.1") | act_mov["nombre"].str.upper().str.contains("CORRIENTE|CIRCULANTE|CAJA|BANCO|CLIENTES|INVENTARIO")
    pas_corr_mask = pas_mov["cod_str"].str.startswith("2.1") | pas_mov["nombre"].str.upper().str.contains("CORRIENTE|CORTO PLAZO|PROVEEDOR|IVA|RETENCION")
    inv_mask = act_mov["nombre"].str.upper().str.contains("INVENTARIO|MERCADERIA|EXISTENCIA")

    act_corr_act = sum(saldos_act.get(int(cid), 0.0) for cid in act_mov[act_corr_mask]["id"])
    act_corr_ant = sum(saldos_ant.get(int(cid), 0.0) for cid in act_mov[act_corr_mask]["id"])
    inventarios_act = sum(saldos_act.get(int(cid), 0.0) for cid in act_mov[inv_mask]["id"])

    tot_act_act = sum(saldos_act.get(int(cid), 0.0) for cid in act_mov["id"])
    tot_act_ant = sum(saldos_ant.get(int(cid), 0.0) for cid in act_mov["id"])
    act_no_corr_act = tot_act_act - act_corr_act

    pas_corr_act = sum(-saldos_act.get(int(cid), 0.0) for cid in pas_mov[pas_corr_mask]["id"])
    pas_corr_ant = sum(-saldos_ant.get(int(cid), 0.0) for cid in pas_mov[pas_corr_mask]["id"])

    tot_pas_act = sum(-saldos_act.get(int(cid), 0.0) for cid in pas_mov["id"])
    tot_pas_ant = sum(-saldos_ant.get(int(cid), 0.0) for cid in pas_mov["id"])
    pas_no_corr_act = tot_pas_act - pas_corr_act

    res_ej_act = sum(-saldos_act.get(int(cid), 0.0) for cid in ing_gast["id"])
    res_ej_ant = sum(-saldos_ant.get(int(cid), 0.0) for cid in ing_gast["id"])

    tot_pat_act = sum(-saldos_act.get(int(cid), 0.0) for cid in pat_mov["id"]) + res_ej_act
    tot_pat_ant = sum(-saldos_ant.get(int(cid), 0.0) for cid in pat_mov["id"]) + res_ej_ant

    # Estado de Resultados del período seleccionado
    df_er, res_er = mod_reportes.construir_estado_resultados(conn, empresa_id, fecha_er_desde, fecha_er_hasta)
    ingresos_per = res_er.get("ingresos", 0.0)
    gastos_per = res_er.get("gastos", 0.0)
    utilidad_per = res_er.get("utilidad", 0.0)

    # Tabla Análisis Horizontal & Vertical del Balance
    filas_hv = []
    for _, r in act_mov.iterrows():
        cid = int(r["id"])
        v_ant = saldos_ant.get(cid, 0.0)
        v_act = saldos_act.get(cid, 0.0)
        if abs(v_ant) >= 0.005 or abs(v_act) >= 0.005:
            var_abs = v_act - v_ant
            var_rel = (var_abs / abs(v_ant) * 100.0) if abs(v_ant) >= 0.01 else (100.0 if abs(v_act) >= 0.01 else 0.0)
            peso_v = (v_act / tot_act_act * 100.0) if abs(tot_act_act) >= 0.01 else 0.0
            filas_hv.append(
                {
                    "Grupo": "ACTIVO",
                    "Código": r["cod_str"],
                    "Cuenta": r["nombre"],
                    "Saldo Anterior (₡)": round(v_ant, 2),
                    "Saldo Actual (₡)": round(v_act, 2),
                    "Variación Abs. (₡)": round(var_abs, 2),
                    "Var. Horizontal (%)": round(var_rel, 2),
                    "Análisis Vertical (%)": round(peso_v, 2),
                }
            )

    base_pas_pat = tot_pas_act + tot_pat_act
    for grupo_nom, df_grp in [("PASIVO", pas_mov), ("PATRIMONIO", pat_mov)]:
        for _, r in df_grp.iterrows():
            cid = int(r["id"])
            v_ant = -saldos_ant.get(cid, 0.0)
            v_act = -saldos_act.get(cid, 0.0)
            if abs(v_ant) >= 0.005 or abs(v_act) >= 0.005:
                var_abs = v_act - v_ant
                var_rel = (var_abs / abs(v_ant) * 100.0) if abs(v_ant) >= 0.01 else (100.0 if abs(v_act) >= 0.01 else 0.0)
                peso_v = (v_act / base_pas_pat * 100.0) if abs(base_pas_pat) >= 0.01 else 0.0
                filas_hv.append(
                    {
                        "Grupo": grupo_nom,
                        "Código": r["cod_str"],
                        "Cuenta": r["nombre"],
                        "Saldo Anterior (₡)": round(v_ant, 2),
                        "Saldo Actual (₡)": round(v_act, 2),
                        "Variación Abs. (₡)": round(var_abs, 2),
                        "Var. Horizontal (%)": round(var_rel, 2),
                        "Análisis Vertical (%)": round(peso_v, 2),
                    }
                )

    df_hv = pd.DataFrame(filas_hv)

    # Cálculo de Razones Financieras
    liquidez_corr = (act_corr_act / pas_corr_act) if pas_corr_act > 0 else 0.0
    prueba_acida = ((act_corr_act - inventarios_act) / pas_corr_act) if pas_corr_act > 0 else 0.0
    capital_trabajo = act_corr_act - pas_corr_act
    endeudamiento_tot = (tot_pas_act / tot_act_act * 100.0) if tot_act_act > 0 else 0.0
    apalancamiento = (tot_pas_act / tot_pat_act) if tot_pat_act > 0 else 0.0
    autonomia_fin = (tot_pat_act / tot_act_act * 100.0) if tot_act_act > 0 else 0.0
    margen_neto = (utilidad_per / ingresos_per * 100.0) if ingresos_per > 0 else 0.0
    roa = (utilidad_per / tot_act_act * 100.0) if tot_act_act > 0 else 0.0
    roe = (utilidad_per / tot_pat_act * 100.0) if tot_pat_act > 0 else 0.0

    indicadores = [
        {
            "Categoría": "Liquidez",
            "Indicador": "Razón Corriente (Liquidez General)",
            "Valor": f"{liquidez_corr:.2f} veces",
            "Referencia Óptima": "1.50 - 2.50 veces",
            "Interpretación": "Capacidad de cubrir pasivos de corto plazo con activos corrientes.",
        },
        {
            "Categoría": "Liquidez",
            "Indicador": "Prueba Ácida",
            "Valor": f"{prueba_acida:.2f} veces",
            "Referencia Óptima": ">= 1.00 vez",
            "Interpretación": "Liquidez inmediata sin depender de la realización de inventarios.",
        },
        {
            "Categoría": "Liquidez",
            "Indicador": "Capital de Trabajo Neto",
            "Valor": f"₡ {capital_trabajo:,.2f}",
            "Referencia Óptima": "> ₡ 0.00",
            "Interpretación": "Recursos operativos disponibles tras liquidar obligaciones corrientes.",
        },
        {
            "Categoría": "Solvencia y Endeudamiento",
            "Indicador": "Nivel de Endeudamiento Total",
            "Valor": f"{endeudamiento_tot:.2f} %",
            "Referencia Óptima": "<= 50.00 %",
            "Interpretación": "Porcentaje de los activos totales financiado con deuda de terceros.",
        },
        {
            "Categoría": "Solvencia y Endeudamiento",
            "Indicador": "Apalancamiento Financiero (Pasivo / Patrimonio)",
            "Valor": f"{apalancamiento:.2f} veces",
            "Referencia Óptima": "<= 1.00 vez",
            "Interpretación": "Relación entre financiamiento externo y recursos propios.",
        },
        {
            "Categoría": "Solvencia y Endeudamiento",
            "Indicador": "Autonomía Financiera (Patrimonio / Activo)",
            "Valor": f"{autonomia_fin:.2f} %",
            "Referencia Óptima": ">= 50.00 %",
            "Interpretación": "Proporción de activos respaldada por patrimonio neto.",
        },
        {
            "Categoría": "Rentabilidad",
            "Indicador": "Margen de Utilidad Neta",
            "Valor": f"{margen_neto:.2f} %",
            "Referencia Óptima": "> 10.00 %",
            "Interpretación": "Utilidad neta generada por cada colón facturado en el período.",
        },
        {
            "Categoría": "Rentabilidad",
            "Indicador": "ROA (Rendimiento sobre Activos)",
            "Valor": f"{roa:.2f} %",
            "Referencia Óptima": "> 5.00 %",
            "Interpretación": "Eficiencia de los activos totales para producir utilidad.",
        },
        {
            "Categoría": "Rentabilidad",
            "Indicador": "ROE (Rendimiento sobre Patrimonio)",
            "Valor": f"{roe:.2f} %",
            "Referencia Óptima": "> 12.00 %",
            "Interpretación": "Retorno obtenido sobre el capital invertido por los socios.",
        },
    ]
    df_indicadores = pd.DataFrame(indicadores)

    # Diagnóstico ejecutivo automático
    hallazgos = []
    if pas_corr_act == 0:
        hallazgos.append("La entidad no registra pasivos corrientes exigibles a la fecha de corte, reflejando holgura de corto plazo.")
    elif liquidez_corr >= 1.5:
        hallazgos.append(f"Liquidez saludable ({liquidez_corr:.2f}x): los activos corrientes cubren ampliamente los compromisos a corto plazo.")
    elif liquidez_corr >= 1.0:
        hallazgos.append(f"Liquidez ajustada ({liquidez_corr:.2f}x): se recomienda monitorear el flujo de caja y rotación de cobros.")
    else:
        hallazgos.append(f"Alerta de liquidez ({liquidez_corr:.2f}x): los pasivos corrientes superan a los activos circulantes (capital de trabajo negativo de ₡ {capital_trabajo:,.2f}).")

    if endeudamiento_tot <= 50.0:
        hallazgos.append(f"Estructura patrimonial sólida con un endeudamiento total del {endeudamiento_tot:.2f}% frente a una autonomía del {autonomia_fin:.2f}%.")
    else:
        hallazgos.append(f"Endeudamiento elevado ({endeudamiento_tot:.2f}%): más de la mitad de los activos está financiada por acreedores.")

    if utilidad_per > 0:
        hallazgos.append(f"Operación rentable en el período analizado con una utilidad de ₡ {utilidad_per:,.2f} (margen neto del {margen_neto:.2f}%).")
    elif utilidad_per < 0:
        hallazgos.append(f"Déficit operativo en el período por ₡ {abs(utilidad_per):,.2f}; conviene revisar estructura de costos y gastos operativos.")
    else:
        hallazgos.append("Punto de equilibrio operativo en el período seleccionado (sin utilidad ni pérdida neta registrada).")

    resumen_kpis = {
        "tot_act_act": round(tot_act_act, 2),
        "tot_act_ant": round(tot_act_ant, 2),
        "act_corr_act": round(act_corr_act, 2),
        "act_no_corr_act": round(act_no_corr_act, 2),
        "tot_pas_act": round(tot_pas_act, 2),
        "tot_pas_ant": round(tot_pas_ant, 2),
        "pas_corr_act": round(pas_corr_act, 2),
        "pas_no_corr_act": round(pas_no_corr_act, 2),
        "tot_pat_act": round(tot_pat_act, 2),
        "tot_pat_ant": round(tot_pat_ant, 2),
        "ingresos_per": round(ingresos_per, 2),
        "gastos_per": round(gastos_per, 2),
        "utilidad_per": round(utilidad_per, 2),
        "liquidez_corr": round(liquidez_corr, 2),
        "prueba_acida": round(prueba_acida, 2),
        "capital_trabajo": round(capital_trabajo, 2),
        "endeudamiento_tot": round(endeudamiento_tot, 2),
        "margen_neto": round(margen_neto, 2),
        "roa": round(roa, 2),
        "roe": round(roe, 2),
    }

    return {
        "kpis": resumen_kpis,
        "df_indicadores": df_indicadores,
        "df_hv": df_hv,
        "df_er": df_er,
        "hallazgos": hallazgos,
    }


def generar_excel_analisis(resultado, empresa_info, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta):
    """Genera libro Excel profesional multipestaña con el análisis financiero."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Pestaña 1: Resumen Ejecutivo y KPIs
        k = resultado["kpis"]
        df_res = pd.DataFrame(
            [
                {"Rubro / Indicador": "Total Activos (Corte Actual)", "Valor": k["tot_act_act"]},
                {"Rubro / Indicador": "Total Activos (Corte Anterior)", "Valor": k["tot_act_ant"]},
                {"Rubro / Indicador": "Activo Corriente", "Valor": k["act_corr_act"]},
                {"Rubro / Indicador": "Activo No Corriente", "Valor": k["act_no_corr_act"]},
                {"Rubro / Indicador": "Total Pasivos (Corte Actual)", "Valor": k["tot_pas_act"]},
                {"Rubro / Indicador": "Pasivo Corriente", "Valor": k["pas_corr_act"]},
                {"Rubro / Indicador": "Pasivo No Corriente", "Valor": k["pas_no_corr_act"]},
                {"Rubro / Indicador": "Total Patrimonio Neto", "Valor": k["tot_pat_act"]},
                {"Rubro / Indicador": "Ingresos del Período", "Valor": k["ingresos_per"]},
                {"Rubro / Indicador": "Costos y Gastos del Período", "Valor": k["gastos_per"]},
                {"Rubro / Indicador": "Utilidad / Excedente Neto", "Valor": k["utilidad_per"]},
                {"Rubro / Indicador": "Capital de Trabajo Neto", "Valor": k["capital_trabajo"]},
            ]
        )
        df_res.to_excel(writer, sheet_name="Resumen_Ejecutivo", index=False, startrow=5)
        ws1 = writer.sheets["Resumen_Ejecutivo"]
        ws1["A1"] = empresa_info.get("razon_social", "")
        ws1["A2"] = f"Cédula: {empresa_info.get('cedula', '')} | Régimen: {empresa_info.get('tipo_contribuyente', '')}"
        ws1["A3"] = f"INFORME DE ANÁLISIS FINANCIERO INTEGRAL (NIIF) - Corte: {fecha_act} vs {fecha_ant}"
        ws1["A4"] = f"Resultados analizados del {fecha_er_desde} al {fecha_er_hasta}"

        # Pestaña 2: Razones Financieras
        resultado["df_indicadores"].to_excel(writer, sheet_name="Razones_Financieras", index=False, startrow=3)
        ws2 = writer.sheets["Razones_Financieras"]
        ws2["A1"] = f"{empresa_info.get('razon_social', '')} - Indicadores y Ratios Financieros"
        ws2["A2"] = f"Fecha de Corte: {fecha_act}"

        # Pestaña 3: Análisis Horizontal y Vertical
        if not resultado["df_hv"].empty:
            resultado["df_hv"].to_excel(writer, sheet_name="Horizontal_y_Vertical", index=False, startrow=3)
            ws3 = writer.sheets["Horizontal_y_Vertical"]
            ws3["A1"] = f"{empresa_info.get('razon_social', '')} - Análisis Horizontal y Vertical"
            ws3["A2"] = f"Comparativo: {fecha_ant} vs {fecha_act}"

        # Pestaña 4: Estado de Resultados
        if not resultado["df_er"].empty:
            resultado["df_er"].to_excel(writer, sheet_name="Estado_Resultados", index=False, startrow=3)
            ws4 = writer.sheets["Estado_Resultados"]
            ws4["A1"] = f"{empresa_info.get('razon_social', '')} - Estado de Resultados"
            ws4["A2"] = f"Período: {fecha_er_desde} al {fecha_er_hasta}"

        # Pestaña 5: Diagnóstico
        df_diag = pd.DataFrame({"Hallazgos y Recomendaciones": resultado["hallazgos"]})
        df_diag.to_excel(writer, sheet_name="Diagnostico_Ejecutivo", index=False, startrow=2)
        writer.sheets["Diagnostico_Ejecutivo"]["A1"] = "DICTAMEN Y OBSERVACIONES FINANCIERAS"

    return output.getvalue()


def generar_word_analisis(resultado, empresa_info, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta, elaborador):
    """Genera un informe ejecutivo profesional en formato Microsoft Word (.docx)."""
    doc = Document()

    # Encabezado corporativo
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tit = p_tit.add_run(f"{empresa_info.get('razon_social', 'EMPRESA').upper()}\n")
    run_tit.bold = True
    run_tit.font.size = Pt(15)
    run_tit.font.color.rgb = RGBColor(27, 54, 93)

    run_sub = p_tit.add_run(
        f"Cédula: {empresa_info.get('cedula', '')} | Régimen: {empresa_info.get('tipo_contribuyente', '')}\n"
        f"INFORME DE ANÁLISIS FINANCIERO INTEGRAL (NIIF)\n"
        f"Corte Balance: {fecha_act} (Comparativo: {fecha_ant}) | Resultados: {fecha_er_desde} al {fecha_er_hasta}"
    )
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(80, 80, 80)

    doc.add_paragraph("_" * 78)

    # 1. Resumen Ejecutivo
    h1 = doc.add_heading("1. Resumen Ejecutivo y Estructura Patrimonial", level=1)
    h1.runs[0].font.color.rgb = RGBColor(27, 54, 93)

    k = resultado["kpis"]
    resumen_items = [
        ("Total Activos", f"₡ {k['tot_act_act']:,.2f}"),
        ("Activo Corriente / No Corriente", f"₡ {k['act_corr_act']:,.2f} / ₡ {k['act_no_corr_act']:,.2f}"),
        ("Total Pasivos", f"₡ {k['tot_pas_act']:,.2f}"),
        ("Total Patrimonio Neto", f"₡ {k['tot_pat_act']:,.2f}"),
        ("Ingresos del Período", f"₡ {k['ingresos_per']:,.2f}"),
        ("Costos y Gastos del Período", f"₡ {k['gastos_per']:,.2f}"),
        ("Utilidad / Excedente Neto", f"₡ {k['utilidad_per']:,.2f}"),
    ]
    tbl_res = doc.add_table(rows=1, cols=2)
    tbl_res.style = "Table Grid"
    tbl_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = tbl_res.rows[0].cells
    hdr[0].text = "Concepto Financiero"
    hdr[1].text = "Monto (CRC)"
    for cell in hdr:
        cell.paragraphs[0].runs[0].bold = True

    for concepto, monto in resumen_items:
        row = tbl_res.add_row().cells
        row[0].text = concepto
        row[1].text = monto
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 2. Razones Financieras
    h2 = doc.add_heading("2. Indicadores y Razones Financieras", level=1)
    h2.runs[0].font.color.rgb = RGBColor(27, 54, 93)

    df_ind = resultado["df_indicadores"]
    tbl_ind = doc.add_table(rows=1, cols=4)
    tbl_ind.style = "Table Grid"
    headers_ind = ["Categoría", "Indicador", "Resultado", "Interpretación"]
    for i, h_txt in enumerate(headers_ind):
        c = tbl_ind.rows[0].cells[i]
        c.text = h_txt
        c.paragraphs[0].runs[0].bold = True

    for _, r in df_ind.iterrows():
        row = tbl_ind.add_row().cells
        row[0].text = str(r["Categoría"])
        row[1].text = str(r["Indicador"])
        row[2].text = str(r["Valor"])
        row[3].text = str(r["Interpretación"])

    # 3. Análisis Horizontal y Vertical
    h3 = doc.add_heading("3. Análisis Horizontal y Vertical del Balance", level=1)
    h3.runs[0].font.color.rgb = RGBColor(27, 54, 93)

    df_hv = resultado["df_hv"]
    if df_hv.empty:
        doc.add_paragraph("No existen saldos registrados para detallar variaciones horizontal y vertical.")
    else:
        tbl_hv = doc.add_table(rows=1, cols=6)
        tbl_hv.style = "Table Grid"
        headers_hv = ["Código", "Cuenta", "Anterior (₡)", "Actual (₡)", "Var. Horiz. (%)", "Vertical (%)"]
        for i, h_txt in enumerate(headers_hv):
            c = tbl_hv.rows[0].cells[i]
            c.text = h_txt
            c.paragraphs[0].runs[0].bold = True

        for _, r in df_hv.iterrows():
            row = tbl_hv.add_row().cells
            row[0].text = str(r["Código"])
            row[1].text = str(r["Cuenta"])
            row[2].text = f"{r['Saldo Anterior (₡)']:,.2f}"
            row[3].text = f"{r['Saldo Actual (₡)']:,.2f}"
            row[4].text = f"{r['Var. Horizontal (%)']:.2f}%"
            row[5].text = f"{r['Análisis Vertical (%)']:.2f}%"

    # 4. Dictamen y Conclusiones
    h4 = doc.add_heading("4. Dictamen Técnico y Recomendaciones", level=1)
    h4.runs[0].font.color.rgb = RGBColor(27, 54, 93)

    for hallazgo in resultado["hallazgos"]:
        doc.add_paragraph(hallazgo, style="List Bullet")

    # Firma
    doc.add_paragraph("\n")
    p_firma = doc.add_paragraph()
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_firma.add_run(
        f"__________________________________________\n"
        f"{elaborador}\n"
        f"Contador Privado Incorporado | Emisión: {datetime.date.today()}"
    )
    r_f.bold = True
    r_f.font.size = Pt(10)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def generar_pdf_analisis(resultado, empresa_info, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta, elaborador):
    """Genera un informe ejecutivo profesional en formato PDF (.pdf) usando fpdf2."""
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Encabezado azul corporativo
    pdf.set_fill_color(27, 54, 93)
    pdf.rect(10, 10, 196, 26, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_xy(14, 13)
    pdf.cell(188, 7, _limpiar_texto_pdf(empresa_info.get("razon_social", "EMPRESA").upper()), ln=True)

    pdf.set_font("Helvetica", "", 9)
    pdf.set_x(14)
    pdf.cell(
        188,
        5,
        _limpiar_texto_pdf(
            f"Cédula: {empresa_info.get('cedula', '')} | Régimen: {empresa_info.get('tipo_contribuyente', '')}"
        ),
        ln=True,
    )
    pdf.set_x(14)
    pdf.cell(
        188,
        5,
        _limpiar_texto_pdf(
            f"INFORME DE ANÁLISIS FINANCIERO (NIIF) | Corte: {fecha_act} vs {fecha_ant} | Resultados: {fecha_er_desde} a {fecha_er_hasta}"
        ),
        ln=True,
    )

    pdf.ln(8)
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, _limpiar_texto_pdf("1. RESUMEN EJECUTIVO DE CIFRAS CLAVE"), ln=True)

    k = resultado["kpis"]
    pdf.set_text_color(30, 30, 30)
    pdf.set_font("Helvetica", "", 9)
    resumen_filas = [
        ("Total Activos", f"CRC {k['tot_act_act']:,.2f}", "Total Pasivos", f"CRC {k['tot_pas_act']:,.2f}"),
        ("Activo Corriente", f"CRC {k['act_corr_act']:,.2f}", "Pasivo Corriente", f"CRC {k['pas_corr_act']:,.2f}"),
        ("Total Patrimonio Neto", f"CRC {k['tot_pat_act']:,.2f}", "Capital de Trabajo Neto", f"CRC {k['capital_trabajo']:,.2f}"),
        ("Ingresos del Período", f"CRC {k['ingresos_per']:,.2f}", "Utilidad / Excedente Neto", f"CRC {k['utilidad_per']:,.2f}"),
    ]
    for c1, v1, c2, v2 in resumen_filas:
        pdf.set_fill_color(245, 247, 250)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(48, 6.5, _limpiar_texto_pdf(c1), border=1, fill=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(50, 6.5, _limpiar_texto_pdf(v1), border=1, align="R")
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(48, 6.5, _limpiar_texto_pdf(c2), border=1, fill=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(50, 6.5, _limpiar_texto_pdf(v2), border=1, align="R", ln=True)

    # Sección 2: Razones Financieras
    pdf.ln(5)
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, _limpiar_texto_pdf("2. INDICADORES Y RAZONES FINANCIERAS"), ln=True)

    pdf.set_fill_color(27, 54, 93)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(34, 6.5, "Categoría", border=1, fill=True)
    pdf.cell(64, 6.5, "Indicador", border=1, fill=True)
    pdf.cell(30, 6.5, "Valor", border=1, fill=True, align="R")
    pdf.cell(68, 6.5, "Referencia", border=1, fill=True, ln=True)

    pdf.set_text_color(30, 30, 30)
    pdf.set_font("Helvetica", "", 8)
    for _, r in resultado["df_indicadores"].iterrows():
        pdf.cell(34, 6, _limpiar_texto_pdf(r["Categoría"])[:20], border=1)
        pdf.cell(64, 6, _limpiar_texto_pdf(r["Indicador"])[:38], border=1)
        pdf.cell(30, 6, _limpiar_texto_pdf(r["Valor"]), border=1, align="R")
        pdf.cell(68, 6, _limpiar_texto_pdf(r["Referencia Óptima"])[:38], border=1, ln=True)

    # Sección 3: Análisis Horizontal y Vertical
    pdf.ln(5)
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, _limpiar_texto_pdf("3. ANÁLISIS HORIZONTAL Y VERTICAL DEL BALANCE"), ln=True)

    df_hv = resultado["df_hv"]
    if df_hv.empty:
        pdf.set_text_color(60, 60, 60)
        pdf.set_font("Helvetica", "I", 9)
        pdf.cell(0, 6, _limpiar_texto_pdf("No hay saldos disponibles en el período seleccionado."), ln=True)
    else:
        pdf.set_fill_color(27, 54, 93)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(22, 6, "Código", border=1, fill=True)
        pdf.cell(74, 6, "Cuenta Contable", border=1, fill=True)
        pdf.cell(30, 6, "Anterior", border=1, fill=True, align="R")
        pdf.cell(30, 6, "Actual", border=1, fill=True, align="R")
        pdf.cell(20, 6, "Horiz. %", border=1, fill=True, align="R")
        pdf.cell(20, 6, "Vert. %", border=1, fill=True, align="R", ln=True)

        pdf.set_text_color(30, 30, 30)
        pdf.set_font("Helvetica", "", 7.5)
        for _, r in df_hv.iterrows():
            pdf.cell(22, 5.5, _limpiar_texto_pdf(r["Código"])[:12], border=1)
            pdf.cell(74, 5.5, _limpiar_texto_pdf(r["Cuenta"])[:42], border=1)
            pdf.cell(30, 5.5, f"{r['Saldo Anterior (₡)']:,.2f}", border=1, align="R")
            pdf.cell(30, 5.5, f"{r['Saldo Actual (₡)']:,.2f}", border=1, align="R")
            pdf.cell(20, 5.5, f"{r['Var. Horizontal (%)']:.1f}%", border=1, align="R")
            pdf.cell(20, 5.5, f"{r['Análisis Vertical (%)']:.1f}%", border=1, align="R", ln=True)

    # Sección 4: Dictamen Ejecutivo
    pdf.ln(5)
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, _limpiar_texto_pdf("4. DICTAMEN EJECUTIVO Y RECOMENDACIONES"), ln=True)

    pdf.set_text_color(30, 30, 30)
    pdf.set_font("Helvetica", "", 9)
    for h_txt in resultado["hallazgos"]:
        pdf.multi_cell(0, 5.5, _limpiar_texto_pdf(f"- {h_txt}"))
        pdf.ln(1)

    # Bloque de Firma
    pdf.ln(8)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 5, "____________________________________________________", align="C", ln=True)
    pdf.cell(0, 5, _limpiar_texto_pdf(elaborador), align="C", ln=True)
    pdf.set_font("Helvetica", "", 8.5)
    pdf.cell(0, 5, _limpiar_texto_pdf(f"Contador Privado Incorporado | Fecha de emisión: {datetime.date.today()}"), align="C", ln=True)

    raw_pdf = pdf.output()
    return bytes(raw_pdf) if isinstance(raw_pdf, (bytes, bytearray)) else raw_pdf.encode("latin-1")


def renderizar_modulo_analisis_financiero(conn, empresa_id, empresa_info):
    st.subheader("📊 Análisis Financiero Integral (NIIF) e Informes Ejecutivos")

    hoy = datetime.date.today()
    c1, c2, c3, c4 = st.columns(4)
    fecha_act = c1.date_input("Corte Balance Actual", value=hoy, key="af_bal_act")
    fecha_ant = c2.date_input("Corte Balance Comparativo", value=datetime.date(hoy.year - 1, 12, 31), key="af_bal_ant")
    fecha_er_desde = c3.date_input("Resultados Desde", value=datetime.date(hoy.year, 1, 1), key="af_er_desde")
    fecha_er_hasta = c4.date_input("Resultados Hasta", value=hoy, key="af_er_hasta")

    elaborador = st.text_input(
        "Profesional Responsable del Informe (Pie de Firma):",
        value="Lic. Roy F. Esquivel González",
        key="af_elaborador",
    )

    resultado = calcular_analisis_integral(
        conn, empresa_id, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta
    )

    if resultado is None:
        st.info("No hay catálogo de cuentas configurado para esta empresa.")
        return

    k = resultado["kpis"]

    st.markdown("### 1. Tablero Ejecutivo de Indicadores Clave")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Liquidez Corriente", f"{k['liquidez_corr']:.2f} x", f"Prueba Ácida: {k['prueba_acida']:.2f} x")
    m2.metric("Capital de Trabajo", f"₡ {k['capital_trabajo']:,.2f}")
    m3.metric("Endeudamiento Total", f"{k['endeudamiento_tot']:.2f} %")
    m4.metric("Margen Neto / ROE", f"{k['margen_neto']:.2f} %", f"ROE: {k['roe']:.2f} %")

    st.divider()

    # Botones de descarga en Excel, Word y PDF
    st.markdown("### 2. Descarga de Informes Profesionales")
    cedula_clean = str(empresa_info.get("cedula", "empresa")).replace(" ", "_")

    excel_data = generar_excel_analisis(resultado, empresa_info, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta)
    word_data = generar_word_analisis(resultado, empresa_info, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta, elaborador)
    pdf_data = generar_pdf_analisis(resultado, empresa_info, fecha_act, fecha_ant, fecha_er_desde, fecha_er_hasta, elaborador)

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            label="📗 Descargar Informe en Excel (.xlsx)",
            data=excel_data,
            file_name=f"Analisis_Financiero_{cedula_clean}_{fecha_act}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            label="📘 Descargar Informe en Word (.docx)",
            data=word_data,
            file_name=f"Analisis_Financiero_{cedula_clean}_{fecha_act}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with d3:
        st.download_button(
            label="📕 Descargar Informe en PDF (.pdf)",
            data=pdf_data,
            file_name=f"Analisis_Financiero_{cedula_clean}_{fecha_act}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    st.divider()

    tab_ind, tab_hv, tab_diag = st.tabs(
        [
            "📐 Razones Financieras",
            "📊 Análisis Horizontal y Vertical",
            "🩺 Dictamen Ejecutivo",
        ]
    )

    with tab_ind:
        st.dataframe(resultado["df_indicadores"], use_container_width=True, hide_index=True)

    with tab_hv:
        if resultado["df_hv"].empty:
            st.info("No hay movimientos o saldos registrados en los cortes seleccionados.")
        else:
            st.dataframe(resultado["df_hv"], use_container_width=True, hide_index=True)

    with tab_diag:
        for h in resultado["hallazgos"]:
            st.info(f"📌 {h}")
