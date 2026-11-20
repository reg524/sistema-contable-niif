import calendar
from datetime import date, datetime
import io
import json
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st


TARIFAS_IVA_CR = [
    ("13%", 0.13),
    ("4%", 0.04),
    ("2%", 0.02),
    ("1%", 0.01),
    ("Exento / 0%", 0.0),
]


def _fmt_moneda(valor):
  if valor is None or valor == "":
    return ""
  val = float(valor)
  if round(val, 2) == 0:
    return "-"
  texto = f"{abs(val):,.2f}".replace(",", " ").replace(".", ",")
  return f"({texto})" if val < 0 else texto


def _fmt_fecha_dmy(fecha_val):
  if fecha_val is None or str(fecha_val).strip() == "":
    return ""
  dt = pd.to_datetime(fecha_val, errors="coerce")
  if pd.isna(dt):
    return str(fecha_val)
  return dt.strftime("%d-%m-%Y")


def _limpiar_pdf(texto):
  if texto is None:
    return ""
  s = str(texto)
  reemplazos = {
      "—": "-",
      "–": "-",
      "₡": "C",
      "•": "-",
      "“": '"',
      "”": '"',
      "‘": "'",
      "’": "'",
      "…": "...",
  }
  for k, v in reemplazos.items():
    s = s.replace(k, v)
  return s.encode("latin-1", errors="replace").decode("latin-1")


def asegurar_tabla_iva(conn):
  cur = conn.cursor()
  cur.execute("""
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
    """)
  conn.commit()


def _detectar_tarifa_texto(texto_cuenta, texto_concepto=""):
  txt = f"{texto_cuenta} {texto_concepto}".upper()
  if "1%" in txt or "CANASTA" in txt:
    return "1%"
  if "2%" in txt or "MEDICAMENT" in txt or "EDUCACION" in txt or "EDUCACIÓN" in txt:
    return "2%"
  if "4%" in txt or "SALUD" in txt or "MEDICO" in txt or "MÉDICO" in txt or "BOLETO" in txt:
    return "4%"
  if "EXENT" in txt or "NO SUJET" in txt or "0%" in txt:
    return "Exento / 0%"
  return "13%"


def calcular_borrador_iva_mensual(df_mov, df_cat, f_ini, f_fin):
  """
  Analiza los movimientos del período para detectar:
  - Cuentas de IVA Débito / IVA por Pagar (Pasivo con 'IVA')
  - Cuentas de IVA Crédito / IVA Soportado (Activo con 'IVA')
  - Cuentas de Ingresos (Ventas gravadas/exentas)
  - Cuentas de Costos y Gastos (Compras gravadas/exentas)
  """
  resumen = {
      t: {
          "Tarifa": t,
          "Tasa": tasa,
          "Base Ventas (₡)": 0.0,
          "IVA Débito (₡)": 0.0,
          "Base Compras (₡)": 0.0,
          "IVA Crédito (₡)": 0.0,
      }
      for t, tasa in TARIFAS_IVA_CR
  }
  ventas_exentas = 0.0
  compras_exentas = 0.0
  auxiliar_docs = []

  if df_mov.empty or df_cat.empty:
    return list(resumen.values()), 0.0, 0.0, pd.DataFrame()

  df = df_mov.copy()
  df["f_dt"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
  df_per = df[
      (df["asiento_id"] > 0) & (df["f_dt"] >= f_ini) & (df["f_dt"] <= f_fin)
  ].copy()

  if df_per.empty:
    return list(resumen.values()), 0.0, 0.0, pd.DataFrame()

  cat_map = df_cat.set_index("id").to_dict(orient="index")

  for aid, grp in df_per.groupby("asiento_id"):
    iva_deb_asiento = 0.0
    iva_cred_asiento = 0.0
    ventas_asiento = 0.0
    compras_asiento = 0.0
    tarifa_detectada = "13%"

    first_row = grp.iloc[0]
    fecha_str = _fmt_fecha_dmy(first_row["fecha"])
    num_as = int(first_row["num_asiento"])
    comp = str(first_row["num_comprobante"] or "")
    conc = str(first_row["concepto"] or "")

    for _, row in grp.iterrows():
      cid = int(row["cuenta_id"])
      cta_info = cat_map.get(cid, {})
      cod = str(cta_info.get("codigo", ""))
      nom = str(cta_info.get("nombre", "")).upper()
      tipo = str(cta_info.get("tipo", "")).upper()
      debe = float(row["debe"] or 0.0)
      haber = float(row["haber"] or 0.0)

      es_cuenta_iva = (
          "IVA" in nom
          or "IMPUESTO AL VALOR AGREGADO" in nom
          or "IMPUESTO VENTAS" in nom
      )

      if es_cuenta_iva:
        t_cta = _detectar_tarifa_texto(nom, conc)
        if t_cta != "13%":
          tarifa_detectada = t_cta
        if tipo == "PASIVO" or (haber > 0 and "POR PAGAR" in nom) or (haber > 0 and "DEBITO" in nom) or (haber > 0 and "DÉBITO" in nom):
          iva_deb_asiento += haber - debe
        elif tipo == "ACTIVO" or (debe > 0 and "FAVOR" in nom) or (debe > 0 and "CREDITO" in nom) or (debe > 0 and "CRÉDITO" in nom) or (debe > 0 and "SOPORTADO" in nom):
          iva_cred_asiento += debe - haber
        else:
          if haber > debe:
            iva_deb_asiento += haber - debe
          else:
            iva_cred_asiento += debe - haber
      elif tipo == "INGRESO" or cod.startswith("4"):
        ventas_asiento += haber - debe
        t_ing = _detectar_tarifa_texto(nom, conc)
        if t_ing != "13%":
          tarifa_detectada = t_ing
      elif tipo in ("COSTO", "GASTO") or cod.startswith(("5", "6")):
        compras_asiento += debe - haber
        t_gas = _detectar_tarifa_texto(nom, conc)
        if t_gas != "13%":
          tarifa_detectada = t_gas

    # Clasificar resultados del asiento
    if round(iva_deb_asiento, 2) > 0:
      t_use = tarifa_detectada if tarifa_detectada != "Exento / 0%" else "13%"
      tasa_u = dict(TARIFAS_IVA_CR)[t_use]
      base_v = (
          ventas_asiento
          if round(ventas_asiento, 2) > 0
          else (round(iva_deb_asiento / tasa_u, 2) if tasa_u > 0 else 0.0)
      )
      resumen[t_use]["Base Ventas (₡)"] += base_v
      resumen[t_use]["IVA Débito (₡)"] += iva_deb_asiento
      auxiliar_docs.append({
          "Fecha": fecha_str,
          "Asiento": f"#{num_as}",
          "Comprobante": comp,
          "Tipo Operación": f"Venta Gravada ({t_use})",
          "Concepto": conc,
          "Base Imponible (₡)": base_v,
          "IVA (₡)": iva_deb_asiento,
      })
    elif round(ventas_asiento, 2) > 0:
      ventas_exentas += ventas_asiento
      auxiliar_docs.append({
          "Fecha": fecha_str,
          "Asiento": f"#{num_as}",
          "Comprobante": comp,
          "Tipo Operación": "Venta Exenta / 0%",
          "Concepto": conc,
          "Base Imponible (₡)": ventas_asiento,
          "IVA (₡)": 0.0,
      })

    if round(iva_cred_asiento, 2) > 0:
      t_use = tarifa_detectada if tarifa_detectada != "Exento / 0%" else "13%"
      tasa_u = dict(TARIFAS_IVA_CR)[t_use]
      base_c = (
          compras_asiento
          if round(compras_asiento, 2) > 0
          else (round(iva_cred_asiento / tasa_u, 2) if tasa_u > 0 else 0.0)
      )
      resumen[t_use]["Base Compras (₡)"] += base_c
      resumen[t_use]["IVA Crédito (₡)"] += iva_cred_asiento
      auxiliar_docs.append({
          "Fecha": fecha_str,
          "Asiento": f"#{num_as}",
          "Comprobante": comp,
          "Tipo Operación": f"Compra / Gasto Gravado ({t_use})",
          "Concepto": conc,
          "Base Imponible (₡)": base_c,
          "IVA (₡)": iva_cred_asiento,
      })
    elif round(compras_asiento, 2) > 0:
      compras_exentas += compras_asiento

  resumen["Exento / 0%"]["Base Ventas (₡)"] = ventas_exentas
  resumen["Exento / 0%"]["Base Compras (₡)"] = compras_exentas

  df_aux = pd.DataFrame(auxiliar_docs)
  return list(resumen.values()), ventas_exentas, compras_exentas, df_aux


def generar_word_iva(datos, empresa, cedula, rep_legal):
  doc = docx.Document()
  for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_tit.paragraph_format.space_after = Pt(6)
  r0 = p_tit.add_run(
      "DECLARACIÓN Y LIQUIDACIÓN MENSUAL DEL IMPUESTO AL VALOR AGREGADO (IVA)\n"
  )
  r0.bold = True
  r0.font.size = Pt(12)
  r1 = p_tit.add_run("ESQUEMA OFICIAL MINISTERIO DE HACIENDA - COSTA RICA (D-104 / TRIBU-CR)\n")
  r1.bold = True
  r1.font.size = Pt(9.5)
  r2 = p_tit.add_run(f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}\n")
  r2.bold = True
  r2.font.size = Pt(10)
  r3 = p_tit.add_run(
      f"Período Fiscal: {datos['mes_texto']} (Del {datos['f_desde_str']} al"
      f" {datos['f_hasta_str']}) | Expresado en colones"
  )
  r3.font.size = Pt(8.5)

  # Tabla 1: Desglose por tarifa
  p_s1 = doc.add_paragraph()
  p_s1.paragraph_format.space_after = Pt(3)
  rs1 = p_s1.add_run("I. DESGLOSE DE OPERACIONES POR TARIFA DE IVA")
  rs1.bold = True
  rs1.font.size = Pt(9.5)

  t1 = doc.add_table(rows=1, cols=5)
  t1.alignment = WD_TABLE_ALIGNMENT.CENTER
  t1.autofit = False
  w1 = [Inches(1.3), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)]
  hdr1 = [
      "Tarifa IVA",
      "Ventas Netas (₡)",
      "IVA Débito (₡)",
      "Compras Netas (₡)",
      "IVA Crédito (₡)",
  ]
  for i, h in enumerate(hdr1):
    t1.rows[0].cells[i].width = w1[i]
    t1.rows[0].cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
    )
    p = t1.rows[0].cells[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(h)
    r.bold = True
    r.font.size = Pt(8)

  for row in datos["tabla_tarifas"]:
    rc = t1.add_row().cells
    vals = [
        row["Tarifa"],
        _fmt_moneda(row["Base Ventas (₡)"]),
        _fmt_moneda(row["IVA Débito (₡)"]),
        _fmt_moneda(row["Base Compras (₡)"]),
        _fmt_moneda(row["IVA Crédito (₡)"]),
    ]
    for i, v in enumerate(vals):
      rc[i].width = w1[i]
      p = rc[i].paragraphs[0]
      p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
      r = p.add_run(str(v))
      r.font.size = Pt(8)

  # Fila totales tabla 1
  rc_tot = t1.add_row().cells
  tot_vals = [
      "TOTALES DEL MES",
      _fmt_moneda(datos["total_ventas_netas"]),
      _fmt_moneda(datos["iva_debito"]),
      _fmt_moneda(datos["total_compras_netas"]),
      _fmt_moneda(datos["iva_credito"]),
  ]
  for i, v in enumerate(tot_vals):
    rc_tot[i].width = w1[i]
    rc_tot[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F8"/>')
    )
    p = rc_tot[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(str(v))
    r.bold = True
    r.font.size = Pt(8)

  # Tabla 2: Liquidación Oficial D-104
  p_s2 = doc.add_paragraph()
  p_s2.paragraph_format.space_before = Pt(8)
  p_s2.paragraph_format.space_after = Pt(3)
  rs2 = p_s2.add_run("II. LIQUIDACIÓN DEL IMPUESTO AL VALOR AGREGADO (D-104)")
  rs2.bold = True
  rs2.font.size = Pt(9.5)

  t2 = doc.add_table(rows=1, cols=2)
  t2.alignment = WD_TABLE_ALIGNMENT.CENTER
  t2.autofit = False
  w2 = [Inches(5.3), Inches(2.0)]
  t2.rows[0].cells[0].width = w2[0]
  t2.rows[0].cells[1].width = w2[1]
  for c in t2.rows[0].cells:
    c._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
    )
  t2.rows[0].cells[0].paragraphs[0].add_run("CASILLA / CONCEPTO TRIBUTARIO").bold = True
  ph = t2.rows[0].cells[1].paragraphs[0]
  ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  ph.add_run("MONTO (₡)").bold = True

  filas_liq = [
      ("Ventas netas gravadas del período", datos["ventas_netas_gravadas"], "NORMAL"),
      ("Ventas netas exentas / no sujetas", datos["ventas_exentas"], "NORMAL"),
      ("Total ventas netas del período", datos["total_ventas_netas"], "SUBTOTAL"),
      ("Total IVA Débito Fiscal (Impuesto generado en ventas)", datos["iva_debito"], "SUBTOTAL"),
      ("Compras netas gravadas con derecho a crédito", datos["compras_netas_gravadas"], "NORMAL"),
      ("Compras netas exentas / sin derecho a crédito", datos["compras_exentas"], "NORMAL"),
      ("(-) Total IVA Crédito Fiscal Aplicable (según prorrata)", datos["iva_credito_aplicable"], "SUBTOTAL"),
      ("Impuesto Neto del Período (Débito - Crédito)", datos["impuesto_neto_mes"], "SUBTOTAL"),
      ("(-) Saldo a favor de períodos anteriores", datos["saldo_favor_anterior"], "NORMAL"),
      ("(-) Retenciones de IVA (Tarjetas / Datáfonos / Agentes)", datos["retenciones_iva"], "NORMAL"),
      ("IMPUESTO FINAL POR PAGAR A HACIENDA", datos["impuesto_por_pagar"], "TOTAL"),
      ("SALDO A FAVOR PARA EL PRÓXIMO PERÍODO", datos["saldo_favor_siguiente"], "TOTAL"),
  ]

  for desc, m_val, tipo in filas_liq:
    rc = t2.add_row().cells
    rc[0].width = w2[0]
    rc[1].width = w2[1]
    p0 = rc[0].paragraphs[0]
    p1 = rc[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    es_b = tipo in ("SUBTOTAL", "TOTAL")
    r0_c = p0.add_run(desc)
    r0_c.bold = es_b
    r0_c.font.size = Pt(8.5)
    r1_c = p1.add_run(_fmt_moneda(m_val))
    r1_c.bold = es_b
    r1_c.font.size = Pt(8.5)
    if es_b:
      for c in rc:
        c._tc.get_or_add_tcPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F8"/>')
        )

  if datos["observaciones"].strip():
    p_obs = doc.add_paragraph()
    p_obs.paragraph_format.space_before = Pt(8)
    ro1 = p_obs.add_run("NOTAS Y OBSERVACIONES TRIBUTARIAS:\n")
    ro1.bold = True
    ro1.font.size = Pt(9)
    ro2 = p_obs.add_run(datos["observaciones"].strip())
    ro2.font.size = Pt(8.5)

  doc.add_paragraph("\n")
  ft = doc.add_table(rows=1, cols=2)
  ft.alignment = WD_TABLE_ALIGNMENT.CENTER
  ft.autofit = False
  ft.rows[0].cells[0].width = Inches(3.65)
  ft.rows[0].cells[1].width = Inches(3.65)

  pi = ft.rows[0].cells[0].paragraphs[0]
  ri = pi.add_run(
      "\n____________________________________\nElaborado por:\nLic. Roy F."
      " Esquivel Gonzalez\nContador Privado Carné 8788"
  )
  ri.bold = True
  ri.font.size = Pt(8.5)

  pd_c = ft.rows[0].cells[1].paragraphs[0]
  pd_c.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  rd = pd_c.add_run(
      f"\n____________________________________\nDeclarante / Representante Legal:\n{rep_legal}\nCédula Jurídica {cedula}"
  )
  rd.bold = True
  rd.font.size = Pt(8.5)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_iva(datos, empresa, cedula, rep_legal):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Liquidación IVA D-104"

  ws["A1"] = "DECLARACIÓN Y LIQUIDACIÓN MENSUAL DEL IVA (COSTA RICA - D-104)"
  ws["A1"].font = Font(bold=True, size=13)
  ws["A2"] = f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}"
  ws["A2"].font = Font(bold=True, size=10)
  ws["A3"] = (
      f"Período Fiscal: {datos['mes_texto']} ({datos['f_desde_str']} al"
      f" {datos['f_hasta_str']})"
  )
  ws["A3"].font = Font(italic=True, size=10)

  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  fill_sub = PatternFill(
      start_color="F2F4F8", end_color="F2F4F8", fill_type="solid"
  )
  borde_tot = Border(top=Side(style="thin"), bottom=Side(style="double"))

  ws.cell(row=5, column=1, value="I. DESGLOSE POR TARIFA DE IVA").font = Font(
      bold=True, size=11
  )
  hdrs1 = [
      "Tarifa IVA",
      "Ventas Netas (₡)",
      "IVA Débito (₡)",
      "Compras Netas (₡)",
      "IVA Crédito (₡)",
  ]
  for idx, h in enumerate(hdrs1, start=1):
    c = ws.cell(row=6, column=idx, value=h)
    c.font = Font(bold=True)
    c.fill = fill_hdr
    if idx > 1:
      c.alignment = Alignment(horizontal="right")

  r = 7
  for row in datos["tabla_tarifas"]:
    ws.cell(row=r, column=1, value=row["Tarifa"])
    for idx_c, k in enumerate(
        ["Base Ventas (₡)", "IVA Débito (₡)", "Base Compras (₡)", "IVA Crédito (₡)"],
        start=2,
    ):
      cm = ws.cell(row=r, column=idx_c, value=float(row[k]))
      cm.number_format = "#,##0.00;(#,##0.00);\"-\""
    r += 1

  ws.cell(row=r, column=1, value="TOTALES DEL MES").font = Font(bold=True)
  ws.cell(row=r, column=1).fill = fill_sub
  for idx_c, val_t in enumerate(
      [
          datos["total_ventas_netas"],
          datos["iva_debito"],
          datos["total_compras_netas"],
          datos["iva_credito"],
      ],
      start=2,
  ):
    cm = ws.cell(row=r, column=idx_c, value=float(val_t))
    cm.font = Font(bold=True)
    cm.fill = fill_sub
    cm.border = borde_tot
    cm.number_format = "#,##0.00;(#,##0.00);\"-\""

  r += 2
  ws.cell(row=r, column=1, value="II. LIQUIDACIÓN OFICIAL D-104").font = Font(
      bold=True, size=11
  )
  r += 1
  ws.cell(row=r, column=1, value="CONCEPTO TRIBUTARIO").font = Font(bold=True)
  ws.cell(row=r, column=1).fill = fill_hdr
  ws.cell(row=r, column=2, value="MONTO (₡)").font = Font(bold=True)
  ws.cell(row=r, column=2).fill = fill_hdr
  ws.cell(row=r, column=2).alignment = Alignment(horizontal="right")
  r += 1

  filas_liq = [
      ("Ventas netas gravadas del período", datos["ventas_netas_gravadas"], "NORMAL"),
      ("Ventas netas exentas / no sujetas", datos["ventas_exentas"], "NORMAL"),
      ("Total ventas netas del período", datos["total_ventas_netas"], "SUBTOTAL"),
      ("Total IVA Débito Fiscal", datos["iva_debito"], "SUBTOTAL"),
      ("Compras netas gravadas", datos["compras_netas_gravadas"], "NORMAL"),
      ("Compras netas exentas", datos["compras_exentas"], "NORMAL"),
      ("(-) Total IVA Crédito Fiscal Aplicable", datos["iva_credito_aplicable"], "SUBTOTAL"),
      ("Impuesto Neto del Período (Débito - Crédito)", datos["impuesto_neto_mes"], "SUBTOTAL"),
      ("(-) Saldo a favor de períodos anteriores", datos["saldo_favor_anterior"], "NORMAL"),
      ("(-) Retenciones de IVA (Tarjetas / Datáfonos)", datos["retenciones_iva"], "NORMAL"),
      ("IMPUESTO FINAL POR PAGAR A HACIENDA", datos["impuesto_por_pagar"], "TOTAL"),
      ("SALDO A FAVOR PARA EL PRÓXIMO PERÍODO", datos["saldo_favor_siguiente"], "TOTAL"),
  ]

  for desc, m_val, tipo in filas_liq:
    c1 = ws.cell(row=r, column=1, value=desc)
    c2 = ws.cell(row=r, column=2, value=float(m_val))
    c2.number_format = "#,##0.00;(#,##0.00);\"-\""
    if tipo in ("SUBTOTAL", "TOTAL"):
      c1.font = Font(bold=True)
      c2.font = Font(bold=True)
      c1.fill = fill_sub
      c2.fill = fill_sub
    if tipo == "TOTAL":
      c1.border = borde_tot
      c2.border = borde_tot
    r += 1

  if datos["observaciones"].strip():
    r += 1
    ws.cell(row=r, column=1, value="OBSERVACIONES:").font = Font(bold=True)
    r += 1
    ws.cell(row=r, column=1, value=datos["observaciones"].strip())

  r += 3
  ws.cell(
      row=r,
      column=1,
      value=(
          "_______________________________\nElaborado por:\nLic. Roy F."
          " Esquivel Gonzalez\nContador Privado Carné 8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)

  ws.cell(
      row=r,
      column=4,
      value=(
          f"_______________________________\nRepresentante Legal:\n{rep_legal}\nCédula"
          f" {cedula}"
      ),
  ).font = Font(bold=True)
  ws.cell(row=r, column=4).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[r].height = 55

  for letra, ancho in zip(["A", "B", "C", "D", "E"], [52, 22, 22, 24, 22]):
    ws.column_dimensions[letra].width = ancho

  # Hoja 2: Auxiliar de Documentos
  if not datos["df_auxiliar"].empty:
    ws2 = wb.create_sheet("Auxiliar Ventas y Compras")
    cols_aux = list(datos["df_auxiliar"].columns)
    for idx, col_name in enumerate(cols_aux, start=1):
      ch = ws2.cell(row=1, column=idx, value=col_name)
      ch.font = Font(bold=True)
      ch.fill = fill_hdr
    for idx_r, row_a in datos["df_auxiliar"].iterrows():
      for idx_c, col_name in enumerate(cols_aux, start=1):
        val = row_a[col_name]
        ca = ws2.cell(row=idx_r + 2, column=idx_c, value=val)
        if "(₡)" in col_name:
          ca.number_format = "#,##0.00"
    for letra, ancho in zip(
        ["A", "B", "C", "D", "E", "F", "G"], [14, 12, 18, 28, 42, 20, 18]
    ):
      ws2.column_dimensions[letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_iva(datos, empresa, cedula, rep_legal):
  pdf = FPDF(orientation="P", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=12)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(
      0,
      5.5,
      "DECLARACION Y LIQUIDACION MENSUAL DEL IVA (COSTA RICA - D-104)",
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5,
      _limpiar_pdf(f"{empresa.upper()} | CEDULA JURIDICA {cedula}"),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "", 8.5)
  pdf.cell(
      0,
      4.5,
      _limpiar_pdf(
          f"Periodo Fiscal: {datos['mes_texto']} (Del {datos['f_desde_str']} al"
          f" {datos['f_hasta_str']}) | Expresado en colones"
      ),
      ln=True,
      align="C",
  )
  pdf.ln(3)

  # Sección I: Tarifas
  pdf.set_font("Helvetica", "B", 8.5)
  pdf.cell(0, 5.5, "I. DESGLOSE DE OPERACIONES POR TARIFA DE IVA", ln=True)
  w1 = [34, 39, 39, 39, 39]
  pdf.set_fill_color(225, 232, 245)
  pdf.set_font("Helvetica", "B", 7.5)
  hdrs1 = [
      "Tarifa IVA",
      "Ventas Netas",
      "IVA Debito",
      "Compras Netas",
      "IVA Credito",
  ]
  for i, h in enumerate(hdrs1):
    pdf.cell(w1[i], 5.5, h, border=1, align="R" if i > 0 else "L", fill=True)
  pdf.ln()

  pdf.set_font("Helvetica", "", 7.5)
  for row in datos["tabla_tarifas"]:
    pdf.cell(w1[0], 5, _limpiar_pdf(row["Tarifa"]), border="B")
    pdf.cell(
        w1[1],
        5,
        _limpiar_pdf(_fmt_moneda(row["Base Ventas (₡)"])),
        border="B",
        align="R",
    )
    pdf.cell(
        w1[2],
        5,
        _limpiar_pdf(_fmt_moneda(row["IVA Débito (₡)"])),
        border="B",
        align="R",
    )
    pdf.cell(
        w1[3],
        5,
        _limpiar_pdf(_fmt_moneda(row["Base Compras (₡)"])),
        border="B",
        align="R",
    )
    pdf.cell(
        w1[4],
        5,
        _limpiar_pdf(_fmt_moneda(row["IVA Crédito (₡)"])),
        border="B",
        align="R",
    )
    pdf.ln()

  pdf.set_font("Helvetica", "B", 7.5)
  pdf.set_fill_color(242, 245, 250)
  pdf.cell(w1[0], 5.5, "TOTALES", border="TB", fill=True)
  for i, v in enumerate(
      [
          datos["total_ventas_netas"],
          datos["iva_debito"],
          datos["total_compras_netas"],
          datos["iva_credito"],
      ],
      start=1,
  ):
    pdf.cell(
        w1[i],
        5.5,
        _limpiar_pdf(_fmt_moneda(v)),
        border="TB",
        align="R",
        fill=True,
    )
  pdf.ln(7)

  # Sección II: Liquidación D-104
  pdf.set_font("Helvetica", "B", 8.5)
  pdf.cell(
      0, 5.5, "II. LIQUIDACION DEL IMPUESTO AL VALOR AGREGADO (D-104)", ln=True
  )
  w2 = [142, 48]
  pdf.set_fill_color(225, 232, 245)
  pdf.set_font("Helvetica", "B", 8)
  pdf.cell(w2[0], 6, "CASILLA / CONCEPTO TRIBUTARIO", border=1, fill=True)
  pdf.cell(w2[1], 6, "MONTO (C)", border=1, align="R", fill=True)
  pdf.ln()

  filas_liq = [
      ("Ventas netas gravadas del periodo", datos["ventas_netas_gravadas"], "NORMAL"),
      ("Ventas netas exentas / no sujetas", datos["ventas_exentas"], "NORMAL"),
      ("Total ventas netas del periodo", datos["total_ventas_netas"], "SUBTOTAL"),
      ("Total IVA Debito Fiscal (Impuesto en ventas)", datos["iva_debito"], "SUBTOTAL"),
      ("Compras netas gravadas con derecho a credito", datos["compras_netas_gravadas"], "NORMAL"),
      ("Compras netas exentas / sin derecho a credito", datos["compras_exentas"], "NORMAL"),
      ("(-) Total IVA Credito Fiscal Aplicable", datos["iva_credito_aplicable"], "SUBTOTAL"),
      ("Impuesto Neto del Periodo (Debito - Credito)", datos["impuesto_neto_mes"], "SUBTOTAL"),
      ("(-) Saldo a favor de periodos anteriores", datos["saldo_favor_anterior"], "NORMAL"),
      ("(-) Retenciones de IVA (Tarjetas / Datafonos)", datos["retenciones_iva"], "NORMAL"),
      ("IMPUESTO FINAL POR PAGAR A HACIENDA", datos["impuesto_por_pagar"], "TOTAL"),
      ("SALDO A FAVOR PARA EL PROXIMO PERIODO", datos["saldo_favor_siguiente"], "TOTAL"),
  ]

  for desc, m_val, tipo in filas_liq:
    es_b = tipo in ("SUBTOTAL", "TOTAL")
    pdf.set_font("Helvetica", "B" if es_b else "", 7.8)
    if es_b:
      pdf.set_fill_color(242, 245, 250)
    pdf.cell(w2[0], 5.2, _limpiar_pdf(desc), border="B", fill=es_b)
    pdf.cell(
        w2[1],
        5.2,
        _limpiar_pdf(_fmt_moneda(m_val)),
        border="B",
        align="R",
        fill=es_b,
    )
    pdf.ln()

  if datos["observaciones"].strip():
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(0, 5, "OBSERVACIONES TRIBUTARIAS:", ln=True)
    pdf.set_font("Helvetica", "", 7.8)
    pdf.multi_cell(0, 4.2, _limpiar_pdf(datos["observaciones"].strip()))

  pdf.ln(12)
  if pdf.get_y() > 240:
    pdf.add_page()
  y_f = pdf.get_y()
  pdf.set_font("Helvetica", "B", 8)
  pdf.set_xy(12, y_f)
  pdf.multi_cell(
      85,
      4.3,
      _limpiar_pdf(
          "_______________________________\nElaborado por:\nLic. Roy F."
          " Esquivel Gonzalez\nContador Privado Carne 8788"
      ),
      align="L",
  )
  pdf.set_xy(115, y_f)
  pdf.multi_cell(
      85,
      4.3,
      _limpiar_pdf(
          f"_______________________________\nRepresentante Legal:\n{rep_legal}\nCedula"
          f" {cedula}"
      ),
      align="R",
  )
  return bytes(pdf.output())


def render_modulo_reporte_iva(
    conn,
    empresa_id,
    datos_emp,
    obtener_movimientos_fn,
    backup_fn,
):
  asegurar_tabla_iva(conn)
  st.header(
      f"Declaración Mensual de IVA - Hacienda Costa Rica (D-104) - {datos_emp['razon_social']}"
  )

  meses_nombres = [
      "Enero",
      "Febrero",
      "Marzo",
      "Abril",
      "Mayo",
      "Junio",
      "Julio",
      "Agosto",
      "Setiembre",
      "Octubre",
      "Noviembre",
      "Diciembre",
  ]

  st.markdown("### 1. Selección del Período Fiscal Mensual")
  c1, c2, c3 = st.columns([1, 1.3, 1.5])
  anio_sel = c1.number_input(
      "Año Fiscal:", min_value=2020, max_value=2050, value=date.today().year
  )
  mes_sel_nombre = c2.selectbox(
      "Mes Declarado:",
      meses_nombres,
      index=date.today().month - 1,
  )
  mes_num = meses_nombres.index(mes_sel_nombre) + 1
  ultimo_dia = calendar.monthrange(int(anio_sel), mes_num)[1]
  f_ini = date(int(anio_sel), mes_num, 1)
  f_fin = date(int(anio_sel), mes_num, ultimo_dia)
  mes_etiqueta = f"{mes_sel_nombre} {int(anio_sel)}"

  c3.info(
      f"📅 **Rango del período:**\nDel {f_ini.strftime('%d-%m-%Y')} al"
      f" {f_fin.strftime('%d-%m-%Y')}"
  )

  # Verificar si existe declaración guardada para este mes
  decl_guardada = pd.read_sql_query(
      "SELECT * FROM declaraciones_iva WHERE empresa_id = ? AND anio = ? AND"
      " mes = ? ORDER BY id DESC LIMIT 1",
      conn,
      params=(empresa_id, int(anio_sel), int(mes_num)),
  )

  # Buscar saldo a favor del mes anterior automáticamente si no hay guardado
  mes_ant = 12 if mes_num == 1 else mes_num - 1
  anio_ant = int(anio_sel) - 1 if mes_num == 1 else int(anio_sel)
  decl_mes_ant = pd.read_sql_query(
      "SELECT saldo_favor_siguiente FROM declaraciones_iva WHERE empresa_id = ?"
      " AND anio = ? AND mes = ? ORDER BY id DESC LIMIT 1",
      conn,
      params=(empresa_id, anio_ant, mes_ant),
  )
  saldo_favor_auto = (
      float(decl_mes_ant.iloc[0]["saldo_favor_siguiente"] or 0.0)
      if not decl_mes_ant.empty
      else 0.0
  )

  df_mov = obtener_movimientos_fn(conn, empresa_id)
  df_cat = pd.read_sql_query(
      "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
      conn,
      params=(empresa_id,),
  )

  tarifas_calc, v_ex_calc, c_ex_calc, df_auxiliar = (
      calcular_borrador_iva_mensual(df_mov, df_cat, f_ini, f_fin)
  )

  usar_guardado = False
  if not decl_guardada.empty:
    dg = decl_guardada.iloc[0]
    st.success(
        f"📂 Existe una declaración guardada para **{mes_etiqueta}**"
        f" (Actualizada: {dg['fecha_actualizacion']})."
    )
    usar_guardado = st.checkbox(
        "Cargar valores guardados anteriormente (desmarque para recalcular desde libros contables)",
        value=True,
        key=f"chk_iva_{empresa_id}_{anio_sel}_{mes_num}",
    )

  if usar_guardado and not decl_guardada.empty:
    dg = decl_guardada.iloc[0]
    try:
      tarifas_base = json.loads(dg["detalle_tarifas_json"] or "[]")
      if not tarifas_base:
        tarifas_base = tarifas_calc
    except Exception:
      tarifas_base = tarifas_calc
    saldo_fav_ant_def = float(dg["saldo_favor_anterior"] or 0.0)
    retenciones_def = float(dg["retenciones_iva"] or 0.0)
    obs_def = str(dg["observaciones"] or "")
  else:
    tarifas_base = tarifas_calc
    saldo_fav_ant_def = saldo_favor_auto
    retenciones_def = 0.0
    obs_def = ""

  st.divider()
  st.markdown(
      "### 2. Desglose de Ventas y Compras por Tarifa de IVA (Editable para Ajuste)"
  )
  st.caption(
      "El sistema lee automáticamente los asientos contables del mes. Puede"
      " ajustar directamente cualquier base imponible o cuota de IVA antes de"
      " guardar y descargar."
  )

  df_tarifas_ed = st.data_editor(
      pd.DataFrame(tarifas_base),
      use_container_width=True,
      hide_index=True,
      disabled=["Tarifa", "Tasa"],
      column_config={
          "Tasa": None,
          "Base Ventas (₡)": st.column_config.NumberColumn(
              min_value=0.0, step=1000.0, format="₡%.2f"
          ),
          "IVA Débito (₡)": st.column_config.NumberColumn(
              min_value=0.0, step=100.0, format="₡%.2f"
          ),
          "Base Compras (₡)": st.column_config.NumberColumn(
              min_value=0.0, step=1000.0, format="₡%.2f"
          ),
          "IVA Crédito (₡)": st.column_config.NumberColumn(
              min_value=0.0, step=100.0, format="₡%.2f"
          ),
      },
      key=f"ed_iva_{empresa_id}_{anio_sel}_{mes_num}_{usar_guardado}",
  )

  # Totales calculados desde la tabla editada
  ventas_gravadas = float(
      df_tarifas_ed[df_tarifas_ed["Tarifa"] != "Exento / 0%"][
          "Base Ventas (₡)"
      ].sum()
  )
  ventas_exentas = float(
      df_tarifas_ed[df_tarifas_ed["Tarifa"] == "Exento / 0%"][
          "Base Ventas (₡)"
      ].sum()
  )
  total_ventas_netas = ventas_gravadas + ventas_exentas
  iva_debito_tot = float(df_tarifas_ed["IVA Débito (₡)"].sum())

  compras_gravadas = float(
      df_tarifas_ed[df_tarifas_ed["Tarifa"] != "Exento / 0%"][
          "Base Compras (₡)"
      ].sum()
  )
  compras_exentas = float(
      df_tarifas_ed[df_tarifas_ed["Tarifa"] == "Exento / 0%"][
          "Base Compras (₡)"
      ].sum()
  )
  total_compras_netas = compras_gravadas + compras_exentas
  iva_credito_tot = float(df_tarifas_ed["IVA Crédito (₡)"].sum())

  st.divider()
  st.markdown("### 3. Créditos Fiscales, Prorrata y Retenciones (Casillas D-104)")
  p1, p2, p3 = st.columns(3)
  prorrata_pct = p1.number_input(
      "Porcentaje de Prorrata Aplicable (%):",
      min_value=0.0,
      max_value=100.0,
      value=100.0,
      step=1.0,
      help="100% si todas sus ventas están gravadas o con derecho pleno a crédito.",
      key=f"pro_iva_{empresa_id}_{anio_sel}_{mes_num}",
  )
  saldo_favor_ant = p2.number_input(
      "(-) Saldo a Favor Períodos Anteriores (₡):",
      min_value=0.0,
      value=saldo_fav_ant_def,
      step=1000.0,
      format="%.2f",
      key=f"sfa_iva_{empresa_id}_{anio_sel}_{mes_num}",
  )
  retenciones_iva = p3.number_input(
      "(-) Retenciones IVA (Datáfonos / Tarjetas) (₡):",
      min_value=0.0,
      value=retenciones_def,
      step=1000.0,
      format="%.2f",
      key=f"ret_iva_{empresa_id}_{anio_sel}_{mes_num}",
  )

  iva_credito_aplicable = round(iva_credito_tot * (prorrata_pct / 100.0), 2)
  impuesto_neto_mes = round(iva_debito_tot - iva_credito_aplicable, 2)
  resultado_liquidacion = round(
      impuesto_neto_mes - saldo_favor_ant - retenciones_iva, 2
  )

  if resultado_liquidacion > 0:
    impuesto_por_pagar = resultado_liquidacion
    saldo_favor_siguiente = 0.0
  else:
    impuesto_por_pagar = 0.0
    saldo_favor_siguiente = abs(resultado_liquidacion)

  st.divider()
  st.markdown("### 4. Resumen Oficial de Liquidación del Mes")
  m1, m2, m3, m4 = st.columns(4)
  m1.metric("IVA Débito (Ventas)", f"₡{iva_debito_tot:,.2f}")
  m2.metric("IVA Crédito Aplicable", f"₡{iva_credito_aplicable:,.2f}")
  m3.metric(
      "Impuesto por Pagar a Hacienda",
      f"₡{impuesto_por_pagar:,.2f}",
      delta="A PAGAR" if impuesto_por_pagar > 0 else None,
      delta_color="inverse",
  )
  m4.metric(
      "Saldo a Favor Próximo Mes",
      f"₡{saldo_favor_siguiente:,.2f}",
      delta="CRÉDITO FISCAL" if saldo_favor_siguiente > 0 else None,
  )

  obs_iva = st.text_area(
      "Observaciones Tributarias para el Reporte D-104:",
      value=obs_def,
      height=85,
      key=f"obs_iva_{empresa_id}_{anio_sel}_{mes_num}",
  )

  if st.button(
      f"💾 Guardar Declaración de IVA ({mes_etiqueta}) en Base de Datos",
      type="primary",
      use_container_width=True,
  ):
    cur = conn.cursor()
    ahora_str = datetime.now().strftime("%d-%m-%Y %H:%M")
    tarifas_json = df_tarifas_ed.to_json(orient="records", force_ascii=False)
    if not decl_guardada.empty:
      id_upd = int(decl_guardada.iloc[0]["id"])
      cur.execute(
          """
          UPDATE declaraciones_iva
          SET fecha_desde=?, fecha_hasta=?, ventas_netas_gravadas=?,
              ventas_exentas=?, iva_debito=?, compras_netas_gravadas=?,
              compras_exentas=?, iva_credito=?, saldo_favor_anterior=?,
              retenciones_iva=?, impuesto_por_pagar=?, saldo_favor_siguiente=?,
              detalle_tarifas_json=?, observaciones=?, fecha_actualizacion=?
          WHERE id=?
          """,
          (
              str(f_ini),
              str(f_fin),
              ventas_gravadas,
              ventas_exentas,
              iva_debito_tot,
              compras_gravadas,
              compras_exentas,
              iva_credito_aplicable,
              float(saldo_favor_ant),
              float(retenciones_iva),
              impuesto_por_pagar,
              saldo_favor_siguiente,
              tarifas_json,
              obs_iva,
              ahora_str,
              id_upd,
          ),
      )
    else:
      cur.execute(
          """
          INSERT INTO declaraciones_iva (
              empresa_id, anio, mes, fecha_desde, fecha_hasta,
              ventas_netas_gravadas, ventas_exentas, iva_debito,
              compras_netas_gravadas, compras_exentas, iva_credito,
              saldo_favor_anterior, retenciones_iva, impuesto_por_pagar,
              saldo_favor_siguiente, detalle_tarifas_json, observaciones,
              fecha_actualizacion
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          """,
          (
              empresa_id,
              int(anio_sel),
              int(mes_num),
              str(f_ini),
              str(f_fin),
              ventas_gravadas,
              ventas_exentas,
              iva_debito_tot,
              compras_gravadas,
              compras_exentas,
              iva_credito_aplicable,
              float(saldo_favor_ant),
              float(retenciones_iva),
              impuesto_por_pagar,
              saldo_favor_siguiente,
              tarifas_json,
              obs_iva,
              ahora_str,
          ),
      )
    conn.commit()
    backup_fn()
    st.success(f"✅ Declaración de IVA de {mes_etiqueta} guardada y respaldada.")

  st.divider()
  st.subheader(
      f"📥 Descargar Reporte de IVA ({mes_etiqueta}) en Word, Excel y PDF"
  )
  datos_export = {
      "mes_texto": mes_etiqueta,
      "f_desde_str": f_ini.strftime("%d-%m-%Y"),
      "f_hasta_str": f_fin.strftime("%d-%m-%Y"),
      "tabla_tarifas": df_tarifas_ed.to_dict(orient="records"),
      "ventas_netas_gravadas": ventas_gravadas,
      "ventas_exentas": ventas_exentas,
      "total_ventas_netas": total_ventas_netas,
      "iva_debito": iva_debito_tot,
      "compras_netas_gravadas": compras_gravadas,
      "compras_exentas": compras_exentas,
      "total_compras_netas": total_compras_netas,
      "iva_credito": iva_credito_tot,
      "iva_credito_aplicable": iva_credito_aplicable,
      "impuesto_neto_mes": impuesto_neto_mes,
      "saldo_favor_anterior": float(saldo_favor_ant),
      "retenciones_iva": float(retenciones_iva),
      "impuesto_por_pagar": impuesto_por_pagar,
      "saldo_favor_siguiente": saldo_favor_siguiente,
      "observaciones": obs_iva,
      "df_auxiliar": df_auxiliar,
  }

  empresa_nom = str(datos_emp["razon_social"] or "")
  cedula_nom = str(datos_emp["cedula"] or "")
  rep_nom = str(datos_emp["representante_legal"] or "Representante Legal")
  sufijo = f"{empresa_nom.replace(' ', '_')}_{mes_etiqueta.replace(' ', '_')}"

  b1, b2, b3 = st.columns(3)
  b1.download_button(
      "📄 Descargar Word (.docx)",
      data=generar_word_iva(datos_export, empresa_nom, cedula_nom, rep_nom),
      file_name=f"Reporte_IVA_D104_{sufijo}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )
  b2.download_button(
      "📊 Descargar Excel (.xlsx)",
      data=generar_excel_iva(datos_export, empresa_nom, cedula_nom, rep_nom),
      file_name=f"Reporte_IVA_D104_{sufijo}.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
      use_container_width=True,
  )
  b3.download_button(
      "📕 Descargar PDF (.pdf)",
      data=generar_pdf_iva(datos_export, empresa_nom, cedula_nom, rep_nom),
      file_name=f"Reporte_IVA_D104_{sufijo}.pdf",
      mime="application/pdf",
      use_container_width=True,
  )

  if not df_auxiliar.empty:
    with st.expander("🔍 Ver Auxiliar de Asientos Gravados/Exentos del Período"):
      st.dataframe(
          df_auxiliar.style.format({
              "Base Imponible (₡)": "₡{:,.2f}",
              "IVA (₡)": "₡{:,.2f}",
          }),
          use_container_width=True,
          hide_index=True,
      )

  st.divider()
  st.subheader("🗂️ Historial de Declaraciones Mensuales de IVA Guardadas")
  df_hist_iva = pd.read_sql_query(
      """
      SELECT anio AS Año, mes AS Mes,
             ventas_netas_gravadas AS Ventas_Gravadas,
             iva_debito AS IVA_Débito,
             iva_credito AS IVA_Crédito,
             retenciones_iva AS Retenciones,
             impuesto_por_pagar AS Por_Pagar,
             saldo_favor_siguiente AS Saldo_Favor,
             fecha_actualizacion AS Guardado
      FROM declaraciones_iva
      WHERE empresa_id = ?
      ORDER BY anio DESC, mes DESC
      """,
      conn,
      params=(empresa_id,),
  )
  if df_hist_iva.empty:
    st.info("Aún no has guardado declaraciones mensuales de IVA para esta empresa.")
  else:
    st.dataframe(
        df_hist_iva.style.format({
            "Ventas_Gravadas": "₡{:,.2f}",
            "IVA_Débito": "₡{:,.2f}",
            "IVA_Crédito": "₡{:,.2f}",
            "Retenciones": "₡{:,.2f}",
            "Por_Pagar": "₡{:,.2f}",
            "Saldo_Favor": "₡{:,.2f}",
        }),
        use_container_width=True,
        hide_index=True,
    )
