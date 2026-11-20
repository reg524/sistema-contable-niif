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


def asegurar_tabla_conciliaciones(conn):
  cur = conn.cursor()
  cur.execute("""
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
    """)
  conn.commit()


def obtener_movimientos_cuenta_bancaria(
    df_mov, cuenta_id, f_ini, f_fin, calcular_saldos_fn
):
  saldos_ini = calcular_saldos_fn(
      df_mov, f_ini, incluir_asientos_del_dia=False
  )
  saldo_lib_ini = float(saldos_ini.get(cuenta_id, 0.0))

  saldos_fin = calcular_saldos_fn(
      df_mov, f_fin, incluir_asientos_del_dia=True
  )
  saldo_lib_fin = float(saldos_fin.get(cuenta_id, 0.0))

  if df_mov.empty:
    return saldo_lib_ini, pd.DataFrame(), saldo_lib_fin

  df = df_mov[df_mov["cuenta_id"] == cuenta_id].copy()
  df["f_dt"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
  df_per = df[
      (df["asiento_id"] > 0) & (df["f_dt"] >= f_ini) & (df["f_dt"] <= f_fin)
  ].sort_values(["f_dt", "num_asiento", "detalle_id"])

  return saldo_lib_ini, df_per, saldo_lib_fin


def generar_word_conciliacion(datos, empresa, cedula, rep_legal):
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
  r0 = p_tit.add_run("CONCILIACIÓN BANCARIA MENSUAL\n")
  r0.bold = True
  r0.font.size = Pt(13)
  r1 = p_tit.add_run(f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}\n")
  r1.bold = True
  r1.font.size = Pt(10)
  r2 = p_tit.add_run(f"Cuenta Contable / Bancaria: {datos['cuenta_etiqueta']}\n")
  r2.bold = True
  r2.font.size = Pt(9.5)
  r3 = p_tit.add_run(
      f"Período Conciliado: {datos['mes_texto']} (Del {datos['f_desde_str']} al"
      f" {datos['f_hasta_str']}) | Expresado en colones"
  )
  r3.font.size = Pt(8.5)

  t = doc.add_table(rows=1, cols=2)
  t.alignment = WD_TABLE_ALIGNMENT.CENTER
  t.autofit = False
  anchos = [Inches(5.3), Inches(2.0)]

  hdr = t.rows[0].cells
  hdr[0].width = anchos[0]
  hdr[1].width = anchos[1]
  for c in hdr:
    c._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
    )
  hdr[0].paragraphs[0].add_run("CONCEPTO DE CONCILIACIÓN").bold = True
  p_h1 = hdr[1].paragraphs[0]
  p_h1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_h1.add_run("MONTO (₡)").bold = True

  filas_cuadre = [
      ("SECCIÓN A: SALDO SEGÚN LIBROS CONTABLES", None, "TITULO"),
      (
          f"Saldo inicial en libros al {datos['f_desde_str']}",
          datos["saldo_libros_ini"],
          "NORMAL",
      ),
      (
          "(+) Total débitos / depósitos contabilizados en el período",
          datos["total_debe_libros"],
          "NORMAL",
      ),
      (
          "(-) Total créditos / cheques y pagos contabilizados",
          -datos["total_haber_libros"],
          "NORMAL",
      ),
      (
          f"Saldo según libros al {datos['f_hasta_str']}",
          datos["saldo_libros_fin"],
          "SUBTOTAL",
      ),
      (
          "(+) Notas de crédito / intereses bancarios no registrados",
          datos["nc_banco"],
          "NORMAL",
      ),
      (
          "(-) Notas de débito / comisiones bancarias no registradas",
          -datos["nd_banco"],
          "NORMAL",
      ),
      (
          "SALDO CONCILIADO SEGÚN LIBROS",
          datos["saldo_conc_libros"],
          "TOTAL",
      ),
      ("SECCIÓN B: SALDO SEGÚN ESTADO DE CUENTA BANCARIO", None, "TITULO"),
      (
          f"Saldo según estado de cuenta bancario al {datos['f_hasta_str']}",
          datos["saldo_banco"],
          "SUBTOTAL",
      ),
      (
          "(+) Depósitos en tránsito (en libros, pendientes en banco)",
          datos["dep_transito"],
          "NORMAL",
      ),
      (
          "(-) Cheques y pagos pendientes de cobro en banco",
          -datos["cheques_pendientes"],
          "NORMAL",
      ),
      (
          "SALDO CONCILIADO SEGÚN BANCO",
          datos["saldo_conc_banco"],
          "TOTAL",
      ),
      (
          "DIFERENCIA POR CONCILIAR (DEBE SER 0,00)",
          datos["diferencia"],
          "TOTAL",
      ),
  ]

  for desc, monto, tipo in filas_cuadre:
    rc = t.add_row().cells
    rc[0].width = anchos[0]
    rc[1].width = anchos[1]
    p0 = rc[0].paragraphs[0]
    p1 = rc[1].paragraphs[0]
    p0.paragraph_format.space_before = Pt(1.5)
    p0.paragraph_format.space_after = Pt(1.5)
    p1.paragraph_format.space_before = Pt(1.5)
    p1.paragraph_format.space_after = Pt(1.5)
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL")
    r_d = p0.add_run(desc)
    r_d.bold = es_b
    r_d.font.size = Pt(8.5)
    r_m = p1.add_run(_fmt_moneda(monto) if monto is not None else "")
    r_m.bold = es_b
    r_m.font.size = Pt(8.5)

    if tipo in ("SUBTOTAL", "TOTAL"):
      for c in rc:
        c._tc.get_or_add_tcPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F8"/>')
        )

  if datos["lista_pendientes"]:
    p_det = doc.add_paragraph()
    p_det.paragraph_format.space_before = Pt(8)
    p_det.paragraph_format.space_after = Pt(3)
    rd_t = p_det.add_run("DETALLE DE PARTIDAS CONCILIATORIAS Y EN TRÁNSITO")
    rd_t.bold = True
    rd_t.font.size = Pt(9.5)

    td = doc.add_table(rows=1, cols=5)
    td.alignment = WD_TABLE_ALIGNMENT.CENTER
    td.autofit = False
    w_d = [
        Inches(0.95),
        Inches(1.0),
        Inches(1.85),
        Inches(2.3),
        Inches(1.2),
    ]
    hd_titles = ["Fecha", "Ref / Doc", "Tipo Partida", "Detalle", "Monto (₡)"]
    for i, h in enumerate(hd_titles):
      td.rows[0].cells[i].width = w_d[i]
      td.rows[0].cells[i]._tc.get_or_add_tcPr().append(
          parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
      )
      p = td.rows[0].cells[i].paragraphs[0]
      p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i == 4 else WD_ALIGN_PARAGRAPH.LEFT
      r = p.add_run(h)
      r.bold = True
      r.font.size = Pt(8)

    for item in datos["lista_pendientes"]:
      rc = td.add_row().cells
      vals = [
          item["fecha"],
          item["ref"],
          item["tipo"],
          item["desc"],
          _fmt_moneda(item["monto"]),
      ]
      for i, v in enumerate(vals):
        rc[i].width = w_d[i]
        p = rc[i].paragraphs[0]
        p.alignment = (
            WD_ALIGN_PARAGRAPH.RIGHT if i == 4 else WD_ALIGN_PARAGRAPH.LEFT
        )
        r = p.add_run(str(v))
        r.font.size = Pt(7.8)

  if datos["observaciones"].strip():
    p_obs = doc.add_paragraph()
    p_obs.paragraph_format.space_before = Pt(8)
    ro1 = p_obs.add_run("OBSERVACIONES DE AUDITORÍA:\n")
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
      f"\n____________________________________\nRevisado y Aprobado por:\n{rep_legal}\nRepresentante"
      " Legal"
  )
  rd.bold = True
  rd.font.size = Pt(8.5)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_conciliacion(datos, empresa, cedula, rep_legal):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Conciliación Bancaria"

  ws["A1"] = "CONCILIACIÓN BANCARIA MENSUAL"
  ws["A1"].font = Font(bold=True, size=14)
  ws["A2"] = f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}"
  ws["A2"].font = Font(bold=True, size=10)
  ws["A3"] = (
      f"Cuenta: {datos['cuenta_etiqueta']} | Período: {datos['mes_texto']}"
      f" ({datos['f_desde_str']} al {datos['f_hasta_str']})"
  )
  ws["A3"].font = Font(italic=True, size=10)

  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  fill_sub = PatternFill(
      start_color="F2F4F8", end_color="F2F4F8", fill_type="solid"
  )
  borde_tot = Border(top=Side(style="thin"), bottom=Side(style="double"))

  ws.cell(row=5, column=1, value="CONCEPTO DE CONCILIACIÓN").font = Font(
      bold=True
  )
  ws.cell(row=5, column=1).fill = fill_hdr
  ws.cell(row=5, column=2, value="MONTO (₡)").font = Font(bold=True)
  ws.cell(row=5, column=2).fill = fill_hdr
  ws.cell(row=5, column=2).alignment = Alignment(horizontal="right")

  filas_cuadre = [
      ("SECCIÓN A: SALDO SEGÚN LIBROS CONTABLES", None, "TITULO"),
      (
          f"Saldo inicial en libros al {datos['f_desde_str']}",
          datos["saldo_libros_ini"],
          "NORMAL",
      ),
      (
          "(+) Total débitos / depósitos contabilizados",
          datos["total_debe_libros"],
          "NORMAL",
      ),
      (
          "(-) Total créditos / cheques y pagos contabilizados",
          -datos["total_haber_libros"],
          "NORMAL",
      ),
      (
          f"Saldo según libros al {datos['f_hasta_str']}",
          datos["saldo_libros_fin"],
          "SUBTOTAL",
      ),
      (
          "(+) Notas de crédito / intereses bancarios no registrados",
          datos["nc_banco"],
          "NORMAL",
      ),
      (
          "(-) Notas de débito / comisiones bancarias no registradas",
          -datos["nd_banco"],
          "NORMAL",
      ),
      (
          "SALDO CONCILIADO SEGÚN LIBROS",
          datos["saldo_conc_libros"],
          "TOTAL",
      ),
      ("SECCIÓN B: SALDO SEGÚN ESTADO DE CUENTA BANCARIO", None, "TITULO"),
      (
          f"Saldo según estado de cuenta bancario al {datos['f_hasta_str']}",
          datos["saldo_banco"],
          "SUBTOTAL",
      ),
      (
          "(+) Depósitos en tránsito",
          datos["dep_transito"],
          "NORMAL",
      ),
      (
          "(-) Cheques y pagos pendientes de cobro",
          -datos["cheques_pendientes"],
          "NORMAL",
      ),
      (
          "SALDO CONCILIADO SEGÚN BANCO",
          datos["saldo_conc_banco"],
          "TOTAL",
      ),
      (
          "DIFERENCIA POR CONCILIAR",
          datos["diferencia"],
          "TOTAL",
      ),
  ]

  r_act = 6
  for desc, monto, tipo in filas_cuadre:
    c1 = ws.cell(row=r_act, column=1, value=desc)
    c2 = ws.cell(row=r_act, column=2)
    if monto is not None:
      c2.value = float(monto)
      c2.number_format = "#,##0.00;(#,##0.00);\"-\""
    if tipo in ("TITULO", "SUBTOTAL", "TOTAL"):
      c1.font = Font(bold=True)
      c2.font = Font(bold=True)
    if tipo in ("SUBTOTAL", "TOTAL"):
      c1.fill = fill_sub
      c2.fill = fill_sub
    if tipo == "TOTAL":
      c1.border = borde_tot
      c2.border = borde_tot
    r_act += 1

  if datos["lista_pendientes"]:
    r_act += 1
    ws.cell(
        row=r_act,
        column=1,
        value="DETALLE DE PARTIDAS CONCILIATORIAS Y EN TRÁNSITO",
    ).font = Font(bold=True, size=11)
    r_act += 1
    hdrs = ["Fecha", "Ref / Doc", "Tipo Partida", "Detalle", "Monto (₡)"]
    for idx_h, h in enumerate(hdrs, start=1):
      ch = ws.cell(row=r_act, column=idx_h, value=h)
      ch.font = Font(bold=True)
      ch.fill = fill_hdr
    r_act += 1
    for item in datos["lista_pendientes"]:
      ws.cell(row=r_act, column=1, value=item["fecha"])
      ws.cell(row=r_act, column=2, value=item["ref"])
      ws.cell(row=r_act, column=3, value=item["tipo"])
      ws.cell(row=r_act, column=4, value=item["desc"])
      cm = ws.cell(row=r_act, column=5, value=float(item["monto"]))
      cm.number_format = "#,##0.00"
      r_act += 1

  if datos["observaciones"].strip():
    r_act += 1
    ws.cell(row=r_act, column=1, value="OBSERVACIONES:").font = Font(bold=True)
    r_act += 1
    ws.cell(row=r_act, column=1, value=datos["observaciones"].strip())

  r_act += 3
  ws.cell(
      row=r_act,
      column=1,
      value=(
          "_______________________________\nElaborado por:\nLic. Roy F."
          " Esquivel Gonzalez\nContador Privado Carné 8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=r_act, column=1).alignment = Alignment(wrap_text=True)

  ws.cell(
      row=r_act,
      column=4,
      value=(
          f"_______________________________\nRevisado y Aprobado por:\n{rep_legal}\nRepresentante"
          " Legal"
      ),
  ).font = Font(bold=True)
  ws.cell(row=r_act, column=4).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[r_act].height = 55

  for letra, ancho in zip(["A", "B", "C", "D", "E"], [52, 22, 28, 38, 18]):
    ws.column_dimensions[letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_conciliacion(datos, empresa, cedula, rep_legal):
  pdf = FPDF(orientation="P", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=12)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 12)
  pdf.cell(0, 6, "CONCILIACION BANCARIA MENSUAL", ln=True, align="C")
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5,
      _limpiar_pdf(f"{empresa.upper()} | CEDULA JURIDICA {cedula}"),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "B", 9)
  pdf.cell(
      0,
      5,
      _limpiar_pdf(f"Cuenta: {datos['cuenta_etiqueta']}"),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "", 8.5)
  pdf.cell(
      0,
      4.5,
      _limpiar_pdf(
          f"Periodo: {datos['mes_texto']} (Del {datos['f_desde_str']} al"
          f" {datos['f_hasta_str']}) | Expresado en colones"
      ),
      ln=True,
      align="C",
  )
  pdf.ln(3)

  w = [142, 48]
  pdf.set_fill_color(225, 232, 245)
  pdf.set_font("Helvetica", "B", 8)
  pdf.cell(w[0], 6, "CONCEPTO DE CONCILIACION", border=1, fill=True)
  pdf.cell(w[1], 6, "MONTO (C)", border=1, align="R", fill=True)
  pdf.ln()

  filas_cuadre = [
      ("SECCION A: SALDO SEGUN LIBROS CONTABLES", None, "TITULO"),
      (
          f"Saldo inicial en libros al {datos['f_desde_str']}",
          datos["saldo_libros_ini"],
          "NORMAL",
      ),
      (
          "(+) Total debitos / depositos contabilizados en el periodo",
          datos["total_debe_libros"],
          "NORMAL",
      ),
      (
          "(-) Total creditos / cheques y pagos contabilizados",
          -datos["total_haber_libros"],
          "NORMAL",
      ),
      (
          f"Saldo segun libros al {datos['f_hasta_str']}",
          datos["saldo_libros_fin"],
          "SUBTOTAL",
      ),
      (
          "(+) Notas de credito / intereses bancarios no registrados",
          datos["nc_banco"],
          "NORMAL",
      ),
      (
          "(-) Notas de debito / comisiones bancarias no registradas",
          -datos["nd_banco"],
          "NORMAL",
      ),
      (
          "SALDO CONCILIADO SEGUN LIBROS",
          datos["saldo_conc_libros"],
          "TOTAL",
      ),
      ("SECCION B: SALDO SEGUN ESTADO DE CUENTA BANCARIO", None, "TITULO"),
      (
          f"Saldo segun estado de cuenta bancario al {datos['f_hasta_str']}",
          datos["saldo_banco"],
          "SUBTOTAL",
      ),
      (
          "(+) Depositos en transito (en libros, pendientes en banco)",
          datos["dep_transito"],
          "NORMAL",
      ),
      (
          "(-) Cheques y pagos pendientes de cobro en banco",
          -datos["cheques_pendientes"],
          "NORMAL",
      ),
      (
          "SALDO CONCILIADO SEGUN BANCO",
          datos["saldo_conc_banco"],
          "TOTAL",
      ),
      (
          "DIFERENCIA POR CONCILIAR (DEBE SER 0,00)",
          datos["diferencia"],
          "TOTAL",
      ),
  ]

  for desc, monto, tipo in filas_cuadre:
    es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL")
    pdf.set_font("Helvetica", "B" if es_b else "", 7.8)
    fill_r = tipo in ("SUBTOTAL", "TOTAL")
    if fill_r:
      pdf.set_fill_color(242, 245, 250)
    pdf.cell(w[0], 5.2, _limpiar_pdf(desc), border="B", fill=fill_r)
    pdf.cell(
        w[1],
        5.2,
        _limpiar_pdf(_fmt_moneda(monto)) if monto is not None else "",
        border="B",
        align="R",
        fill=fill_r,
    )
    pdf.ln()

  if datos["lista_pendientes"]:
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(0, 5.5, "DETALLE DE PARTIDAS CONCILIATORIAS Y EN TRANSITO", ln=True)
    wd = [22, 26, 46, 66, 30]
    pdf.set_fill_color(225, 232, 245)
    pdf.set_font("Helvetica", "B", 7.2)
    for i, h in enumerate(["Fecha", "Ref / Doc", "Tipo", "Detalle", "Monto"]):
      pdf.cell(
          wd[i], 5.5, h, border=1, align="R" if i == 4 else "L", fill=True
      )
    pdf.ln()

    pdf.set_font("Helvetica", "", 7)
    for item in datos["lista_pendientes"]:
      pdf.cell(wd[0], 4.8, _limpiar_pdf(item["fecha"]), border="B")
      pdf.cell(wd[1], 4.8, _limpiar_pdf(item["ref"])[:15], border="B")
      pdf.cell(wd[2], 4.8, _limpiar_pdf(item["tipo"])[:28], border="B")
      pdf.cell(wd[3], 4.8, _limpiar_pdf(item["desc"])[:40], border="B")
      pdf.cell(
          wd[4],
          4.8,
          _limpiar_pdf(_fmt_moneda(item["monto"])),
          border="B",
          align="R",
      )
      pdf.ln()

  if datos["observaciones"].strip():
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(0, 5, "OBSERVACIONES DE AUDITORIA:", ln=True)
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
          f"_______________________________\nRevisado y Aprobado por:\n{rep_legal}\nRepresentante"
          " Legal"
      ),
      align="R",
  )
  return bytes(pdf.output())


def render_modulo_conciliacion_bancaria(
    conn,
    empresa_id,
    datos_emp,
    obtener_movimientos_fn,
    obtener_cuentas_fn,
    calcular_saldos_fn,
    backup_fn,
):
  asegurar_tabla_conciliaciones(conn)
  st.header(
      f"Conciliación Bancaria Mensual / Por Rango - {datos_emp['razon_social']}"
  )

  df_cat = pd.read_sql_query(
      "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
      conn,
      params=(empresa_id,),
  )
  cat_limpio = obtener_cuentas_fn(df_cat)
  cat_hojas = cat_limpio[cat_limpio["es_hoja_real"] == True].copy()
  cat_hojas["nom_up"] = cat_hojas["nombre"].astype(str).str.upper()

  ctas_banco = cat_hojas[
      cat_hojas["nom_up"].str.contains(
          "BANCO|CUENTA CORRIENTE|AHORRO|CAJA|EFECTIVO"
      )
      | cat_hojas["cod_str"].str.startswith(("1.1.1", "1.1.01", "101", "111"))
  ]
  if ctas_banco.empty:
    ctas_banco = cat_hojas[cat_hojas["cod_str"].str.startswith("1")]
  if ctas_banco.empty:
    ctas_banco = cat_hojas

  opciones_bco = [
      f"{r['codigo']} - {r['nombre']}" for _, r in ctas_banco.iterrows()
  ]
  mapa_bco_id = dict(zip(opciones_bco, ctas_banco["id"]))

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

  st.markdown("### 1. Selección de Cuenta Bancaria y Período (Mes a Mes o Rango)")
  col_b1, col_b2, col_b3, col_b4 = st.columns([2.2, 1.1, 0.9, 1.2])
  cta_bco_sel = col_b1.selectbox("Cuenta Bancaria Contable:", opciones_bco)
  cuenta_bco_id = int(mapa_bco_id[cta_bco_sel])

  modo_periodo = col_b2.radio(
      "Modo de Selección:", ["Mes a Mes", "Rango Personalizado"]
  )
  anio_sel = col_b3.number_input(
      "Año:", min_value=2020, max_value=2050, value=date.today().year
  )
  mes_sel_nombre = col_b4.selectbox(
      "Mes a Conciliar:",
      meses_nombres,
      index=date.today().month - 1,
      disabled=(modo_periodo != "Mes a Mes"),
  )
  mes_num = meses_nombres.index(mes_sel_nombre) + 1

  if modo_periodo == "Mes a Mes":
    ultimo_dia = calendar.monthrange(int(anio_sel), mes_num)[1]
    f_conc_ini = date(int(anio_sel), mes_num, 1)
    f_conc_fin = date(int(anio_sel), mes_num, ultimo_dia)
    mes_etiqueta = f"{mes_sel_nombre} {int(anio_sel)}"
  else:
    rc1, rc2 = st.columns(2)
    f_conc_ini = rc1.date_input(
        "Fecha Desde:",
        value=date(int(anio_sel), mes_num, 1),
        format="DD-MM-YYYY",
        key="cb_f1",
    )
    f_conc_fin = rc2.date_input(
        "Fecha Hasta:",
        value=date.today(),
        format="DD-MM-YYYY",
        key="cb_f2",
    )
    mes_etiqueta = (
        f"{meses_nombres[f_conc_fin.month - 1]} {f_conc_fin.year} (Rango)"
    )

  conc_guardada = pd.read_sql_query(
      "SELECT * FROM conciliaciones_bancarias WHERE empresa_id = ? AND"
      " cuenta_id = ? AND anio = ? AND mes = ? ORDER BY id DESC LIMIT 1",
      conn,
      params=(
          empresa_id,
          cuenta_bco_id,
          int(f_conc_fin.year),
          int(f_conc_fin.month),
      ),
  )

  saldo_bco_default = 0.0
  obs_default = ""
  conciliados_previos = {}
  partidas_bco_previas = [
      {
          "Fecha": f_conc_fin.strftime("%d-%m-%Y"),
          "Referencia": "",
          "Tipo": "Nota de Débito / Comisión Bancaria (-)",
          "Descripción": "Comisión mantenimiento / transferencias",
          "Monto": 0.0,
      }
  ]

  if not conc_guardada.empty:
    cg = conc_guardada.iloc[0]
    saldo_bco_default = float(cg["saldo_banco"] or 0.0)
    obs_default = str(cg["observaciones"] or "")
    try:
      conciliados_previos = json.loads(cg["detalle_libros_json"] or "{}")
    except Exception:
      conciliados_previos = {}
    try:
      pb_loaded = json.loads(cg["partidas_banco_json"] or "[]")
      if pb_loaded:
        partidas_bco_previas = pb_loaded
    except Exception:
      pass
    st.info(
        f"📂 Conciliación guardada cargada para **{mes_etiqueta}**"
        f" (Actualizada: {cg['fecha_actualizacion']})."
    )

  df_mov = obtener_movimientos_fn(conn, empresa_id)
  saldo_lib_ini, df_cta_mov, saldo_lib_fin = (
      obtener_movimientos_cuenta_bancaria(
          df_mov, cuenta_bco_id, f_conc_ini, f_conc_fin, calcular_saldos_fn
      )
  )

  st.markdown("### 2. Saldo según Estado de Cuenta Bancario al Cierre")
  sb1, sb2, sb3 = st.columns([1.3, 1, 1])
  saldo_banco_in = sb1.number_input(
      f"Saldo según Estado de Cuenta al {f_conc_fin.strftime('%d-%m-%Y')} (₡):",
      value=saldo_bco_default,
      step=1000.0,
      format="%.2f",
      key=f"sb_{empresa_id}_{cuenta_bco_id}_{f_conc_ini}_{f_conc_fin}",
  )
  sb2.metric(
      f"Saldo Inicial Libros ({f_conc_ini.strftime('%d-%m-%Y')})",
      f"₡{saldo_lib_ini:,.2f}",
  )
  sb3.metric(
      f"Saldo Final Libros ({f_conc_fin.strftime('%d-%m-%Y')})",
      f"₡{saldo_lib_fin:,.2f}",
  )

  st.divider()
  st.markdown(
      "### 3. Cotejo de Movimientos en Libros (Marque los que aparecen en el Estado de Cuenta)"
  )
  st.caption(
      "Los movimientos **NO marcados** se sumarán automáticamente como"
      " **Depósitos en Tránsito** (Debe) o **Cheques / Pagos Pendientes de"
      " Cobro** (Haber)."
  )

  dep_transito = 0.0
  cheques_pendientes = 0.0
  total_debe_libros = 0.0
  total_haber_libros = 0.0
  mapa_estado_libros = {}
  lista_pendientes_export = []

  if df_cta_mov.empty:
    st.warning(
        "No existen asientos contables en esta cuenta bancaria dentro del rango"
        " seleccionado."
    )
  else:
    total_debe_libros = float(df_cta_mov["debe"].sum())
    total_haber_libros = float(df_cta_mov["haber"].sum())

    filas_ed_lib = []
    for _, r in df_cta_mov.iterrows():
      det_id_str = str(int(r["detalle_id"]))
      marcado = bool(conciliados_previos.get(det_id_str, True))
      filas_ed_lib.append({
          "detalle_id": int(r["detalle_id"]),
          "Conciliado": marcado,
          "Asiento": f"#{int(r['num_asiento'])}",
          "Fecha": _fmt_fecha_dmy(r["fecha"]),
          "Comprobante": str(r["num_comprobante"] or ""),
          "Concepto": str(r["concepto"] or ""),
          "Depósito / Debe (₡)": float(r["debe"]),
          "Cheque / Haber (₡)": float(r["haber"]),
      })

    df_ed_lib = st.data_editor(
        pd.DataFrame(filas_ed_lib),
        use_container_width=True,
        hide_index=True,
        disabled=[
            "detalle_id",
            "Asiento",
            "Fecha",
            "Comprobante",
            "Concepto",
            "Depósito / Debe (₡)",
            "Cheque / Haber (₡)",
        ],
        column_config={
            "detalle_id": None,
            "Conciliado": st.column_config.CheckboxColumn(
                "Conciliado en Banco", default=True
            ),
            "Depósito / Debe (₡)": st.column_config.NumberColumn(
                format="₡%.2f"
            ),
            "Cheque / Haber (₡)": st.column_config.NumberColumn(
                format="₡%.2f"
            ),
        },
        key=f"ed_lib_{empresa_id}_{cuenta_bco_id}_{f_conc_ini}_{f_conc_fin}",
    )

    for _, rl in df_ed_lib.iterrows():
      did_s = str(int(rl["detalle_id"]))
      es_conc = bool(rl["Conciliado"])
      mapa_estado_libros[did_s] = es_conc
      if not es_conc:
        if rl["Depósito / Debe (₡)"] > 0:
          dep_transito += float(rl["Depósito / Debe (₡)"])
          lista_pendientes_export.append({
              "fecha": rl["Fecha"],
              "ref": rl["Comprobante"] or rl["Asiento"],
              "tipo": "(+) Depósito en Tránsito",
              "desc": rl["Concepto"],
              "monto": float(rl["Depósito / Debe (₡)"]),
          })
        if rl["Cheque / Haber (₡)"] > 0:
          cheques_pendientes += float(rl["Cheque / Haber (₡)"])
          lista_pendientes_export.append({
              "fecha": rl["Fecha"],
              "ref": rl["Comprobante"] or rl["Asiento"],
              "tipo": "(-) Cheque / Pago Pendiente",
              "desc": rl["Concepto"],
              "monto": float(rl["Cheque / Haber (₡)"]),
          })

  st.divider()
  st.markdown(
      "### 4. Partidas del Estado de Cuenta Bancario No Registradas en Libros"
  )
  df_pb_edit = st.data_editor(
      pd.DataFrame(partidas_bco_previas),
      use_container_width=True,
      hide_index=True,
      num_rows="dynamic",
      column_config={
          "Fecha": st.column_config.TextColumn("Fecha (DD-MM-AAAA)"),
          "Referencia": st.column_config.TextColumn("N° Documento / Ref."),
          "Tipo": st.column_config.SelectboxColumn(
              "Tipo de Ajuste",
              options=[
                  "Nota de Crédito / Intereses (+)",
                  "Nota de Débito / Comisión Bancaria (-)",
              ],
              required=True,
          ),
          "Descripción": st.column_config.TextColumn("Detalle / Concepto"),
          "Monto": st.column_config.NumberColumn(
              "Monto (₡)", min_value=0.0, step=100.0, format="₡%.2f"
          ),
      },
      key=f"ed_pb_{empresa_id}_{cuenta_bco_id}_{f_conc_ini}_{f_conc_fin}",
  )

  nc_banco = 0.0
  nd_banco = 0.0
  partidas_bco_guardar = []
  for _, rpb in df_pb_edit.iterrows():
    m_val = float(rpb.get("Monto") or 0.0)
    tipo_val = str(rpb.get("Tipo") or "")
    partidas_bco_guardar.append({
        "Fecha": str(rpb.get("Fecha") or ""),
        "Referencia": str(rpb.get("Referencia") or ""),
        "Tipo": tipo_val,
        "Descripción": str(rpb.get("Descripción") or ""),
        "Monto": m_val,
    })
    if m_val > 0:
      if "Crédito" in tipo_val or "(+)" in tipo_val:
        nc_banco += m_val
        etq_t = "(+) Nota Crédito Banco"
      else:
        nd_banco += m_val
        etq_t = "(-) Nota Débito / Comisión"
      lista_pendientes_export.append({
          "fecha": str(rpb.get("Fecha") or ""),
          "ref": str(rpb.get("Referencia") or "Banco"),
          "tipo": etq_t,
          "desc": str(rpb.get("Descripción") or ""),
          "monto": m_val,
      })

  saldo_conc_libros = saldo_lib_fin + nc_banco - nd_banco
  saldo_conc_banco = float(saldo_banco_in) + dep_transito - cheques_pendientes
  diferencia_conc = round(saldo_conc_libros - saldo_conc_banco, 2)

  st.divider()
  st.markdown("### 5. Resumen Oficial de Cuadre Bancario")
  km1, km2, km3 = st.columns(3)
  km1.metric("Saldo Conciliado según Libros", f"₡{saldo_conc_libros:,.2f}")
  km2.metric("Saldo Conciliado según Banco", f"₡{saldo_conc_banco:,.2f}")
  km3.metric(
      "Diferencia por Conciliar",
      f"₡{diferencia_conc:,.2f}",
      delta="CUADRADO" if abs(diferencia_conc) < 0.01 else "DESCUADRADO",
      delta_color="normal" if abs(diferencia_conc) < 0.01 else "inverse",
  )

  obs_conc = st.text_area(
      "Observaciones / Notas de Auditoría para el Informe:",
      value=obs_default,
      height=90,
      key=f"obs_cb_{empresa_id}_{cuenta_bco_id}_{f_conc_ini}_{f_conc_fin}",
  )

  if st.button(
      f"💾 Guardar Conciliación de {mes_etiqueta} en Base de Datos",
      type="primary",
      use_container_width=True,
  ):
    cur = conn.cursor()
    ahora_str = datetime.now().strftime("%d-%m-%Y %H:%M")
    if not conc_guardada.empty:
      id_upd = int(conc_guardada.iloc[0]["id"])
      cur.execute(
          """
          UPDATE conciliaciones_bancarias
          SET fecha_desde=?, fecha_hasta=?, saldo_banco=?, saldo_libros_ini=?,
              saldo_libros_fin=?, saldo_conciliado=?, diferencia=?,
              detalle_libros_json=?, partidas_banco_json=?, observaciones=?,
              fecha_actualizacion=?
          WHERE id=?
          """,
          (
              str(f_conc_ini),
              str(f_conc_fin),
              float(saldo_banco_in),
              float(saldo_lib_ini),
              float(saldo_lib_fin),
              float(saldo_conc_libros),
              float(diferencia_conc),
              json.dumps(mapa_estado_libros),
              json.dumps(partidas_bco_guardar),
              obs_conc,
              ahora_str,
              id_upd,
          ),
      )
    else:
      cur.execute(
          """
          INSERT INTO conciliaciones_bancarias (
              empresa_id, cuenta_id, anio, mes, fecha_desde, fecha_hasta,
              saldo_banco, saldo_libros_ini, saldo_libros_fin, saldo_conciliado,
              diferencia, detalle_libros_json, partidas_banco_json,
              observaciones, fecha_actualizacion
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          """,
          (
              empresa_id,
              cuenta_bco_id,
              int(f_conc_fin.year),
              int(f_conc_fin.month),
              str(f_conc_ini),
              str(f_conc_fin),
              float(saldo_banco_in),
              float(saldo_lib_ini),
              float(saldo_lib_fin),
              float(saldo_conc_libros),
              float(diferencia_conc),
              json.dumps(mapa_estado_libros),
              json.dumps(partidas_bco_guardar),
              obs_conc,
              ahora_str,
          ),
      )
    conn.commit()
    backup_fn()
    st.success(f"✅ Conciliación bancaria de {mes_etiqueta} guardada.")

  st.divider()
  st.subheader(
      f"📥 Descargar Conciliación Bancaria ({mes_etiqueta}) en Word, Excel y PDF"
  )
  datos_export = {
      "cuenta_etiqueta": cta_bco_sel,
      "f_desde_str": f_conc_ini.strftime("%d-%m-%Y"),
      "f_hasta_str": f_conc_fin.strftime("%d-%m-%Y"),
      "mes_texto": mes_etiqueta,
      "saldo_libros_ini": saldo_lib_ini,
      "total_debe_libros": total_debe_libros,
      "total_haber_libros": total_haber_libros,
      "saldo_libros_fin": saldo_lib_fin,
      "nc_banco": nc_banco,
      "nd_banco": nd_banco,
      "saldo_conc_libros": saldo_conc_libros,
      "saldo_banco": float(saldo_banco_in),
      "dep_transito": dep_transito,
      "cheques_pendientes": cheques_pendientes,
      "saldo_conc_banco": saldo_conc_banco,
      "diferencia": diferencia_conc,
      "lista_pendientes": lista_pendientes_export,
      "observaciones": obs_conc,
  }

  empresa_nom = str(datos_emp["razon_social"] or "")
  cedula_nom = str(datos_emp["cedula"] or "")
  rep_nom = str(datos_emp["representante_legal"] or "Representante Legal")
  sufijo = f"{empresa_nom.replace(' ', '_')}_{mes_etiqueta.replace(' ', '_')}"

  b1, b2, b3 = st.columns(3)
  b1.download_button(
      "📄 Descargar Word (.docx)",
      data=generar_word_conciliacion(
          datos_export, empresa_nom, cedula_nom, rep_nom
      ),
      file_name=f"Conciliacion_Bancaria_{sufijo}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )
  b2.download_button(
      "📊 Descargar Excel (.xlsx)",
      data=generar_excel_conciliacion(
          datos_export, empresa_nom, cedula_nom, rep_nom
      ),
      file_name=f"Conciliacion_Bancaria_{sufijo}.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
      use_container_width=True,
  )
  b3.download_button(
      "📕 Descargar PDF (.pdf)",
      data=generar_pdf_conciliacion(
          datos_export, empresa_nom, cedula_nom, rep_nom
      ),
      file_name=f"Conciliacion_Bancaria_{sufijo}.pdf",
      mime="application/pdf",
      use_container_width=True,
  )

  st.divider()
  st.subheader("🗂️ Historial de Conciliaciones Mensuales Guardadas")
  df_hist_cb = pd.read_sql_query(
      """
      SELECT cat.codigo || ' - ' || cat.nombre AS Cuenta,
             cb.anio AS Año, cb.mes AS Mes, cb.fecha_desde AS Desde,
             cb.fecha_hasta AS Hasta, cb.saldo_banco AS Saldo_Banco,
             cb.saldo_conciliado AS Saldo_Conciliado, cb.diferencia AS Diferencia,
             cb.fecha_actualizacion AS Guardado
      FROM conciliaciones_bancarias cb
      JOIN catalogo_cuentas cat ON cb.cuenta_id = cat.id
      WHERE cb.empresa_id = ?
      ORDER BY cb.anio DESC, cb.mes DESC
      """,
      conn,
      params=(empresa_id,),
  )
  if df_hist_cb.empty:
    st.info("Aún no has guardado conciliaciones mensuales para esta empresa.")
  else:
    df_hist_cb["Desde"] = df_hist_cb["Desde"].apply(_fmt_fecha_dmy)
    df_hist_cb["Hasta"] = df_hist_cb["Hasta"].apply(_fmt_fecha_dmy)
    st.dataframe(
        df_hist_cb.style.format({
            "Saldo_Banco": "₡{:,.2f}",
            "Saldo_Conciliado": "₡{:,.2f}",
            "Diferencia": "₡{:,.2f}",
        }),
        use_container_width=True,
        hide_index=True,
    )
