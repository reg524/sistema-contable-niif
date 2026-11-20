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


def asegurar_tabla_renta(conn):
  cur = conn.cursor()
  cur.execute("""
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
    """)
  conn.commit()


def calcular_impuesto_renta_cr(
    ingresos_brutos,
    renta_neta,
    regimen,
    tope_pyme=119174000.0,
    tramo1_jur=5621000.0,
    tramo2_jur=8433000.0,
    tramo3_jur=11243000.0,
    tramo1_fis=4181000.0,
    tramo2_fis=6244000.0,
    tramo3_fis=10414000.0,
    tramo4_fis=20872000.0,
):
  """
  Calcula el Impuesto sobre la Renta según Ley N° 7092 y Ley N° 9635 (Costa Rica):
  - Persona Jurídica: Si Ingresos Brutos <= tope_pyme aplica escala progresiva
    (5%, 10%, 15%, 20% sobre renta neta). Si supera el tope, aplica 30% pleno.
  - Persona Física Lucrativa: Escala progresiva sobre renta neta (0%, 10%, 15%, 20%, 25%).
  """
  detalle_tramos = []
  if renta_neta <= 0:
    return 0.0, detalle_tramos, "Sin renta neta imponible (pérdida o cero)"

  if "Jurídica" in regimen:
    if ingresos_brutos > tope_pyme:
      imp = round(renta_neta * 0.30, 2)
      detalle_tramos.append({
          "Tramo / Base": f"Tarifa Plena (Ingresos Brutos > ₡{tope_pyme:,.0f})",
          "Base Gravada (₡)": renta_neta,
          "Tarifa": "30%",
          "Impuesto (₡)": imp,
      })
      return imp, detalle_tramos, "Tarifa general 30% (Supera tope PYME)"
    else:
      imp_tot = 0.0
      t1 = min(renta_neta, tramo1_jur)
      i1 = round(t1 * 0.05, 2)
      imp_tot += i1
      detalle_tramos.append({
          "Tramo / Base": f"Hasta ₡{tramo1_jur:,.0f}",
          "Base Gravada (₡)": t1,
          "Tarifa": "5%",
          "Impuesto (₡)": i1,
      })
      if renta_neta > tramo1_jur:
        t2 = min(renta_neta, tramo2_jur) - tramo1_jur
        i2 = round(t2 * 0.10, 2)
        imp_tot += i2
        detalle_tramos.append({
            "Tramo / Base": f"Exceso de ₡{tramo1_jur:,.0f} hasta ₡{tramo2_jur:,.0f}",
            "Base Gravada (₡)": t2,
            "Tarifa": "10%",
            "Impuesto (₡)": i2,
        })
      if renta_neta > tramo2_jur:
        t3 = min(renta_neta, tramo3_jur) - tramo2_jur
        i3 = round(t3 * 0.15, 2)
        imp_tot += i3
        detalle_tramos.append({
            "Tramo / Base": f"Exceso de ₡{tramo2_jur:,.0f} hasta ₡{tramo3_jur:,.0f}",
            "Base Gravada (₡)": t3,
            "Tarifa": "15%",
            "Impuesto (₡)": i3,
        })
      if renta_neta > tramo3_jur:
        t4 = renta_neta - tramo3_jur
        i4 = round(t4 * 0.20, 2)
        imp_tot += i4
        detalle_tramos.append({
            "Tramo / Base": f"Exceso sobre ₡{tramo3_jur:,.0f}",
            "Base Gravada (₡)": t4,
            "Tarifa": "20%",
            "Impuesto (₡)": i4,
        })
      return (
          round(imp_tot, 2),
          detalle_tramos,
          "Escala progresiva PYME (5% - 20%)",
      )
  else:
    # Persona Física con Actividad Lucrativa
    imp_tot = 0.0
    t1 = min(renta_neta, tramo1_fis)
    detalle_tramos.append({
        "Tramo / Base": f"Exento hasta ₡{tramo1_fis:,.0f}",
        "Base Gravada (₡)": t1,
        "Tarifa": "0%",
        "Impuesto (₡)": 0.0,
    })
    if renta_neta > tramo1_fis:
      t2 = min(renta_neta, tramo2_fis) - tramo1_fis
      i2 = round(t2 * 0.10, 2)
      imp_tot += i2
      detalle_tramos.append({
          "Tramo / Base": f"De ₡{tramo1_fis:,.0f} a ₡{tramo2_fis:,.0f}",
          "Base Gravada (₡)": t2,
          "Tarifa": "10%",
          "Impuesto (₡)": i2,
      })
    if renta_neta > tramo2_fis:
      t3 = min(renta_neta, tramo3_fis) - tramo2_fis
      i3 = round(t3 * 0.15, 2)
      imp_tot += i3
      detalle_tramos.append({
          "Tramo / Base": f"De ₡{tramo2_fis:,.0f} a ₡{tramo3_fis:,.0f}",
          "Base Gravada (₡)": t3,
          "Tarifa": "15%",
          "Impuesto (₡)": i3,
      })
    if renta_neta > tramo3_fis:
      t4 = min(renta_neta, tramo4_fis) - tramo3_fis
      i4 = round(t4 * 0.20, 2)
      imp_tot += i4
      detalle_tramos.append({
          "Tramo / Base": f"De ₡{tramo3_fis:,.0f} a ₡{tramo4_fis:,.0f}",
          "Base Gravada (₡)": t4,
          "Tarifa": "20%",
          "Impuesto (₡)": i4,
      })
    if renta_neta > tramo4_fis:
      t5 = renta_neta - tramo4_fis
      i5 = round(t5 * 0.25, 2)
      imp_tot += i5
      detalle_tramos.append({
          "Tramo / Base": f"Exceso sobre ₡{tramo4_fis:,.0f}",
          "Base Gravada (₡)": t5,
          "Tarifa": "25%",
          "Impuesto (₡)": i5,
      })
    return (
        round(imp_tot, 2),
        detalle_tramos,
        "Escala progresiva Persona Física (0% - 25%)",
    )


def obtener_resumen_contable_renta(df_mov, df_cat, f_ini, f_fin):
  ingresos = 0.0
  costos = 0.0
  gastos_op = 0.0
  gastos_fin = 0.0
  retenciones_detectadas = 0.0

  if df_mov.empty or df_cat.empty:
    return {
        "ingresos": 0.0,
        "costos": 0.0,
        "gastos_op": 0.0,
        "gastos_fin": 0.0,
        "retenciones_2pct": 0.0,
    }

  df = df_mov.copy()
  df["f_dt"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
  df_per = df[
      (df["asiento_id"] > 0) & (df["f_dt"] >= f_ini) & (df["f_dt"] <= f_fin)
  ].copy()

  if df_per.empty:
    return {
        "ingresos": 0.0,
        "costos": 0.0,
        "gastos_op": 0.0,
        "gastos_fin": 0.0,
        "retenciones_2pct": 0.0,
    }

  cat_map = df_cat.set_index("id").to_dict(orient="index")
  for _, row in df_per.iterrows():
    cid = int(row["cuenta_id"])
    cta = cat_map.get(cid, {})
    cod = str(cta.get("codigo", ""))
    nom = str(cta.get("nombre", "")).upper()
    tipo = str(cta.get("tipo", "")).upper()
    debe = float(row["debe"] or 0.0)
    haber = float(row["haber"] or 0.0)

    if tipo == "INGRESO" or cod.startswith("4"):
      ingresos += haber - debe
    elif tipo == "COSTO" or cod.startswith("5"):
      costos += debe - haber
    elif tipo == "GASTO" or cod.startswith("6"):
      if "FINANCIER" in nom or "INTERES" in nom or "COMISION BANC" in nom:
        gastos_fin += debe - haber
      else:
        gastos_op += debe - haber
    elif tipo == "ACTIVO" and ("RETENCION" in nom or "RETENCIÓN" in nom) and ("2%" in nom or "RENTA" in nom):
      retenciones_detectadas += debe - haber

  return {
      "ingresos": round(ingresos, 2),
      "costos": round(costos, 2),
      "gastos_op": round(gastos_op, 2),
      "gastos_fin": round(gastos_fin, 2),
      "retenciones_2pct": round(max(0.0, retenciones_detectadas), 2),
  }


def generar_word_renta(datos, empresa, cedula, rep_legal):
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
      "DECLARACIÓN JURADA Y LIQUIDACIÓN DEL IMPUESTO SOBRE LA RENTA\n"
  )
  r0.bold = True
  r0.font.size = Pt(12)
  r1 = p_tit.add_run(
      "ESQUEMA OFICIAL MINISTERIO DE HACIENDA - COSTA RICA (D-101 / TRIBU-CR)\n"
  )
  r1.bold = True
  r1.font.size = Pt(9.5)
  r2 = p_tit.add_run(f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}\n")
  r2.bold = True
  r2.font.size = Pt(10)
  r3 = p_tit.add_run(
      f"Período Fiscal {datos['anio_fiscal']} (Del {datos['f_desde_str']} al"
      f" {datos['f_hasta_str']}) | {datos['regimen']} | Expresado en colones"
  )
  r3.font.size = Pt(8.5)

  # Tabla 1: Conciliación Fiscal y Liquidación D-101
  t = doc.add_table(rows=1, cols=2)
  t.alignment = WD_TABLE_ALIGNMENT.CENTER
  t.autofit = False
  w = [Inches(5.3), Inches(2.0)]
  t.rows[0].cells[0].width = w[0]
  t.rows[0].cells[1].width = w[1]
  for c in t.rows[0].cells:
    c._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
    )
  t.rows[0].cells[0].paragraphs[0].add_run("CASILLA / CONCEPTO D-101").bold = True
  ph = t.rows[0].cells[1].paragraphs[0]
  ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  ph.add_run("MONTO (₡)").bold = True

  filas = [
      ("I. DETERMINACIÓN DE LA UTILIDAD CONTABLE", None, "TITULO"),
      ("Total ingresos brutos gravables y no gravables", datos["ingresos_brutos"], "NORMAL"),
      ("(-) Costo de ventas y servicios prestados", -datos["costos"], "NORMAL"),
      ("Utilidad bruta en operaciones", datos["utilidad_bruta"], "SUBTOTAL"),
      ("(-) Gastos administrativos y de operación", -datos["gastos_op"], "NORMAL"),
      ("(-) Gastos financieros", -datos["gastos_fin"], "NORMAL"),
      ("Utilidad (o Pérdida) Contable antes de Impuesto", datos["utilidad_contable"], "SUBTOTAL"),
      ("II. CONCILIACIÓN FISCAL (ART. 8 Y 9 LEY N° 7092)", None, "TITULO"),
      ("(+) Gastos no deducibles fiscales", datos["gastos_no_deducibles"], "NORMAL"),
      ("(-) Ingresos no gravables / rentas exentas", -datos["ingresos_no_gravables"], "NORMAL"),
      ("RENTA NETA IMPONIBLE (BASE GRAVABLE)", datos["renta_neta"], "SUBTOTAL"),
      ("III. LIQUIDACIÓN DEL IMPUESTO Y CRÉDITOS", None, "TITULO"),
      (f"Impuesto sobre la renta determinado ({datos['nota_tarifa']})", datos["impuesto_determinado"], "SUBTOTAL"),
      ("(-) Pagos parciales realizados en el período", -datos["pagos_parciales"], "NORMAL"),
      ("(-) Retenciones en la fuente aplicadas (2% u otras)", -datos["retenciones_2pct"], "NORMAL"),
      ("(-) Otros créditos fiscales aplicables", -datos["creditos_fiscales"], "NORMAL"),
      ("IMPUESTO FINAL POR PAGAR A HACIENDA", datos["impuesto_por_pagar"], "TOTAL"),
      ("SALDO A FAVOR DEL CONTRIBUYENTE", datos["saldo_favor"], "TOTAL"),
  ]

  for desc, m_val, tipo in filas:
    rc = t.add_row().cells
    rc[0].width = w[0]
    rc[1].width = w[1]
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
    r_m = p1.add_run(_fmt_moneda(m_val) if m_val is not None else "")
    r_m.bold = es_b
    r_m.font.size = Pt(8.5)
    if tipo in ("SUBTOTAL", "TOTAL"):
      for c in rc:
        c._tc.get_or_add_tcPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F8"/>')
        )

  if datos["detalle_tramos"]:
    p_tr = doc.add_paragraph()
    p_tr.paragraph_format.space_before = Pt(8)
    p_tr.paragraph_format.space_after = Pt(3)
    rt = p_tr.add_run("IV. CÁLCULO POR TRAMOS DE RENTA APLICADOS")
    rt.bold = True
    rt.font.size = Pt(9)

    tt = doc.add_table(rows=1, cols=4)
    tt.alignment = WD_TABLE_ALIGNMENT.CENTER
    tt.autofit = False
    wt = [Inches(3.3), Inches(1.6), Inches(0.8), Inches(1.6)]
    for i, h in enumerate(["Tramo / Base", "Base Gravada (₡)", "Tarifa", "Impuesto (₡)"]):
      tt.rows[0].cells[i].width = wt[i]
      tt.rows[0].cells[i]._tc.get_or_add_tcPr().append(
          parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
      )
      p = tt.rows[0].cells[i].paragraphs[0]
      p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i in (1, 3) else WD_ALIGN_PARAGRAPH.LEFT
      r = p.add_run(h)
      r.bold = True
      r.font.size = Pt(8)
    for tr in datos["detalle_tramos"]:
      rc = tt.add_row().cells
      vals = [
          tr["Tramo / Base"],
          _fmt_moneda(tr["Base Gravada (₡)"]),
          tr["Tarifa"],
          _fmt_moneda(tr["Impuesto (₡)"]),
      ]
      for i, v in enumerate(vals):
        rc[i].width = wt[i]
        p = rc[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i in (1, 3) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(str(v))
        r.font.size = Pt(8)

  if datos["observaciones"].strip():
    p_obs = doc.add_paragraph()
    p_obs.paragraph_format.space_before = Pt(8)
    ro1 = p_obs.add_run("NOTAS DE CONCILIACIÓN FISCAL:\n")
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


def generar_excel_renta(datos, empresa, cedula, rep_legal):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Declaración Renta D-101"

  ws["A1"] = "DECLARACIÓN JURADA Y LIQUIDACIÓN DEL IMPUESTO SOBRE LA RENTA (D-101)"
  ws["A1"].font = Font(bold=True, size=13)
  ws["A2"] = f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}"
  ws["A2"].font = Font(bold=True, size=10)
  ws["A3"] = (
      f"Período Fiscal {datos['anio_fiscal']} ({datos['f_desde_str']} al"
      f" {datos['f_hasta_str']}) | {datos['regimen']}"
  )
  ws["A3"].font = Font(italic=True, size=10)

  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  fill_sub = PatternFill(
      start_color="F2F4F8", end_color="F2F4F8", fill_type="solid"
  )
  borde_tot = Border(top=Side(style="thin"), bottom=Side(style="double"))

  ws.cell(row=5, column=1, value="CASILLA / CONCEPTO D-101").font = Font(bold=True)
  ws.cell(row=5, column=1).fill = fill_hdr
  ws.cell(row=5, column=2, value="MONTO (₡)").font = Font(bold=True)
  ws.cell(row=5, column=2).fill = fill_hdr
  ws.cell(row=5, column=2).alignment = Alignment(horizontal="right")

  filas = [
      ("I. DETERMINACIÓN DE LA UTILIDAD CONTABLE", None, "TITULO"),
      ("Total ingresos brutos", datos["ingresos_brutos"], "NORMAL"),
      ("(-) Costo de ventas y servicios", -datos["costos"], "NORMAL"),
      ("Utilidad bruta en operaciones", datos["utilidad_bruta"], "SUBTOTAL"),
      ("(-) Gastos administrativos y de operación", -datos["gastos_op"], "NORMAL"),
      ("(-) Gastos financieros", -datos["gastos_fin"], "NORMAL"),
      ("Utilidad (o Pérdida) Contable antes de Impuesto", datos["utilidad_contable"], "SUBTOTAL"),
      ("II. CONCILIACIÓN FISCAL", None, "TITULO"),
      ("(+) Gastos no deducibles fiscales", datos["gastos_no_deducibles"], "NORMAL"),
      ("(-) Ingresos no gravables / exentos", -datos["ingresos_no_gravables"], "NORMAL"),
      ("RENTA NETA IMPONIBLE", datos["renta_neta"], "SUBTOTAL"),
      ("III. LIQUIDACIÓN DEL IMPUESTO Y CRÉDITOS", None, "TITULO"),
      (f"Impuesto determinado ({datos['nota_tarifa']})", datos["impuesto_determinado"], "SUBTOTAL"),
      ("(-) Pagos parciales del período", -datos["pagos_parciales"], "NORMAL"),
      ("(-) Retenciones en la fuente (2% u otras)", -datos["retenciones_2pct"], "NORMAL"),
      ("(-) Otros créditos fiscales", -datos["creditos_fiscales"], "NORMAL"),
      ("IMPUESTO FINAL POR PAGAR A HACIENDA", datos["impuesto_por_pagar"], "TOTAL"),
      ("SALDO A FAVOR DEL CONTRIBUYENTE", datos["saldo_favor"], "TOTAL"),
  ]

  r = 6
  for desc, m_val, tipo in filas:
    c1 = ws.cell(row=r, column=1, value=desc)
    c2 = ws.cell(row=r, column=2)
    if m_val is not None:
      c2.value = float(m_val)
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
    r += 1

  if datos["detalle_tramos"]:
    r += 1
    ws.cell(row=r, column=1, value="IV. CÁLCULO POR TRAMOS APLICADOS").font = Font(
        bold=True, size=11
    )
    r += 1
    for idx_h, h in enumerate(
        ["Tramo / Base", "Base Gravada (₡)", "Tarifa", "Impuesto (₡)"], start=1
    ):
      ch = ws.cell(row=r, column=idx_h, value=h)
      ch.font = Font(bold=True)
      ch.fill = fill_hdr
    r += 1
    for tr in datos["detalle_tramos"]:
      ws.cell(row=r, column=1, value=tr["Tramo / Base"])
      cb = ws.cell(row=r, column=2, value=float(tr["Base Gravada (₡)"]))
      cb.number_format = "#,##0.00"
      ws.cell(row=r, column=3, value=tr["Tarifa"])
      ci = ws.cell(row=r, column=4, value=float(tr["Impuesto (₡)"]))
      ci.number_format = "#,##0.00"
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

  for letra, ancho in zip(["A", "B", "C", "D"], [54, 22, 14, 22]):
    ws.column_dimensions[letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_renta(datos, empresa, cedula, rep_legal):
  pdf = FPDF(orientation="P", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=12)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(
      0,
      5.5,
      "DECLARACION JURADA Y LIQUIDACION DEL IMPUESTO SOBRE LA RENTA (D-101)",
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
          f"Periodo Fiscal {datos['anio_fiscal']} (Del {datos['f_desde_str']} al"
          f" {datos['f_hasta_str']}) | {datos['regimen']}"
      ),
      ln=True,
      align="C",
  )
  pdf.ln(3)

  w = [142, 48]
  pdf.set_fill_color(225, 232, 245)
  pdf.set_font("Helvetica", "B", 8)
  pdf.cell(w[0], 6, "CASILLA / CONCEPTO TRIBUTARIO D-101", border=1, fill=True)
  pdf.cell(w[1], 6, "MONTO (C)", border=1, align="R", fill=True)
  pdf.ln()

  filas = [
      ("I. DETERMINACION DE LA UTILIDAD CONTABLE", None, "TITULO"),
      ("Total ingresos brutos", datos["ingresos_brutos"], "NORMAL"),
      ("(-) Costo de ventas y servicios", -datos["costos"], "NORMAL"),
      ("Utilidad bruta en operaciones", datos["utilidad_bruta"], "SUBTOTAL"),
      ("(-) Gastos administrativos y de operacion", -datos["gastos_op"], "NORMAL"),
      ("(-) Gastos financieros", -datos["gastos_fin"], "NORMAL"),
      ("Utilidad (o Perdida) Contable antes de Impuesto", datos["utilidad_contable"], "SUBTOTAL"),
      ("II. CONCILIACION FISCAL (LEY N 7092)", None, "TITULO"),
      ("(+) Gastos no deducibles fiscales", datos["gastos_no_deducibles"], "NORMAL"),
      ("(-) Ingresos no gravables / exentos", -datos["ingresos_no_gravables"], "NORMAL"),
      ("RENTA NETA IMPONIBLE (BASE GRAVABLE)", datos["renta_neta"], "SUBTOTAL"),
      ("III. LIQUIDACION DEL IMPUESTO Y CREDITOS", None, "TITULO"),
      (f"Impuesto determinado ({datos['nota_tarifa']})", datos["impuesto_determinado"], "SUBTOTAL"),
      ("(-) Pagos parciales realizados en el periodo", -datos["pagos_parciales"], "NORMAL"),
      ("(-) Retenciones en la fuente aplicadas (2%)", -datos["retenciones_2pct"], "NORMAL"),
      ("(-) Otros creditos fiscales aplicables", -datos["creditos_fiscales"], "NORMAL"),
      ("IMPUESTO FINAL POR PAGAR A HACIENDA", datos["impuesto_por_pagar"], "TOTAL"),
      ("SALDO A FAVOR DEL CONTRIBUYENTE", datos["saldo_favor"], "TOTAL"),
  ]

  for desc, m_val, tipo in filas:
    es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL")
    pdf.set_font("Helvetica", "B" if es_b else "", 7.8)
    fill_r = tipo in ("SUBTOTAL", "TOTAL")
    if fill_r:
      pdf.set_fill_color(242, 245, 250)
    pdf.cell(w[0], 5.1, _limpiar_pdf(desc), border="B", fill=fill_r)
    pdf.cell(
        w[1],
        5.1,
        _limpiar_pdf(_fmt_moneda(m_val)) if m_val is not None else "",
        border="B",
        align="R",
        fill=fill_r,
    )
    pdf.ln()

  if datos["detalle_tramos"]:
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(0, 5.5, "IV. CALCULO POR TRAMOS DE RENTA APLICADOS", ln=True)
    wt = [95, 40, 18, 37]
    pdf.set_fill_color(225, 232, 245)
    pdf.set_font("Helvetica", "B", 7.3)
    for i, h in enumerate(["Tramo / Base", "Base Gravada", "Tarifa", "Impuesto"]):
      pdf.cell(
          wt[i],
          5.3,
          h,
          border=1,
          align="R" if i in (1, 3) else "L",
          fill=True,
      )
    pdf.ln()
    pdf.set_font("Helvetica", "", 7.2)
    for tr in datos["detalle_tramos"]:
      pdf.cell(wt[0], 4.8, _limpiar_pdf(tr["Tramo / Base"]), border="B")
      pdf.cell(
          wt[1],
          4.8,
          _limpiar_pdf(_fmt_moneda(tr["Base Gravada (₡)"])),
          border="B",
          align="R",
      )
      pdf.cell(wt[2], 4.8, _limpiar_pdf(tr["Tarifa"]), border="B")
      pdf.cell(
          wt[3],
          4.8,
          _limpiar_pdf(_fmt_moneda(tr["Impuesto (₡)"])),
          border="B",
          align="R",
      )
      pdf.ln()

  if datos["observaciones"].strip():
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(0, 5, "NOTAS DE CONCILIACION FISCAL:", ln=True)
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


def render_modulo_reporte_renta(
    conn,
    empresa_id,
    datos_emp,
    obtener_movimientos_fn,
    backup_fn,
):
  asegurar_tabla_renta(conn)
  st.header(
      f"Declaración Jurada Impuesto sobre la Renta (D-101) - {datos_emp['razon_social']}"
  )

  st.markdown("### 1. Período Fiscal y Régimen Tributario")
  c1, c2, c3, c4 = st.columns([0.9, 1.1, 1.1, 1.9])
  anio_fisc = c1.number_input(
      "Año Fiscal:", min_value=2020, max_value=2050, value=date.today().year
  )
  f_ini = c2.date_input(
      "Desde:",
      value=date(int(anio_fisc), 1, 1),
      format="DD-MM-YYYY",
      key=f"r_f1_{anio_fisc}",
  )
  f_fin = c3.date_input(
      "Hasta:",
      value=date(int(anio_fisc), 12, 31),
      format="DD-MM-YYYY",
      key=f"r_f2_{anio_fisc}",
  )
  regimen = c4.selectbox(
      "Régimen del Contribuyente:",
      [
          "Persona Jurídica (PYME Escalonada / 30% General)",
          "Persona Física con Actividad Lucrativa (Tramos 0% - 25%)",
      ],
  )

  decl_guardada = pd.read_sql_query(
      "SELECT * FROM declaraciones_renta WHERE empresa_id = ? AND anio_fiscal = ?"
      " ORDER BY id DESC LIMIT 1",
      conn,
      params=(empresa_id, int(anio_fisc)),
  )

  df_mov = obtener_movimientos_fn(conn, empresa_id)
  df_cat = pd.read_sql_query(
      "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
      conn,
      params=(empresa_id,),
  )
  res_libros = obtener_resumen_contable_renta(df_mov, df_cat, f_ini, f_fin)

  gnd_def = 0.0
  ing_ng_def = 0.0
  pp_def = 0.0
  ret_def = res_libros["retenciones_2pct"]
  cred_def = 0.0
  obs_def = ""
  ajustes_previos = [
      {
          "Tipo Ajuste": "(+) Gasto No Deducible",
          "Concepto / Referencia": "Multas, recargos o gastos sin respaldo electrónico",
          "Monto (₡)": 0.0,
      },
      {
          "Tipo Ajuste": "(-) Ingreso No Gravable",
          "Concepto / Referencia": "Dividendos recibidos / rentas sujetas a retención única",
          "Monto (₡)": 0.0,
      },
  ]

  if not decl_guardada.empty:
    dg = decl_guardada.iloc[0]
    st.info(
        f"📂 Declaración guardada cargada para el período fiscal **{int(anio_fisc)}**"
        f" (Actualizada: {dg['fecha_actualizacion']})."
    )
    gnd_def = float(dg["gastos_no_deducibles"] or 0.0)
    ing_ng_def = float(dg["ingresos_no_gravables"] or 0.0)
    pp_def = float(dg["pagos_parciales"] or 0.0)
    ret_def = float(dg["retenciones_2pct"] or 0.0)
    cred_def = float(dg["creditos_fiscales"] or 0.0)
    obs_def = str(dg["observaciones"] or "")
    try:
      aj_loaded = json.loads(dg["detalle_ajustes_json"] or "[]")
      if aj_loaded:
        ajustes_previos = aj_loaded
    except Exception:
      pass

  st.divider()
  st.markdown("### 2. Resultados Contables del Período (Lectura Automática de Libros)")
  r1, r2, r3, r4 = st.columns(4)
  ingresos_in = r1.number_input(
      "Ingresos Brutos (₡):",
      value=float(res_libros["ingresos"]),
      step=10000.0,
      format="%.2f",
      key=f"r_ing_{empresa_id}_{anio_fisc}",
  )
  costos_in = r2.number_input(
      "Costo de Ventas (₡):",
      value=float(res_libros["costos"]),
      step=10000.0,
      format="%.2f",
      key=f"r_cos_{empresa_id}_{anio_fisc}",
  )
  gastos_op_in = r3.number_input(
      "Gastos Operativos (₡):",
      value=float(res_libros["gastos_op"]),
      step=10000.0,
      format="%.2f",
      key=f"r_gop_{empresa_id}_{anio_fisc}",
  )
  gastos_fin_in = r4.number_input(
      "Gastos Financieros (₡):",
      value=float(res_libros["gastos_fin"]),
      step=5000.0,
      format="%.2f",
      key=f"r_gfi_{empresa_id}_{anio_fisc}",
  )

  utilidad_bruta = round(ingresos_in - costos_in, 2)
  utilidad_contable = round(utilidad_bruta - gastos_op_in - gastos_fin_in, 2)

  st.divider()
  st.markdown("### 3. Conciliación Fiscal (Partidas No Deducibles y No Gravables)")
  df_aj_ed = st.data_editor(
      pd.DataFrame(ajustes_previos),
      use_container_width=True,
      hide_index=True,
      num_rows="dynamic",
      column_config={
          "Tipo Ajuste": st.column_config.SelectboxColumn(
              "Tipo de Partida Fiscal",
              options=["(+) Gasto No Deducible", "(-) Ingreso No Gravable"],
              required=True,
          ),
          "Concepto / Referencia": st.column_config.TextColumn(
              "Descripción del Ajuste Fiscal"
          ),
          "Monto (₡)": st.column_config.NumberColumn(
              "Monto (₡)", min_value=0.0, step=1000.0, format="₡%.2f"
          ),
      },
      key=f"ed_aj_r_{empresa_id}_{anio_fisc}",
  )

  gastos_no_ded = 0.0
  ingresos_no_grav = 0.0
  ajustes_guardar = []
  for _, ra in df_aj_ed.iterrows():
    m_v = float(ra.get("Monto (₡)") or 0.0)
    t_v = str(ra.get("Tipo Ajuste") or "")
    ajustes_guardar.append({
        "Tipo Ajuste": t_v,
        "Concepto / Referencia": str(ra.get("Concepto / Referencia") or ""),
        "Monto (₡)": m_v,
    })
    if "No Deducible" in t_v or "(+)" in t_v:
      gastos_no_ded += m_v
    else:
      ingresos_no_grav += m_v

  if gastos_no_ded == 0 and gnd_def > 0 and decl_guardada.empty:
    gastos_no_ded = gnd_def
  if ingresos_no_grav == 0 and ing_ng_def > 0 and decl_guardada.empty:
    ingresos_no_grav = ing_ng_def

  renta_neta = round(utilidad_contable + gastos_no_ded - ingresos_no_grav, 2)

  with st.expander("⚙️ Parámetros de Tramos Vigentes Hacienda CR (Editables por Año)"):
    pt1, pt2, pt3, pt4 = st.columns(4)
    tope_pyme = pt1.number_input(
        "Tope Ingresos PYME (₡):", value=119174000.0, step=100000.0
    )
    tramo1_j = pt2.number_input(
        "Tramo 1 Jurídica 5% (₡):", value=5621000.0, step=10000.0
    )
    tramo2_j = pt3.number_input(
        "Tramo 2 Jurídica 10% (₡):", value=8433000.0, step=10000.0
    )
    tramo3_j = pt4.number_input(
        "Tramo 3 Jurídica 15% (₡):", value=11243000.0, step=10000.0
    )

  imp_determinado, detalle_tramos, nota_tarifa = calcular_impuesto_renta_cr(
      ingresos_in,
      renta_neta,
      regimen,
      tope_pyme=tope_pyme,
      tramo1_jur=tramo1_j,
      tramo2_jur=tramo2_j,
      tramo3_jur=tramo3_j,
  )

  st.divider()
  st.markdown("### 4. Pagos Parciales, Retenciones y Créditos Fiscales")
  cr1, cr2, cr3 = st.columns(3)
  pagos_parciales = cr1.number_input(
      "(-) Pagos Parciales Trimestrales (₡):",
      min_value=0.0,
      value=pp_def,
      step=5000.0,
      format="%.2f",
      key=f"r_pp_{empresa_id}_{anio_fisc}",
  )
  retenciones_2pct = cr2.number_input(
      "(-) Retenciones 2% / Fuente Aplicadas (₡):",
      min_value=0.0,
      value=ret_def,
      step=1000.0,
      format="%.2f",
      key=f"r_ret_{empresa_id}_{anio_fisc}",
  )
  creditos_fisc = cr3.number_input(
      "(-) Otros Créditos Fiscales / Saldos Favor (₡):",
      min_value=0.0,
      value=cred_def,
      step=1000.0,
      format="%.2f",
      key=f"r_cred_{empresa_id}_{anio_fisc}",
  )

  liq_neta = round(
      imp_determinado - pagos_parciales - retenciones_2pct - creditos_fisc, 2
  )
  if liq_neta > 0:
    imp_por_pagar = liq_neta
    saldo_favor = 0.0
  else:
    imp_por_pagar = 0.0
    saldo_favor = abs(liq_neta)

  st.divider()
  st.markdown("### 5. Resumen Oficial de Liquidación D-101")
  m1, m2, m3, m4 = st.columns(4)
  m1.metric("Utilidad Contable", f"₡{utilidad_contable:,.2f}")
  m2.metric("Renta Neta Imponible", f"₡{renta_neta:,.2f}")
  m3.metric("Impuesto Determinado", f"₡{imp_determinado:,.2f}")
  m4.metric(
      "Neto a Pagar / (Saldo Favor)",
      f"₡{imp_por_pagar:,.2f}" if imp_por_pagar > 0 else f"Favor ₡{saldo_favor:,.2f}",
      delta="POR PAGAR" if imp_por_pagar > 0 else "SALDO A FAVOR",
      delta_color="inverse" if imp_por_pagar > 0 else "normal",
  )

  if detalle_tramos:
    st.dataframe(
        pd.DataFrame(detalle_tramos).style.format({
            "Base Gravada (₡)": "₡{:,.2f}",
            "Impuesto (₡)": "₡{:,.2f}",
        }),
        use_container_width=True,
        hide_index=True,
    )

  obs_renta = st.text_area(
      "Notas y Observaciones de Conciliación Fiscal para D-101:",
      value=obs_def,
      height=85,
      key=f"obs_r_{empresa_id}_{anio_fisc}",
  )

  if st.button(
      f"💾 Guardar Declaración de Renta {int(anio_fisc)} en Base de Datos",
      type="primary",
      use_container_width=True,
  ):
    cur = conn.cursor()
    ahora_str = datetime.now().strftime("%d-%m-%Y %H:%M")
    costos_gastos_tot = round(costos_in + gastos_op_in + gastos_fin_in, 2)
    aj_json = json.dumps(ajustes_guardar, ensure_ascii=False)
    if not decl_guardada.empty:
      id_upd = int(decl_guardada.iloc[0]["id"])
      cur.execute(
          """
          UPDATE declaraciones_renta
          SET fecha_desde=?, fecha_hasta=?, regimen=?, ingresos_brutos=?,
              costos_gastos_contables=?, utilidad_contable=?,
              gastos_no_deducibles=?, ingresos_no_gravables=?,
              renta_neta_imponible=?, impuesto_determinado=?,
              pagos_parciales=?, retenciones_2pct=?, creditos_fiscales=?,
              impuesto_por_pagar=?, saldo_favor=?, detalle_ajustes_json=?,
              observaciones=?, fecha_actualizacion=?
          WHERE id=?
          """,
          (
              str(f_ini),
              str(f_fin),
              regimen,
              float(ingresos_in),
              costos_gastos_tot,
              utilidad_contable,
              gastos_no_ded,
              ingresos_no_grav,
              renta_neta,
              imp_determinado,
              float(pagos_parciales),
              float(retenciones_2pct),
              float(creditos_fisc),
              imp_por_pagar,
              saldo_favor,
              aj_json,
              obs_renta,
              ahora_str,
              id_upd,
          ),
      )
    else:
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
              empresa_id,
              int(anio_fisc),
              str(f_ini),
              str(f_fin),
              regimen,
              float(ingresos_in),
              costos_gastos_tot,
              utilidad_contable,
              gastos_no_ded,
              ingresos_no_grav,
              renta_neta,
              imp_determinado,
              float(pagos_parciales),
              float(retenciones_2pct),
              float(creditos_fisc),
              imp_por_pagar,
              saldo_favor,
              aj_json,
              obs_renta,
              ahora_str,
          ),
      )
    conn.commit()
    backup_fn()
    st.success(
        f"✅ Declaración de Impuesto sobre la Renta {int(anio_fisc)} guardada y respaldada."
    )

  st.divider()
  st.subheader(
      f"📥 Descargar Reporte de Renta ({int(anio_fisc)}) en Word, Excel y PDF"
  )
  datos_export = {
      "anio_fiscal": int(anio_fisc),
      "f_desde_str": f_ini.strftime("%d-%m-%Y"),
      "f_hasta_str": f_fin.strftime("%d-%m-%Y"),
      "regimen": regimen,
      "ingresos_brutos": float(ingresos_in),
      "costos": float(costos_in),
      "utilidad_bruta": utilidad_bruta,
      "gastos_op": float(gastos_op_in),
      "gastos_fin": float(gastos_fin_in),
      "utilidad_contable": utilidad_contable,
      "gastos_no_deducibles": gastos_no_ded,
      "ingresos_no_gravables": ingresos_no_grav,
      "renta_neta": renta_neta,
      "nota_tarifa": nota_tarifa,
      "impuesto_determinado": imp_determinado,
      "pagos_parciales": float(pagos_parciales),
      "retenciones_2pct": float(retenciones_2pct),
      "creditos_fiscales": float(creditos_fisc),
      "impuesto_por_pagar": imp_por_pagar,
      "saldo_favor": saldo_favor,
      "detalle_tramos": detalle_tramos,
      "observaciones": obs_renta,
  }

  empresa_nom = str(datos_emp["razon_social"] or "")
  cedula_nom = str(datos_emp["cedula"] or "")
  rep_nom = str(datos_emp["representante_legal"] or "Representante Legal")
  sufijo = f"{empresa_nom.replace(' ', '_')}_{int(anio_fisc)}"

  b1, b2, b3 = st.columns(3)
  b1.download_button(
      "📄 Descargar Word (.docx)",
      data=generar_word_renta(datos_export, empresa_nom, cedula_nom, rep_nom),
      file_name=f"Declaracion_Renta_D101_{sufijo}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )
  b2.download_button(
      "📊 Descargar Excel (.xlsx)",
      data=generar_excel_renta(datos_export, empresa_nom, cedula_nom, rep_nom),
      file_name=f"Declaracion_Renta_D101_{sufijo}.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
      use_container_width=True,
  )
  b3.download_button(
      "📕 Descargar PDF (.pdf)",
      data=generar_pdf_renta(datos_export, empresa_nom, cedula_nom, rep_nom),
      file_name=f"Declaracion_Renta_D101_{sufijo}.pdf",
      mime="application/pdf",
      use_container_width=True,
  )

  st.divider()
  st.subheader("🗂️ Historial de Declaraciones de Renta Guardadas")
  df_hist_r = pd.read_sql_query(
      """
      SELECT anio_fiscal AS Año_Fiscal,
             ingresos_brutos AS Ingresos_Brutos,
             utilidad_contable AS Utilidad_Contable,
             renta_neta_imponible AS Renta_Neta,
             impuesto_determinado AS Impuesto_Determinado,
             pagos_parciales AS Pagos_Parciales,
             impuesto_por_pagar AS Por_Pagar,
             saldo_favor AS Saldo_Favor,
             fecha_actualizacion AS Guardado
      FROM declaraciones_renta
      WHERE empresa_id = ?
      ORDER BY anio_fiscal DESC
      """,
      conn,
      params=(empresa_id,),
  )
  if df_hist_r.empty:
    st.info("Aún no has guardado declaraciones de Renta para esta empresa.")
  else:
    st.dataframe(
        df_hist_r.style.format({
            "Ingresos_Brutos": "₡{:,.2f}",
            "Utilidad_Contable": "₡{:,.2f}",
            "Renta_Neta": "₡{:,.2f}",
            "Impuesto_Determinado": "₡{:,.2f}",
            "Pagos_Parciales": "₡{:,.2f}",
            "Por_Pagar": "₡{:,.2f}",
            "Saldo_Favor": "₡{:,.2f}",
        }),
        use_container_width=True,
        hide_index=True,
    )
