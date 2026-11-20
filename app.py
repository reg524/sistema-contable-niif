from gestion_empresas import render_gestion_empresas
from sync_sheets import inicializar_y_descargar_bd, subir_bd_a_sheets
from conciliacion import render_modulo_conciliacion_bancaria
from reporte_iva import render_modulo_reporte_iva
from reporte_renta import render_modulo_reporte_renta
from datetime import date, datetime
import glob
import io
import os
import sqlite3
import zipfile
import docx
from docx.enum.section import WD_ORIENT
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

st.set_page_config(
    page_title="Sistema Contable NIIF", layout="wide", page_icon="📊"
)

DB_PATH = "sistema_contable.db"
BACKUP_DIR = "backups"


def get_conn():
  return sqlite3.connect(DB_PATH, check_same_thread=False)


def fmt_moneda(valor):
  if valor is None or valor == "":
    return ""
  val = float(valor)
  if round(val, 2) == 0:
    return "-"
  texto = f"{abs(val):,.2f}".replace(",", " ").replace(".", ",")
  return f"({texto})" if val < 0 else texto


def fmt_moneda_entero(valor):
  if valor is None or valor == "":
    return ""
  val = float(valor)
  if round(val, 0) == 0:
    return "-"
  texto = f"{abs(val):,.0f}".replace(",", ".")
  return f"({texto})" if val < 0 else texto


def fmt_pct(valor):
  if valor is None:
    return "-"
  val = float(valor)
  if round(val, 2) == 0:
    return "0.00%"
  return f"{val:,.2f}%"


def fmt_fecha_dmy(fecha_val):
  if fecha_val is None or str(fecha_val).strip() == "":
    return ""
  dt = pd.to_datetime(fecha_val, errors="coerce")
  if pd.isna(dt):
    return str(fecha_val)
  return dt.strftime("%d-%m-%Y")


def parse_fecha_input(fecha_str):
  s = str(fecha_str).strip()
  for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
    try:
      return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
    except ValueError:
      continue
  return s


def fecha_texto_es(dt_val):
  meses = [
      "enero",
      "febrero",
      "marzo",
      "abril",
      "mayo",
      "junio",
      "julio",
      "agosto",
      "setiembre",
      "octubre",
      "noviembre",
      "diciembre",
  ]
  return f"{dt_val.day} de {meses[dt_val.month - 1]} {dt_val.year}"


def limpiar_texto_pdf(texto):
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


# ==============================================================================
# 0. SISTEMA DE BACKUP AUTOMÁTICO (BASE DE DATOS + CÓDIGO FUENTE)
# ==============================================================================
def ejecutar_backup_automatico(max_copias=20):
  os.makedirs(BACKUP_DIR, exist_ok=True)
  timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
  nombre_zip = os.path.join(BACKUP_DIR, f"Backup_Contable_{timestamp}.zip")

  try:
    with zipfile.ZipFile(nombre_zip, "w", zipfile.ZIP_DEFLATED) as zf:
      if os.path.exists(DB_PATH):
        temp_db = os.path.join(BACKUP_DIR, f"temp_{timestamp}.db")
        src = sqlite3.connect(DB_PATH)
        dst = sqlite3.connect(temp_db)
        src.backup(dst)
        dst.close()
        src.close()
        zf.write(temp_db, arcname=DB_PATH)
        os.remove(temp_db)

      for archivo_py in glob.glob("*.py"):
        zf.write(archivo_py, arcname=archivo_py)

    archivos_backup = sorted(
        glob.glob(os.path.join(BACKUP_DIR, "Backup_Contable_*.zip"))
    )
    if len(archivos_backup) > max_copias:
      for viejo in archivos_backup[:-max_copias]:
        os.remove(viejo)

    return nombre_zip
  except Exception as e:
    return f"Error: {e}"


# ==============================================================================
# 1. CONSULTA UNIFICADA Y CÁLCULO DE CORTES HISTÓRICOS SIN DOBLE CONTEO
# ==============================================================================
def obtener_movimientos_empresa(conn, empresa_id):
  q_asientos = """
        SELECT 
            c.id AS asiento_id,
            c.num_asiento,
            c.fecha,
            c.concepto,
            c.num_comprobante,
            d.id AS detalle_id,
            d.cuenta_id,
            cat.codigo,
            cat.nombre AS cuenta_nombre,
            cat.tipo,
            d.debe,
            d.haber
        FROM asientos_cabecera c
        JOIN asientos_detalle d ON c.id = d.asiento_id
        JOIN catalogo_cuentas cat ON d.cuenta_id = cat.id
        WHERE c.empresa_id = ?
    """
  df_as = pd.read_sql_query(q_asientos, conn, params=(empresa_id,))

  q_ini = """
        SELECT 
            0 AS asiento_id,
            0 AS num_asiento,
            s.fecha_inicio AS fecha,
            'Saldo Inicial / Histórico' AS concepto,
            '' AS num_comprobante,
            s.id AS detalle_id,
            s.cuenta_id,
            cat.codigo,
            cat.nombre AS cuenta_nombre,
            cat.tipo,
            s.debe,
            s.haber
        FROM saldos_iniciales s
        JOIN catalogo_cuentas cat ON s.cuenta_id = cat.id
        WHERE s.empresa_id = ?
    """
  df_ini = pd.read_sql_query(q_ini, conn, params=(empresa_id,))

  if not df_ini.empty:
    df_ini["debe"] = pd.to_numeric(df_ini["debe"], errors="coerce").fillna(0.0)
    df_ini["haber"] = pd.to_numeric(df_ini["haber"], errors="coerce").fillna(
        0.0
    )
    df_ini = df_ini.drop_duplicates(
        subset=["cuenta_id", "fecha", "debe", "haber"], keep="first"
    )

  if not df_ini.empty and not df_as.empty:
    df_as["debe"] = pd.to_numeric(df_as["debe"], errors="coerce").fillna(0.0)
    df_as["haber"] = pd.to_numeric(df_as["haber"], errors="coerce").fillna(
        0.0
    )
    fechas_si = set(df_ini["fecha"].astype(str).str.strip())

    asientos_excluir = set()
    for aid, grp in df_as.groupby("asiento_id"):
      f_as_str = str(grp["fecha"].iloc[0]).strip()
      conc_str = str(grp["concepto"].iloc[0] or "").lower()
      if f_as_str in fechas_si and (
          "saldo" in conc_str
          or "inicial" in conc_str
          or "apertura" in conc_str
          or "cierre" in conc_str
      ):
        asientos_excluir.add(aid)
      elif f_as_str in fechas_si and len(grp) == len(df_ini):
        if round(grp["debe"].sum(), 2) == round(
            df_ini["debe"].sum(), 2
        ) and round(grp["haber"].sum(), 2) == round(df_ini["haber"].sum(), 2):
          asientos_excluir.add(aid)

    if asientos_excluir:
      df_as = df_as[~df_as["asiento_id"].isin(asientos_excluir)]

  df_total = pd.concat([df_ini, df_as], ignore_index=True)
  if not df_total.empty:
    df_total["debe"] = pd.to_numeric(df_total["debe"], errors="coerce").fillna(
        0.0
    )
    df_total["haber"] = pd.to_numeric(
        df_total["haber"], errors="coerce"
    ).fillna(0.0)
  return df_total


def calcular_saldos_a_fecha(df_mov, fecha_corte, incluir_asientos_del_dia=True):
  if df_mov.empty:
    return {}

  df = df_mov.copy()
  df["f_dt"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
  df = df[df["f_dt"].notna()]

  df_si = df[df["asiento_id"] == 0]
  df_as = df[df["asiento_id"] > 0]

  saldos_finales = {}
  todas_cuentas = df["cuenta_id"].unique()

  for cid in todas_cuentas:
    si_cta = df_si[(df_si["cuenta_id"] == cid) & (df_si["f_dt"] <= fecha_corte)]
    if not si_cta.empty:
      max_fecha_si = si_cta["f_dt"].max()
      si_foto = si_cta[si_cta["f_dt"] == max_fecha_si]
      base_saldo = si_foto["debe"].sum() - si_foto["haber"].sum()

      if incluir_asientos_del_dia:
        as_cta = df_as[
            (df_as["cuenta_id"] == cid)
            & (df_as["f_dt"] > max_fecha_si)
            & (df_as["f_dt"] <= fecha_corte)
        ]
      else:
        as_cta = df_as[
            (df_as["cuenta_id"] == cid)
            & (df_as["f_dt"] > max_fecha_si)
            & (df_as["f_dt"] < fecha_corte)
        ]
    else:
      base_saldo = 0.0
      if incluir_asientos_del_dia:
        as_cta = df_as[
            (df_as["cuenta_id"] == cid) & (df_as["f_dt"] <= fecha_corte)
        ]
      else:
        as_cta = df_as[
            (df_as["cuenta_id"] == cid) & (df_as["f_dt"] < fecha_corte)
        ]

    mov_asientos = as_cta["debe"].sum() - as_cta["haber"].sum()
    saldos_finales[cid] = base_saldo + mov_asientos

  return saldos_finales


def calcular_movimientos_periodo(df_mov, f_ini, f_fin, naturaleza_acreedora=True):
  if df_mov.empty:
    return {}

  df = df_mov.copy()
  df["f_dt"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
  df_as = df[
      (df["asiento_id"] > 0) & (df["f_dt"] >= f_ini) & (df["f_dt"] <= f_fin)
  ]

  if df_as.empty:
    df_si = df[
        (df["asiento_id"] == 0) & (df["f_dt"] >= f_ini) & (df["f_dt"] <= f_fin)
    ]
    if df_si.empty:
      return {}
    ult_f = df_si["f_dt"].max()
    df_as = df_si[df_si["f_dt"] == ult_f]

  if naturaleza_acreedora:
    return (
        df_as.groupby("cuenta_id")
        .apply(lambda x: x["haber"].sum() - x["debe"].sum())
        .to_dict()
    )
  return (
      df_as.groupby("cuenta_id")
      .apply(lambda x: x["debe"].sum() - x["haber"].sum())
      .to_dict()
  )


def obtener_cuentas_analiticas_sin_duplicar(df_catalogo):
  cat = df_catalogo.copy()
  cat["cod_str"] = cat["codigo"].astype(str).str.strip()
  todos_codigos = set(cat["cod_str"])

  def es_hoja_real(row):
    cod = row["cod_str"]
    tiene_hijos = any(
        otro != cod
        and (otro.startswith(cod + ".") or otro.startswith(cod + "-"))
        for otro in todos_codigos
    )
    return (int(row["acepta_movimiento"]) == 1) and (not tiene_hijos)

  cat["es_hoja_real"] = cat.apply(es_hoja_real, axis=1)
  return cat


# ==============================================================================
# 2. BALANCE DE COMPROBACIÓN DE 6 COLUMNAS
# ==============================================================================
def construir_balance_comprobacion_6col(df_mov, df_cat, f_ini, f_fin):
  df_cat = obtener_cuentas_analiticas_sin_duplicar(df_cat)
  df_cat = df_cat.sort_values("cod_str").reset_index(drop=True)

  ant_d, ant_c, mov_d, mov_c = {}, {}, {}, {}

  if not df_mov.empty:
    saldos_ant_neto = calcular_saldos_a_fecha(
        df_mov, f_ini, incluir_asientos_del_dia=False
    )
    for cid, neto in saldos_ant_neto.items():
      ant_d[cid] = neto if round(neto, 2) > 0 else 0.0
      ant_c[cid] = -neto if round(neto, 2) < 0 else 0.0

    df_m = df_mov.copy()
    df_m["f_dt"] = pd.to_datetime(df_m["fecha"], errors="coerce").dt.date
    df_si = df_m[df_m["asiento_id"] == 0]
    df_as = df_m[df_m["asiento_id"] > 0]

    si_per = df_si[(df_si["f_dt"] > f_ini) & (df_si["f_dt"] <= f_fin)]
    as_per = df_as[(df_as["f_dt"] >= f_ini) & (df_as["f_dt"] <= f_fin)]
    df_per = pd.concat([si_per, as_per], ignore_index=True)

    for cid, grp in df_per.groupby("cuenta_id"):
      mov_d[cid] = grp["debe"].sum()
      mov_c[cid] = grp["haber"].sum()

  cuentas_hoja = df_cat[df_cat["es_hoja_real"] == True].copy()
  valores_hoja = {}
  for _, r in cuentas_hoja.iterrows():
    cid = r["id"]
    ad = ant_d.get(cid, 0.0)
    ac = ant_c.get(cid, 0.0)
    md = mov_d.get(cid, 0.0)
    mc = mov_c.get(cid, 0.0)

    neto_final = (ad - ac) + (md - mc)
    fd = neto_final if round(neto_final, 2) > 0 else 0.0
    fc = -neto_final if round(neto_final, 2) < 0 else 0.0
    valores_hoja[cid] = (ad, ac, md, mc, fd, fc)

  filas = []
  for _, r in df_cat.iterrows():
    cid = r["id"]
    cod = r["cod_str"]
    es_hoja = bool(r["es_hoja_real"])

    if es_hoja:
      ad, ac, md, mc, fd, fc = valores_hoja.get(
          cid, (0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
      )
    else:
      hijos = cuentas_hoja[
          cuentas_hoja["cod_str"].str.startswith(cod + ".")
          | cuentas_hoja["cod_str"].str.startswith(cod + "-")
          | (cuentas_hoja["cod_str"] == cod)
      ]
      if hijos.empty:
        hijos = cuentas_hoja[cuentas_hoja["cod_str"].str.startswith(cod)]
      ad = sum(valores_hoja[h][0] for h in hijos["id"])
      ac = sum(valores_hoja[h][1] for h in hijos["id"])
      md = sum(valores_hoja[h][2] for h in hijos["id"])
      mc = sum(valores_hoja[h][3] for h in hijos["id"])
      fd = sum(valores_hoja[h][4] for h in hijos["id"])
      fc = sum(valores_hoja[h][5] for h in hijos["id"])

    filas.append({
        "cuenta_id": cid,
        "es_hoja": es_hoja,
        "codigo": cod,
        "nombre": r["nombre"],
        "ant_d": ad,
        "ant_c": ac,
        "mov_d": md,
        "mov_c": mc,
        "fin_d": fd,
        "fin_c": fc,
    })

  df_res = pd.DataFrame(filas)
  df_hojas_res = df_res[df_res["es_hoja"] == True]
  totales = {
      "ant_d": df_hojas_res["ant_d"].sum(),
      "ant_c": df_hojas_res["ant_c"].sum(),
      "mov_d": df_hojas_res["mov_d"].sum(),
      "mov_c": df_hojas_res["mov_c"].sum(),
      "fin_d": df_hojas_res["fin_d"].sum(),
      "fin_c": df_hojas_res["fin_c"].sum(),
  }
  return df_res, totales


def generar_word_balance_6col(
    df_res, totales, empresa, cedula, f_ini, f_fin, rep_legal
):
  doc = docx.Document()
  for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
  p_tit.paragraph_format.space_after = Pt(4)
  r1 = p_tit.add_run(f"{empresa.upper()}\n")
  r1.bold = True
  r1.font.size = Pt(10)
  r2 = p_tit.add_run(f"CEDULA {cedula}\n")
  r2.font.size = Pt(9)
  r3 = p_tit.add_run(
      f"Balance de comprobación Del {f_ini.strftime('%d-%m-%Y')} al"
      f" {f_fin.strftime('%d-%m-%Y')}"
  )
  r3.bold = True
  r3.font.size = Pt(9.5)

  table = doc.add_table(rows=2, cols=8)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER
  table.autofit = False
  anchos = [
      Inches(0.85),
      Inches(2.05),
      Inches(0.80),
      Inches(0.80),
      Inches(0.80),
      Inches(0.80),
      Inches(0.80),
      Inches(0.80),
  ]

  r0 = table.rows[0].cells
  for i in range(8):
    r0[i].width = anchos[i]
    r0[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8EEF5"/>')
    )

  r0[2].merge(r0[3])
  r0[4].merge(r0[5])
  r0[6].merge(r0[7])

  for idx_grp, titulo_grp in [
      (2, "Saldo anterior"),
      (4, "Movimientos"),
      (6, "Saldo final"),
  ]:
    p = r0[idx_grp].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(titulo_grp)
    run.bold = True
    run.font.size = Pt(7.5)

  subheaders = [
      "Número\nde cuenta",
      "Cuenta contable",
      "Débito",
      "Crédito",
      "Débito",
      "Crédito",
      "Débito",
      "Crédito",
  ]
  r1_cells = table.rows[1].cells
  for i, sh in enumerate(subheaders):
    r1_cells[i].width = anchos[i]
    r1_cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8EEF5"/>')
    )
    p = r1_cells[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(sh)
    run.bold = True
    run.font.size = Pt(7)

  for _, row in df_res.iterrows():
    cells = table.add_row().cells
    es_grupo = not row["es_hoja"]
    vals = [
        str(row["codigo"]),
        str(row["nombre"]),
        fmt_moneda(row["ant_d"]),
        fmt_moneda(row["ant_c"]),
        fmt_moneda(row["mov_d"]),
        fmt_moneda(row["mov_c"]),
        fmt_moneda(row["fin_d"]),
        fmt_moneda(row["fin_c"]),
    ]
    for i, val in enumerate(vals):
      cells[i].width = anchos[i]
      p = cells[i].paragraphs[0]
      p.paragraph_format.space_before = Pt(0.5)
      p.paragraph_format.space_after = Pt(0.5)
      p.alignment = (
          WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
      )
      run = p.add_run(val)
      run.bold = es_grupo
      run.font.size = Pt(6.8)

  tot_cells = table.add_row().cells
  tot_vals = [
      "",
      "TOTALES IGUALES",
      fmt_moneda(totales["ant_d"]),
      fmt_moneda(totales["ant_c"]),
      fmt_moneda(totales["mov_d"]),
      fmt_moneda(totales["mov_c"]),
      fmt_moneda(totales["fin_d"]),
      fmt_moneda(totales["fin_c"]),
  ]
  for i, val in enumerate(tot_vals):
    tot_cells[i].width = anchos[i]
    tot_cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(
            f'<w:tcBorders {nsdecls("w")}><w:top w:val="single"'
            ' w:sz="6"/><w:bottom w:val="double" w:sz="6"/></w:tcBorders>'
        )
    )
    p = tot_cells[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(val)
    run.bold = True
    run.font.size = Pt(7)

  doc.add_paragraph("\n")
  firma_table = doc.add_table(rows=1, cols=2)
  firma_table.alignment = WD_TABLE_ALIGNMENT.CENTER
  firma_table.autofit = False
  firma_table.rows[0].cells[0].width = Inches(3.85)
  firma_table.rows[0].cells[1].width = Inches(3.85)

  p_izq = firma_table.rows[0].cells[0].paragraphs[0]
  p_izq.alignment = WD_ALIGN_PARAGRAPH.LEFT
  r_izq = p_izq.add_run(
      "\n____________________________________\nLic. Roy F. Esquivel"
      " Gonzalez\nContador Privado Carné 8788"
  )
  r_izq.bold = True
  r_izq.font.size = Pt(8.5)

  p_der = firma_table.rows[0].cells[1].paragraphs[0]
  p_der.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_der = p_der.add_run(
      f"\n____________________________________\n{rep_legal}\nRepresentante"
      " Legal"
  )
  r_der.bold = True
  r_der.font.size = Pt(8.5)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_balance_6col(
    df_res, totales, empresa, cedula, f_ini, f_fin, rep_legal
):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Balance Comprobación"

  ws.merge_cells("A1:H1")
  ws["A1"] = empresa.upper()
  ws["A1"].font = Font(bold=True, size=12)

  ws.merge_cells("A2:H2")
  ws["A2"] = f"CEDULA {cedula}"
  ws["A2"].font = Font(size=10)

  ws.merge_cells("A3:H3")
  ws["A3"] = (
      f"Balance de comprobación Del {f_ini.strftime('%d-%m-%Y')} al"
      f" {f_fin.strftime('%d-%m-%Y')}"
  )
  ws["A3"].font = Font(bold=True, size=10)

  ws.append([])
  ws.merge_cells("C5:D5")
  ws["C5"] = "Saldo anterior"
  ws.merge_cells("E5:F5")
  ws["E5"] = "Movimientos"
  ws.merge_cells("G5:H5")
  ws["G5"] = "Saldo final"

  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  for col in range(1, 9):
    c = ws.cell(row=5, column=col)
    c.font = Font(bold=True)
    c.fill = fill_hdr
    c.alignment = Alignment(horizontal="center")

  subh = [
      "Número de cuenta",
      "Cuenta contable",
      "Débito",
      "Crédito",
      "Débito",
      "Crédito",
      "Débito",
      "Crédito",
  ]
  ws.append(subh)
  for col in range(1, 9):
    c = ws.cell(row=6, column=col)
    c.font = Font(bold=True)
    c.fill = fill_hdr
    c.alignment = Alignment(horizontal="right" if col >= 3 else "left")

  fila = 7
  for _, r in df_res.iterrows():
    es_grupo = not r["es_hoja"]
    ws.cell(row=fila, column=1, value=str(r["codigo"]))
    ws.cell(row=fila, column=2, value=str(r["nombre"]))
    vals = [
        r["ant_d"],
        r["ant_c"],
        r["mov_d"],
        r["mov_c"],
        r["fin_d"],
        r["fin_c"],
    ]
    for idx_v, v in enumerate(vals, start=3):
      cell = ws.cell(row=fila, column=idx_v, value=float(v))
      cell.number_format = "#,##0.00"
    if es_grupo:
      for col in range(1, 9):
        ws.cell(row=fila, column=col).font = Font(bold=True)
    fila += 1

  ws.cell(row=fila, column=2, value="TOTALES IGUALES").font = Font(bold=True)
  borde_tot = Border(top=Side(style="thin"), bottom=Side(style="double"))
  tot_list = [
      totales["ant_d"],
      totales["ant_c"],
      totales["mov_d"],
      totales["mov_c"],
      totales["fin_d"],
      totales["fin_c"],
  ]
  for idx_v, v in enumerate(tot_list, start=3):
    cell = ws.cell(row=fila, column=idx_v, value=float(v))
    cell.font = Font(bold=True)
    cell.number_format = "#,##0.00"
    cell.border = borde_tot

  fila += 3
  ws.cell(
      row=fila,
      column=1,
      value=(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=fila, column=1).alignment = Alignment(
      wrap_text=True, horizontal="left"
  )

  ws.cell(
      row=fila,
      column=8,
      value=f"_______________________________\n{rep_legal}\nRepresentante Legal",
  ).font = Font(bold=True)
  ws.cell(row=fila, column=8).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[fila].height = 45

  for letra, ancho in zip(
      ["A", "B", "C", "D", "E", "F", "G", "H"],
      [16, 38, 16, 16, 16, 16, 16, 16],
  ):
    ws.column_dimensions[letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_balance_6col(
    df_res, totales, empresa, cedula, f_ini, f_fin, rep_legal
):
  pdf = FPDF(orientation="P", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=12)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 10)
  pdf.cell(0, 5, limpiar_texto_pdf(empresa.upper()), ln=True)
  pdf.set_font("Helvetica", "", 8.5)
  pdf.cell(0, 4.5, limpiar_texto_pdf(f"CEDULA {cedula}"), ln=True)
  pdf.set_font("Helvetica", "B", 9)
  pdf.cell(
      0,
      5,
      limpiar_texto_pdf(
          f"Balance de comprobación Del {f_ini.strftime('%d-%m-%Y')} al"
          f" {f_fin.strftime('%d-%m-%Y')}"
      ),
      ln=True,
  )
  pdf.ln(2)

  w = [20, 54, 20, 20, 20, 20, 20, 20]
  pdf.set_fill_color(230, 235, 245)
  pdf.set_font("Helvetica", "B", 7)

  pdf.cell(w[0] + w[1], 5, "", border=1, fill=True)
  pdf.cell(w[2] + w[3], 5, "Saldo anterior", border=1, align="C", fill=True)
  pdf.cell(w[4] + w[5], 5, "Movimientos", border=1, align="C", fill=True)
  pdf.cell(w[6] + w[7], 5, "Saldo final", border=1, align="C", fill=True)
  pdf.ln()

  subh = [
      "N° Cuenta",
      "Cuenta contable",
      "Débito",
      "Crédito",
      "Débito",
      "Crédito",
      "Débito",
      "Crédito",
  ]
  for i, h in enumerate(subh):
    pdf.cell(
        w[i],
        5,
        limpiar_texto_pdf(h),
        border=1,
        align="R" if i >= 2 else "L",
        fill=True,
    )
  pdf.ln()

  for _, r in df_res.iterrows():
    es_grupo = not r["es_hoja"]
    pdf.set_font("Helvetica", "B" if es_grupo else "", 6.5)
    pdf.cell(w[0], 4.5, limpiar_texto_pdf(r["codigo"])[:14], border="B")
    pdf.cell(w[1], 4.5, limpiar_texto_pdf(r["nombre"])[:34], border="B")
    vals = [
        r["ant_d"],
        r["ant_c"],
        r["mov_d"],
        r["mov_c"],
        r["fin_d"],
        r["fin_c"],
    ]
    for idx_v, v in enumerate(vals, start=2):
      pdf.cell(w[idx_v], 4.5, fmt_moneda(v), border="B", align="R")
    pdf.ln()

  pdf.set_font("Helvetica", "B", 6.8)
  pdf.cell(w[0] + w[1], 5.5, "TOTALES IGUALES", border="TB")
  tot_list = [
      totales["ant_d"],
      totales["ant_c"],
      totales["mov_d"],
      totales["mov_c"],
      totales["fin_d"],
      totales["fin_c"],
  ]
  for idx_v, v in enumerate(tot_list, start=2):
    pdf.cell(w[idx_v], 5.5, fmt_moneda(v), border="TB", align="R")
  pdf.ln(14)

  if pdf.get_y() > 245:
    pdf.add_page()
  y_f = pdf.get_y()
  pdf.set_font("Helvetica", "B", 8)
  pdf.set_xy(10, y_f)
  pdf.multi_cell(
      85,
      4.5,
      limpiar_texto_pdf(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
      align="L",
  )
  pdf.set_xy(115, y_f)
  pdf.multi_cell(
      85,
      4.5,
      limpiar_texto_pdf(
          f"_______________________________\n{rep_legal}\nRepresentante Legal"
      ),
      align="R",
  )
  return bytes(pdf.output())


# ==============================================================================
# 3. GENERADORES DE REPORTE IMPRIMIBLE DE ASIENTOS ELABORADOS
# ==============================================================================
def generar_word_asientos(df_asientos_listado, empresa, cedula, f_ini, f_fin):
  doc = docx.Document()
  for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_tit.add_run("REPORTE DE ASIENTOS CONTABLES ELABORADOS\n").bold = True
  p_tit.add_run(f"{empresa.upper()}\n").bold = True
  p_tit.add_run(f"CÉDULA JURÍDICA {cedula}\n")
  p_tit.add_run(
      f"Del {f_ini.strftime('%d-%m-%Y')} al {f_fin.strftime('%d-%m-%Y')}\n"
  ).bold = True
  p_tit.add_run("Expresado en colones").italic = True

  table = doc.add_table(rows=1, cols=6)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER
  table.autofit = False
  anchos = [
      Inches(0.65),
      Inches(0.95),
      Inches(0.85),
      Inches(2.55),
      Inches(1.2),
      Inches(1.2),
  ]

  headers = ["Asiento", "Fecha", "Código", "Cuenta / Concepto", "Debe", "Haber"]
  hdr_cells = table.rows[0].cells
  for i, h in enumerate(headers):
    hdr_cells[i].width = anchos[i]
    p = hdr_cells[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 4 else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(h)
    r.bold = True
    r.font.size = Pt(8.5)
    hdr_cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8EEF5"/>')
    )

  for _, row in df_asientos_listado.iterrows():
    cells = table.add_row().cells
    for i in range(6):
      cells[i].width = anchos[i]
      cells[i].paragraphs[0].paragraph_format.space_before = Pt(1)
      cells[i].paragraphs[0].paragraph_format.space_after = Pt(1)

    vals = [
        f"#{row['num_asiento']}",
        fmt_fecha_dmy(row["fecha"]),
        str(row["codigo"]),
        f"{row['cuenta_nombre']} ({row['concepto']})",
        fmt_moneda(row["debe"]) if row["debe"] > 0 else "",
        fmt_moneda(row["haber"]) if row["haber"] > 0 else "",
    ]
    for i, val in enumerate(vals):
      p = cells[i].paragraphs[0]
      p.alignment = (
          WD_ALIGN_PARAGRAPH.RIGHT if i >= 4 else WD_ALIGN_PARAGRAPH.LEFT
      )
      r = p.add_run(val)
      r.font.size = Pt(8)

  tot_cells = table.add_row().cells
  for i in range(6):
    tot_cells[i].width = anchos[i]
  tot_cells[3].paragraphs[0].add_run("TOTALES GENERALES").bold = True
  p_td = tot_cells[4].paragraphs[0]
  p_td.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_td = p_td.add_run(fmt_moneda(df_asientos_listado["debe"].sum()))
  r_td.bold = True
  r_td.font.size = Pt(8.5)

  p_th = tot_cells[5].paragraphs[0]
  p_th.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_th = p_th.add_run(fmt_moneda(df_asientos_listado["haber"].sum()))
  r_th.bold = True
  r_th.font.size = Pt(8.5)

  doc.add_paragraph("\n")
  p_firma = doc.add_paragraph()
  p_firma.alignment = WD_ALIGN_PARAGRAPH.LEFT
  r_f = p_firma.add_run(
      "\n____________________________________\nLic. Roy F. Esquivel"
      " Gonzalez\nContador Privado Carné 8788"
  )
  r_f.bold = True
  r_f.font.size = Pt(9)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_asientos(df_asientos_listado, empresa, cedula, f_ini, f_fin):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Asientos Elaborados"

  ws.merge_cells("A1:H1")
  ws["A1"] = "REPORTE DE ASIENTOS CONTABLES ELABORADOS"
  ws["A1"].font = Font(bold=True, size=13)
  ws["A1"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A2:H2")
  ws["A2"] = (
      f"{empresa.upper()} - CÉDULA {cedula} | Del {f_ini.strftime('%d-%m-%Y')}"
      f" al {f_fin.strftime('%d-%m-%Y')}"
  )
  ws["A2"].font = Font(bold=True, size=10)
  ws["A2"].alignment = Alignment(horizontal="center")

  ws.append([])
  headers = [
      "N° Asiento",
      "Fecha",
      "Comprobante",
      "Código",
      "Cuenta Contable",
      "Concepto",
      "Debe",
      "Haber",
  ]
  ws.append(headers)
  fill_h = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  for col in range(1, 9):
    c = ws.cell(row=4, column=col)
    c.font = Font(bold=True)
    c.fill = fill_h

  fila = 5
  for _, r in df_asientos_listado.iterrows():
    ws.cell(row=fila, column=1, value=int(r["num_asiento"]))
    ws.cell(row=fila, column=2, value=fmt_fecha_dmy(r["fecha"]))
    ws.cell(row=fila, column=3, value=str(r["num_comprobante"] or ""))
    ws.cell(row=fila, column=4, value=str(r["codigo"]))
    ws.cell(row=fila, column=5, value=str(r["cuenta_nombre"]))
    ws.cell(row=fila, column=6, value=str(r["concepto"] or ""))
    cd = ws.cell(row=fila, column=7, value=float(r["debe"]))
    ch = ws.cell(row=fila, column=8, value=float(r["haber"]))
    cd.number_format = "#,##0.00"
    ch.number_format = "#,##0.00"
    fila += 1

  ws.cell(row=fila, column=6, value="TOTALES").font = Font(bold=True)
  c_td = ws.cell(
      row=fila, column=7, value=float(df_asientos_listado["debe"].sum())
  )
  c_th = ws.cell(
      row=fila, column=8, value=float(df_asientos_listado["haber"].sum())
  )
  c_td.font = Font(bold=True)
  c_th.font = Font(bold=True)
  c_td.number_format = "#,##0.00"
  c_th.number_format = "#,##0.00"

  fila += 3
  ws.cell(
      row=fila,
      column=1,
      value=(
          "___________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=fila, column=1).alignment = Alignment(
      wrap_text=True, horizontal="left"
  )
  ws.row_dimensions[fila].height = 45

  for col_letra, ancho in zip(
      ["A", "B", "C", "D", "E", "F", "G", "H"],
      [12, 14, 16, 12, 32, 35, 18, 18],
  ):
    ws.column_dimensions[col_letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_asientos(df_asientos_listado, empresa, cedula, f_ini, f_fin):
  pdf = FPDF(orientation="P", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=15)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(0, 6, "REPORTE DE ASIENTOS CONTABLES ELABORADOS", ln=True, align="C")
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5,
      limpiar_texto_pdf(f"{empresa.upper()} - CÉDULA JURÍDICA {cedula}"),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "", 8.5)
  pdf.cell(
      0,
      5,
      f"Del {f_ini.strftime('%d-%m-%Y')} al {f_fin.strftime('%d-%m-%Y')}",
      ln=True,
      align="C",
  )
  pdf.ln(3)

  w = [16, 22, 20, 74, 30, 30]
  pdf.set_fill_color(230, 235, 245)
  pdf.set_font("Helvetica", "B", 8)
  headers = ["Asiento", "Fecha", "Código", "Cuenta / Concepto", "Debe", "Haber"]
  for i, h in enumerate(headers):
    pdf.cell(
        w[i],
        6,
        limpiar_texto_pdf(h),
        border=1,
        align="R" if i >= 4 else "L",
        fill=True,
    )
  pdf.ln()

  pdf.set_font("Helvetica", "", 7.5)
  for _, r in df_asientos_listado.iterrows():
    desc = limpiar_texto_pdf(f"{r['cuenta_nombre']} - {r['concepto']}")[:44]
    pdf.cell(w[0], 5, f"#{r['num_asiento']}", border="B")
    pdf.cell(w[1], 5, fmt_fecha_dmy(r["fecha"]), border="B")
    pdf.cell(w[2], 5, limpiar_texto_pdf(r["codigo"]), border="B")
    pdf.cell(w[3], 5, desc, border="B")
    pdf.cell(
        w[4],
        5,
        fmt_moneda(r["debe"]) if r["debe"] > 0 else "",
        border="B",
        align="R",
    )
    pdf.cell(
        w[5],
        5,
        fmt_moneda(r["haber"]) if r["haber"] > 0 else "",
        border="B",
        align="R",
    )
    pdf.ln()

  pdf.set_font("Helvetica", "B", 8)
  pdf.cell(sum(w[:4]), 6, "TOTALES GENERALES", border="T", align="R")
  pdf.cell(
      w[4],
      6,
      fmt_moneda(df_asientos_listado["debe"].sum()),
      border="T",
      align="R",
  )
  pdf.cell(
      w[5],
      6,
      fmt_moneda(df_asientos_listado["haber"].sum()),
      border="T",
      align="R",
  )
  pdf.ln(14)

  if pdf.get_y() > 245:
    pdf.add_page()
  pdf.set_xy(12, pdf.get_y())
  pdf.multi_cell(
      85,
      4.5,
      limpiar_texto_pdf(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
      align="L",
  )
  return bytes(pdf.output())


# ==============================================================================
# 4. MODAL DRILL-DOWN PARA EDITAR O BORRAR ASIENTOS DESDE EL BALANCE
# ==============================================================================
@st.dialog("Auditoría y Edición de Asientos", width="large")
def modal_editar_movimientos(empresa_id, cuenta_id, codigo, nombre):
  conn = get_conn()
  st.subheader(f"Cuenta: {codigo} - {nombre}")

  q_ids = """
        SELECT DISTINCT c.id, c.num_asiento, c.fecha, c.concepto
        FROM asientos_cabecera c
        JOIN asientos_detalle d ON c.id = d.asiento_id
        WHERE c.empresa_id = ? AND d.cuenta_id = ?
        ORDER BY c.fecha, c.num_asiento
    """
  asientos = pd.read_sql_query(q_ids, conn, params=(empresa_id, int(cuenta_id)))

  df_si = pd.read_sql_query(
      "SELECT * FROM saldos_iniciales WHERE empresa_id = ? AND cuenta_id = ?"
      " ORDER BY fecha_inicio",
      conn,
      params=(empresa_id, int(cuenta_id)),
  )
  if not df_si.empty:
    for _, si in df_si.iterrows():
      st.info(
          f"📌 **Saldo Histórico / Inicial ({fmt_fecha_dmy(si['fecha_inicio'])}):** Debe:"
          f" ₡{si['debe']:,.2f} | Haber: ₡{si['haber']:,.2f}"
      )

  if asientos.empty:
    if df_si.empty:
      st.info("No se encontraron asientos registrados para esta cuenta.")
    conn.close()
    return

  for _, as_row in asientos.iterrows():
    asiento_id = int(as_row["id"])
    lineas = pd.read_sql_query(
        """
            SELECT d.id AS detalle_id, d.cuenta_id, cat.codigo, cat.nombre, d.debe, d.haber
            FROM asientos_detalle d
            JOIN catalogo_cuentas cat ON d.cuenta_id = cat.id
            WHERE d.asiento_id = ?
        """,
        conn,
        params=(asiento_id,),
    )
    total_asiento = lineas["debe"].sum()

    with st.expander(
        f"Asiento #{as_row['num_asiento']} | Fecha:"
        f" {fmt_fecha_dmy(as_row['fecha'])} | {as_row['concepto']} | Total:"
        f" ₡{total_asiento:,.2f}"
    ):
      with st.form(key=f"form_as_{asiento_id}"):
        c_f, c_c = st.columns([1, 2])
        nueva_fecha = c_f.text_input(
            "Fecha (DD-MM-AAAA)", value=fmt_fecha_dmy(as_row["fecha"])
        )
        nuevo_concepto = c_c.text_input(
            "Concepto", value=str(as_row["concepto"])
        )

        cambios = []
        for _, l_row in lineas.iterrows():
          st.markdown(f"**{l_row['codigo']} - {l_row['nombre']}**")
          col1, col2 = st.columns(2)
          nd = col1.number_input(
              "Debe (₡)",
              value=float(l_row["debe"]),
              min_value=0.0,
              step=100.0,
              key=f"d_{l_row['detalle_id']}",
          )
          nh = col2.number_input(
              "Haber (₡)",
              value=float(l_row["haber"]),
              min_value=0.0,
              step=100.0,
              key=f"h_{l_row['detalle_id']}",
          )
          cambios.append(
              {"detalle_id": int(l_row["detalle_id"]), "debe": nd, "haber": nh}
          )
          st.divider()

        b1, b2 = st.columns(2)
        guardar = b1.form_submit_button(
            "💾 Guardar Cambios", type="primary", use_container_width=True
        )
        borrar = b2.form_submit_button(
            "🗑️ Eliminar Asiento Completo", use_container_width=True
        )

        if guardar:
          sd = sum(x["debe"] for x in cambios)
          sh = sum(x["haber"] for x in cambios)
          if round(sd, 2) != round(sh, 2):
            st.error(
                f"❌ El asiento está descuadrado: Debe (₡{sd:,.2f}) ≠ Haber"
                f" (₡{sh:,.2f})"
            )
          else:
            fecha_iso = parse_fecha_input(nueva_fecha)
            cur = conn.cursor()
            cur.execute(
                "UPDATE asientos_cabecera SET fecha = ?, periodo = ?, concepto"
                " = ? WHERE id = ?",
                (fecha_iso, fecha_iso[:7], nuevo_concepto, asiento_id),
            )
            for item in cambios:
              cur.execute(
                  "UPDATE asientos_detalle SET debe = ?, haber = ? WHERE id = ?",
                  (item["debe"], item["haber"], item["detalle_id"]),
              )
            conn.commit()
            conn.close()
            ejecutar_backup_automatico()
            st.success("✅ Asiento actualizado correctamente.")
            st.rerun()

        if borrar:
          cur = conn.cursor()
          cur.execute(
              "DELETE FROM asientos_detalle WHERE asiento_id = ?", (asiento_id,)
          )
          cur.execute(
              "DELETE FROM asientos_cabecera WHERE id = ?", (asiento_id,)
          )
          conn.commit()
          conn.close()
          ejecutar_backup_automatico()
          st.warning(f"🗑️ Asiento #{as_row['num_asiento']} eliminado.")
          st.rerun()
  conn.close()


# ==============================================================================
# 5. ESTADO DE LA SITUACIÓN FINANCIERA COMPARATIVO
# ==============================================================================
def construir_estado_situacion_comparativo(
    df_mov, df_catalogo, fecha_ant, fecha_act
):
  saldos_ant = calcular_saldos_a_fecha(
      df_mov, fecha_ant, incluir_asientos_del_dia=True
  )
  saldos_act = calcular_saldos_a_fecha(
      df_mov, fecha_act, incluir_asientos_del_dia=True
  )

  cat_limpio = obtener_cuentas_analiticas_sin_duplicar(df_catalogo)
  cat = (
      cat_limpio[cat_limpio["es_hoja_real"] == True]
      .sort_values("cod_str")
      .copy()
  )
  cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()

  activos = cat[
      cat["tipo_up"].str.contains("ACTIVO")
      | cat["cod_str"].str.startswith("1")
  ]
  act_circ = activos[
      activos["cod_str"].str.startswith(("1.1", "11", "1-1", "101"))
  ]
  act_fijo = activos[~activos.index.isin(act_circ.index)]
  if act_circ.empty and not activos.empty:
    act_circ = activos

  pasivos = cat[
      (
          cat["tipo_up"].str.contains("PASIVO")
          | cat["cod_str"].str.startswith("2")
      )
      & (~cat.index.isin(activos.index))
  ]
  patrimonio = cat[
      (
          cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL")
          | cat["cod_str"].str.startswith("3")
      )
      & (~cat.index.isin(activos.index))
      & (~cat.index.isin(pasivos.index))
  ]

  ing_gast = cat[
      cat["tipo_up"].str.contains("INGRESO|GASTO|COSTO|EGRESO")
      | cat["cod_str"].str.startswith(("4", "5", "6"))
  ]
  res_ej_ant = sum(-saldos_ant.get(cid, 0.0) for cid in ing_gast["id"])
  res_ej_act = sum(-saldos_act.get(cid, 0.0) for cid in ing_gast["id"])

  estructura = []
  nota_idx = 1

  estructura.append(("TITULO", "ACTIVOS", "", None, None))
  estructura.append(("SUBTITULO", "CIRCULANTE", "", None, None))

  sub_ac_ant, sub_ac_act = 0.0, 0.0
  for _, cta in act_circ.iterrows():
    v_ant = saldos_ant.get(cta["id"], 0.0)
    v_act = saldos_act.get(cta["id"], 0.0)
    if round(v_ant, 2) != 0 or round(v_act, 2) != 0:
      sub_ac_ant += v_ant
      sub_ac_act += v_act
      estructura.append(("CUENTA", cta["nombre"], str(nota_idx), v_ant, v_act))
      nota_idx += 1
  estructura.append(
      ("SUBTOTAL", "Total de Circulante", "", sub_ac_ant, sub_ac_act)
  )

  if not act_fijo.empty:
    estructura.append(("SUBTITULO", "FIJOS", str(nota_idx), None, None))
    nota_idx += 1
    sub_af_ant, sub_af_act = 0.0, 0.0
    for _, cta in act_fijo.iterrows():
      v_ant = saldos_ant.get(cta["id"], 0.0)
      v_act = saldos_act.get(cta["id"], 0.0)
      if round(v_ant, 2) != 0 or round(v_act, 2) != 0:
        sub_af_ant += v_ant
        sub_af_act += v_act
        estructura.append(("CUENTA", cta["nombre"], "", v_ant, v_act))
    estructura.append(("SUBTOTAL", "Total de Fijos", "", sub_af_ant, sub_af_act))
  else:
    sub_af_ant, sub_af_act = 0.0, 0.0

  tot_activo_ant = sub_ac_ant + sub_af_ant
  tot_activo_act = sub_ac_act + sub_af_act
  estructura.append(
      ("TOTAL", "TOTAL DE ACTIVOS", "", tot_activo_ant, tot_activo_act)
  )

  estructura.append(("TITULO", "PASIVOS", "", None, None))
  estructura.append(("SUBTITULO", "CIRCULANTES", "", None, None))
  sub_pc_ant, sub_pc_act = 0.0, 0.0
  for _, cta in pasivos.iterrows():
    v_ant = -saldos_ant.get(cta["id"], 0.0)
    v_act = -saldos_act.get(cta["id"], 0.0)
    if round(v_ant, 2) != 0 or round(v_act, 2) != 0:
      sub_pc_ant += v_ant
      sub_pc_act += v_act
      estructura.append(("CUENTA", cta["nombre"], str(nota_idx), v_ant, v_act))
      nota_idx += 1
  estructura.append(
      ("SUBTOTAL", "Total de Circulante", "", sub_pc_ant, sub_pc_act)
  )
  estructura.append(("TOTAL", "TOTAL DE PASIVOS", "", sub_pc_ant, sub_pc_act))

  estructura.append(("TITULO", "PATRIMONIO", "", None, None))
  sub_pat_ant, sub_pat_act = 0.0, 0.0
  for _, cta in patrimonio.iterrows():
    v_ant = -saldos_ant.get(cta["id"], 0.0)
    v_act = -saldos_act.get(cta["id"], 0.0)
    if round(v_ant, 2) != 0 or round(v_act, 2) != 0:
      sub_pat_ant += v_ant
      sub_pat_act += v_act
      estructura.append(("CUENTA", cta["nombre"], str(nota_idx), v_ant, v_act))
      nota_idx += 1

  if round(res_ej_ant, 2) != 0 or round(res_ej_act, 2) != 0:
    sub_pat_ant += res_ej_ant
    sub_pat_act += res_ej_act
    estructura.append((
        "CUENTA",
        "Excedentes / Resultado del Período",
        str(nota_idx),
        res_ej_ant,
        res_ej_act,
    ))

  estructura.append(
      ("SUBTOTAL", "Total de Patrimonio", "", sub_pat_ant, sub_pat_act)
  )
  estructura.append((
      "TOTAL",
      "TOTAL DE PASIVO Y PATRIMONIO",
      "",
      sub_pc_ant + sub_pat_ant,
      sub_pc_act + sub_pat_act,
  ))

  return estructura


# ==============================================================================
# 5B. ESTADO DE RESULTADOS COMPARATIVO (NIIF)
# ==============================================================================
def construir_estado_resultados_comparativo(
    df_mov,
    df_catalogo,
    f1_ini,
    f1_fin,
    f2_ini,
    f2_fin,
    desglosar_subcuentas=False,
):
  saldos_p1 = calcular_movimientos_periodo(
      df_mov, f1_ini, f1_fin, naturaleza_acreedora=True
  )
  saldos_p2 = calcular_movimientos_periodo(
      df_mov, f2_ini, f2_fin, naturaleza_acreedora=True
  )

  cat_limpio = obtener_cuentas_analiticas_sin_duplicar(df_catalogo)
  cat_hojas = (
      cat_limpio[cat_limpio["es_hoja_real"] == True]
      .sort_values("cod_str")
      .copy()
  )
  cat_hojas["tipo_up"] = cat_hojas["tipo"].astype(str).str.upper().str.strip()

  cat_padres = cat_limpio[cat_limpio["es_hoja_real"] == False].copy()
  mapa_nombre_padre = dict(zip(cat_padres["cod_str"], cat_padres["nombre"]))

  def obtener_cuenta_principal(cod_str):
    partes = cod_str.replace("-", ".").split(".")
    for longitud in range(len(partes) - 1, 0, -1):
      prefijo = ".".join(partes[:longitud])
      if prefijo in mapa_nombre_padre and longitud >= 2:
        return prefijo, mapa_nombre_padre[prefijo]
    return cod_str, None

  ingresos = cat_hojas[
      cat_hojas["tipo_up"].str.contains("INGRESO")
      | cat_hojas["cod_str"].str.startswith("4")
  ]
  costos = cat_hojas[
      (
          cat_hojas["tipo_up"].str.contains("COSTO")
          | cat_hojas["cod_str"].str.startswith("5")
      )
      & (~cat_hojas.index.isin(ingresos.index))
  ]
  gastos_op = cat_hojas[
      (
          cat_hojas["tipo_up"].str.contains("GASTO|EGRESO")
          | cat_hojas["cod_str"].str.startswith("6")
      )
      & (~cat_hojas.index.isin(ingresos.index))
      & (~cat_hojas.index.isin(costos.index))
  ]
  otros_res = cat_hojas[
      cat_hojas["cod_str"].str.startswith(("7", "8"))
      & (~cat_hojas.index.isin(ingresos.index))
      & (~cat_hojas.index.isin(costos.index))
      & (~cat_hojas.index.isin(gastos_op.index))
  ]

  def agrupar_bloque(df_bloque, es_acreedor=True):
    items = []
    if desglosar_subcuentas:
      for _, cta in df_bloque.iterrows():
        v1 = saldos_p1.get(cta["id"], 0.0) * (1 if es_acreedor else -1)
        v2 = saldos_p2.get(cta["id"], 0.0) * (1 if es_acreedor else -1)
        if round(v1, 2) != 0 or round(v2, 2) != 0:
          items.append((cta["nombre"], v1, v2))
    else:
      agrupado = {}
      orden_claves = []
      for _, cta in df_bloque.iterrows():
        v1 = saldos_p1.get(cta["id"], 0.0) * (1 if es_acreedor else -1)
        v2 = saldos_p2.get(cta["id"], 0.0) * (1 if es_acreedor else -1)
        if round(v1, 2) == 0 and round(v2, 2) == 0:
          continue
        _, nom_p = obtener_cuenta_principal(cta["cod_str"])
        etiqueta = nom_p if nom_p else cta["nombre"]
        if etiqueta not in agrupado:
          agrupado[etiqueta] = [0.0, 0.0]
          orden_claves.append(etiqueta)
        agrupado[etiqueta][0] += v1
        agrupado[etiqueta][1] += v2
      for etq in orden_claves:
        items.append((etq, agrupado[etq][0], agrupado[etq][1]))
    return items

  estructura = []
  nota_idx = 1

  estructura.append(("TITULO", "INGRESOS", "", None, None))
  items_ing = agrupar_bloque(ingresos, es_acreedor=True)
  tot_ing_1 = sum(x[1] for x in items_ing)
  tot_ing_2 = sum(x[2] for x in items_ing)
  for nom, v1, v2 in items_ing:
    estructura.append(("CUENTA", nom, str(nota_idx), v1, v2))
    nota_idx += 1
  estructura.append(("SUBTOTAL", "TOTAL INGRESOS", "", tot_ing_1, tot_ing_2))

  items_cos = agrupar_bloque(costos, es_acreedor=False)
  tot_cos_1 = sum(x[1] for x in items_cos)
  tot_cos_2 = sum(x[2] for x in items_cos)
  if items_cos:
    estructura.append(("TITULO", "COSTOS DE OPERACIÓN", "", None, None))
    for nom, v1, v2 in items_cos:
      estructura.append(("CUENTA", nom, str(nota_idx), v1, v2))
      nota_idx += 1
    estructura.append(("SUBTOTAL", "TOTAL COSTOS", "", tot_cos_1, tot_cos_2))
    estructura.append((
        "SUBTOTAL",
        "UTILIDAD / EXCEDENTE BRUTO",
        "",
        tot_ing_1 - tot_cos_1,
        tot_ing_2 - tot_cos_2,
    ))

  estructura.append(("TITULO", "GASTOS DE OPERACIÓN", "", None, None))
  items_gas = agrupar_bloque(gastos_op, es_acreedor=False)
  tot_gas_1 = sum(x[1] for x in items_gas)
  tot_gas_2 = sum(x[2] for x in items_gas)
  for nom, v1, v2 in items_gas:
    estructura.append(("CUENTA", nom, str(nota_idx), v1, v2))
    nota_idx += 1
  estructura.append(
      ("SUBTOTAL", "TOTAL GASTOS DE OPERACIÓN", "", tot_gas_1, tot_gas_2)
  )

  items_otr = agrupar_bloque(otros_res, es_acreedor=True)
  tot_otr_1 = sum(x[1] for x in items_otr)
  tot_otr_2 = sum(x[2] for x in items_otr)
  if items_otr:
    estructura.append(
        ("TITULO", "OTROS INGRESOS Y GASTOS NETOS", "", None, None)
    )
    for nom, v1, v2 in items_otr:
      estructura.append(("CUENTA", nom, str(nota_idx), v1, v2))
      nota_idx += 1
    estructura.append(
        ("SUBTOTAL", "TOTAL OTROS INGRESOS / GASTOS", "", tot_otr_1, tot_otr_2)
    )

  res_neto_1 = (tot_ing_1 - tot_cos_1) - tot_gas_1 + tot_otr_1
  res_neto_2 = (tot_ing_2 - tot_cos_2) - tot_gas_2 + tot_otr_2
  estructura.append((
      "TOTAL",
      "EXCEDENTES / RESULTADO DEL PERÍODO",
      "",
      res_neto_1,
      res_neto_2,
  ))

  return estructura


# ==============================================================================
# 5C. ESTADO DE FLUJOS DE EFECTIVO COMPARATIVO (MODELO ADJUNTO NIIF)
# ==============================================================================
def construir_estado_flujo_efectivo_comparativo(
    df_mov, df_catalogo, f1_ini, f1_fin, f2_ini, f2_fin
):
  cat_limpio = obtener_cuentas_analiticas_sin_duplicar(df_catalogo)
  cat = (
      cat_limpio[cat_limpio["es_hoja_real"] == True]
      .sort_values("cod_str")
      .copy()
  )
  cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()
  cat["nom_up"] = cat["nombre"].astype(str).str.upper().str.strip()
  cat["cod_norm"] = cat["cod_str"].str.lstrip("0")

  s1_ini = calcular_saldos_a_fecha(
      df_mov, f1_ini, incluir_asientos_del_dia=False
  )
  s1_fin = calcular_saldos_a_fecha(
      df_mov, f1_fin, incluir_asientos_del_dia=True
  )
  s2_ini = calcular_saldos_a_fecha(
      df_mov, f2_ini, incluir_asientos_del_dia=False
  )
  s2_fin = calcular_saldos_a_fecha(
      df_mov, f2_fin, incluir_asientos_del_dia=True
  )

  mov_res_1 = calcular_movimientos_periodo(
      df_mov, f1_ini, f1_fin, naturaleza_acreedora=True
  )
  mov_res_2 = calcular_movimientos_periodo(
      df_mov, f2_ini, f2_fin, naturaleza_acreedora=True
  )

  es_efectivo = cat["nom_up"].str.contains(
      "CAJA|BANCO|EFECTIVO|DISPONIBLE|DISPONIBILIDAD|FONDO"
  ) | cat["cod_norm"].str.startswith(("1.1.1", "1.1.01", "101", "111", "1-1-1"))

  es_activo = (
      cat["tipo_up"].str.contains("ACTIVO")
      | cat["cod_norm"].str.startswith("1")
  ) & (~es_efectivo)

  es_depreciacion_acum = es_activo & cat["nom_up"].str.contains(
      "DEPRECIACI|AMORTIZACI"
  )
  es_activo_fijo = (
      es_activo
      & (~es_depreciacion_acum)
      & (
          cat["cod_norm"].str.startswith(
              ("1.2", "12", "1-2", "102", "1.3", "13")
          )
          | cat["nom_up"].str.contains(
              "PROPIEDAD|PLANTA|EQUIPO|MOBILIARIO|VEHICULO|VEHÍCULO|EDIFICIO|TERRENO|INTANGIBLE|INVERSI"
          )
      )
  )
  es_activo_op = es_activo & (~es_depreciacion_acum) & (~es_activo_fijo)

  es_pasivo = (
      (
          cat["tipo_up"].str.contains("PASIVO")
          | cat["cod_norm"].str.startswith("2")
      )
      & (~es_activo)
      & (~es_efectivo)
  )
  es_pasivo_lp = es_pasivo & (
      cat["cod_norm"].str.startswith(("2.2", "22", "2-2", "202"))
      | cat["nom_up"].str.contains(
          "LARGO PLAZO|PRESTAMO|PRÉSTAMO|HIPOTECA|FINANCIAMIENTO"
      )
  )
  es_pasivo_op = es_pasivo & (~es_pasivo_lp)

  es_patrimonio = (
      (
          cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL")
          | cat["cod_norm"].str.startswith("3")
      )
      & (~es_activo)
      & (~es_pasivo)
      & (~es_efectivo)
  )

  es_resultado = cat["tipo_up"].str.contains(
      "INGRESO|GASTO|COSTO|EGRESO"
  ) | cat["cod_norm"].str.startswith(("4", "5", "6", "7", "8"))

  exc_1 = sum(mov_res_1.get(cid, 0.0) for cid in cat[es_resultado]["id"])
  exc_2 = sum(mov_res_2.get(cid, 0.0) for cid in cat[es_resultado]["id"])

  ctas_dep_gasto = cat[
      es_resultado & cat["nom_up"].str.contains("DEPRECIACI|AMORTIZACI")
  ]
  dep_1 = sum(-mov_res_1.get(cid, 0.0) for cid in ctas_dep_gasto["id"])
  dep_2 = sum(-mov_res_2.get(cid, 0.0) for cid in ctas_dep_gasto["id"])

  ctas_renta = cat[es_resultado & cat["nom_up"].str.contains("RENTA|IMPUESTO")]
  renta_1 = sum(-mov_res_1.get(cid, 0.0) for cid in ctas_renta["id"])
  renta_2 = sum(-mov_res_2.get(cid, 0.0) for cid in ctas_renta["id"])

  def var_activo_op(df_sub):
    v1 = sum(
        -(s1_fin.get(cid, 0.0) - s1_ini.get(cid, 0.0)) for cid in df_sub["id"]
    )
    v2 = sum(
        -(s2_fin.get(cid, 0.0) - s2_ini.get(cid, 0.0)) for cid in df_sub["id"]
    )
    return v1, v2

  def var_pasivo_op(df_sub):
    v1 = sum(
        -(s1_fin.get(cid, 0.0) - s1_ini.get(cid, 0.0)) for cid in df_sub["id"]
    )
    v2 = sum(
        -(s2_fin.get(cid, 0.0) - s2_ini.get(cid, 0.0)) for cid in df_sub["id"]
    )
    return v1, v2

  cxc = cat[
      es_activo_op & cat["nom_up"].str.contains("COBRAR|CLIENTE|DEUDOR")
  ]
  inv = cat[
      es_activo_op & cat["nom_up"].str.contains("INVENTARIO|MERCADER|ALMACEN")
  ]
  otros_ac = cat[
      es_activo_op & (~cat.index.isin(cxc.index)) & (~cat.index.isin(inv.index))
  ]

  cxp = cat[
      es_pasivo_op & cat["nom_up"].str.contains("PAGAR|PROVEEDOR|ACREEDOR")
  ]
  otros_pc = cat[es_pasivo_op & (~cat.index.isin(cxp.index))]

  v_cxc_1, v_cxc_2 = var_activo_op(cxc)
  v_inv_1, v_inv_2 = var_activo_op(inv)
  v_oac_1, v_oac_2 = var_activo_op(otros_ac)
  v_cxp_1, v_cxp_2 = var_pasivo_op(cxp)
  v_opc_1, v_opc_2 = var_pasivo_op(otros_pc)

  flujo_op_1 = (
      exc_1
      + dep_1
      + renta_1
      + v_cxc_1
      + v_inv_1
      + v_oac_1
      + v_cxp_1
      + v_opc_1
  )
  flujo_op_2 = (
      exc_2
      + dep_2
      + renta_2
      + v_cxc_2
      + v_inv_2
      + v_oac_2
      + v_cxp_2
      + v_opc_2
  )

  v_inv_af_1, v_inv_af_2 = var_activo_op(cat[es_activo_fijo])

  v_fin_lp_1, v_fin_lp_2 = var_pasivo_op(cat[es_pasivo_lp])
  v_fin_pat_1, v_fin_pat_2 = var_pasivo_op(cat[es_patrimonio])
  flujo_fin_1 = v_fin_lp_1 + v_fin_pat_1
  flujo_fin_2 = v_fin_lp_2 + v_fin_pat_2

  ef_ini_1 = sum(s1_ini.get(cid, 0.0) for cid in cat[es_efectivo]["id"])
  ef_fin_1 = sum(s1_fin.get(cid, 0.0) for cid in cat[es_efectivo]["id"])

  ef_ini_2 = sum(s2_ini.get(cid, 0.0) for cid in cat[es_efectivo]["id"])
  ef_fin_2 = sum(s2_fin.get(cid, 0.0) for cid in cat[es_efectivo]["id"])

  aumento_neto_1 = ef_fin_1 - ef_ini_1
  aumento_neto_2 = ef_fin_2 - ef_ini_2

  return [
      (
          "TITULO",
          "Flujo de efectivo por actividades de operación:",
          "",
          None,
          None,
      ),
      ("CUENTA", "Excedentes del Período", "", exc_1, exc_2),
      ("CUENTA", "Depreciación del Período", "", dep_1, dep_2),
      (
          "CUENTA",
          "Impuesto sobre la renta y provisiones",
          "",
          renta_1,
          renta_2,
      ),
      (
          "SUBTITULO",
          "Cambios en activos y pasivos operativos:",
          "",
          None,
          None,
      ),
      (
          "CUENTA",
          "Aumento/disminución en cuentas por cobrar",
          "",
          v_cxc_1,
          v_cxc_2,
      ),
      ("CUENTA", "Aumento/disminución en inventarios", "", v_inv_1, v_inv_2),
      (
          "CUENTA",
          "Aumento/disminución en otros activos circulantes",
          "",
          v_oac_1,
          v_oac_2,
      ),
      ("CUENTA", "Aumento/disminución en cuentas por pagar", "", v_cxp_1, v_cxp_2),
      (
          "CUENTA",
          "Aumento/disminución en otros pasivos circulantes",
          "",
          v_opc_1,
          v_opc_2,
      ),
      (
          "SUBTOTAL",
          "Efectivo neto proveniente de las actividades de operación",
          "",
          flujo_op_1,
          flujo_op_2,
      ),
      (
          "TITULO",
          "Flujo de efectivo por actividades de inversión:",
          "",
          None,
          None,
      ),
      (
          "CUENTA",
          "Adquisición / variación en propiedad, planta y equipo",
          "",
          v_inv_af_1,
          v_inv_af_2,
      ),
      (
          "SUBTOTAL",
          "Efectivo neto usado en las actividades de inversión",
          "",
          v_inv_af_1,
          v_inv_af_2,
      ),
      (
          "TITULO",
          "Flujo de efectivo por actividades de financiamiento:",
          "",
          None,
          None,
      ),
      (
          "CUENTA",
          "Variación en préstamos y obligaciones financieras",
          "",
          v_fin_lp_1,
          v_fin_lp_2,
      ),
      (
          "CUENTA",
          "Aportes de capital, reservas y patrimonio",
          "",
          v_fin_pat_1,
          v_fin_pat_2,
      ),
      (
          "SUBTOTAL",
          "Efectivo neto proveniente de actividades de financiamiento",
          "",
          flujo_fin_1,
          flujo_fin_2,
      ),
      (
          "SUBTOTAL",
          "Aumento neto en efectivo",
          "",
          aumento_neto_1,
          aumento_neto_2,
      ),
      ("CUENTA", "Efectivo al inicio del período", "", ef_ini_1, ef_ini_2),
      ("TOTAL", "Efectivo al final del período", "", ef_fin_1, ef_fin_2),
  ]


# ==============================================================================
# 5D. ESTADO DE CAMBIOS EN EL PATRIMONIO COMPARATIVO (NIIF)
# ==============================================================================
def construir_estado_cambios_patrimonio(
    df_mov, df_catalogo, f1_ini, f1_fin, f2_ini, f2_fin
):
  cat_limpio = obtener_cuentas_analiticas_sin_duplicar(df_catalogo)
  cat = (
      cat_limpio[cat_limpio["es_hoja_real"] == True]
      .sort_values("cod_str")
      .copy()
  )
  cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()
  cat["nom_up"] = cat["nombre"].astype(str).str.upper().str.strip()
  cat["cod_norm"] = cat["cod_str"].str.lstrip("0")

  es_pat = (
      cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL")
      | cat["cod_norm"].str.startswith("3")
  )
  es_res = cat["tipo_up"].str.contains(
      "INGRESO|GASTO|COSTO|EGRESO"
  ) | cat["cod_norm"].str.startswith(("4", "5", "6", "7", "8"))

  cat_pat = cat[es_pat].copy()
  cat_res = cat[es_res].copy()

  es_cap = cat_pat["nom_up"].str.contains("CAPITAL|APORTE|ACCION|CUOTA")
  es_resv = (~es_cap) & cat_pat["nom_up"].str.contains(
      "RESERVA|SUPERAVIT|SUPERÁVIT|REVALUACI"
  )
  es_util_acum = (~es_cap) & (~es_resv)

  s1_ini_deudor = calcular_saldos_a_fecha(
      df_mov, f1_ini, incluir_asientos_del_dia=False
  )
  s1_ini = {cid: -val for cid, val in s1_ini_deudor.items()}

  cap_ini_1 = sum(s1_ini.get(cid, 0.0) for cid in cat_pat[es_cap]["id"])
  resv_ini_1 = sum(s1_ini.get(cid, 0.0) for cid in cat_pat[es_resv]["id"])
  util_ini_1 = sum(s1_ini.get(cid, 0.0) for cid in cat_pat[es_util_acum]["id"])
  res_prev_1 = sum(s1_ini.get(cid, 0.0) for cid in cat_res["id"])
  util_ini_1 += res_prev_1
  tot_ini_1 = cap_ini_1 + resv_ini_1 + util_ini_1

  m1 = calcular_movimientos_periodo(
      df_mov, f1_ini, f1_fin, naturaleza_acreedora=True
  )
  d_cap_1 = sum(m1.get(cid, 0.0) for cid in cat_pat[es_cap]["id"])
  d_resv_1 = sum(m1.get(cid, 0.0) for cid in cat_pat[es_resv]["id"])
  d_util_1 = sum(m1.get(cid, 0.0) for cid in cat_pat[es_util_acum]["id"])
  exc_p1 = sum(m1.get(cid, 0.0) for cid in cat_res["id"])

  cap_fin_1 = cap_ini_1 + d_cap_1
  resv_fin_1 = resv_ini_1 + d_resv_1
  util_fin_1 = util_ini_1 + d_util_1 + exc_p1
  tot_fin_1 = cap_fin_1 + resv_fin_1 + util_fin_1

  m2 = calcular_movimientos_periodo(
      df_mov, f2_ini, f2_fin, naturaleza_acreedora=True
  )
  d_cap_2 = sum(m2.get(cid, 0.0) for cid in cat_pat[es_cap]["id"])
  d_resv_2 = sum(m2.get(cid, 0.0) for cid in cat_pat[es_resv]["id"])
  d_util_2 = sum(m2.get(cid, 0.0) for cid in cat_pat[es_util_acum]["id"])
  exc_p2 = sum(m2.get(cid, 0.0) for cid in cat_res["id"])

  cap_fin_2 = cap_fin_1 + d_cap_2
  resv_fin_2 = resv_fin_1 + d_resv_2
  util_fin_2 = util_fin_1 + d_util_2 + exc_p2
  tot_fin_2 = cap_fin_2 + resv_fin_2 + util_fin_2

  return [
      (
          "SALDO_INI",
          f"Saldos al {f1_ini.strftime('%d-%m-%Y')}",
          "",
          cap_ini_1,
          resv_ini_1,
          util_ini_1,
          tot_ini_1,
      ),
      (
          "MOV",
          "Aportes / aumentos de capital social",
          "",
          d_cap_1,
          0.0,
          0.0,
          d_cap_1,
      ),
      (
          "MOV",
          "Constitución / movimiento de reservas",
          "",
          0.0,
          d_resv_1,
          0.0,
          d_resv_1,
      ),
      (
          "MOV",
          "Ajustes / dividendos / movimientos en utilidades acumuladas",
          "",
          0.0,
          0.0,
          d_util_1,
          d_util_1,
      ),
      (
          "MOV",
          f"Excedentes / Resultado neto del período {f1_fin.year}",
          "",
          0.0,
          0.0,
          exc_p1,
          exc_p1,
      ),
      (
          "SUBTOTAL",
          f"Saldos al {f1_fin.strftime('%d-%m-%Y')}",
          "",
          cap_fin_1,
          resv_fin_1,
          util_fin_1,
          tot_fin_1,
      ),
      (
          "MOV",
          "Aportes / aumentos de capital social",
          "",
          d_cap_2,
          0.0,
          0.0,
          d_cap_2,
      ),
      (
          "MOV",
          "Constitución / movimiento de reservas",
          "",
          0.0,
          d_resv_2,
          0.0,
          d_resv_2,
      ),
      (
          "MOV",
          "Ajustes / dividendos / movimientos en utilidades acumuladas",
          "",
          0.0,
          0.0,
          d_util_2,
          d_util_2,
      ),
      (
          "MOV",
          f"Excedentes / Resultado neto del período {f2_fin.year}",
          "",
          0.0,
          0.0,
          exc_p2,
          exc_p2,
      ),
      (
          "TOTAL",
          f"Saldos al {f2_fin.strftime('%d-%m-%Y')}",
          "",
          cap_fin_2,
          resv_fin_2,
          util_fin_2,
          tot_fin_2,
      ),
  ]


def generar_word_cambios_patrimonio(
    subtitulo_periodo, estructura, empresa, cedula, rep_legal, cargo_rep
):
  doc = docx.Document()
  for section in doc.sections:
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.0)
    section.page_height = Inches(8.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_tit.paragraph_format.space_after = Pt(6)
  p_tit.add_run(f"{empresa.upper()}\n").bold = True
  p_tit.add_run(f"CÉDULA JURÍDICA {cedula}\n")
  p_tit.add_run("ESTADO DE CAMBIOS EN EL PATRIMONIO\n").bold = True
  p_tit.add_run(f"{subtitulo_periodo}\n").bold = True
  p_tit.add_run("(Expresado en colones)").italic = True

  table = doc.add_table(rows=1, cols=6)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER
  table.autofit = False
  anchos = [
      Inches(3.4),
      Inches(0.45),
      Inches(1.5),
      Inches(1.5),
      Inches(1.55),
      Inches(1.6),
  ]
  headers = [
      "CONCEPTO",
      "N",
      "Capital Social",
      "Reservas",
      "Excedentes / Utilidades Acumuladas",
      "Total Patrimonio",
  ]
  hdr_cells = table.rows[0].cells
  for i, h in enumerate(headers):
    hdr_cells[i].width = anchos[i]
    p = hdr_cells[i].paragraphs[0]
    p.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT
        if i >= 2
        else (WD_ALIGN_PARAGRAPH.CENTER if i == 1 else WD_ALIGN_PARAGRAPH.LEFT)
    )
    run = p.add_run(h)
    run.bold = True
    run.font.size = Pt(8.5)
    hdr_cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8EEF5"/>')
    )

  for tipo, concepto, nota, v_cap, v_res, v_uti, v_tot in estructura:
    row_cells = table.add_row().cells
    for i in range(6):
      row_cells[i].width = anchos[i]
      row_cells[i].paragraphs[0].paragraph_format.space_before = Pt(2)
      row_cells[i].paragraphs[0].paragraph_format.space_after = Pt(2)

    vals = [
        concepto,
        nota,
        fmt_moneda(v_cap),
        fmt_moneda(v_res),
        fmt_moneda(v_uti),
        fmt_moneda(v_tot),
    ]
    es_negrita = tipo in ("SALDO_INI", "SUBTOTAL", "TOTAL")
    for i, val in enumerate(vals):
      p = row_cells[i].paragraphs[0]
      p.alignment = (
          WD_ALIGN_PARAGRAPH.RIGHT
          if i >= 2
          else (WD_ALIGN_PARAGRAPH.CENTER if i == 1 else WD_ALIGN_PARAGRAPH.LEFT)
      )
      r = p.add_run(str(val))
      r.bold = es_negrita
      r.font.size = Pt(8.5)

    if tipo == "SUBTOTAL":
      for c in row_cells:
        c._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}><w:top w:val="single"'
                ' w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tcBorders>'
            )
        )
    elif tipo == "TOTAL":
      for c in row_cells:
        c._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}><w:top w:val="single"'
                ' w:sz="6"/><w:bottom w:val="double" w:sz="6"/></w:tcBorders>'
            )
        )

  doc.add_paragraph("\n")
  firma_table = doc.add_table(rows=1, cols=2)
  firma_table.alignment = WD_TABLE_ALIGNMENT.CENTER
  firma_table.autofit = False
  firma_table.rows[0].cells[0].width = Inches(5.0)
  firma_table.rows[0].cells[1].width = Inches(5.0)

  p_izq = firma_table.rows[0].cells[0].paragraphs[0]
  p_izq.alignment = WD_ALIGN_PARAGRAPH.LEFT
  r_izq = p_izq.add_run(
      "\n____________________________________\nLic. Roy F. Esquivel"
      " Gonzalez\nContador Privado Carné 8788"
  )
  r_izq.bold = True
  r_izq.font.size = Pt(9)

  p_der = firma_table.rows[0].cells[1].paragraphs[0]
  p_der.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_der = p_der.add_run(
      f"\n____________________________________\n{rep_legal}\n{cargo_rep}"
  )
  r_der.bold = True
  r_der.font.size = Pt(9)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_cambios_patrimonio(
    subtitulo_periodo, estructura, empresa, cedula, rep_legal, cargo_rep
):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Cambios en el Patrimonio"

  ws.merge_cells("A1:F1")
  ws["A1"] = empresa.upper()
  ws["A1"].font = Font(bold=True, size=13)
  ws["A1"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A2:F2")
  ws["A2"] = f"CÉDULA JURÍDICA {cedula}"
  ws["A2"].font = Font(size=10)
  ws["A2"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A3:F3")
  ws["A3"] = "ESTADO DE CAMBIOS EN EL PATRIMONIO"
  ws["A3"].font = Font(bold=True, size=12)
  ws["A3"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A4:F4")
  ws["A4"] = f"{subtitulo_periodo} (Expresado en colones)"
  ws["A4"].font = Font(italic=True, size=10)
  ws["A4"].alignment = Alignment(horizontal="center")

  ws.append([])
  headers = [
      "CONCEPTO",
      "N",
      "Capital Social",
      "Reservas",
      "Excedentes / Utilidades Acumuladas",
      "Total Patrimonio",
  ]
  ws.append(headers)
  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  for col in range(1, 7):
    c = ws.cell(row=6, column=col)
    c.font = Font(bold=True)
    c.fill = fill_hdr
    c.alignment = Alignment(
        horizontal="right" if col >= 3 else ("center" if col == 2 else "left")
    )

  fila = 7
  borde_sub = Border(top=Side(style="thin"), bottom=Side(style="thin"))
  borde_tot = Border(top=Side(style="thin"), bottom=Side(style="double"))

  for tipo, concepto, nota, v_cap, v_res, v_uti, v_tot in estructura:
    ws.cell(row=fila, column=1, value=concepto)
    ws.cell(row=fila, column=2, value=nota if nota else "").alignment = (
        Alignment(horizontal="center")
    )
    for idx_v, v in enumerate([v_cap, v_res, v_uti, v_tot], start=3):
      cell = ws.cell(row=fila, column=idx_v, value=float(v))
      cell.number_format = "#,##0.00;(#,##0.00);\"-\""

    if tipo in ("SALDO_INI", "SUBTOTAL", "TOTAL"):
      for col in range(1, 7):
        ws.cell(row=fila, column=col).font = Font(bold=True)
    if tipo == "SUBTOTAL":
      for col in range(1, 7):
        ws.cell(row=fila, column=col).border = borde_sub
    elif tipo == "TOTAL":
      for col in range(1, 7):
        ws.cell(row=fila, column=col).border = borde_tot
    fila += 1

  fila += 3
  ws.cell(
      row=fila,
      column=1,
      value=(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=fila, column=1).alignment = Alignment(
      wrap_text=True, horizontal="left"
  )

  ws.cell(
      row=fila,
      column=6,
      value=f"_______________________________\n{rep_legal}\n{cargo_rep}",
  ).font = Font(bold=True)
  ws.cell(row=fila, column=6).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[fila].height = 45

  for letra, ancho in zip(
      ["A", "B", "C", "D", "E", "F"], [48, 8, 22, 22, 26, 24]
  ):
    ws.column_dimensions[letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_cambios_patrimonio(
    subtitulo_periodo, estructura, empresa, cedula, rep_legal, cargo_rep
):
  pdf = FPDF(orientation="L", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=14)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(0, 6, limpiar_texto_pdf(empresa.upper()), ln=True, align="C")
  pdf.set_font("Helvetica", "", 9)
  pdf.cell(
      0, 5, limpiar_texto_pdf(f"CÉDULA JURÍDICA {cedula}"), ln=True, align="C"
  )
  pdf.set_font("Helvetica", "B", 10.5)
  pdf.cell(0, 5.5, "ESTADO DE CAMBIOS EN EL PATRIMONIO", ln=True, align="C")
  pdf.set_font("Helvetica", "B", 9)
  pdf.cell(0, 5, limpiar_texto_pdf(subtitulo_periodo), ln=True, align="C")
  pdf.set_font("Helvetica", "I", 8.5)
  pdf.cell(0, 5, "(Expresado en colones)", ln=True, align="C")
  pdf.ln(3)

  w = [88, 12, 38, 38, 42, 40]
  pdf.set_fill_color(230, 235, 245)
  pdf.set_font("Helvetica", "B", 7.8)
  headers = [
      "CONCEPTO",
      "N",
      "Capital Social",
      "Reservas",
      "Utilidades Acumuladas",
      "Total Patrimonio",
  ]
  for i, h in enumerate(headers):
    pdf.cell(
        w[i],
        6.5,
        limpiar_texto_pdf(h),
        border=1,
        align="R" if i >= 2 else ("C" if i == 1 else "L"),
        fill=True,
    )
  pdf.ln()

  for tipo, concepto, nota, v_cap, v_res, v_uti, v_tot in estructura:
    is_bold = tipo in ("SALDO_INI", "SUBTOTAL", "TOTAL")
    pdf.set_font("Helvetica", "B" if is_bold else "", 8)
    borde = "TB" if tipo in ("SUBTOTAL", "TOTAL") else 0
    pdf.cell(w[0], 5.5, limpiar_texto_pdf(concepto)[:52], border=borde)
    pdf.cell(
        w[1],
        5.5,
        limpiar_texto_pdf(nota) if nota else "",
        border=borde,
        align="C",
    )
    for idx_v, v in enumerate([v_cap, v_res, v_uti, v_tot], start=2):
      pdf.cell(w[idx_v], 5.5, fmt_moneda(v), border=borde, align="R")
    pdf.ln()

  pdf.ln(14)
  if pdf.get_y() > 175:
    pdf.add_page()
  y_firma = pdf.get_y()
  pdf.set_font("Helvetica", "B", 8.5)
  pdf.set_xy(12, y_firma)
  pdf.multi_cell(
      95,
      4.5,
      limpiar_texto_pdf(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
      align="L",
  )
  pdf.set_xy(170, y_firma)
  pdf.multi_cell(
      95,
      4.5,
      limpiar_texto_pdf(
          f"_______________________________\n{rep_legal}\n{cargo_rep}"
      ),
      align="R",
  )
  return bytes(pdf.output())


# ==============================================================================
# 5E. GENERADORES WORD, EXCEL Y PDF PARA ESTADOS FINANCIEROS VERTICALES
# ==============================================================================
def generar_word_estado_financiero(
    titulo_estado,
    subtitulo_periodo,
    estructura,
    empresa,
    cedula,
    col1_hdr,
    col2_hdr,
    rep_legal,
    cargo_rep,
):
  doc = docx.Document()
  for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_tit.paragraph_format.space_after = Pt(6)
  p_tit.add_run(f"{empresa.upper()}\n").bold = True
  p_tit.add_run(f"CÉDULA JURÍDICA {cedula}\n")
  p_tit.add_run(f"{titulo_estado.upper()}\n").bold = True
  p_tit.add_run(f"{subtitulo_periodo}\n").bold = True
  p_tit.add_run("(Expresado en colones)").italic = True

  table = doc.add_table(rows=1, cols=4)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER
  table.autofit = False
  anchos = [Inches(3.9), Inches(0.45), Inches(1.48), Inches(1.48)]

  hdr_cells = table.rows[0].cells
  headers = ["DESCRIPCIÓN", "N", col1_hdr, col2_hdr]
  for i, texto in enumerate(headers):
    hdr_cells[i].width = anchos[i]
    p = hdr_cells[i].paragraphs[0]
    p.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT
        if i >= 2
        else (WD_ALIGN_PARAGRAPH.CENTER if i == 1 else WD_ALIGN_PARAGRAPH.LEFT)
    )
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(8.5)
    hdr_cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8EEF5"/>')
    )

  for tipo, nombre, nota, v_ant, v_act in estructura:
    row_cells = table.add_row().cells
    for i in range(4):
      row_cells[i].width = anchos[i]
      row_cells[i].paragraphs[0].paragraph_format.space_before = Pt(1.5)
      row_cells[i].paragraphs[0].paragraph_format.space_after = Pt(1.5)

    p0, p1, p2, p3 = [c.paragraphs[0] for c in row_cells]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    r0 = p0.add_run(str(nombre))
    r1 = p1.add_run(str(nota) if nota else "")
    r2 = p2.add_run(fmt_moneda(v_ant))
    r3 = p3.add_run(fmt_moneda(v_act))

    es_negrita = tipo in ("TITULO", "SUBTITULO", "SUBTOTAL", "TOTAL")
    for r in (r0, r1, r2, r3):
      r.bold = es_negrita
      r.font.size = Pt(9 if tipo == "TOTAL" else 8.5)

    if tipo == "SUBTOTAL":
      for c in row_cells:
        c._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}><w:top w:val="single"'
                ' w:sz="4"/></w:tcBorders>'
            )
        )
    elif tipo == "TOTAL":
      for c in row_cells:
        c._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}><w:top w:val="single"'
                ' w:sz="6"/><w:bottom w:val="double" w:sz="6"/></w:tcBorders>'
            )
        )

  doc.add_paragraph("\n")
  firma_table = doc.add_table(rows=1, cols=2)
  firma_table.alignment = WD_TABLE_ALIGNMENT.CENTER
  firma_table.autofit = False
  firma_table.rows[0].cells[0].width = Inches(3.65)
  firma_table.rows[0].cells[1].width = Inches(3.65)

  p_izq = firma_table.rows[0].cells[0].paragraphs[0]
  p_izq.alignment = WD_ALIGN_PARAGRAPH.LEFT
  r_izq = p_izq.add_run(
      "\n____________________________________\nLic. Roy F. Esquivel"
      " Gonzalez\nContador Privado Carné 8788"
  )
  r_izq.bold = True
  r_izq.font.size = Pt(9)

  p_der = firma_table.rows[0].cells[1].paragraphs[0]
  p_der.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_der = p_der.add_run(
      f"\n____________________________________\n{rep_legal}\n{cargo_rep}"
  )
  r_der.bold = True
  r_der.font.size = Pt(9)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_estado_financiero(
    titulo_estado,
    subtitulo_periodo,
    estructura,
    empresa,
    cedula,
    col1_hdr,
    col2_hdr,
    rep_legal,
    cargo_rep,
):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = titulo_estado[:30]

  ws.merge_cells("A1:D1")
  ws["A1"] = empresa.upper()
  ws["A1"].font = Font(bold=True, size=13)
  ws["A1"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A2:D2")
  ws["A2"] = f"CÉDULA JURÍDICA {cedula}"
  ws["A2"].font = Font(size=10)
  ws["A2"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A3:D3")
  ws["A3"] = titulo_estado.upper()
  ws["A3"].font = Font(bold=True, size=12)
  ws["A3"].alignment = Alignment(horizontal="center")

  ws.merge_cells("A4:D4")
  ws["A4"] = f"{subtitulo_periodo} (Expresado en colones)"
  ws["A4"].font = Font(italic=True, size=10)
  ws["A4"].alignment = Alignment(horizontal="center")

  ws.append([])
  ws.append(["DESCRIPCIÓN", "N", col1_hdr, col2_hdr])
  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  for col in range(1, 5):
    cell = ws.cell(row=6, column=col)
    cell.font = Font(bold=True)
    cell.fill = fill_hdr
    cell.alignment = Alignment(
        horizontal="right" if col >= 3 else ("center" if col == 2 else "left")
    )

  fila = 7
  borde_sub = Border(top=Side(style="thin"))
  borde_tot = Border(top=Side(style="thin"), bottom=Side(style="double"))

  for tipo, nombre, nota, v_ant, v_act in estructura:
    ws.cell(row=fila, column=1, value=nombre)
    ws.cell(row=fila, column=2, value=nota if nota else "").alignment = (
        Alignment(horizontal="center")
    )
    c_ant = ws.cell(row=fila, column=3, value=v_ant if v_ant is not None else "")
    c_act = ws.cell(row=fila, column=4, value=v_act if v_act is not None else "")

    if v_ant is not None:
      c_ant.number_format = "#,##0.00;(#,##0.00);\"-\""
      c_act.number_format = "#,##0.00;(#,##0.00);\"-\""

    if tipo in ("TITULO", "SUBTITULO", "SUBTOTAL", "TOTAL"):
      for col in range(1, 5):
        ws.cell(row=fila, column=col).font = Font(bold=True)
    if tipo == "SUBTOTAL":
      for col in range(1, 5):
        ws.cell(row=fila, column=col).border = borde_sub
    elif tipo == "TOTAL":
      for col in range(1, 5):
        ws.cell(row=fila, column=col).border = borde_tot
    fila += 1

  fila += 3
  ws.cell(
      row=fila,
      column=1,
      value=(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=fila, column=1).alignment = Alignment(
      wrap_text=True, horizontal="left"
  )

  ws.cell(
      row=fila,
      column=4,
      value=f"_______________________________\n{rep_legal}\n{cargo_rep}",
  ).font = Font(bold=True)
  ws.cell(row=fila, column=4).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[fila].height = 45

  ws.column_dimensions["A"].width = 54
  ws.column_dimensions["B"].width = 8
  ws.column_dimensions["C"].width = 24
  ws.column_dimensions["D"].width = 24

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_estado_financiero(
    titulo_estado,
    subtitulo_periodo,
    estructura,
    empresa,
    cedula,
    col1_hdr,
    col2_hdr,
    rep_legal,
    cargo_rep,
):
  pdf = FPDF(orientation="P", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=15)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(0, 6, limpiar_texto_pdf(empresa.upper()), ln=True, align="C")
  pdf.set_font("Helvetica", "", 9)
  pdf.cell(
      0, 5, limpiar_texto_pdf(f"CÉDULA JURÍDICA {cedula}"), ln=True, align="C"
  )
  pdf.set_font("Helvetica", "B", 10.5)
  pdf.cell(0, 5.5, limpiar_texto_pdf(titulo_estado.upper()), ln=True, align="C")
  pdf.set_font("Helvetica", "B", 9)
  pdf.cell(0, 5, limpiar_texto_pdf(subtitulo_periodo), ln=True, align="C")
  pdf.set_font("Helvetica", "I", 8.5)
  pdf.cell(0, 5, "(Expresado en colones)", ln=True, align="C")
  pdf.ln(3)

  w = [96, 12, 42, 42]
  pdf.set_fill_color(230, 235, 245)
  pdf.set_font("Helvetica", "B", 7.8)
  pdf.cell(w[0], 6.5, "DESCRIPCION", border=1, fill=True)
  pdf.cell(w[1], 6.5, "N", border=1, align="C", fill=True)
  pdf.cell(w[2], 6.5, limpiar_texto_pdf(col1_hdr), border=1, align="R", fill=True)
  pdf.cell(w[3], 6.5, limpiar_texto_pdf(col2_hdr), border=1, align="R", fill=True)
  pdf.ln()

  for tipo, nombre, nota, v_ant, v_act in estructura:
    is_bold = tipo in ("TITULO", "SUBTITULO", "SUBTOTAL", "TOTAL")
    pdf.set_font("Helvetica", "B" if is_bold else "", 8.2)
    borde = "T" if tipo == "SUBTOTAL" else ("TB" if tipo == "TOTAL" else 0)
    pdf.cell(w[0], 5.2, limpiar_texto_pdf(nombre)[:56], border=borde)
    pdf.cell(
        w[1],
        5.2,
        limpiar_texto_pdf(nota) if nota else "",
        border=borde,
        align="C",
    )
    pdf.cell(w[2], 5.2, fmt_moneda(v_ant), border=borde, align="R")
    pdf.cell(w[3], 5.2, fmt_moneda(v_act), border=borde, align="R")
    pdf.ln()

  pdf.ln(14)
  if pdf.get_y() > 245:
    pdf.add_page()
  y_firma = pdf.get_y()
  pdf.set_font("Helvetica", "B", 8.5)
  pdf.set_xy(12, y_firma)
  pdf.multi_cell(
      85,
      4.5,
      limpiar_texto_pdf(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Carné 8788"
      ),
      align="L",
  )
  pdf.set_xy(115, y_firma)
  pdf.multi_cell(
      85,
      4.5,
      limpiar_texto_pdf(
          f"_______________________________\n{rep_legal}\n{cargo_rep}"
      ),
      align="R",
  )
  return bytes(pdf.output())


# ==============================================================================
# 5F. MÓDULO 10: FLUJO DE CAJA PROYECTADO (AÑO 1 MENSUAL + AÑOS 2..N ANUAL)
# ==============================================================================
def obtener_base_historica_proyeccion(df_mov, df_catalogo, f_base_ini, f_base_fin):
  cat_limpio = obtener_cuentas_analiticas_sin_duplicar(df_catalogo)
  cat = cat_limpio[cat_limpio["es_hoja_real"] == True].copy()
  cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()
  cat["nom_up"] = cat["nombre"].astype(str).str.upper().str.strip()
  cat["cod_norm"] = cat["cod_str"].str.lstrip("0")

  saldos_fin = calcular_saldos_a_fecha(
      df_mov, f_base_fin, incluir_asientos_del_dia=True
  )
  mov_per = calcular_movimientos_periodo(
      df_mov, f_base_ini, f_base_fin, naturaleza_acreedora=True
  )

  es_efectivo = cat["nom_up"].str.contains(
      "CAJA|BANCO|EFECTIVO|DISPONIBLE|DISPONIBILIDAD|FONDO"
  ) | cat["cod_norm"].str.startswith(("1.1.1", "1.1.01", "101", "111", "1-1-1"))
  caja_actual = sum(saldos_fin.get(cid, 0.0) for cid in cat[es_efectivo]["id"])

  ingresos = cat[
      cat["tipo_up"].str.contains("INGRESO")
      | cat["cod_norm"].str.startswith("4")
  ]
  costos = cat[
      (
          cat["tipo_up"].str.contains("COSTO")
          | cat["cod_norm"].str.startswith("5")
      )
      & (~cat.index.isin(ingresos.index))
  ]
  gastos = cat[
      (
          cat["tipo_up"].str.contains("GASTO|EGRESO")
          | cat["cod_norm"].str.startswith("6")
      )
      & (~cat.index.isin(ingresos.index))
      & (~cat.index.isin(costos.index))
  ]

  ventas_base = sum(mov_per.get(cid, 0.0) for cid in ingresos["id"])
  costos_base = sum(-mov_per.get(cid, 0.0) for cid in costos["id"])
  gastos_base = sum(-mov_per.get(cid, 0.0) for cid in gastos["id"])

  return {
      "caja_actual": float(caja_actual),
      "ventas_base": float(ventas_base),
      "costos_base": float(costos_base),
      "gastos_base": float(gastos_base),
  }


def calcular_cuota_nivelada(monto_prestamo, tasa_anual_pct, plazo_anos):
  if monto_prestamo <= 0 or plazo_anos <= 0:
    return 0.0
  n_meses = int(plazo_anos) * 12
  i_mensual = (float(tasa_anual_pct) / 100.0) / 12.0
  if i_mensual == 0:
    return monto_prestamo / n_meses
  factor = (i_mensual * ((1 + i_mensual) ** n_meses)) / (
      ((1 + i_mensual) ** n_meses) - 1
  )
  return monto_prestamo * factor


def construir_flujo_caja_proyectado(
    anio_inicio_proy,
    anios_proyectar,
    caja_inicial,
    ventas_ano1_total,
    crec_mensual_ventas_pct,
    crec_anual_ventas_pct,
    pct_costo_ventas,
    gastos_mensuales_ano1,
    crec_anual_gastos_pct,
    monto_prestamo,
    tasa_anual_pct,
    plazo_prestamo_anos,
):
  cuota_mensual = calcular_cuota_nivelada(
      monto_prestamo, tasa_anual_pct, plazo_prestamo_anos
  )

  g_m = crec_mensual_ventas_pct / 100.0
  if round(g_m, 8) == 0:
    v_mes1 = ventas_ano1_total / 12.0
  else:
    suma_factores = sum((1 + g_m) ** m for m in range(12))
    v_mes1 = ventas_ano1_total / suma_factores

  ventas_m = [v_mes1 * ((1 + g_m) ** m) for m in range(12)]
  costos_m = [v * (pct_costo_ventas / 100.0) for v in ventas_m]
  gastos_m = [float(gastos_mensuales_ano1) for _ in range(12)]
  tot_egr_m = [costos_m[i] + gastos_m[i] for i in range(12)]
  ebitda_m = [ventas_m[i] - tot_egr_m[i] for i in range(12)]
  cuota_m = [cuota_mensual if plazo_prestamo_anos >= 1 else 0.0 for _ in range(12)]
  flujo_neto_m = [ebitda_m[i] - cuota_m[i] for i in range(12)]

  saldo_ini_m = []
  saldo_fin_m = []
  acum = float(caja_inicial)
  for i in range(12):
    saldo_ini_m.append(acum)
    acum += flujo_neto_m[i]
    saldo_fin_m.append(acum)

  icsd_m = [
      (ebitda_m[i] / cuota_m[i]) if cuota_m[i] > 0 else 0.0 for i in range(12)
  ]

  tot_v_1 = sum(ventas_m)
  tot_c_1 = sum(costos_m)
  tot_g_1 = sum(gastos_m)
  tot_e_1 = sum(tot_egr_m)
  tot_eb_1 = sum(ebitda_m)
  tot_cu_1 = sum(cuota_m)
  tot_fn_1 = sum(flujo_neto_m)
  icsd_tot_1 = (tot_eb_1 / tot_cu_1) if tot_cu_1 > 0 else 0.0

  sec1_filas = [
      ("TITULO", "INGRESOS OPERATIVOS", [None] * 13),
      ("CUENTA", "Ventas Netas Mensuales", ventas_m + [tot_v_1]),
      ("SUBTOTAL", "TOTAL INGRESOS OPERATIVOS", ventas_m + [tot_v_1]),
      ("TITULO", "EGRESOS OPERATIVOS DE CAJA", [None] * 13),
      (
          "CUENTA",
          f"Costo de Ventas ({pct_costo_ventas:.2f}%)",
          costos_m + [tot_c_1],
      ),
      ("CUENTA", "Gastos Operativos/Admin", gastos_m + [tot_g_1]),
      ("SUBTOTAL", "TOTAL EGRESOS OPERATIVOS", tot_egr_m + [tot_e_1]),
      ("SUBTOTAL", "FLUJO OPERATIVO BRUTO (EBITDA)", ebitda_m + [tot_eb_1]),
      (
          "TITULO",
          f"SERVICIO DE LA DEUDA BANCARIA ({plazo_prestamo_anos} AÑOS PLAZO)",
          [None] * 13,
      ),
      (
          "CUENTA",
          f"Cuota Mensual ({tasa_anual_pct:.2f}% Anual)",
          cuota_m + [tot_cu_1],
      ),
      ("SUBTOTAL", "FLUJO DE CAJA NETO MENSUAL", flujo_neto_m + [tot_fn_1]),
      ("CUENTA", "Saldo Inicial de Efectivo", saldo_ini_m + [caja_inicial]),
      (
          "TOTAL",
          "SALDO FINAL ACUMULADO CAJA",
          saldo_fin_m + [saldo_fin_m[-1]],
      ),
      ("RATIO", "ÍNDICE COBERTURA DEUDA (ICSD)", icsd_m + [icsd_tot_1]),
  ]

  n_anios = max(2, int(anios_proyectar))
  cols_sec2 = [
      f"Año {k} ({anio_inicio_proy + k - 1})" for k in range(2, n_anios + 1)
  ]

  ventas_a, costos_a, gastos_a, tot_egr_a, ebitda_a = [], [], [], [], []
  cuota_a, flujo_neto_a, saldo_ini_a, saldo_fin_a, icsd_a = [], [], [], [], []

  v_prev = tot_v_1
  g_prev = tot_g_1
  acum_a = saldo_fin_m[-1]

  for k in range(2, n_anios + 1):
    v_curr = v_prev * (1 + crec_anual_ventas_pct / 100.0)
    c_curr = v_curr * (pct_costo_ventas / 100.0)
    g_curr = g_prev * (1 + crec_anual_gastos_pct / 100.0)
    egr_curr = c_curr + g_curr
    eb_curr = v_curr - egr_curr
    cu_curr = (
        (cuota_mensual * 12.0)
        if (monto_prestamo > 0 and k <= int(plazo_prestamo_anos))
        else 0.0
    )
    fn_curr = eb_curr - cu_curr

    s_ini_curr = acum_a
    acum_a += fn_curr
    s_fin_curr = acum_a
    ic_curr = (eb_curr / cu_curr) if cu_curr > 0 else 0.0

    ventas_a.append(v_curr)
    costos_a.append(c_curr)
    gastos_a.append(g_curr)
    tot_egr_a.append(egr_curr)
    ebitda_a.append(eb_curr)
    cuota_a.append(cu_curr)
    flujo_neto_a.append(fn_curr)
    saldo_ini_a.append(s_ini_curr)
    saldo_fin_a.append(s_fin_curr)
    icsd_a.append(ic_curr)

    v_prev = v_curr
    g_prev = g_curr

  n_c2 = len(cols_sec2)
  sec2_filas = [
      ("TITULO", "INGRESOS OPERATIVOS", [None] * n_c2),
      (
          "CUENTA",
          f"Ventas Netas (+{crec_anual_ventas_pct:.1f}% anual)",
          ventas_a,
      ),
      ("SUBTOTAL", "TOTAL INGRESOS", ventas_a),
      ("TITULO", "EGRESOS OPERATIVOS DE CAJA", [None] * n_c2),
      ("CUENTA", f"Costo de Ventas ({pct_costo_ventas:.2f}%)", costos_a),
      (
          "CUENTA",
          f"Gastos Operativos (+{crec_anual_gastos_pct:.1f}%)",
          gastos_a,
      ),
      ("SUBTOTAL", "TOTAL EGRESOS OPERATIVOS", tot_egr_a),
      ("SUBTOTAL", "FLUJO OPERATIVO (EBITDA)", ebitda_a),
      (
          "TITULO",
          f"SERVICIO DE LA DEUDA BANCARIA ({plazo_prestamo_anos} AÑOS PLAZO)",
          [None] * n_c2,
      ),
      ("CUENTA", "Amortización e Intereses", cuota_a),
      ("SUBTOTAL", "FLUJO CAJA NETO ANUAL", flujo_neto_a),
      ("CUENTA", "Saldo Inicial Efectivo", saldo_ini_a),
      ("TOTAL", "SALDO FINAL ACUMULADO CAJA", saldo_fin_a),
      ("RATIO", "ÍNDICE COBERTURA (ICSD)", icsd_a),
  ]

  return {
      "cuota_mensual": cuota_mensual,
      "sec1_cols": [f"Mes {i}" for i in range(1, 13)] + ["TOTAL AÑO 1"],
      "sec1_filas": sec1_filas,
      "sec2_cols": cols_sec2,
      "sec2_filas": sec2_filas,
      "caja_final_proyectada": saldo_fin_a[-1] if saldo_fin_a else saldo_fin_m[-1],
  }


def generar_word_flujo_proyectado(
    res_proy,
    empresa,
    cedula,
    rep_legal,
    anios_proyectar,
    monto_prestamo,
    tasa_anual_pct,
    plazo_prestamo_anos,
):
  doc = docx.Document()
  for section in doc.sections:
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.0)
    section.page_height = Inches(8.5)
    section.top_margin = Inches(0.38)
    section.bottom_margin = Inches(0.38)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
  p_tit.paragraph_format.space_after = Pt(4)
  r0 = p_tit.add_run(
      f"MODELO DE FLUJO DE CAJA PROYECTADO A {anios_proyectar} AÑOS\n"
  )
  r0.bold = True
  r0.font.size = Pt(11)
  r1 = p_tit.add_run(
      f"Año 1 Desglosado Mensualmente + Proyección Anualizada Años 2 a"
      f" {anios_proyectar} (Tasa {tasa_anual_pct:.2f}% Anual a"
      f" {plazo_prestamo_anos} Años Plazo)\n"
  )
  r1.font.size = Pt(8.5)
  r2 = p_tit.add_run(
      f"ENTIDAD: {empresa.upper()} | CÉDULA JURÍDICA: {cedula} | PRÉSTAMO:"
      f" ₡{monto_prestamo:,.2f} | CUOTA MENSUAL:"
      f" ₡{res_proy['cuota_mensual']:,.2f}"
  )
  r2.bold = True
  r2.font.size = Pt(8.5)

  def agregar_tabla_seccion(titulo_sec, cols, filas, ancho_primera=1.85):
    p_s = doc.add_paragraph()
    p_s.paragraph_format.space_before = Pt(4)
    p_s.paragraph_format.space_after = Pt(2)
    rs = p_s.add_run(titulo_sec)
    rs.bold = True
    rs.font.size = Pt(9)

    n_cols = len(cols) + 1
    ancho_restante = 10.3 - ancho_primera
    ancho_num = ancho_restante / max(1, len(cols))

    t = doc.add_table(rows=1, cols=n_cols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False

    hdr = t.rows[0].cells
    hdr[0].width = Inches(ancho_primera)
    hdr[0]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="1F4E78"/>')
    )
    ph0 = hdr[0].paragraphs[0]
    rh0 = ph0.add_run("Concepto / Periodo")
    rh0.bold = True
    rh0.font.size = Pt(6.5)

    for idx_c, c_name in enumerate(cols, start=1):
      hdr[idx_c].width = Inches(ancho_num)
      hdr[idx_c]._tc.get_or_add_tcPr().append(
          parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
      )
      ph = hdr[idx_c].paragraphs[0]
      ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
      rh = ph.add_run(c_name)
      rh.bold = True
      rh.font.size = Pt(6.2)

    for tipo, concepto, vals in filas:
      rc = t.add_row().cells
      rc[0].width = Inches(ancho_primera)
      p0 = rc[0].paragraphs[0]
      p0.paragraph_format.space_before = Pt(0.5)
      p0.paragraph_format.space_after = Pt(0.5)
      r_c = p0.add_run(concepto)
      es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL", "RATIO")
      r_c.bold = es_b
      r_c.font.size = Pt(6.3)

      for idx_v, val in enumerate(vals, start=1):
        rc[idx_v].width = Inches(ancho_num)
        pv = rc[idx_v].paragraphs[0]
        pv.paragraph_format.space_before = Pt(0.5)
        pv.paragraph_format.space_after = Pt(0.5)
        pv.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if val is None:
          txt = ""
        elif tipo == "RATIO":
          txt = f"{val:.2f}x" if val > 0 else "N/A"
        else:
          txt = fmt_moneda_entero(val)
        rv = pv.add_run(txt)
        rv.bold = es_b
        rv.font.size = Pt(6.2)

      if tipo in ("SUBTOTAL", "TOTAL"):
        for c in rc:
          c._tc.get_or_add_tcPr().append(
              parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F8"/>')
          )

  agregar_tabla_seccion(
      "SECCIÓN I: FLUJO DE CAJA DETALLADO MES A MES - AÑO 1 (EN COLONES)",
      res_proy["sec1_cols"],
      res_proy["sec1_filas"],
  )
  doc.add_page_break()
  agregar_tabla_seccion(
      f"SECCIÓN II: FLUJO DE CAJA PROYECTADO ANUALIZADO (AÑOS 2 A"
      f" {anios_proyectar} - EN COLONES)",
      res_proy["sec2_cols"],
      res_proy["sec2_filas"],
  )

  doc.add_paragraph("\n")
  ft = doc.add_table(rows=1, cols=2)
  ft.alignment = WD_TABLE_ALIGNMENT.CENTER
  ft.autofit = False
  ft.rows[0].cells[0].width = Inches(5.1)
  ft.rows[0].cells[1].width = Inches(5.1)

  pi = ft.rows[0].cells[0].paragraphs[0]
  ri = pi.add_run(
      "\n____________________________________\nLic. Roy F. Esquivel"
      " Gonzalez\nContador Privado Incorporado #8788"
  )
  ri.bold = True
  ri.font.size = Pt(8.5)

  pd_cell = ft.rows[0].cells[1].paragraphs[0]
  pd_cell.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  rd = pd_cell.add_run(
      f"\n____________________________________\n{rep_legal}\nGerente General"
      " y Representante Legal"
  )
  rd.bold = True
  rd.font.size = Pt(8.5)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_flujo_proyectado(
    res_proy,
    empresa,
    cedula,
    rep_legal,
    anios_proyectar,
    monto_prestamo,
    tasa_anual_pct,
    plazo_prestamo_anos,
):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Flujo Caja Proyectado"

  ws["A1"] = f"MODELO DE FLUJO DE CAJA PROYECTADO A {anios_proyectar} AÑOS"
  ws["A1"].font = Font(bold=True, size=13)
  ws["A2"] = (
      f"{empresa.upper()} | CÉDULA {cedula} | PRÉSTAMO: ₡{monto_prestamo:,.2f} |"
      f" TASA: {tasa_anual_pct:.2f}% | PLAZO: {plazo_prestamo_anos} AÑOS |"
      f" CUOTA MENSUAL: ₡{res_proy['cuota_mensual']:,.2f}"
  )
  ws["A2"].font = Font(bold=True, size=10)

  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  fill_sub = PatternFill(
      start_color="F2F4F8", end_color="F2F4F8", fill_type="solid"
  )

  def escribir_bloque(fila_ini, titulo, cols, filas):
    ws.cell(row=fila_ini, column=1, value=titulo).font = Font(
        bold=True, size=11
    )
    r_hdr = fila_ini + 1
    ws.cell(row=r_hdr, column=1, value="Concepto / Periodo").font = Font(
        bold=True
    )
    ws.cell(row=r_hdr, column=1).fill = fill_hdr
    for idx_c, cname in enumerate(cols, start=2):
      c = ws.cell(row=r_hdr, column=idx_c, value=cname)
      c.font = Font(bold=True)
      c.fill = fill_hdr
      c.alignment = Alignment(horizontal="right")

    r_act = r_hdr + 1
    for tipo, concepto, vals in filas:
      c0 = ws.cell(row=r_act, column=1, value=concepto)
      es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL", "RATIO")
      if es_b:
        c0.font = Font(bold=True)
      for idx_v, v in enumerate(vals, start=2):
        cv = ws.cell(row=r_act, column=idx_v)
        if v is None:
          cv.value = ""
        elif tipo == "RATIO":
          cv.value = f"{v:.2f}x" if v > 0 else "N/A"
          cv.alignment = Alignment(horizontal="right")
        else:
          cv.value = float(v)
          cv.number_format = "#,##0"
        if es_b:
          cv.font = Font(bold=True)
        if tipo in ("SUBTOTAL", "TOTAL"):
          cv.fill = fill_sub
      r_act += 1
    return r_act + 2

  sig_fila = escribir_bloque(
      4,
      "SECCIÓN I: FLUJO DE CAJA DETALLADO MES A MES - AÑO 1",
      res_proy["sec1_cols"],
      res_proy["sec1_filas"],
  )
  fin_fila = escribir_bloque(
      sig_fila,
      f"SECCIÓN II: FLUJO DE CAJA PROYECTADO ANUALIZADO (AÑOS 2 A"
      f" {anios_proyectar})",
      res_proy["sec2_cols"],
      res_proy["sec2_filas"],
  )

  ws.cell(
      row=fin_fila + 1,
      column=1,
      value=(
          "_______________________________\nLic. Roy F. Esquivel"
          " Gonzalez\nContador Privado Incorporado #8788"
      ),
  ).font = Font(bold=True)
  ws.cell(row=fin_fila + 1, column=1).alignment = Alignment(wrap_text=True)

  ws.cell(
      row=fin_fila + 1,
      column=8,
      value=(
          f"_______________________________\n{rep_legal}\nRepresentante Legal"
      ),
  ).font = Font(bold=True)
  ws.cell(row=fin_fila + 1, column=8).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[fin_fila + 1].height = 45

  ws.column_dimensions["A"].width = 38
  for col_i in range(2, 16):
    letra = openpyxl.utils.get_column_letter(col_i)
    ws.column_dimensions[letra].width = 16

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_flujo_proyectado(
    res_proy,
    empresa,
    cedula,
    rep_legal,
    anios_proyectar,
    monto_prestamo,
    tasa_anual_pct,
    plazo_prestamo_anos,
):
  pdf = FPDF(orientation="L", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=10)

  def dibujar_seccion(titulo_sec, cols, filas, incluir_firmas=False):
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.cell(
        0,
        5,
        limpiar_texto_pdf(
            f"MODELO DE FLUJO DE CAJA PROYECTADO A {anios_proyectar} AÑOS"
        ),
        ln=True,
    )
    pdf.set_font("Helvetica", "", 8)
    pdf.cell(
        0,
        4.5,
        limpiar_texto_pdf(
            f"{empresa.upper()} | CÉDULA {cedula} | PRÉSTAMO:"
            f" C{monto_prestamo:,.0f} | TASA: {tasa_anual_pct:.2f}% | PLAZO:"
            f" {plazo_prestamo_anos} AÑOS | CUOTA:"
            f" C{res_proy['cuota_mensual']:,.0f}"
        ),
        ln=True,
    )
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(0, 5, limpiar_texto_pdf(titulo_sec), ln=True)
    pdf.ln(1)

    w0 = 52
    wn = (260 - w0) / max(1, len(cols))
    pdf.set_fill_color(225, 232, 245)
    pdf.set_font("Helvetica", "B", 5.8)
    pdf.cell(w0, 5.5, "Concepto / Periodo", border=1, fill=True)
    for cname in cols:
      pdf.cell(
          wn, 5.5, limpiar_texto_pdf(cname)[:12], border=1, align="R", fill=True
      )
    pdf.ln()

    for tipo, concepto, vals in filas:
      es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL", "RATIO")
      pdf.set_font("Helvetica", "B" if es_b else "", 5.6)
      fill_r = tipo in ("SUBTOTAL", "TOTAL")
      if fill_r:
        pdf.set_fill_color(242, 245, 250)
      pdf.cell(w0, 4.8, limpiar_texto_pdf(concepto)[:36], border="B", fill=fill_r)
      for v in vals:
        if v is None:
          txt = ""
        elif tipo == "RATIO":
          txt = f"{v:.2f}x" if v > 0 else "N/A"
        else:
          txt = fmt_moneda_entero(v)
        pdf.cell(wn, 4.8, txt, border="B", align="R", fill=fill_r)
      pdf.ln()

    if incluir_firmas:
      pdf.ln(10)
      y_f = pdf.get_y()
      pdf.set_font("Helvetica", "B", 8)
      pdf.set_xy(10, y_f)
      pdf.multi_cell(
          95,
          4.2,
          limpiar_texto_pdf(
              "_______________________________\nLic. Roy F. Esquivel"
              " Gonzalez\nContador Privado Incorporado #8788"
          ),
          align="L",
      )
      pdf.set_xy(170, y_f)
      pdf.multi_cell(
          95,
          4.2,
          limpiar_texto_pdf(
              f"_______________________________\n{rep_legal}\nRepresentante Legal"
          ),
          align="R",
      )

  dibujar_seccion(
      "SECCIÓN I: FLUJO DE CAJA DETALLADO MES A MES - AÑO 1 (EN COLONES)",
      res_proy["sec1_cols"],
      res_proy["sec1_filas"],
      incluir_firmas=False,
  )
  dibujar_seccion(
      f"SECCIÓN II: FLUJO DE CAJA PROYECTADO ANUALIZADO (AÑOS 2 A"
      f" {anios_proyectar} - EN COLONES)",
      res_proy["sec2_cols"],
      res_proy["sec2_filas"],
      incluir_firmas=True,
  )
  return bytes(pdf.output())


# ==============================================================================
# 5G. MÓDULO 11: ANÁLISIS FINANCIERO INTEGRAL (VERTICAL, HORIZONTAL, RAZONES)
# ==============================================================================
def construir_analisis_financiero_integral(
    df_mov, df_catalogo, p1_ini, p1_fin, p2_ini, p2_fin
):
  cat_limpio = obtener_cuentas_analiticas_sin_duplicar(df_catalogo)
  cat = (
      cat_limpio[cat_limpio["es_hoja_real"] == True]
      .sort_values("cod_str")
      .copy()
  )
  cat["tipo_up"] = cat["tipo"].astype(str).str.upper().str.strip()
  cat["nom_up"] = cat["nombre"].astype(str).str.upper().str.strip()
  cat["cod_norm"] = cat["cod_str"].str.lstrip("0")

  s_ant = calcular_saldos_a_fecha(
      df_mov, p1_fin, incluir_asientos_del_dia=True
  )
  s_act = calcular_saldos_a_fecha(
      df_mov, p2_fin, incluir_asientos_del_dia=True
  )
  m_ant = calcular_movimientos_periodo(
      df_mov, p1_ini, p1_fin, naturaleza_acreedora=True
  )
  m_act = calcular_movimientos_periodo(
      df_mov, p2_ini, p2_fin, naturaleza_acreedora=True
  )

  # --- Clasificación Patrimonial ---
  activos = cat[
      cat["tipo_up"].str.contains("ACTIVO")
      | cat["cod_norm"].str.startswith("1")
  ]
  act_circ = activos[
      activos["cod_norm"].str.startswith(("1.1", "11", "1-1", "101"))
  ]
  act_no_circ = activos[~activos.index.isin(act_circ.index)]
  if act_circ.empty and not activos.empty:
    act_circ = activos

  efectivo = act_circ[
      act_circ["nom_up"].str.contains(
          "CAJA|BANCO|EFECTIVO|DISPONIBLE|DISPONIBILIDAD|FONDO"
      )
      | act_circ["cod_norm"].str.startswith(("1.1.1", "1.1.01", "101", "111"))
  ]
  cxc = act_circ[
      act_circ["nom_up"].str.contains("COBRAR|CLIENTE|DEUDOR")
      & (~act_circ.index.isin(efectivo.index))
  ]
  inventarios = act_circ[
      act_circ["nom_up"].str.contains("INVENTARIO|MERCADER|ALMACEN")
      & (~act_circ.index.isin(efectivo.index))
  ]

  pasivos = cat[
      (
          cat["tipo_up"].str.contains("PASIVO")
          | cat["cod_norm"].str.startswith("2")
      )
      & (~cat.index.isin(activos.index))
  ]
  pas_circ = pasivos[
      pasivos["cod_norm"].str.startswith(("2.1", "21", "2-1", "201"))
  ]
  pas_lp = pasivos[~pasivos.index.isin(pas_circ.index)]
  if pas_circ.empty and not pasivos.empty:
    pas_circ = pasivos

  patrimonio = cat[
      (
          cat["tipo_up"].str.contains("PATRIMONIO|CAPITAL")
          | cat["cod_norm"].str.startswith("3")
      )
      & (~cat.index.isin(activos.index))
      & (~cat.index.isin(pasivos.index))
  ]

  ingresos = cat[
      cat["tipo_up"].str.contains("INGRESO")
      | cat["cod_norm"].str.startswith("4")
  ]
  costos = cat[
      (
          cat["tipo_up"].str.contains("COSTO")
          | cat["cod_norm"].str.startswith("5")
      )
      & (~cat.index.isin(ingresos.index))
  ]
  gastos_op = cat[
      (
          cat["tipo_up"].str.contains("GASTO|EGRESO")
          | cat["cod_norm"].str.startswith("6")
      )
      & (~cat.index.isin(ingresos.index))
      & (~cat.index.isin(costos.index))
  ]
  gastos_fin = gastos_op[
      gastos_op["nom_up"].str.contains("INTERES|INTERÉS|FINANCIER|BANCARI")
  ]
  otros_res = cat[
      cat["cod_norm"].str.startswith(("7", "8"))
      & (~cat.index.isin(ingresos.index))
      & (~cat.index.isin(costos.index))
      & (~cat.index.isin(gastos_op.index))
  ]
  ing_gast_all = cat[
      cat.index.isin(ingresos.index)
      | cat.index.isin(costos.index)
      | cat.index.isin(gastos_op.index)
      | cat.index.isin(otros_res.index)
  ]

  def suma_deudora(df_sub, dic_s):
    return sum(dic_s.get(cid, 0.0) for cid in df_sub["id"])

  def suma_acreedora(df_sub, dic_s):
    return sum(-dic_s.get(cid, 0.0) for cid in df_sub["id"])

  ef_1, ef_2 = suma_deudora(efectivo, s_ant), suma_deudora(efectivo, s_act)
  cxc_1, cxc_2 = suma_deudora(cxc, s_ant), suma_deudora(cxc, s_act)
  inv_1, inv_2 = suma_deudora(inventarios, s_ant), suma_deudora(inventarios, s_act)
  ac_1, ac_2 = suma_deudora(act_circ, s_ant), suma_deudora(act_circ, s_act)
  anc_1, anc_2 = suma_deudora(act_no_circ, s_ant), suma_deudora(act_no_circ, s_act)
  at_1, at_2 = ac_1 + anc_1, ac_2 + anc_2

  pc_1, pc_2 = suma_acreedora(pas_circ, s_ant), suma_acreedora(pas_circ, s_act)
  plp_1, plp_2 = suma_acreedora(pas_lp, s_ant), suma_acreedora(pas_lp, s_act)
  pt_1, pt_2 = pc_1 + plp_1, pc_2 + plp_2

  pat_cuentas_1 = suma_acreedora(patrimonio, s_ant)
  pat_cuentas_2 = suma_acreedora(patrimonio, s_act)
  res_acum_1 = suma_acreedora(ing_gast_all, s_ant)
  res_acum_2 = suma_acreedora(ing_gast_all, s_act)
  pat_1 = pat_cuentas_1 + res_acum_1
  pat_2 = pat_cuentas_2 + res_acum_2
  pas_pat_1 = pt_1 + pat_1
  pas_pat_2 = pt_2 + pat_2

  v_1 = sum(m_ant.get(cid, 0.0) for cid in ingresos["id"])
  v_2 = sum(m_act.get(cid, 0.0) for cid in ingresos["id"])
  cv_1 = sum(-m_ant.get(cid, 0.0) for cid in costos["id"])
  cv_2 = sum(-m_act.get(cid, 0.0) for cid in costos["id"])
  ub_1, ub_2 = v_1 - cv_1, v_2 - cv_2
  go_1 = sum(-m_ant.get(cid, 0.0) for cid in gastos_op["id"])
  go_2 = sum(-m_act.get(cid, 0.0) for cid in gastos_op["id"])
  gf_1 = sum(-m_ant.get(cid, 0.0) for cid in gastos_fin["id"])
  gf_2 = sum(-m_act.get(cid, 0.0) for cid in gastos_fin["id"])
  uo_1, uo_2 = ub_1 - go_1, ub_2 - go_2
  otr_1 = sum(m_ant.get(cid, 0.0) for cid in otros_res["id"])
  otr_2 = sum(m_act.get(cid, 0.0) for cid in otros_res["id"])
  un_1, un_2 = uo_1 + otr_1, uo_2 + otr_2

  def calc_vh(val1, val2, base1, base2):
    av1 = (val1 / base1 * 100.0) if round(base1, 2) != 0 else 0.0
    av2 = (val2 / base2 * 100.0) if round(base2, 2) != 0 else 0.0
    ah_abs = val2 - val1
    ah_rel = (
        ((val2 - val1) / abs(val1) * 100.0)
        if round(val1, 2) != 0
        else (100.0 if round(val2, 2) != 0 else 0.0)
    )
    return av1, av2, ah_abs, ah_rel

  rubros_sf = [
      ("TITULO", "ACTIVOS", None, None),
      ("CUENTA", "Efectivo y Equivalentes", ef_1, ef_2),
      ("CUENTA", "Cuentas por Cobrar", cxc_1, cxc_2),
      ("CUENTA", "Inventarios", inv_1, inv_2),
      ("SUBTOTAL", "Total Activo Circulante", ac_1, ac_2),
      (
          "CUENTA",
          "Activo No Circulante / Propiedad, Planta y Equipo",
          anc_1,
          anc_2,
      ),
      ("TOTAL", "TOTAL ACTIVOS", at_1, at_2),
      ("TITULO", "PASIVOS Y PATRIMONIO", None, None),
      ("CUENTA", "Pasivo Circulante (Corto Plazo)", pc_1, pc_2),
      ("CUENTA", "Pasivo No Circulante (Largo Plazo)", plp_1, plp_2),
      ("SUBTOTAL", "Total Pasivos", pt_1, pt_2),
      ("CUENTA", "Capital Social y Reservas", pat_cuentas_1, pat_cuentas_2),
      ("CUENTA", "Resultados / Excedentes Acumulados", res_acum_1, res_acum_2),
      ("SUBTOTAL", "Total Patrimonio", pat_1, pat_2),
      ("TOTAL", "TOTAL PASIVO Y PATRIMONIO", pas_pat_1, pas_pat_2),
  ]

  tabla_sf = []
  for tipo, nombre, x1, x2 in rubros_sf:
    if x1 is None:
      tabla_sf.append((tipo, nombre, None, None, None, None, None, None))
    else:
      av1, av2, ah_abs, ah_rel = calc_vh(x1, x2, at_1, at_2)
      tabla_sf.append((tipo, nombre, x1, av1, x2, av2, ah_abs, ah_rel))

  rubros_er = [
      ("SUBTOTAL", "Ingresos / Ventas Netas", v_1, v_2),
      ("CUENTA", "Costo de Ventas / Operación", cv_1, cv_2),
      ("SUBTOTAL", "Utilidad / Excedente Bruto", ub_1, ub_2),
      ("CUENTA", "Gastos Operativos y Administrativos", go_1, go_2),
      ("SUBTOTAL", "Utilidad Operativa (EBIT)", uo_1, uo_2),
      ("CUENTA", "Otros Ingresos / Gastos Netos", otr_1, otr_2),
      ("TOTAL", "Utilidad / Excedente Neto del Período", un_1, un_2),
  ]

  tabla_er = []
  for tipo, nombre, x1, x2 in rubros_er:
    av1, av2, ah_abs, ah_rel = calc_vh(x1, x2, v_1, v_2)
    tabla_er.append((tipo, nombre, x1, av1, x2, av2, ah_abs, ah_rel))

  dias_p1 = max(1, (p1_fin - p1_ini).days + 1)
  dias_p2 = max(1, (p2_fin - p2_ini).days + 1)

  rc_1 = (ac_1 / pc_1) if round(pc_1, 2) > 0 else 0.0
  rc_2 = (ac_2 / pc_2) if round(pc_2, 2) > 0 else 0.0

  pa_1 = ((ac_1 - inv_1) / pc_1) if round(pc_1, 2) > 0 else 0.0
  pa_2 = ((ac_2 - inv_2) / pc_2) if round(pc_2, 2) > 0 else 0.0

  ctn_1 = ac_1 - pc_1
  ctn_2 = ac_2 - pc_2

  end_1 = (pt_1 / at_1 * 100.0) if round(at_1, 2) > 0 else 0.0
  end_2 = (pt_2 / at_2 * 100.0) if round(at_2, 2) > 0 else 0.0

  dp_1 = (pt_1 / pat_1) if round(pat_1, 2) != 0 else 0.0
  dp_2 = (pt_2 / pat_2) if round(pat_2, 2) != 0 else 0.0

  cob_int_1 = (uo_1 / gf_1) if round(gf_1, 2) > 0 else 0.0
  cob_int_2 = (uo_2 / gf_2) if round(gf_2, 2) > 0 else 0.0

  rot_cxc_1 = (v_1 / cxc_1) if round(cxc_1, 2) > 0 else 0.0
  rot_cxc_2 = (v_2 / cxc_2) if round(cxc_2, 2) > 0 else 0.0

  dias_cxc_1 = (dias_p1 / rot_cxc_1) if rot_cxc_1 > 0 else 0.0
  dias_cxc_2 = (dias_p2 / rot_cxc_2) if rot_cxc_2 > 0 else 0.0

  rot_inv_1 = (cv_1 / inv_1) if round(inv_1, 2) > 0 else 0.0
  rot_inv_2 = (cv_2 / inv_2) if round(inv_2, 2) > 0 else 0.0

  rot_act_1 = (v_1 / at_1) if round(at_1, 2) > 0 else 0.0
  rot_act_2 = (v_2 / at_2) if round(at_2, 2) > 0 else 0.0

  mb_1 = (ub_1 / v_1 * 100.0) if round(v_1, 2) > 0 else 0.0
  mb_2 = (ub_2 / v_2 * 100.0) if round(v_2, 2) > 0 else 0.0

  mo_1 = (uo_1 / v_1 * 100.0) if round(v_1, 2) > 0 else 0.0
  mo_2 = (uo_2 / v_2 * 100.0) if round(v_2, 2) > 0 else 0.0

  mn_1 = (un_1 / v_1 * 100.0) if round(v_1, 2) > 0 else 0.0
  mn_2 = (un_2 / v_2 * 100.0) if round(v_2, 2) > 0 else 0.0

  roa_1 = (un_1 / at_1 * 100.0) if round(at_1, 2) > 0 else 0.0
  roa_2 = (un_2 / at_2 * 100.0) if round(at_2, 2) > 0 else 0.0

  roe_1 = (un_1 / pat_1 * 100.0) if round(pat_1, 2) != 0 else 0.0
  roe_2 = (un_2 / pat_2 * 100.0) if round(pat_2, 2) != 0 else 0.0

  razones = [
      (
          "LIQUIDEZ",
          "Razón Corriente (Activo Circ. / Pasivo Circ.)",
          f"{rc_1:.2f}x",
          f"{rc_2:.2f}x",
          "Mide capacidad de pago a corto plazo (Óptimo >= 1.50x)",
      ),
      (
          "LIQUIDEZ",
          "Prueba Ácida ((Activo Circ. - Inv.) / Pasivo Circ.)",
          f"{pa_1:.2f}x",
          f"{pa_2:.2f}x",
          "Liquidez inmediata sin depender de inventarios (Óptimo >= 1.00x)",
      ),
      (
          "LIQUIDEZ",
          "Capital de Trabajo Neto (₡)",
          f"₡{fmt_moneda(ctn_1)}",
          f"₡{fmt_moneda(ctn_2)}",
          "Recursos operativos disponibles tras cubrir pasivos corrientes",
      ),
      (
          "ENDEUDAMIENTO",
          "Índice de Endeudamiento Total (Pasivo / Activo)",
          fmt_pct(end_1),
          fmt_pct(end_2),
          "Porcentaje de los activos financiado por terceros",
      ),
      (
          "ENDEUDAMIENTO",
          "Apalancamiento Deuda / Patrimonio",
          f"{dp_1:.2f}x",
          f"{dp_2:.2f}x",
          "Relación entre deuda total y capital propio",
      ),
      (
          "ENDEUDAMIENTO",
          "Cobertura de Intereses (EBIT / Gasto Financiero)",
          f"{cob_int_1:.2f}x" if cob_int_1 > 0 else "N/A",
          f"{cob_int_2:.2f}x" if cob_int_2 > 0 else "N/A",
          "Veces que la utilidad operativa cubre intereses",
      ),
      (
          "ACTIVIDAD",
          "Rotación de Cuentas por Cobrar",
          f"{rot_cxc_1:.2f} veces",
          f"{rot_cxc_2:.2f} veces",
          "Frecuencia de recuperación de cartera en el período",
      ),
      (
          "ACTIVIDAD",
          "Período Promedio de Cobro (Días)",
          f"{dias_cxc_1:.0f} días" if dias_cxc_1 > 0 else "N/A",
          f"{dias_cxc_2:.0f} días" if dias_cxc_2 > 0 else "N/A",
          "Días promedio en convertir ventas a crédito en efectivo",
      ),
      (
          "ACTIVIDAD",
          "Rotación de Inventarios",
          f"{rot_inv_1:.2f} veces" if rot_inv_1 > 0 else "N/A",
          f"{rot_inv_2:.2f} veces" if rot_inv_2 > 0 else "N/A",
          "Veces que se renueva el inventario en el período",
      ),
      (
          "ACTIVIDAD",
          "Rotación de Activos Totales",
          f"{rot_act_1:.2f} veces",
          f"{rot_act_2:.2f} veces",
          "Eficiencia de los activos para generar ingresos",
      ),
      (
          "RENTABILIDAD",
          "Margen de Utilidad Bruta",
          fmt_pct(mb_1),
          fmt_pct(mb_2),
          "Rentabilidad directa tras descontar costo de ventas",
      ),
      (
          "RENTABILIDAD",
          "Margen de Utilidad Operativa",
          fmt_pct(mo_1),
          fmt_pct(mo_2),
          "Eficiencia operativa antes de partidas financieras",
      ),
      (
          "RENTABILIDAD",
          "Margen de Utilidad Neta",
          fmt_pct(mn_1),
          fmt_pct(mn_2),
          "Porcentaje de ganancia neta por cada colón vendido",
      ),
      (
          "RENTABILIDAD",
          "Rendimiento sobre Activos (ROA)",
          fmt_pct(roa_1),
          fmt_pct(roa_2),
          "Capacidad de los activos totales para generar utilidad neta",
      ),
      (
          "RENTABILIDAD",
          "Rendimiento sobre Patrimonio (ROE)",
          fmt_pct(roe_1),
          fmt_pct(roe_2),
          "Rentabilidad obtenida sobre el capital de los socios",
      ),
  ]

  conclusiones = []
  recomendaciones = []

  var_v_pct = ((v_2 - v_1) / abs(v_1) * 100.0) if round(v_1, 2) != 0 else 0.0
  var_a_pct = ((at_2 - at_1) / abs(at_1) * 100.0) if round(at_1, 2) != 0 else 0.0
  conclusiones.append(
      f"- Estructura Patrimonial (Análisis Vertical): Al cierre de"
      f" {p2_fin.strftime('%d-%m-%Y')}, el Activo Circulante representa el"
      f" {(ac_2 / at_2 * 100.0 if at_2 else 0):.2f}% del Total de Activos"
      f" (₡{fmt_moneda(at_2)}), mientras que el financiamiento externo (Pasivo"
      f" Total) equivale al {end_2:.2f}% y el Patrimonio cubre el"
      f" {(100.0 - end_2 if at_2 else 0):.2f}% de la estructura."
  )
  conclusiones.append(
      f"- Evolución Comparativa (Análisis Horizontal): Los activos totales"
      f" presentaron una variación de {var_a_pct:+.2f}%"
      f" (₡{fmt_moneda(at_2 - at_1)}) respecto al período anterior. Por su"
      f" parte, los ingresos operativos mostraron una variación de"
      f" {var_v_pct:+.2f}% (₡{fmt_moneda(v_2 - v_1)}), generando un resultado"
      f" neto de ₡{fmt_moneda(un_2)} en el período actual frente a"
      f" ₡{fmt_moneda(un_1)} en el período comparativo."
  )

  if pc_2 == 0 and ac_2 > 0:
    conclusiones.append(
        "- Liquidez: La entidad no registra pasivos exigibles a corto plazo al"
        " corte, disponiendo de un Capital de Trabajo Neto positivo de"
        f" ₡{fmt_moneda(ctn_2)}."
    )
    recomendaciones.append(
        "- Gestión de Tesorería: Evaluar instrumentos financieros de bajo riesgo"
        " o inversiones a la vista para rentabilizar los excedentes de liquidez"
        " ociosa en caja y bancos."
    )
  elif rc_2 >= 1.5:
    conclusiones.append(
        f"- Liquidez: La Razón Corriente se ubica en {rc_2:.2f}x (Prueba Ácida:"
        f" {pa_2:.2f}x) con un Capital de Trabajo Neto de ₡{fmt_moneda(ctn_2)},"
        " lo que evidencia holgura para atender compromisos a corto plazo."
    )
    recomendaciones.append(
        "- Optimización de Capital de Trabajo: Mantener el monitoreo periódico"
        f" de la cartera por cobrar (rotación actual de {rot_cxc_2:.2f} veces)"
        " para preservar el flujo de efectivo operativo."
    )
  else:
    conclusiones.append(
        f"- Liquidez: La Razón Corriente se sitúa en {rc_2:.2f}x (Prueba Ácida:"
        f" {pa_2:.2f}x), nivel ajustado respecto al estándar prudencial de"
        " 1.50x, lo que exige disciplina en el calce de vencimientos."
    )
    recomendaciones.append(
        "- Fortalecimiento de Liquidez: Acelerar la gestión de cobro de"
        " cartera, optimizar niveles de inventario y evitar financiar activos"
        " fijos con deuda de corto plazo."
    )

  if mn_2 >= 10.0:
    conclusiones.append(
        f"- Rentabilidad: El Margen Neto alcanza el {mn_2:.2f}% (Margen"
        f" Operativo: {mo_2:.2f}%), con un ROE de {roe_2:.2f}%, reflejando una"
        " sólida capacidad de generación de utilidades."
    )
  elif mn_2 > 0:
    conclusiones.append(
        f"- Rentabilidad: La operación es rentable con un Margen Neto de"
        f" {mn_2:.2f}% y un ROA de {roa_2:.2f}%, aunque con espacio para mejorar"
        " la eficiencia en costos y gastos operativos"
        f" ({(go_2 / v_2 * 100.0 if v_2 else 0):.2f}% sobre ventas)."
    )
  else:
    conclusiones.append(
        f"- Rentabilidad: El período cierra con margen neto negativo"
        f" ({mn_2:.2f}%), producto de que los costos y gastos operativos"
        " superaron los ingresos generados en el corte analizado."
    )

  if end_2 > 60.0:
    recomendaciones.append(
        f"- Estructura de Endeudamiento: Dado que el endeudamiento total"
        f" alcanza el {end_2:.2f}%, se sugiere priorizar la amortización de"
        " pasivos onerosos y capitalizar utilidades para fortalecer el"
        " patrimonio."
    )
  else:
    recomendaciones.append(
        "- Control Presupuestario y Costos: Establecer presupuestos mensuales"
        f" por centro de gasto para sostener el Margen Bruto ({mb_2:.2f}%) y"
        " proteger el flujo operativo ante variaciones en insumos."
    )

  recomendaciones.append(
      "- Seguimiento Contable NIIF: Continuar con conciliaciones bancarias"
      " mensuales y revisiones trimestrales de indicadores financieros para"
      " respaldar la toma de decisiones gerenciales y bancarias."
  )

  return {
      "tabla_sf": tabla_sf,
      "tabla_er": tabla_er,
      "razones": razones,
      "texto_conclusiones": "\n\n".join(conclusiones),
      "texto_recomendaciones": "\n\n".join(recomendaciones),
      "kpis": {
          "rc_2": rc_2,
          "end_2": end_2,
          "mn_2": mn_2,
          "roe_2": roe_2,
      },
  }


def generar_word_analisis_financiero(
    res_af,
    empresa,
    cedula,
    p1_ini,
    p1_fin,
    p2_ini,
    p2_fin,
    conclusiones_txt,
    recomendaciones_txt,
    rep_legal,
):
  doc = docx.Document()
  for section in doc.sections:
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.0)
    section.page_height = Inches(8.5)
    section.top_margin = Inches(0.42)
    section.bottom_margin = Inches(0.42)
    section.left_margin = Inches(0.42)
    section.right_margin = Inches(0.42)

  p_tit = doc.add_paragraph()
  p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_tit.paragraph_format.space_after = Pt(6)
  r0 = p_tit.add_run("INFORME PROFESIONAL DE ANÁLISIS FINANCIERO\n")
  r0.bold = True
  r0.font.size = Pt(13)
  r1 = p_tit.add_run(f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}\n")
  r1.bold = True
  r1.font.size = Pt(10)
  r2 = p_tit.add_run(
      "Análisis Vertical, Análisis Horizontal y Razones Financieras\n"
      f"Períodos Comparativos: {p1_fin.strftime('%d-%m-%Y')} vs"
      f" {p2_fin.strftime('%d-%m-%Y')} (Expresado en colones)\n"
  )
  r2.font.size = Pt(9)
  r3 = p_tit.add_run(
      "Elaborado por: Lic. Roy F. Esquivel Gonzalez - Contador Privado Carné 8788"
  )
  r3.bold = True
  r3.font.size = Pt(9)

  def agregar_tabla_vh(titulo_sec, filas_vh):
    p_s = doc.add_paragraph()
    p_s.paragraph_format.space_before = Pt(6)
    p_s.paragraph_format.space_after = Pt(2)
    rs = p_s.add_run(titulo_sec)
    rs.bold = True
    rs.font.size = Pt(10)

    t = doc.add_table(rows=1, cols=7)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    anchos = [
        Inches(3.25),
        Inches(1.25),
        Inches(0.95),
        Inches(1.25),
        Inches(0.95),
        Inches(1.30),
        Inches(1.05),
    ]
    headers = [
        "Rubro / Cuenta",
        f"Al {p1_fin.strftime('%d-%m-%Y')}",
        "Vert. %",
        f"Al {p2_fin.strftime('%d-%m-%Y')}",
        "Vert. %",
        "Var. Absoluta (₡)",
        "Var. Rel. %",
    ]
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
      hdr[i].width = anchos[i]
      hdr[i]._tc.get_or_add_tcPr().append(
          parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
      )
      p = hdr[i].paragraphs[0]
      p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
      r = p.add_run(h)
      r.bold = True
      r.font.size = Pt(7.5)

    for tipo, nom, x1, av1, x2, av2, ah_abs, ah_rel in filas_vh:
      rc = t.add_row().cells
      es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL")
      vals = [
          nom,
          fmt_moneda(x1) if x1 is not None else "",
          fmt_pct(av1),
          fmt_moneda(x2) if x2 is not None else "",
          fmt_pct(av2),
          fmt_moneda(ah_abs) if ah_abs is not None else "",
          fmt_pct(ah_rel),
      ]
      for i, v in enumerate(vals):
        rc[i].width = anchos[i]
        p = rc[i].paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(v)
        r.bold = es_b
        r.font.size = Pt(7.5)
      if tipo in ("SUBTOTAL", "TOTAL"):
        for c in rc:
          c._tc.get_or_add_tcPr().append(
              parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F8"/>')
          )

  agregar_tabla_vh(
      "1. ANÁLISIS VERTICAL Y HORIZONTAL - ESTADO DE SITUACIÓN FINANCIERA",
      res_af["tabla_sf"],
  )
  agregar_tabla_vh(
      "2. ANÁLISIS VERTICAL Y HORIZONTAL - ESTADO DE RESULTADOS",
      res_af["tabla_er"],
  )

  doc.add_page_break()
  p_r = doc.add_paragraph()
  p_r.paragraph_format.space_after = Pt(3)
  rr = p_r.add_run("3. RAZONES E INDICADORES FINANCIEROS")
  rr.bold = True
  rr.font.size = Pt(10.5)

  tr = doc.add_table(rows=1, cols=5)
  tr.alignment = WD_TABLE_ALIGNMENT.CENTER
  tr.autofit = False
  anchos_r = [
      Inches(1.35),
      Inches(3.15),
      Inches(1.15),
      Inches(1.15),
      Inches(3.20),
  ]
  hdr_r = [
      "Categoría",
      "Indicador Financiero",
      p1_fin.strftime("%d-%m-%Y"),
      p2_fin.strftime("%d-%m-%Y"),
      "Interpretación Técnica",
  ]
  for i, h in enumerate(hdr_r):
    tr.rows[0].cells[i].width = anchos_r[i]
    tr.rows[0].cells[i]._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>')
    )
    p = tr.rows[0].cells[i].paragraphs[0]
    p.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT
        if i in (2, 3)
        else WD_ALIGN_PARAGRAPH.LEFT
    )
    r = p.add_run(h)
    r.bold = True
    r.font.size = Pt(7.5)

  for cat_r, ind_r, v1_r, v2_r, interp_r in res_af["razones"]:
    rc = tr.add_row().cells
    vals = [cat_r, ind_r, v1_r, v2_r, interp_r]
    for i, v in enumerate(vals):
      rc[i].width = anchos_r[i]
      p = rc[i].paragraphs[0]
      p.paragraph_format.space_before = Pt(1)
      p.paragraph_format.space_after = Pt(1)
      p.alignment = (
          WD_ALIGN_PARAGRAPH.RIGHT
          if i in (2, 3)
          else WD_ALIGN_PARAGRAPH.LEFT
      )
      r = p.add_run(v)
      r.bold = i in (0, 3)
      r.font.size = Pt(7.5)

  p_c = doc.add_paragraph()
  p_c.paragraph_format.space_before = Pt(8)
  p_c.paragraph_format.space_after = Pt(2)
  rc_tit = p_c.add_run("4. CONCLUSIONES DEL ANÁLISIS FINANCIERO")
  rc_tit.bold = True
  rc_tit.font.size = Pt(10.5)

  p_c_body = doc.add_paragraph()
  r_cb = p_c_body.add_run(conclusiones_txt)
  r_cb.font.size = Pt(8.5)

  p_rec = doc.add_paragraph()
  p_rec.paragraph_format.space_before = Pt(6)
  p_rec.paragraph_format.space_after = Pt(2)
  rrec_tit = p_rec.add_run("5. RECOMENDACIONES PROFESIONALES")
  rrec_tit.bold = True
  rrec_tit.font.size = Pt(10.5)

  p_rec_body = doc.add_paragraph()
  r_rb = p_rec_body.add_run(recomendaciones_txt)
  r_rb.font.size = Pt(8.5)

  doc.add_paragraph("\n")
  ft = doc.add_table(rows=1, cols=2)
  ft.alignment = WD_TABLE_ALIGNMENT.CENTER
  ft.autofit = False
  ft.rows[0].cells[0].width = Inches(5.0)
  ft.rows[0].cells[1].width = Inches(5.0)

  pi = ft.rows[0].cells[0].paragraphs[0]
  ri = pi.add_run(
      "\n____________________________________\nElaborado por:\nLic. Roy F."
      " Esquivel Gonzalez\nContador Privado Carné 8788"
  )
  ri.bold = True
  ri.font.size = Pt(8.5)

  pd_cell = ft.rows[0].cells[1].paragraphs[0]
  pd_cell.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  rd = pd_cell.add_run(
      f"\n____________________________________\nRecibido por:\n{rep_legal}\nRepresentante"
      " Legal"
  )
  rd.bold = True
  rd.font.size = Pt(8.5)

  buf = io.BytesIO()
  doc.save(buf)
  buf.seek(0)
  return buf


def generar_excel_analisis_financiero(
    res_af,
    empresa,
    cedula,
    p1_ini,
    p1_fin,
    p2_ini,
    p2_fin,
    conclusiones_txt,
    recomendaciones_txt,
    rep_legal,
):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Análisis Financiero"

  ws["A1"] = "INFORME PROFESIONAL DE ANÁLISIS FINANCIERO"
  ws["A1"].font = Font(bold=True, size=14)
  ws["A2"] = (
      f"{empresa.upper()} | CÉDULA JURÍDICA {cedula} | PERÍODOS:"
      f" {p1_fin.strftime('%d-%m-%Y')} vs {p2_fin.strftime('%d-%m-%Y')}"
  )
  ws["A2"].font = Font(bold=True, size=10)
  ws["A3"] = (
      "Elaborado por: Lic. Roy F. Esquivel Gonzalez - Contador Privado Carné 8788"
  )
  ws["A3"].font = Font(italic=True, bold=True, size=10)

  fill_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  fill_sub = PatternFill(
      start_color="F2F4F8", end_color="F2F4F8", fill_type="solid"
  )

  def escribir_vh(fila_ini, titulo, filas_vh):
    ws.cell(row=fila_ini, column=1, value=titulo).font = Font(
        bold=True, size=11
    )
    r_h = fila_ini + 1
    headers = [
        "Rubro / Cuenta",
        f"Al {p1_fin.strftime('%d-%m-%Y')}",
        "Vertical %",
        f"Al {p2_fin.strftime('%d-%m-%Y')}",
        "Vertical %",
        "Variación Absoluta (₡)",
        "Variación Relativa %",
    ]
    for idx_c, h in enumerate(headers, start=1):
      c = ws.cell(row=r_h, column=idx_c, value=h)
      c.font = Font(bold=True)
      c.fill = fill_hdr
      c.alignment = Alignment(horizontal="right" if idx_c >= 2 else "left")

    r_act = r_h + 1
    for tipo, nom, x1, av1, x2, av2, ah_abs, ah_rel in filas_vh:
      es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL")
      ws.cell(row=r_act, column=1, value=nom)
      if x1 is not None:
        c2 = ws.cell(row=r_act, column=2, value=float(x1))
        c3 = ws.cell(row=r_act, column=3, value=float(av1) / 100.0)
        c4 = ws.cell(row=r_act, column=4, value=float(x2))
        c5 = ws.cell(row=r_act, column=5, value=float(av2) / 100.0)
        c6 = ws.cell(row=r_act, column=6, value=float(ah_abs))
        c7 = ws.cell(row=r_act, column=7, value=float(ah_rel) / 100.0)
        for cm in (c2, c4, c6):
          cm.number_format = "#,##0.00;(#,##0.00);\"-\""
        for cp in (c3, c5, c7):
          cp.number_format = "0.00%"
      if es_b:
        for col in range(1, 8):
          ws.cell(row=r_act, column=col).font = Font(bold=True)
      if tipo in ("SUBTOTAL", "TOTAL"):
        for col in range(1, 8):
          ws.cell(row=r_act, column=col).fill = fill_sub
      r_act += 1
    return r_act + 2

  sig_f = escribir_vh(
      5,
      "1. ANÁLISIS VERTICAL Y HORIZONTAL - ESTADO DE SITUACIÓN FINANCIERA",
      res_af["tabla_sf"],
  )
  sig_f = escribir_vh(
      sig_f,
      "2. ANÁLISIS VERTICAL Y HORIZONTAL - ESTADO DE RESULTADOS",
      res_af["tabla_er"],
  )

  ws.cell(row=sig_f, column=1, value="3. RAZONES FINANCIERAS").font = Font(
      bold=True, size=11
  )
  r_hr = sig_f + 1
  hdr_r = [
      "Categoría",
      "Indicador",
      p1_fin.strftime("%d-%m-%Y"),
      p2_fin.strftime("%d-%m-%Y"),
      "Interpretación",
  ]
  for idx_c, h in enumerate(hdr_r, start=1):
    c = ws.cell(row=r_hr, column=idx_c, value=h)
    c.font = Font(bold=True)
    c.fill = fill_hdr
  r_act = r_hr + 1
  for cat_r, ind_r, v1_r, v2_r, interp_r in res_af["razones"]:
    ws.cell(row=r_act, column=1, value=cat_r).font = Font(bold=True)
    ws.cell(row=r_act, column=2, value=ind_r)
    ws.cell(row=r_act, column=3, value=v1_r).alignment = Alignment(
        horizontal="right"
    )
    ws.cell(row=r_act, column=4, value=v2_r).alignment = Alignment(
        horizontal="right"
    )
    ws.cell(row=r_act, column=4).font = Font(bold=True)
    ws.cell(row=r_act, column=5, value=interp_r)
    r_act += 1

  r_act += 1
  ws.cell(
      row=r_act, column=1, value="4. CONCLUSIONES DEL ANÁLISIS FINANCIERO"
  ).font = Font(bold=True, size=11)
  r_act += 1
  ws.merge_cells(start_row=r_act, start_column=1, end_row=r_act, end_column=7)
  ws.cell(row=r_act, column=1, value=conclusiones_txt).alignment = Alignment(
      wrap_text=True, vertical="top"
  )
  ws.row_dimensions[r_act].height = 110

  r_act += 2
  ws.cell(
      row=r_act, column=1, value="5. RECOMENDACIONES PROFESIONALES"
  ).font = Font(bold=True, size=11)
  r_act += 1
  ws.merge_cells(start_row=r_act, start_column=1, end_row=r_act, end_column=7)
  ws.cell(row=r_act, column=1, value=recomendaciones_txt).alignment = Alignment(
      wrap_text=True, vertical="top"
  )
  ws.row_dimensions[r_act].height = 110

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
      column=7,
      value=(
          f"_______________________________\nRecibido por:\n{rep_legal}\nRepresentante"
          " Legal"
      ),
  ).font = Font(bold=True)
  ws.cell(row=r_act, column=7).alignment = Alignment(
      wrap_text=True, horizontal="right"
  )
  ws.row_dimensions[r_act].height = 55

  for letra, ancho in zip(
      ["A", "B", "C", "D", "E", "F", "G"], [44, 18, 14, 18, 42, 20, 16]
  ):
    ws.column_dimensions[letra].width = ancho

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def generar_pdf_analisis_financiero(
    res_af,
    empresa,
    cedula,
    p1_ini,
    p1_fin,
    p2_ini,
    p2_fin,
    conclusiones_txt,
    recomendaciones_txt,
    rep_legal,
):
  pdf = FPDF(orientation="L", unit="mm", format="Letter")
  pdf.set_auto_page_break(auto=True, margin=12)
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 12)
  pdf.cell(
      0,
      6,
      limpiar_texto_pdf("INFORME PROFESIONAL DE ANÁLISIS FINANCIERO"),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5,
      limpiar_texto_pdf(f"{empresa.upper()} | CÉDULA JURÍDICA {cedula}"),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "", 8.5)
  pdf.cell(
      0,
      4.5,
      limpiar_texto_pdf(
          f"Análisis Vertical, Horizontal y Razones Financieras | Períodos:"
          f" {p1_fin.strftime('%d-%m-%Y')} vs {p2_fin.strftime('%d-%m-%Y')}"
      ),
      ln=True,
      align="C",
  )
  pdf.set_font("Helvetica", "B", 8.5)
  pdf.cell(
      0,
      4.5,
      limpiar_texto_pdf(
          "Elaborado por: Lic. Roy F. Esquivel Gonzalez - Contador Privado Carné"
          " 8788"
      ),
      ln=True,
      align="C",
  )
  pdf.ln(3)

  def dibujar_tabla_vh_pdf(titulo_sec, filas_vh):
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 5.5, limpiar_texto_pdf(titulo_sec), ln=True)
    w = [84, 32, 20, 32, 20, 34, 24]
    pdf.set_fill_color(225, 232, 245)
    pdf.set_font("Helvetica", "B", 7)
    headers = [
        "Rubro / Cuenta",
        f"Al {p1_fin.strftime('%d-%m-%Y')}",
        "Vert. %",
        f"Al {p2_fin.strftime('%d-%m-%Y')}",
        "Vert. %",
        "Var. Abs. (C)",
        "Var. %",
    ]
    for i, h in enumerate(headers):
      pdf.cell(
          w[i],
          5.5,
          limpiar_texto_pdf(h),
          border=1,
          align="R" if i >= 1 else "L",
          fill=True,
      )
    pdf.ln()

    for tipo, nom, x1, av1, x2, av2, ah_abs, ah_rel in filas_vh:
      es_b = tipo in ("TITULO", "SUBTOTAL", "TOTAL")
      pdf.set_font("Helvetica", "B" if es_b else "", 6.8)
      fill_r = tipo in ("SUBTOTAL", "TOTAL")
      if fill_r:
        pdf.set_fill_color(242, 245, 250)
      pdf.cell(w[0], 4.8, limpiar_texto_pdf(nom)[:48], border="B", fill=fill_r)
      vals = [
          fmt_moneda(x1) if x1 is not None else "",
          fmt_pct(av1),
          fmt_moneda(x2) if x2 is not None else "",
          fmt_pct(av2),
          fmt_moneda(ah_abs) if ah_abs is not None else "",
          fmt_pct(ah_rel),
      ]
      for i, v in enumerate(vals, start=1):
        pdf.cell(
            w[i],
            4.8,
            limpiar_texto_pdf(v),
            border="B",
            align="R",
            fill=fill_r,
        )
      pdf.ln()
    pdf.ln(3)

  dibujar_tabla_vh_pdf(
      "1. ANÁLISIS VERTICAL Y HORIZONTAL - ESTADO DE SITUACIÓN FINANCIERA",
      res_af["tabla_sf"],
  )
  dibujar_tabla_vh_pdf(
      "2. ANÁLISIS VERTICAL Y HORIZONTAL - ESTADO DE RESULTADOS",
      res_af["tabla_er"],
  )

  pdf.add_page()
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5.5,
      limpiar_texto_pdf("3. RAZONES E INDICADORES FINANCIEROS"),
      ln=True,
  )
  wr = [32, 82, 26, 26, 80]
  pdf.set_fill_color(225, 232, 245)
  pdf.set_font("Helvetica", "B", 7)
  hdr_r = [
      "Categoría",
      "Indicador Financiero",
      p1_fin.strftime("%d-%m-%Y"),
      p2_fin.strftime("%d-%m-%Y"),
      "Interpretación",
  ]
  for i, h in enumerate(hdr_r):
    pdf.cell(
        wr[i],
        5.5,
        limpiar_texto_pdf(h),
        border=1,
        align="R" if i in (2, 3) else "L",
        fill=True,
    )
  pdf.ln()

  for cat_r, ind_r, v1_r, v2_r, interp_r in res_af["razones"]:
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.cell(wr[0], 4.8, limpiar_texto_pdf(cat_r), border="B")
    pdf.set_font("Helvetica", "", 6.5)
    pdf.cell(wr[1], 4.8, limpiar_texto_pdf(ind_r)[:48], border="B")
    pdf.cell(wr[2], 4.8, limpiar_texto_pdf(v1_r), border="B", align="R")
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.cell(wr[3], 4.8, limpiar_texto_pdf(v2_r), border="B", align="R")
    pdf.set_font("Helvetica", "", 6.3)
    pdf.cell(wr[4], 4.8, limpiar_texto_pdf(interp_r)[:52], border="B")
    pdf.ln()

  pdf.ln(4)
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5.5,
      limpiar_texto_pdf("4. CONCLUSIONES DEL ANÁLISIS FINANCIERO"),
      ln=True,
  )
  pdf.set_font("Helvetica", "", 8)
  pdf.multi_cell(0, 4.3, limpiar_texto_pdf(conclusiones_txt))

  pdf.ln(3)
  pdf.set_font("Helvetica", "B", 9.5)
  pdf.cell(
      0,
      5.5,
      limpiar_texto_pdf("5. RECOMENDACIONES PROFESIONALES"),
      ln=True,
  )
  pdf.set_font("Helvetica", "", 8)
  pdf.multi_cell(0, 4.3, limpiar_texto_pdf(recomendaciones_txt))

  pdf.ln(10)
  if pdf.get_y() > 175:
    pdf.add_page()
  y_f = pdf.get_y()
  pdf.set_font("Helvetica", "B", 8)
  pdf.set_xy(12, y_f)
  pdf.multi_cell(
      95,
      4.2,
      limpiar_texto_pdf(
          "_______________________________\nElaborado por:\nLic. Roy F."
          " Esquivel Gonzalez\nContador Privado Carné 8788"
      ),
      align="L",
  )
  pdf.set_xy(170, y_f)
  pdf.multi_cell(
      95,
      4.2,
      limpiar_texto_pdf(
          f"_______________________________\nRecibido por:\n{rep_legal}\nRepresentante"
          " Legal"
      ),
      align="R",
  )
  return bytes(pdf.output())


# ==============================================================================
# 6. INTERFAZ PRINCIPAL DE USUARIO
# ==============================================================================
def main():
  if "backup_inicial_hecho" not in st.session_state:
    ejecutar_backup_automatico()
    st.session_state.backup_inicial_hecho = True

  conn = get_conn()
  inicializar_y_descargar_bd(conn)
  df_empresas = pd.read_sql_query("SELECT * FROM empresas", conn)

  if df_empresas.empty:
    st.warning("No hay empresas registradas aún. Crea una en el menú lateral.")
    empresa_id = None
  else:
    st.sidebar.title("🏢 Empresa Activa")
    df_empresas["etiqueta"] = df_empresas.apply(
        lambda r: f"{r['razon_social']} ({r['cedula']})", axis=1
    )
    mapa_emp = dict(zip(df_empresas["etiqueta"], df_empresas["id"]))
    emp_sel = st.sidebar.selectbox("Seleccione Empresa:", list(mapa_emp.keys()))
    empresa_id = int(mapa_emp[emp_sel])
    datos_emp = df_empresas[df_empresas["id"] == empresa_id].iloc[0]

  menu = st.sidebar.radio(
      "Módulos del Sistema:",
      [
          "1. Empresas y Configuración",
          "2. Catálogo de Cuentas y Saldos Iniciales",
          "3. Elaboración y Gestión de Asientos",
          "4. Libro Diario y Mayor",
          "5. Balance de Comprobación",
          "6. Estado de la Situación Financiera",
          "7. Estado de Resultados",
          "8. Estado de Flujo de Efectivo",
          "9. Estado de Cambios en el Patrimonio",
          "10. Flujo de Caja Proyectado",
          "11. Análisis Financiero Integral",
          "12. Conciliación Bancaria",
          "13. Reporte de IVA (Hacienda CR)",
          "14. Reporte de Renta (Hacienda CR)",
          "15. Administración de Empresas",
      ],
  )

  st.sidebar.divider()
  st.sidebar.subheader("🛡️ Seguridad y Respaldo")
  if st.sidebar.button("💾 Crear Backup Ahora", use_container_width=True):
    ruta_zip = ejecutar_backup_automatico()
    st.sidebar.success(f"Backup guardado: {os.path.basename(ruta_zip)}")

  lista_zips = sorted(
      glob.glob(os.path.join(BACKUP_DIR, "Backup_Contable_*.zip"))
  )
  if lista_zips:
    ultimo_zip = lista_zips[-1]
    with open(ultimo_zip, "rb") as f_zip:
      st.sidebar.download_button(
          label="📥 Descargar Último Backup (.zip)",
          data=f_zip.read(),
          file_name=os.path.basename(ultimo_zip),
          mime="application/zip",
          use_container_width=True,
      )

  # --- 1. EMPRESAS ---
  if menu == "1. Empresas y Configuración":
    st.header("Gestión de Empresas")
    st.dataframe(
        df_empresas[[
            "id",
            "cedula",
            "razon_social",
            "nombre_comercial",
            "representante_legal",
        ]],
        use_container_width=True,
        hide_index=True,
    )

    if empresa_id is not None:
      st.subheader("Editar Empresa Seleccionada")
      with st.form("edit_emp"):
        rz = st.text_input(
            "Razón Social", value=str(datos_emp["razon_social"] or "")
        )
        nc = st.text_input(
            "Nombre Comercial", value=str(datos_emp["nombre_comercial"] or "")
        )
        ced = st.text_input(
            "Cédula Jurídica", value=str(datos_emp["cedula"] or "")
        )
        rep = st.text_input(
            "Representante Legal",
            value=str(datos_emp["representante_legal"] or ""),
        )
        act = st.text_input(
            "Actividad Económica",
            value=str(datos_emp["actividad_economica"] or ""),
        )
        if st.form_submit_button("💾 Guardar Cambios", type="primary"):
          cur = conn.cursor()
          cur.execute(
              "UPDATE empresas SET razon_social=?, nombre_comercial=?,"
              " cedula=?, representante_legal=?, actividad_economica=? WHERE"
              " id=?",
              (rz, nc, ced, rep, act, empresa_id),
          )
          conn.commit()
          ejecutar_backup_automatico()
          st.success("Datos actualizados.")
          st.rerun()

    st.subheader("Registrar Nueva Empresa")
    with st.form("new_emp"):
      n_rz = st.text_input("Razón Social")
      n_ced = st.text_input("Cédula Jurídica")
      n_rep = st.text_input("Representante Legal")
      if st.form_submit_button("➕ Crear Empresa"):
        if n_rz:
          cur = conn.cursor()
          cur.execute(
              "INSERT INTO empresas (razon_social, cedula,"
              " representante_legal) VALUES (?, ?, ?)",
              (n_rz, n_ced, n_rep),
          )
          new_id = cur.lastrowid
          cur.execute(
              "INSERT INTO catalogo_cuentas (empresa_id, codigo, nombre, tipo,"
              " acepta_movimiento) SELECT ?, codigo, nombre, tipo,"
              " acepta_movimiento FROM plantilla_catalogo_cuentas",
              (new_id,),
          )
          conn.commit()
          ejecutar_backup_automatico()
          st.success("Empresa creada con catálogo base.")
          st.rerun()

  # --- 2. CATÁLOGO Y SALDOS INICIALES / HISTÓRICOS ---
  elif menu == "2. Catálogo de Cuentas y Saldos Iniciales" and empresa_id:
    st.header(
        f"Catálogo de Cuentas y Saldos Históricos - {datos_emp['razon_social']}"
    )
    df_cat = pd.read_sql_query(
        "SELECT id, codigo, nombre, tipo, acepta_movimiento FROM"
        " catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )
    st.dataframe(df_cat, use_container_width=True, hide_index=True)

    c_izq, c_der = st.columns([1, 1.4])
    with c_izq:
      with st.form("add_cta"):
        st.subheader("➕ Agregar Cuenta al Catálogo")
        cod = st.text_input("Código (ej. 1.1.01.05)")
        nom = st.text_input("Nombre de Cuenta")
        tip = st.selectbox(
            "Tipo",
            ["ACTIVO", "PASIVO", "PATRIMONIO", "INGRESO", "COSTO", "GASTO"],
        )
        ac = st.checkbox("Acepta Movimiento", value=True)
        if st.form_submit_button(
            "💾 Guardar Cuenta", type="primary", use_container_width=True
        ):
          if cod and nom:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO catalogo_cuentas (empresa_id, codigo, nombre,"
                " tipo, acepta_movimiento) VALUES (?, ?, ?, ?, ?)",
                (empresa_id, cod.strip(), nom.strip(), tip, 1 if ac else 0),
            )
            conn.commit()
            ejecutar_backup_automatico()
            st.success(
                f"Cuenta '{cod.strip()} - {nom.strip()}' agregada. Ya puedes"
                " asignarle saldos históricos a la derecha."
            )
            st.rerun()

    with c_der:
      st.subheader("📅 Registrar Saldo Inicial o Histórico (Años Anteriores)")
      df_ctas_mov = df_cat[df_cat["acepta_movimiento"] == 1].copy()
      opciones_cta_si = [
          f"{r['codigo']} - {r['nombre']} ({r['tipo']})"
          for _, r in df_ctas_mov.iterrows()
      ]
      mapa_cta_si = dict(zip(opciones_cta_si, df_ctas_mov["id"]))

      with st.form("add_saldo_historico"):
        sel_cta_si = st.selectbox(
            "Cuenta Contable (Existente o Nueva):",
            opciones_cta_si,
            key="sel_cta_hist",
        )
        sc1, sc2, sc3 = st.columns(3)
        f_hist = sc1.date_input(
            "Fecha de Corte / Inicio:",
            value=date(2025, 12, 31),
            format="DD-MM-YYYY",
        )
        d_hist = sc2.number_input(
            "Saldo Debe (₡)", min_value=0.0, step=100.0, format="%.2f"
        )
        h_hist = sc3.number_input(
            "Saldo Haber (₡)", min_value=0.0, step=100.0, format="%.2f"
        )

        if st.form_submit_button(
            "➕ Registrar Saldo Histórico / Inicial",
            type="primary",
            use_container_width=True,
        ):
          if round(d_hist, 2) == 0 and round(h_hist, 2) == 0:
            st.error("Debe digitar un monto mayor a cero en Debe o Haber.")
          elif round(d_hist, 2) > 0 and round(h_hist, 2) > 0:
            st.error(
                "Digita el saldo únicamente en Debe o únicamente en Haber según"
                " la naturaleza de la cuenta."
            )
          else:
            cid_sel = int(mapa_cta_si[sel_cta_si])
            f_iso_hist = str(f_hist)
            cur = conn.cursor()
            cur.execute(
                "SELECT id FROM saldos_iniciales WHERE empresa_id = ? AND"
                " cuenta_id = ? AND fecha_inicio = ?",
                (empresa_id, cid_sel, f_iso_hist),
            )
            existente = cur.fetchone()
            if existente:
              cur.execute(
                  "UPDATE saldos_iniciales SET debe = ?, haber = ? WHERE id = ?",
                  (float(d_hist), float(h_hist), int(existente[0])),
              )
              msg_ok = "Saldo histórico actualizado para esa fecha."
            else:
              cur.execute(
                  "INSERT INTO saldos_iniciales (empresa_id, cuenta_id,"
                  " fecha_inicio, debe, haber) VALUES (?, ?, ?, ?, ?)",
                  (empresa_id, cid_sel, f_iso_hist, float(d_hist), float(h_hist)),
              )
              msg_ok = "Saldo histórico registrado correctamente."
            conn.commit()
            ejecutar_backup_automatico()
            st.success(msg_ok)
            st.rerun()

    st.divider()
    st.subheader("📚 Histórico de Saldos Iniciales y Años Anteriores Registrados")
    df_si = pd.read_sql_query(
        "SELECT s.id, c.codigo, c.nombre, c.tipo, s.fecha_inicio, s.debe,"
        " s.haber FROM saldos_iniciales s JOIN catalogo_cuentas c ON"
        " s.cuenta_id = c.id WHERE s.empresa_id = ? ORDER BY s.fecha_inicio"
        " DESC, c.codigo ASC",
        conn,
        params=(empresa_id,),
    )
    if not df_si.empty:
      fechas_disp = ["Todas las fechas"] + sorted(
          df_si["fecha_inicio"].apply(fmt_fecha_dmy).unique().tolist(),
          reverse=True,
      )
      f_filtro_si = st.selectbox(
          "Filtrar por fecha de corte histórico:", fechas_disp
      )

      df_si_vista = df_si.copy()
      df_si_vista["fecha_inicio"] = df_si_vista["fecha_inicio"].apply(
          fmt_fecha_dmy
      )
      if f_filtro_si != "Todas las fechas":
        df_si_vista = df_si_vista[df_si_vista["fecha_inicio"] == f_filtro_si]

      tot_d_si = df_si_vista["debe"].sum()
      tot_h_si = df_si_vista["haber"].sum()
      m1, m2, m3 = st.columns(3)
      m1.metric("Total Debe (Vista Actual)", f"₡{tot_d_si:,.2f}")
      m2.metric("Total Haber (Vista Actual)", f"₡{tot_h_si:,.2f}")
      m3.metric(
          "Diferencia / Cuadre",
          f"₡{abs(tot_d_si - tot_h_si):,.2f}",
      )

      df_si_edit = st.data_editor(
          df_si_vista,
          use_container_width=True,
          hide_index=True,
          num_rows="dynamic",
          disabled=["id", "codigo", "nombre", "tipo"],
          column_config={
              "fecha_inicio": st.column_config.TextColumn(
                  "Fecha Corte (DD-MM-AAAA)"
              ),
              "debe": st.column_config.NumberColumn("Debe", format="₡%.2f"),
              "haber": st.column_config.NumberColumn("Haber", format="₡%.2f"),
          },
          key=f"ed_si_{empresa_id}_{f_filtro_si}",
      )
      b_si1, b_si2 = st.columns(2)
      if b_si1.button(
          "💾 Guardar Cambios en Tabla Histórica",
          type="primary",
          use_container_width=True,
      ):
        cur = conn.cursor()
        ids_vigentes = set(df_si_edit["id"].dropna().astype(int).tolist())
        ids_originales = set(df_si_vista["id"].astype(int).tolist())
        for id_borrar in ids_originales - ids_vigentes:
          cur.execute(
              "DELETE FROM saldos_iniciales WHERE id = ?", (int(id_borrar),)
          )
        for _, r_si in df_si_edit.dropna(subset=["id"]).iterrows():
          f_iso = parse_fecha_input(r_si["fecha_inicio"])
          cur.execute(
              "UPDATE saldos_iniciales SET fecha_inicio = ?, debe = ?,"
              " haber = ? WHERE id = ?",
              (
                  f_iso,
                  float(r_si["debe"]),
                  float(r_si["haber"]),
                  int(r_si["id"]),
              ),
          )
        conn.commit()
        ejecutar_backup_automatico()
        st.success("Saldos históricos actualizados.")
        st.rerun()

      if b_si2.button(
          "🧹 Eliminar Filas Duplicadas", use_container_width=True
      ):
        cur = conn.cursor()
        cur.execute(
            """
            DELETE FROM saldos_iniciales
            WHERE empresa_id = ? AND id NOT IN (
                SELECT MIN(id)
                FROM saldos_iniciales
                WHERE empresa_id = ?
                GROUP BY cuenta_id, fecha_inicio, debe, haber
            )
            """,
            (empresa_id, empresa_id),
        )
        conn.commit()
        ejecutar_backup_automatico()
        st.success("Duplicados eliminados.")
        st.rerun()
    else:
      st.info("No hay saldos iniciales ni históricos registrados aún.")

  # --- 3. ELABORACIÓN Y GESTIÓN DE ASIENTOS ---
  elif menu == "3. Elaboración y Gestión de Asientos" and empresa_id:
    st.header(f"Asientos Contables - {datos_emp['razon_social']}")

    tab_crear, tab_consultar = st.tabs([
        "✍️ Registrar Nuevo Asiento",
        "📋 Consultar, Modificar, Borrar e Imprimir Asientos",
    ])

    df_cat = pd.read_sql_query(
        "SELECT id, codigo, nombre FROM catalogo_cuentas WHERE empresa_id = ?"
        " AND acepta_movimiento = 1 ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )
    opciones = [f"{r['codigo']} - {r['nombre']}" for _, r in df_cat.iterrows()]
    mapa_id = dict(zip(opciones, df_cat["id"]))
    mapa_id_a_texto = dict(zip(df_cat["id"], opciones))

    with tab_crear:
      cur = conn.cursor()
      cur.execute(
          "SELECT COALESCE(MAX(num_asiento), 0) + 1 FROM asientos_cabecera"
          " WHERE empresa_id = ?",
          (empresa_id,),
      )
      sig_num = cur.fetchone()[0]

      c1, c2, c3, c4 = st.columns([1, 1, 1.5, 2])
      num_as = c1.number_input("N° Asiento", value=int(sig_num), min_value=1)
      f_as = c2.date_input("Fecha", value=date.today(), format="DD-MM-YYYY")
      comp = c3.text_input("N° Comprobante / Ref")
      conc = c4.text_input("Concepto General")

      if "lineas_asiento" not in st.session_state:
        st.session_state.lineas_asiento = [
            {
                "cta": opciones[0] if opciones else "",
                "debe": 0.0,
                "haber": 0.0,
            },
            {
                "cta": opciones[0] if opciones else "",
                "debe": 0.0,
                "haber": 0.0,
            },
        ]

      col_btn_add, col_btn_rem = st.columns([1, 5])
      if col_btn_add.button("➕ Agregar Línea"):
        st.session_state.lineas_asiento.append(
            {"cta": opciones[0] if opciones else "", "debe": 0.0, "haber": 0.0}
        )
      if len(st.session_state.lineas_asiento) > 2 and col_btn_rem.button(
          "➖ Quitar Última Línea"
      ):
        st.session_state.lineas_asiento.pop()

      for i in range(len(st.session_state.lineas_asiento)):
        col1, col2, col3 = st.columns([3, 1, 1])
        st.session_state.lineas_asiento[i]["cta"] = col1.selectbox(
            f"Cuenta #{i+1}", opciones, key=f"sel_c_{i}"
        )
        st.session_state.lineas_asiento[i]["debe"] = col2.number_input(
            "Debe (₡)", min_value=0.0, step=100.0, key=f"in_d_{i}"
        )
        st.session_state.lineas_asiento[i]["haber"] = col3.number_input(
            "Haber (₡)", min_value=0.0, step=100.0, key=f"in_h_{i}"
        )

      td = sum(x["debe"] for x in st.session_state.lineas_asiento)
      th = sum(x["haber"] for x in st.session_state.lineas_asiento)
      st.markdown(
          f"**Debe:** `₡{td:,.2f}` | **Haber:** `₡{th:,.2f}` |"
          f" **Diferencia:** `₡{abs(td - th):,.2f}`"
      )

      if st.button("💾 Guardar Asiento", type="primary"):
        if td == 0 and th == 0:
          st.error("Debes digitar montos mayores a cero.")
        elif round(td, 2) != round(th, 2):
          st.error("El asiento está descuadrado (Debe ≠ Haber).")
        else:
          cur.execute(
              "INSERT INTO asientos_cabecera (empresa_id, num_asiento, fecha,"
              " periodo, concepto, num_comprobante, fecha_registro) VALUES (?,"
              " ?, ?, ?, ?, ?, ?)",
              (
                  empresa_id,
                  int(num_as),
                  str(f_as),
                  str(f_as)[:7],
                  conc,
                  comp,
                  str(date.today()),
              ),
          )
          cab_id = cur.lastrowid
          for l in st.session_state.lineas_asiento:
            if l["debe"] > 0 or l["haber"] > 0:
              cur.execute(
                  "INSERT INTO asientos_detalle (asiento_id, cuenta_id, debe,"
                  " haber) VALUES (?, ?, ?, ?)",
                  (
                      cab_id,
                      int(mapa_id[l["cta"]]),
                      float(l["debe"]),
                      float(l["haber"]),
                  ),
              )
          conn.commit()
          ejecutar_backup_automatico()
          st.session_state.lineas_asiento = [
              {"cta": opciones[0], "debe": 0.0, "haber": 0.0},
              {"cta": opciones[0], "debe": 0.0, "haber": 0.0},
          ]
          st.success(f"Asiento #{num_as} registrado exitosamente.")
          st.rerun()

    with tab_consultar:
      st.subheader("Consulta de Asientos Elaborados")
      fc1, fc2, fc3 = st.columns([1, 1, 2])
      f_desde = fc1.date_input(
          "Fecha Desde:",
          value=date(2025, 1, 1),
          format="DD-MM-YYYY",
          key="as_f_desde",
      )
      f_hasta = fc2.date_input(
          "Fecha Hasta:",
          value=date(2026, 12, 31),
          format="DD-MM-YYYY",
          key="as_f_hasta",
      )
      busqueda = fc3.text_input(
          "Buscar por N° Asiento, Concepto o Comprobante:", key="as_busq"
      )

      q_listado = """
                SELECT 
                    c.id AS asiento_id,
                    c.num_asiento,
                    c.fecha,
                    c.num_comprobante,
                    c.concepto,
                    d.id AS detalle_id,
                    d.cuenta_id,
                    cat.codigo,
                    cat.nombre AS cuenta_nombre,
                    d.debe,
                    d.haber
                FROM asientos_cabecera c
                JOIN asientos_detalle d ON c.id = d.asiento_id
                JOIN catalogo_cuentas cat ON d.cuenta_id = cat.id
                WHERE c.empresa_id = ? AND c.fecha BETWEEN ? AND ?
                ORDER BY c.fecha DESC, c.num_asiento DESC, d.id ASC
            """
      df_listado = pd.read_sql_query(
          q_listado, conn, params=(empresa_id, str(f_desde), str(f_hasta))
      )

      if busqueda.strip():
        b_low = busqueda.strip().lower()
        df_listado = df_listado[
            df_listado["num_asiento"].astype(str).str.contains(b_low)
            | df_listado["concepto"]
            .fillna("")
            .str.lower()
            .str.contains(b_low)
            | df_listado["num_comprobante"]
            .fillna("")
            .str.lower()
            .str.contains(b_low)
        ]

      if df_listado.empty:
        st.info("No se encontraron asientos en el rango de fechas seleccionado.")
      else:
        st.markdown("#### 🖨️ Imprimir / Descargar Reporte de Asientos")
        dl1, dl2, dl3 = st.columns(3)
        dl1.download_button(
            "📄 Descargar Word (.docx)",
            data=generar_word_asientos(
                df_listado,
                datos_emp["razon_social"],
                datos_emp["cedula"],
                f_desde,
                f_hasta,
            ),
            file_name=f"Reporte_Asientos_{f_desde.strftime('%d-%m-%Y')}_{f_hasta.strftime('%d-%m-%Y')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        dl2.download_button(
            "📊 Descargar Excel (.xlsx)",
            data=generar_excel_asientos(
                df_listado,
                datos_emp["razon_social"],
                datos_emp["cedula"],
                f_desde,
                f_hasta,
            ),
            file_name=f"Reporte_Asientos_{f_desde.strftime('%d-%m-%Y')}_{f_hasta.strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        dl3.download_button(
            "📕 Descargar PDF (.pdf)",
            data=generar_pdf_asientos(
                df_listado,
                datos_emp["razon_social"],
                datos_emp["cedula"],
                f_desde,
                f_hasta,
            ),
            file_name=f"Reporte_Asientos_{f_desde.strftime('%d-%m-%Y')}_{f_hasta.strftime('%d-%m-%Y')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

        st.divider()
        st.markdown(
            "#### ✏️ Selecciona un asiento para Modificarlo o Eliminarlo"
        )

        asientos_unicos = df_listado[
            ["asiento_id", "num_asiento", "fecha", "num_comprobante", "concepto"]
        ].drop_duplicates()

        for _, cab in asientos_unicos.iterrows():
          aid = int(cab["asiento_id"])
          detalles = df_listado[df_listado["asiento_id"] == aid]
          tot_as = detalles["debe"].sum()

          titulo_exp = (
              f"Asiento #{cab['num_asiento']} | Fecha:"
              f" {fmt_fecha_dmy(cab['fecha'])} | Comp:"
              f" {cab['num_comprobante'] or 'S/N'} | {cab['concepto'] or ''} |"
              f" Total: ₡{tot_as:,.2f}"
          )
          with st.expander(titulo_exp):
            with st.form(key=f"edit_full_as_{aid}"):
              ec1, ec2, ec3 = st.columns([1, 1.2, 2])
              ed_fecha = ec1.text_input(
                  "Fecha (DD-MM-AAAA)", value=fmt_fecha_dmy(cab["fecha"])
              )
              ed_comp = ec2.text_input(
                  "Comprobante", value=str(cab["num_comprobante"] or "")
              )
              ed_conc = ec3.text_input(
                  "Concepto", value=str(cab["concepto"] or "")
              )

              lineas_editadas = []
              for _, drow in detalles.iterrows():
                did = int(drow["detalle_id"])
                cta_actual_txt = mapa_id_a_texto.get(
                    drow["cuenta_id"], opciones[0] if opciones else ""
                )
                idx_def = (
                    opciones.index(cta_actual_txt)
                    if cta_actual_txt in opciones
                    else 0
                )

                lc1, lc2, lc3 = st.columns([3, 1, 1])
                nueva_cta_txt = lc1.selectbox(
                    "Cuenta", opciones, index=idx_def, key=f"ed_cta_{did}"
                )
                nuevo_d = lc2.number_input(
                    "Debe (₡)",
                    value=float(drow["debe"]),
                    min_value=0.0,
                    step=100.0,
                    key=f"ed_d_{did}",
                )
                nuevo_h = lc3.number_input(
                    "Haber (₡)",
                    value=float(drow["haber"]),
                    min_value=0.0,
                    step=100.0,
                    key=f"ed_h_{did}",
                )
                lineas_editadas.append({
                    "detalle_id": did,
                    "cuenta_id": int(mapa_id[nueva_cta_txt]),
                    "debe": nuevo_d,
                    "haber": nuevo_h,
                })

              btn_c1, btn_c2 = st.columns(2)
              guardar_ed = btn_c1.form_submit_button(
                  "💾 Guardar Cambios del Asiento",
                  type="primary",
                  use_container_width=True,
              )
              borrar_ed = btn_c2.form_submit_button(
                  "🗑️ Eliminar Asiento Completo", use_container_width=True
              )

              if guardar_ed:
                sd = sum(x["debe"] for x in lineas_editadas)
                sh = sum(x["haber"] for x in lineas_editadas)
                if round(sd, 2) != round(sh, 2):
                  st.error(
                      f"❌ Descuadre: Debe (₡{sd:,.2f}) ≠ Haber (₡{sh:,.2f})"
                  )
                else:
                  fecha_iso = parse_fecha_input(ed_fecha)
                  cur = conn.cursor()
                  cur.execute(
                      "UPDATE asientos_cabecera SET fecha=?, periodo=?,"
                      " num_comprobante=?, concepto=? WHERE id=?",
                      (fecha_iso, fecha_iso[:7], ed_comp, ed_conc, aid),
                  )
                  for item in lineas_editadas:
                    cur.execute(
                        "UPDATE asientos_detalle SET cuenta_id=?, debe=?,"
                        " haber=? WHERE id=?",
                        (
                            item["cuenta_id"],
                            item["debe"],
                            item["haber"],
                            item["detalle_id"],
                        ),
                    )
                  conn.commit()
                  ejecutar_backup_automatico()
                  st.success(f"✅ Asiento #{cab['num_asiento']} actualizado.")
                  st.rerun()

              if borrar_ed:
                cur = conn.cursor()
                cur.execute(
                    "DELETE FROM asientos_detalle WHERE asiento_id=?", (aid,)
                )
                cur.execute("DELETE FROM asientos_cabecera WHERE id=?", (aid,))
                conn.commit()
                ejecutar_backup_automatico()
                st.warning(f"🗑️ Asiento #{cab['num_asiento']} eliminado.")
                st.rerun()

  # --- 4. LIBRO DIARIO Y MAYOR ---
  elif menu == "4. Libro Diario y Mayor" and empresa_id:
    st.header("Libro Diario y Mayor")
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    if not df_mov.empty:
      df_mov_vista = df_mov.sort_values(["fecha", "num_asiento"]).copy()
      df_mov_vista["fecha"] = df_mov_vista["fecha"].apply(fmt_fecha_dmy)
      st.dataframe(
          df_mov_vista[[
              "num_asiento",
              "fecha",
              "codigo",
              "cuenta_nombre",
              "concepto",
              "debe",
              "haber",
          ]],
          use_container_width=True,
          hide_index=True,
      )
    else:
      st.info("No hay movimientos registrados.")

  # --- 5. BALANCE DE COMPROBACIÓN DE 6 COLUMNAS ---
  elif menu == "5. Balance de Comprobación" and empresa_id:
    st.header(f"Balance de Comprobación - {datos_emp['razon_social']}")
    st.info(
        "💡 Haz clic en el cuadro izquierdo de cualquier cuenta con movimientos"
        " para ver, editar o eliminar sus asientos."
    )

    c1, c2, c3 = st.columns([1, 1, 2])
    f_ini = c1.date_input(
        "Del:", value=date(2026, 1, 1), format="DD-MM-YYYY"
    )
    f_fin = c2.date_input(
        "Al:", value=date(2026, 6, 30), format="DD-MM-YYYY"
    )
    rep_leg_bc = c3.text_input(
        "Representante Legal (Firma derecha):",
        value=str(datos_emp["representante_legal"] or ""),
    )

    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT id, codigo, nombre, tipo, acepta_movimiento FROM"
        " catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    df_res, totales = construir_balance_comprobacion_6col(
        df_mov, df_cat, f_ini, f_fin
    )

    vista_bc = df_res[[
        "codigo",
        "nombre",
        "ant_d",
        "ant_c",
        "mov_d",
        "mov_c",
        "fin_d",
        "fin_c",
    ]]

    sel = st.dataframe(
        vista_bc,
        use_container_width=True,
        hide_index=True,
        height=540,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "codigo": "Número de cuenta",
            "nombre": "Cuenta contable",
            "ant_d": st.column_config.NumberColumn(
                "Saldo Ant. Débito", format="₡%.2f"
            ),
            "ant_c": st.column_config.NumberColumn(
                "Saldo Ant. Crédito", format="₡%.2f"
            ),
            "mov_d": st.column_config.NumberColumn(
                "Movimientos Débito", format="₡%.2f"
            ),
            "mov_c": st.column_config.NumberColumn(
                "Movimientos Crédito", format="₡%.2f"
            ),
            "fin_d": st.column_config.NumberColumn(
                "Saldo Final Débito", format="₡%.2f"
            ),
            "fin_c": st.column_config.NumberColumn(
                "Saldo Final Crédito", format="₡%.2f"
            ),
        },
    )

    t1, t2, t3 = st.columns(3)
    t1.metric(
        "Saldo Anterior (Débito / Crédito)",
        f"₡{totales['ant_d']:,.2f} | ₡{totales['ant_c']:,.2f}",
    )
    t2.metric(
        "Movimientos (Débito / Crédito)",
        f"₡{totales['mov_d']:,.2f} | ₡{totales['mov_c']:,.2f}",
    )
    t3.metric(
        "Saldo Final (Débito / Crédito)",
        f"₡{totales['fin_d']:,.2f} | ₡{totales['fin_c']:,.2f}",
    )

    if sel.selection.rows:
      idx = sel.selection.rows[0]
      fila_sel = df_res.iloc[idx]
      if fila_sel["es_hoja"] and (
          fila_sel["mov_d"] > 0
          or fila_sel["mov_c"] > 0
          or fila_sel["ant_d"] > 0
          or fila_sel["ant_c"] > 0
      ):
        modal_editar_movimientos(
            empresa_id,
            fila_sel["cuenta_id"],
            fila_sel["codigo"],
            fila_sel["nombre"],
        )

    st.subheader("📥 Descargar Balance de Comprobación")
    b1, b2, b3 = st.columns(3)
    b1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_balance_6col(
            df_res,
            totales,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            f_ini,
            f_fin,
            rep_leg_bc,
        ),
        file_name=f"Balance_Comprobacion_{f_fin.strftime('%d-%m-%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    b2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_balance_6col(
            df_res,
            totales,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            f_ini,
            f_fin,
            rep_leg_bc,
        ),
        file_name=f"Balance_Comprobacion_{f_fin.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    b3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_balance_6col(
            df_res,
            totales,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            f_ini,
            f_fin,
            rep_leg_bc,
        ),
        file_name=f"Balance_Comprobacion_{f_fin.strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  # --- 6. ESTADO DE LA SITUACIÓN FINANCIERA COMPARATIVO ---
  elif menu == "6. Estado de la Situación Financiera" and empresa_id:
    st.header("Estado de la Situación Financiera Comparativo")
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    with st.expander("⚙️ Verificar Encabezado y Firmantes", expanded=False):
      c1, c2 = st.columns([2, 1])
      emp_nombre = c1.text_input(
          "Entidad", value=str(datos_emp["razon_social"] or "")
      )
      emp_ced = c2.text_input(
          "Cédula Jurídica", value=str(datos_emp["cedula"] or "")
      )
      c3, c4 = st.columns(2)
      rep_leg = c3.text_input(
          "Representante Legal",
          value=str(datos_emp["representante_legal"] or ""),
      )
      cargo_rep = c4.text_input(
          "Cargo", value="Presidente / Representante Legal"
      )

    col_f1, col_f2 = st.columns(2)
    fecha_ant = col_f1.date_input(
        "Fecha Corte Comparativo (Año Anterior)",
        value=date(2025, 12, 31),
        format="DD-MM-YYYY",
    )
    fecha_act = col_f2.date_input(
        "Fecha Corte Actual", value=date(2026, 6, 30), format="DD-MM-YYYY"
    )

    str_f_ant = fecha_ant.strftime("%d-%m-%Y")
    str_f_act = fecha_act.strftime("%d-%m-%Y")
    sub_periodo = f"COMPARATIVOS {str_f_ant[-4:]} - {str_f_act[-4:]}"

    estructura_base = construir_estado_situacion_comparativo(
        df_mov, df_cat, fecha_ant, fecha_act
    )

    st.info(
        "✏️ **Numeración de Notas (Columna N):** Haz doble clic en cualquier"
        " celda de la columna **N** para modificar, agregar o borrar el número"
        " de nota antes de descargar."
    )

    df_vista = pd.DataFrame([
        {
            "Descripción": nombre,
            "N": nota,
            str_f_ant: fmt_moneda(v_ant),
            str_f_act: fmt_moneda(v_act),
        }
        for _, nombre, nota, v_ant, v_act in estructura_base
    ])

    df_editado = st.data_editor(
        df_vista,
        use_container_width=True,
        hide_index=True,
        height=540,
        disabled=["Descripción", str_f_ant, str_f_act],
        column_config={
            "Descripción": st.column_config.TextColumn(
                "Descripción", width="large"
            ),
            "N": st.column_config.TextColumn("N (Nota)", width="small"),
            str_f_ant: st.column_config.TextColumn(str_f_ant, width="medium"),
            str_f_act: st.column_config.TextColumn(str_f_act, width="medium"),
        },
        key=f"editor_notas_sf_{empresa_id}_{fecha_ant}_{fecha_act}",
    )

    notas_actualizadas = df_editado["N"].fillna("").astype(str).tolist()
    estructura_final = [
        (tipo, nombre, notas_actualizadas[idx], v_ant, v_act)
        for idx, (tipo, nombre, _, v_ant, v_act) in enumerate(estructura_base)
    ]

    st.subheader("📥 Descargar Reporte Oficial")
    b1, b2, b3 = st.columns(3)
    b1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_estado_financiero(
            "ESTADO DE LA SITUACIÓN FINANCIERA",
            sub_periodo,
            estructura_final,
            emp_nombre,
            emp_ced,
            str_f_ant,
            str_f_act,
            rep_leg,
            cargo_rep,
        ),
        file_name=f"Estado_Situacion_Financiera_{str_f_act}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    b2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_estado_financiero(
            "ESTADO DE LA SITUACIÓN FINANCIERA",
            sub_periodo,
            estructura_final,
            emp_nombre,
            emp_ced,
            str_f_ant,
            str_f_act,
            rep_leg,
            cargo_rep,
        ),
        file_name=f"Estado_Situacion_Financiera_{str_f_act}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    b3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_estado_financiero(
            "ESTADO DE LA SITUACIÓN FINANCIERA",
            sub_periodo,
            estructura_final,
            emp_nombre,
            emp_ced,
            str_f_ant,
            str_f_act,
            rep_leg,
            cargo_rep,
        ),
        file_name=f"Estado_Situacion_Financiera_{str_f_act}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  # --- 7. ESTADO DE RESULTADOS COMPARATIVO (NIIF) ---
  elif menu == "7. Estado de Resultados" and empresa_id:
    st.header("Estado de Resultados Comparativo")
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    with st.expander("⚙️ Verificar Encabezado y Firmantes", expanded=False):
      c1, c2 = st.columns([2, 1])
      emp_nombre_er = c1.text_input(
          "Entidad",
          value=str(datos_emp["razon_social"] or ""),
          key="er_emp",
      )
      emp_ced_er = c2.text_input(
          "Cédula Jurídica",
          value=str(datos_emp["cedula"] or ""),
          key="er_ced",
      )
      c3, c4 = st.columns(2)
      rep_leg_er = c3.text_input(
          "Representante Legal",
          value=str(datos_emp["representante_legal"] or ""),
          key="er_rep",
      )
      cargo_rep_er = c4.text_input(
          "Cargo",
          value="Presidente / Representante Legal",
          key="er_cargo",
      )

    st.markdown("#### 📅 Rangos de Fechas Comparativos y Nivel de Detalle")
    r_col1, r_col2, r_col3, r_col4, r_col5 = st.columns([1, 1, 1, 1, 1.4])
    p1_ini = r_col1.date_input(
        "Período 1 Desde (Comp.):",
        value=date(2025, 1, 1),
        format="DD-MM-YYYY",
        key="er_p1_ini",
    )
    p1_fin = r_col2.date_input(
        "Período 1 Hasta (Comp.):",
        value=date(2025, 12, 31),
        format="DD-MM-YYYY",
        key="er_p1_fin",
    )
    p2_ini = r_col3.date_input(
        "Período 2 Desde (Actual):",
        value=date(2026, 1, 1),
        format="DD-MM-YYYY",
        key="er_p2_ini",
    )
    p2_fin = r_col4.date_input(
        "Período 2 Hasta (Actual):",
        value=date(2026, 6, 30),
        format="DD-MM-YYYY",
        key="er_p2_fin",
    )
    modo_desglose = r_col5.radio(
        "Nivel de Cuentas:",
        ["Cuentas Principales (Resumido)", "Desglosar Subcuentas (Detallado)"],
        key="er_desglose",
    )
    desglosar = modo_desglose == "Desglosar Subcuentas (Detallado)"

    col1_hdr = f"AL {p1_fin.strftime('%d-%m-%Y')}"
    col2_hdr = f"AL {p2_fin.strftime('%d-%m-%Y')}"
    sub_periodo_er = (
        f"DEL {p2_ini.strftime('%d-%m-%Y')} AL {p2_fin.strftime('%d-%m-%Y')}"
        f" (COMPARATIVO {p1_fin.strftime('%Y')} - {p2_fin.strftime('%Y')})"
    )

    estructura_er = construir_estado_resultados_comparativo(
        df_mov,
        df_cat,
        p1_ini,
        p1_fin,
        p2_ini,
        p2_fin,
        desglosar_subcuentas=desglosar,
    )

    st.info(
        "✏️ **Numeración de Notas (Columna N):** Haz doble clic en cualquier"
        " celda de la columna **N** para ajustar los números de nota antes de"
        " exportar."
    )

    df_vista_er = pd.DataFrame([
        {
            "Descripción": nombre,
            "N": nota,
            col1_hdr: fmt_moneda(v1),
            col2_hdr: fmt_moneda(v2),
        }
        for _, nombre, nota, v1, v2 in estructura_er
    ])

    df_editado_er = st.data_editor(
        df_vista_er,
        use_container_width=True,
        hide_index=True,
        height=520,
        disabled=["Descripción", col1_hdr, col2_hdr],
        column_config={
            "Descripción": st.column_config.TextColumn(
                "Descripción", width="large"
            ),
            "N": st.column_config.TextColumn("N (Nota)", width="small"),
            col1_hdr: st.column_config.TextColumn(col1_hdr, width="medium"),
            col2_hdr: st.column_config.TextColumn(col2_hdr, width="medium"),
        },
        key=(
            f"editor_notas_er_{empresa_id}_{p1_ini}_{p1_fin}_{p2_ini}_{p2_fin}_{desglosar}"
        ),
    )

    notas_er_act = df_editado_er["N"].fillna("").astype(str).tolist()
    estructura_er_final = [
        (tipo, nombre, notas_er_act[idx], v1, v2)
        for idx, (tipo, nombre, _, v1, v2) in enumerate(estructura_er)
    ]

    st.subheader("📥 Descargar Estado de Resultados Oficial")
    eb1, eb2, eb3 = st.columns(3)
    eb1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_estado_financiero(
            "ESTADO DE RESULTADOS",
            sub_periodo_er,
            estructura_er_final,
            emp_nombre_er,
            emp_ced_er,
            col1_hdr,
            col2_hdr,
            rep_leg_er,
            cargo_rep_er,
        ),
        file_name=f"Estado_Resultados_{p2_fin.strftime('%d-%m-%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    eb2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_estado_financiero(
            "ESTADO DE RESULTADOS",
            sub_periodo_er,
            estructura_er_final,
            emp_nombre_er,
            emp_ced_er,
            col1_hdr,
            col2_hdr,
            rep_leg_er,
            cargo_rep_er,
        ),
        file_name=f"Estado_Resultados_{p2_fin.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    eb3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_estado_financiero(
            "ESTADO DE RESULTADOS",
            sub_periodo_er,
            estructura_er_final,
            emp_nombre_er,
            emp_ced_er,
            col1_hdr,
            col2_hdr,
            rep_leg_er,
            cargo_rep_er,
        ),
        file_name=f"Estado_Resultados_{p2_fin.strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  # --- 8. ESTADO DE FLUJOS DE EFECTIVO COMPARATIVO ---
  elif menu == "8. Estado de Flujo de Efectivo" and empresa_id:
    st.header("Estado de Flujos de Efectivo Comparativo")
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    with st.expander("⚙️ Verificar Encabezado y Firmantes", expanded=False):
      c1, c2 = st.columns([2, 1])
      emp_nombre_fe = c1.text_input(
          "Entidad",
          value=str(datos_emp["razon_social"] or ""),
          key="fe_emp",
      )
      emp_ced_fe = c2.text_input(
          "Cédula Jurídica",
          value=str(datos_emp["cedula"] or ""),
          key="fe_ced",
      )
      c3, c4 = st.columns(2)
      rep_leg_fe = c3.text_input(
          "Representante Legal",
          value=str(datos_emp["representante_legal"] or ""),
          key="fe_rep",
      )
      cargo_rep_fe = c4.text_input(
          "Cargo",
          value="Presidente / Representante Legal",
          key="fe_cargo",
      )

    st.markdown("#### 📅 Rangos de Fechas Comparativos por Año")
    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    fe_p1_ini = f_col1.date_input(
        "Período 1 Desde (Año Base):",
        value=date(2025, 1, 1),
        format="DD-MM-YYYY",
        key="fe_p1_ini",
    )
    fe_p1_fin = f_col2.date_input(
        "Período 1 Hasta (Año Base):",
        value=date(2025, 12, 31),
        format="DD-MM-YYYY",
        key="fe_p1_fin",
    )
    fe_p2_ini = f_col3.date_input(
        "Período 2 Desde (Año Actual):",
        value=date(2026, 1, 1),
        format="DD-MM-YYYY",
        key="fe_p2_ini",
    )
    fe_p2_fin = f_col4.date_input(
        "Período 2 Hasta (Año Actual):",
        value=date(2026, 6, 30),
        format="DD-MM-YYYY",
        key="fe_p2_fin",
    )

    col1_fe_hdr = f"AL {fe_p1_fin.strftime('%d-%m-%Y')}"
    col2_fe_hdr = f"AL {fe_p2_fin.strftime('%d-%m-%Y')}"
    sub_periodo_fe = (
        f"DEL {fecha_texto_es(fe_p1_ini).upper()} AL"
        f" {fecha_texto_es(fe_p1_fin).upper()} Y DEL"
        f" {fecha_texto_es(fe_p2_ini).upper()} AL"
        f" {fecha_texto_es(fe_p2_fin).upper()}"
    )

    estructura_fe = construir_estado_flujo_efectivo_comparativo(
        df_mov, df_cat, fe_p1_ini, fe_p1_fin, fe_p2_ini, fe_p2_fin
    )

    st.info(
        "✏️ **Numeración de Notas (Columna N):** Haz doble clic en cualquier"
        " celda de la columna **N** si deseas referenciar una nota explicativa."
    )

    df_vista_fe = pd.DataFrame([
        {
            "Descripción": nombre,
            "N": nota,
            col1_fe_hdr: fmt_moneda(v1),
            col2_fe_hdr: fmt_moneda(v2),
        }
        for _, nombre, nota, v1, v2 in estructura_fe
    ])

    df_editado_fe = st.data_editor(
        df_vista_fe,
        use_container_width=True,
        hide_index=True,
        height=560,
        disabled=["Descripción", col1_fe_hdr, col2_fe_hdr],
        column_config={
            "Descripción": st.column_config.TextColumn(
                "Descripción", width="large"
            ),
            "N": st.column_config.TextColumn("N (Nota)", width="small"),
            col1_fe_hdr: st.column_config.TextColumn(
                col1_fe_hdr, width="medium"
            ),
            col2_fe_hdr: st.column_config.TextColumn(
                col2_fe_hdr, width="medium"
            ),
        },
        key=f"editor_notas_fe_{empresa_id}_{fe_p1_fin}_{fe_p2_fin}",
    )

    notas_fe_act = df_editado_fe["N"].fillna("").astype(str).tolist()
    estructura_fe_final = [
        (tipo, nombre, notas_fe_act[idx], v1, v2)
        for idx, (tipo, nombre, _, v1, v2) in enumerate(estructura_fe)
    ]

    st.subheader("📥 Descargar Estado de Flujos de Efectivo Oficial")
    feb1, feb2, feb3 = st.columns(3)
    feb1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_estado_financiero(
            "ESTADO DE FLUJOS DE EFECTIVO",
            sub_periodo_fe,
            estructura_fe_final,
            emp_nombre_fe,
            emp_ced_fe,
            col1_fe_hdr,
            col2_fe_hdr,
            rep_leg_fe,
            cargo_rep_fe,
        ),
        file_name=f"Estado_Flujos_Efectivo_{fe_p2_fin.strftime('%d-%m-%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    feb2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_estado_financiero(
            "ESTADO DE FLUJOS DE EFECTIVO",
            sub_periodo_fe,
            estructura_fe_final,
            emp_nombre_fe,
            emp_ced_fe,
            col1_fe_hdr,
            col2_fe_hdr,
            rep_leg_fe,
            cargo_rep_fe,
        ),
        file_name=f"Estado_Flujos_Efectivo_{fe_p2_fin.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    feb3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_estado_financiero(
            "ESTADO DE FLUJOS DE EFECTIVO",
            sub_periodo_fe,
            estructura_fe_final,
            emp_nombre_fe,
            emp_ced_fe,
            col1_fe_hdr,
            col2_fe_hdr,
            rep_leg_fe,
            cargo_rep_fe,
        ),
        file_name=f"Estado_Flujos_Efectivo_{fe_p2_fin.strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  # --- 9. ESTADO DE CAMBIOS EN EL PATRIMONIO COMPARATIVO ---
  elif menu == "9. Estado de Cambios en el Patrimonio" and empresa_id:
    st.header("Estado de Cambios en el Patrimonio Comparativo")
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    with st.expander("⚙️ Verificar Encabezado y Firmantes", expanded=False):
      c1, c2 = st.columns([2, 1])
      emp_nombre_cp = c1.text_input(
          "Entidad",
          value=str(datos_emp["razon_social"] or ""),
          key="cp_emp",
      )
      emp_ced_cp = c2.text_input(
          "Cédula Jurídica",
          value=str(datos_emp["cedula"] or ""),
          key="cp_ced",
      )
      c3, c4 = st.columns(2)
      rep_leg_cp = c3.text_input(
          "Representante Legal",
          value=str(datos_emp["representante_legal"] or ""),
          key="cp_rep",
      )
      cargo_rep_cp = c4.text_input(
          "Cargo",
          value="Presidente / Representante Legal",
          key="cp_cargo",
      )

    st.markdown("#### 📅 Rangos de Fechas Comparativos por Período")
    cp_col1, cp_col2, cp_col3, cp_col4 = st.columns(4)
    cp_p1_ini = cp_col1.date_input(
        "Período 1 Desde (Comp.):",
        value=date(2025, 1, 1),
        format="DD-MM-YYYY",
        key="cp_p1_ini",
    )
    cp_p1_fin = cp_col2.date_input(
        "Período 1 Hasta (Comp.):",
        value=date(2025, 12, 31),
        format="DD-MM-YYYY",
        key="cp_p1_fin",
    )
    cp_p2_ini = cp_col3.date_input(
        "Período 2 Desde (Actual):",
        value=date(2026, 1, 1),
        format="DD-MM-YYYY",
        key="cp_p2_ini",
    )
    cp_p2_fin = cp_col4.date_input(
        "Período 2 Hasta (Actual):",
        value=date(2026, 6, 30),
        format="DD-MM-YYYY",
        key="cp_p2_fin",
    )

    sub_periodo_cp = (
        f"POR LOS PERÍODOS TERMINADOS AL {cp_p1_fin.strftime('%d-%m-%Y')} Y AL"
        f" {cp_p2_fin.strftime('%d-%m-%Y')}"
    )

    estructura_cp = construir_estado_cambios_patrimonio(
        df_mov, df_cat, cp_p1_ini, cp_p1_fin, cp_p2_ini, cp_p2_fin
    )

    st.info(
        "✏️ **Numeración de Notas (Columna N):** Haz doble clic en cualquier"
        " celda de la columna **N** para agregar o editar referencias a notas"
        " contables."
    )

    df_vista_cp = pd.DataFrame([
        {
            "Concepto": concepto,
            "N": nota,
            "Capital Social": fmt_moneda(v_cap),
            "Reservas": fmt_moneda(v_res),
            "Excedentes / Utilidades Acumuladas": fmt_moneda(v_uti),
            "Total Patrimonio": fmt_moneda(v_tot),
        }
        for _, concepto, nota, v_cap, v_res, v_uti, v_tot in estructura_cp
    ])

    df_editado_cp = st.data_editor(
        df_vista_cp,
        use_container_width=True,
        hide_index=True,
        height=440,
        disabled=[
            "Concepto",
            "Capital Social",
            "Reservas",
            "Excedentes / Utilidades Acumuladas",
            "Total Patrimonio",
        ],
        column_config={
            "Concepto": st.column_config.TextColumn("Concepto", width="large"),
            "N": st.column_config.TextColumn("N (Nota)", width="small"),
        },
        key=f"editor_notas_cp_{empresa_id}_{cp_p1_fin}_{cp_p2_fin}",
    )

    notas_cp_act = df_editado_cp["N"].fillna("").astype(str).tolist()
    estructura_cp_final = [
        (tipo, concepto, notas_cp_act[idx], v_cap, v_res, v_uti, v_tot)
        for idx, (tipo, concepto, _, v_cap, v_res, v_uti, v_tot) in enumerate(
            estructura_cp
        )
    ]

    st.subheader("📥 Descargar Estado de Cambios en el Patrimonio Oficial")
    cpb1, cpb2, cpb3 = st.columns(3)
    cpb1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_cambios_patrimonio(
            sub_periodo_cp,
            estructura_cp_final,
            emp_nombre_cp,
            emp_ced_cp,
            rep_leg_cp,
            cargo_rep_cp,
        ),
        file_name=f"Estado_Cambios_Patrimonio_{cp_p2_fin.strftime('%d-%m-%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    cpb2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_cambios_patrimonio(
            sub_periodo_cp,
            estructura_cp_final,
            emp_nombre_cp,
            emp_ced_cp,
            rep_leg_cp,
            cargo_rep_cp,
        ),
        file_name=f"Estado_Cambios_Patrimonio_{cp_p2_fin.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    cpb3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_cambios_patrimonio(
            sub_periodo_cp,
            estructura_cp_final,
            emp_nombre_cp,
            emp_ced_cp,
            rep_leg_cp,
            cargo_rep_cp,
        ),
        file_name=f"Estado_Cambios_Patrimonio_{cp_p2_fin.strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  # --- 10. FLUJO DE CAJA PROYECTADO ---
  elif menu == "10. Flujo de Caja Proyectado" and empresa_id:
    st.header(
        "Modelo de Flujo de Caja Proyectado (Año 1 Mensual + Años 2..N Anual)"
    )
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    st.markdown("#### 1️⃣ Período Base Contable para Extracción Histórica")
    b_col1, b_col2, b_col3 = st.columns([1, 1, 1.5])
    f_base_ini = b_col1.date_input(
        "Base Histórica Desde:",
        value=date(2025, 1, 1),
        format="DD-MM-YYYY",
        key="proy_b_ini",
    )
    f_base_fin = b_col2.date_input(
        "Base Histórica Hasta (Corte Caja):",
        value=date(2025, 12, 31),
        format="DD-MM-YYYY",
        key="proy_b_fin",
    )
    rep_leg_proy = b_col3.text_input(
        "Representante Legal (Firma):",
        value=str(datos_emp["representante_legal"] or ""),
        key="proy_rep",
    )

    base_hist = obtener_base_historica_proyeccion(
        df_mov, df_cat, f_base_ini, f_base_fin
    )
    meses_base = max(
        1,
        (f_base_fin.year - f_base_ini.year) * 12
        + (f_base_fin.month - f_base_ini.month)
        + 1,
    )
    ventas_anualizadas_def = (base_hist["ventas_base"] / meses_base) * 12.0
    gastos_mensuales_def = base_hist["gastos_base"] / meses_base
    pct_costo_def = (
        (base_hist["costos_base"] / base_hist["ventas_base"] * 100.0)
        if base_hist["ventas_base"] > 0
        else 33.0
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric(
        "Caja/Bancos Real al Corte", f"₡{base_hist['caja_actual']:,.2f}"
    )
    m2.metric(
        "Ventas Históricas Base", f"₡{base_hist['ventas_base']:,.2f}"
    )
    m3.metric(
        "Costos Históricos Base", f"₡{base_hist['costos_base']:,.2f}"
    )
    m4.metric(
        "Gastos Históricos Base", f"₡{base_hist['gastos_base']:,.2f}"
    )

    st.markdown("#### 2️⃣ Supuestos de Proyección y Condiciones Financieras")
    p1, p2, p3, p4 = st.columns(4)
    anio_inicio_proy = p1.number_input(
        "Año 1 de Proyección:",
        min_value=2024,
        max_value=2040,
        value=f_base_fin.year + 1,
        step=1,
    )
    anios_proyectar = p2.number_input(
        "Plazo Proyección (Años):", min_value=2, max_value=15, value=5, step=1
    )
    caja_ini_proy = p3.number_input(
        "Saldo Inicial Efectivo Año 1 (₡):",
        value=float(max(0.0, base_hist["caja_actual"])),
        step=100000.0,
        format="%.2f",
    )
    ventas_ano1_proy = p4.number_input(
        "Ventas Proyectadas Año 1 (₡):",
        value=float(
            ventas_anualizadas_def if ventas_anualizadas_def > 0 else 12000000.0
        ),
        step=500000.0,
        format="%.2f",
    )

    s1, s2, s3, s4 = st.columns(4)
    crec_mensual_v = s1.number_input(
        "Crecimiento Mensual Ventas Año 1 (%):",
        value=0.50,
        min_value=-10.0,
        max_value=50.0,
        step=0.10,
        format="%.2f",
    )
    crec_anual_v = s2.number_input(
        "Crecimiento Anual Ventas Años 2..N (%):",
        value=8.00,
        min_value=-20.0,
        max_value=100.0,
        step=0.50,
        format="%.2f",
    )
    pct_costo_v = s3.number_input(
        "Costo de Ventas (% s/ Ventas):",
        value=float(round(pct_costo_def, 2)),
        min_value=0.0,
        max_value=100.0,
        step=0.50,
        format="%.2f",
    )
    gastos_mens_a1 = s4.number_input(
        "Gastos Operativos Mensuales Año 1 (₡):",
        value=float(
            gastos_mensuales_def if gastos_mensuales_def > 0 else 500000.0
        ),
        step=50000.0,
        format="%.2f",
    )

    d1, d2, d3, d4 = st.columns(4)
    crec_anual_g = d1.number_input(
        "Crecimiento Anual Gastos / Inflación (%):",
        value=5.00,
        min_value=0.0,
        max_value=50.0,
        step=0.50,
        format="%.2f",
    )
    monto_prestamo = d2.number_input(
        "Monto Préstamo / Financiamiento (₡):",
        value=0.0,
        min_value=0.0,
        step=1000000.0,
        format="%.2f",
    )
    tasa_anual_pct = d3.number_input(
        "Tasa de Interés Anual (%):",
        value=12.00,
        min_value=0.0,
        max_value=50.0,
        step=0.25,
        format="%.2f",
    )
    plazo_prestamo_anos = d4.number_input(
        "Plazo del Préstamo (Años):", min_value=1, max_value=30, value=5, step=1
    )

    res_proy = construir_flujo_caja_proyectado(
        anio_inicio_proy,
        anios_proyectar,
        caja_ini_proy,
        ventas_ano1_proy,
        crec_mensual_v,
        crec_anual_v,
        pct_costo_v,
        gastos_mens_a1,
        crec_anual_g,
        monto_prestamo,
        tasa_anual_pct,
        plazo_prestamo_anos,
    )

    k1, k2 = st.columns(2)
    k1.metric(
        "Cuota Bancaria Nivelada Mensual",
        f"₡{res_proy['cuota_mensual']:,.2f}",
    )
    k2.metric(
        f"Saldo Final Acumulado de Caja al Año {anios_proyectar}",
        f"₡{res_proy['caja_final_proyectada']:,.2f}",
    )

    st.subheader(
        f"Sección I: Flujo de Caja Detallado Mes a Mes - Año 1 ({anio_inicio_proy})"
    )
    filas_df1 = []
    for tipo, concepto, vals in res_proy["sec1_filas"]:
      fila_dict = {"Concepto / Periodo": concepto}
      for cname, v in zip(res_proy["sec1_cols"], vals):
        if v is None:
          fila_dict[cname] = ""
        elif tipo == "RATIO":
          fila_dict[cname] = f"{v:.2f}x" if v > 0 else "N/A"
        else:
          fila_dict[cname] = fmt_moneda_entero(v)
      filas_df1.append(fila_dict)
    st.dataframe(
        pd.DataFrame(filas_df1),
        use_container_width=True,
        hide_index=True,
        height=490,
    )

    st.subheader(
        f"Sección II: Flujo de Caja Proyectado Anualizado (Años 2 a"
        f" {anios_proyectar})"
    )
    filas_df2 = []
    for tipo, concepto, vals in res_proy["sec2_filas"]:
      fila_dict = {"Concepto / Periodo": concepto}
      for cname, v in zip(res_proy["sec2_cols"], vals):
        if v is None:
          fila_dict[cname] = ""
        elif tipo == "RATIO":
          fila_dict[cname] = f"{v:.2f}x" if v > 0 else "N/A"
        else:
          fila_dict[cname] = fmt_moneda_entero(v)
      filas_df2.append(fila_dict)
    st.dataframe(
        pd.DataFrame(filas_df2),
        use_container_width=True,
        hide_index=True,
        height=490,
    )

    st.subheader("📥 Descargar Modelo de Flujo de Caja Proyectado")
    pb1, pb2, pb3 = st.columns(3)
    pb1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_flujo_proyectado(
            res_proy,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            rep_leg_proy,
            anios_proyectar,
            monto_prestamo,
            tasa_anual_pct,
            plazo_prestamo_anos,
        ),
        file_name=f"Flujo_Caja_Proyectado_{anios_proyectar}_Anos.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    pb2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_flujo_proyectado(
            res_proy,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            rep_leg_proy,
            anios_proyectar,
            monto_prestamo,
            tasa_anual_pct,
            plazo_prestamo_anos,
        ),
        file_name=f"Flujo_Caja_Proyectado_{anios_proyectar}_Anos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    pb3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_flujo_proyectado(
            res_proy,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            rep_leg_proy,
            anios_proyectar,
            monto_prestamo,
            tasa_anual_pct,
            plazo_prestamo_anos,
        ),
        file_name=f"Flujo_Caja_Proyectado_{anios_proyectar}_Anos.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  # --- 11. ANÁLISIS FINANCIERO INTEGRAL (VERTICAL, HORIZONTAL Y RAZONES) ---
  elif menu == "11. Análisis Financiero Integral" and empresa_id:
    st.header("Informe Profesional de Análisis Financiero Integral")
    df_mov = obtener_movimientos_empresa(conn, empresa_id)
    df_cat = pd.read_sql_query(
        "SELECT * FROM catalogo_cuentas WHERE empresa_id = ? ORDER BY codigo",
        conn,
        params=(empresa_id,),
    )

    af_c1, af_c2, af_c3, af_c4, af_c5 = st.columns([1, 1, 1, 1, 1.5])
    af_p1_ini = af_c1.date_input(
        "Período 1 Desde (Base):",
        value=date(2025, 1, 1),
        format="DD-MM-YYYY",
        key="af_p1_ini",
    )
    af_p1_fin = af_c2.date_input(
        "Período 1 Hasta (Corte 1):",
        value=date(2025, 12, 31),
        format="DD-MM-YYYY",
        key="af_p1_fin",
    )
    af_p2_ini = af_c3.date_input(
        "Período 2 Desde (Actual):",
        value=date(2026, 1, 1),
        format="DD-MM-YYYY",
        key="af_p2_ini",
    )
    af_p2_fin = af_c4.date_input(
        "Período 2 Hasta (Corte 2):",
        value=date(2026, 6, 30),
        format="DD-MM-YYYY",
        key="af_p2_fin",
    )
    rep_leg_af = af_c5.text_input(
        "Representante Legal (Firma):",
        value=str(datos_emp["representante_legal"] or ""),
        key="af_rep",
    )

    res_af = construir_analisis_financiero_integral(
        df_mov, df_cat, af_p1_ini, af_p1_fin, af_p2_ini, af_p2_fin
    )

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Razón Corriente Actual", f"{res_af['kpis']['rc_2']:.2f}x")
    k2.metric("Endeudamiento Total", f"{res_af['kpis']['end_2']:.2f}%")
    k3.metric("Margen Neto", f"{res_af['kpis']['mn_2']:.2f}%")
    k4.metric("Rendimiento Patrimonio (ROE)", f"{res_af['kpis']['roe_2']:.2f}%")

    tab_vh, tab_raz, tab_con = st.tabs([
        "📊 1. Análisis Vertical y Horizontal",
        "📐 2. Razones Financieras",
        "📝 3. Conclusiones y Recomendaciones",
    ])

    with tab_vh:
      st.subheader(
          "Análisis Vertical y Horizontal - Estado de Situación Financiera"
      )
      df_sf_vh = pd.DataFrame([
          {
              "Rubro / Cuenta": nom,
              f"Al {af_p1_fin.strftime('%d-%m-%Y')}": (
                  fmt_moneda(x1) if x1 is not None else ""
              ),
              "Vert. % (1)": fmt_pct(av1),
              f"Al {af_p2_fin.strftime('%d-%m-%Y')}": (
                  fmt_moneda(x2) if x2 is not None else ""
              ),
              "Vert. % (2)": fmt_pct(av2),
              "Variación Absoluta (₡)": (
                  fmt_moneda(ah_abs) if ah_abs is not None else ""
              ),
              "Variación %": fmt_pct(ah_rel),
          }
          for _, nom, x1, av1, x2, av2, ah_abs, ah_rel in res_af["tabla_sf"]
      ])
      st.dataframe(df_sf_vh, use_container_width=True, hide_index=True)

      st.subheader("Análisis Vertical y Horizontal - Estado de Resultados")
      df_er_vh = pd.DataFrame([
          {
              "Rubro / Cuenta": nom,
              f"Al {af_p1_fin.strftime('%d-%m-%Y')}": (
                  fmt_moneda(x1) if x1 is not None else ""
              ),
              "Vert. % (1)": fmt_pct(av1),
              f"Al {af_p2_fin.strftime('%d-%m-%Y')}": (
                  fmt_moneda(x2) if x2 is not None else ""
              ),
              "Vert. % (2)": fmt_pct(av2),
              "Variación Absoluta (₡)": (
                  fmt_moneda(ah_abs) if ah_abs is not None else ""
              ),
              "Variación %": fmt_pct(ah_rel),
          }
          for _, nom, x1, av1, x2, av2, ah_abs, ah_rel in res_af["tabla_er"]
      ])
      st.dataframe(df_er_vh, use_container_width=True, hide_index=True)

    with tab_raz:
      st.subheader("Indicadores de Liquidez, Endeudamiento, Actividad y Rentabilidad")
      df_raz_vista = pd.DataFrame(
          res_af["razones"],
          columns=[
              "Categoría",
              "Indicador Financiero",
              af_p1_fin.strftime("%d-%m-%Y"),
              af_p2_fin.strftime("%d-%m-%Y"),
              "Interpretación Técnica",
          ],
      )
      st.dataframe(
          df_raz_vista,
          use_container_width=True,
          hide_index=True,
          height=540,
      )

    with tab_con:
      st.info(
          "✏️ Puedes editar o complementar las conclusiones y recomendaciones"
          " generadas automáticamente antes de exportar el informe oficial."
      )
      conclusiones_edit = st.text_area(
          "Conclusiones del Análisis Financiero:",
          value=res_af["texto_conclusiones"],
          height=220,
          key=f"af_conc_{empresa_id}_{af_p1_fin}_{af_p2_fin}",
      )
      recomendaciones_edit = st.text_area(
          "Recomendaciones Profesionales:",
          value=res_af["texto_recomendaciones"],
          height=220,
          key=f"af_rec_{empresa_id}_{af_p1_fin}_{af_p2_fin}",
      )

    st.subheader("📥 Descargar Informe Profesional de Análisis Financiero")
    afb1, afb2, afb3 = st.columns(3)
    afb1.download_button(
        "📄 Descargar Word (.docx)",
        data=generar_word_analisis_financiero(
            res_af,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            af_p1_ini,
            af_p1_fin,
            af_p2_ini,
            af_p2_fin,
            conclusiones_edit,
            recomendaciones_edit,
            rep_leg_af,
        ),
        file_name=f"Informe_Analisis_Financiero_{af_p2_fin.strftime('%d-%m-%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    afb2.download_button(
        "📊 Descargar Excel (.xlsx)",
        data=generar_excel_analisis_financiero(
            res_af,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            af_p1_ini,
            af_p1_fin,
            af_p2_ini,
            af_p2_fin,
            conclusiones_edit,
            recomendaciones_edit,
            rep_leg_af,
        ),
        file_name=f"Informe_Analisis_Financiero_{af_p2_fin.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    afb3.download_button(
        "📕 Descargar PDF (.pdf)",
        data=generar_pdf_analisis_financiero(
            res_af,
            datos_emp["razon_social"],
            datos_emp["cedula"],
            af_p1_ini,
            af_p1_fin,
            af_p2_ini,
            af_p2_fin,
            conclusiones_edit,
            recomendaciones_edit,
            rep_leg_af,
        ),
        file_name=f"Informe_Analisis_Financiero_{af_p2_fin.strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
  elif menu == "12. Conciliación Bancaria" and empresa_id:
      render_modulo_conciliacion_bancaria(
          conn,
          empresa_id,
          datos_emp,
          obtener_movimientos_empresa,
          obtener_cuentas_analiticas_sin_duplicar,
          calcular_saldos_a_fecha,
          ejecutar_backup_automatico,
      )
  elif menu == "13. Reporte de IVA (Hacienda CR)" and empresa_id:
      render_modulo_reporte_iva(
          conn,
          empresa_id,
          datos_emp,
          obtener_movimientos_empresa,
          ejecutar_backup_automatico,
      )
  elif menu == "14. Reporte de Renta (Hacienda CR)" and empresa_id:
      render_modulo_reporte_renta(
          conn,
          empresa_id,
          datos_emp,
          obtener_movimientos_empresa,
          ejecutar_backup_automatico,
      )
  elif menu == "15. Administración de Empresas":
   render_gestion_empresas(conn)
subir_bd_a_sheets(conn)
conn.close()


if __name__ == "__main__":
  main()
