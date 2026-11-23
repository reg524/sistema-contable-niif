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
        df_ini["haber"] = pd.to_numeric(df_ini["haber"], errors="coerce").fillna(0.0)
        df_ini = df_ini.drop_duplicates(
            subset=["cuenta_id", "fecha", "debe", "haber"], keep="first"
        )

    if not df_ini.empty and not df_as.empty:
        df_as["debe"] = pd.to_numeric(df_as["debe"], errors="coerce").fillna(0.0)
        df_as["haber"] = pd.to_numeric(df_as["haber"], errors="coerce").fillna(0.0)
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
        df_total["debe"] = pd.to_numeric(df_total["debe"], errors="coerce").fillna(0.0)
        df_total["haber"] = pd.to_numeric(df_total["haber"], errors="coerce").fillna(0.0)
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
# 2. BALANCE DE COMPROBACIÓN DE 6 COLUMNAS (CÁLCULO Y EXPORTADORES)
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
            "_______________________________\nLic. Roy F. Esquivel"
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
# 4. GENERADORES WORD, EXCEL Y PDF PARA ESTADOS FINANCIEROS VERTICALES
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
            c_ant.number_format = '#,##0.00;(#,##0.00);"-"'
            c_act.number_format = '#,##0.00;(#,##0.00);"-"'

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
# 5. GENERADORES WORD, EXCEL Y PDF PARA CAMBIOS EN EL PATRIMONIO
# ==============================================================================
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
            cell.number_format = '#,##0.00;(#,##0.00);"-"'

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
# 6. GENERADORES WORD, EXCEL Y PDF PARA FLUJO DE CAJA PROYECTADO
# ==============================================================================
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
# 7. GENERADORES WORD, EXCEL Y PDF PARA ANÁLISIS FINANCIERO INTEGRAL
# ==============================================================================
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
                    cm.number_format = '#,##0.00;(#,##0.00);"-"'
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
