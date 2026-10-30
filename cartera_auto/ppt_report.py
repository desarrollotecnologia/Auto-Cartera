# -*- coding: utf-8 -*-
"""
Genera PPT para junta directiva:
- Conserva plantilla COLBEEF (Tesorería + títulos)
- En Cartera: tablas editables + gráficos estadísticos (sin pantallazos)
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

from .config import AGE_BUCKETS, ROOT, UNIT_ORDER
from .parsers import format_corte_label
from .transform import money_mm, pct

TEMPLATE = ROOT / "templates" / "plantilla_ig_tesoreria.pptx"

GREEN = "185E22"
GREEN_HDR = "289B38"
GREEN_SOFT = "70AD47"
GREEN_LIGHT = "C6E0B4"
RED = "C01E1E"
AMBER = "D4A017"
WHITE = "FFFFFF"
BLACK = "1A1A1A"
MUTED = "5A5A5A"
ROW_ALT = "F3F9F3"
BG_CARD = "F7FBF7"


def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def generate_pptx(model: dict, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"No está la plantilla PPT: {TEMPLATE}")

    shutil.copy2(TEMPLATE, output_path)
    prs = Presentation(str(output_path))

    corte_label = format_corte_label(model["meta"].get("corte"))
    k = model["kpis"]
    _set_slide_body_text(prs.slides[0], corte_label)

    for slide in prs.slides:
        title = _slide_title(slide).upper()
        pic = _first_picture(slide)
        if not pic:
            continue
        box = (pic.left, pic.top, pic.width, pic.height)

        if title == "CARTERA POR EDADES":
            _remove_shape(pic)
            _slide_edades_dashboard(slide, box, model)
        elif "INDICADOR CARTERA VENCIDA" in title:
            _remove_shape(pic)
            _slide_indicador(slide, box, model)
        elif "SIN OLIMPICA" in title:
            _remove_shape(pic)
            _slide_sin_olimpica(slide, box, model)
        elif "UNIDAD DE NEGOCIO" in title:
            _remove_shape(pic)
            _slide_unidades(slide, box, model)
        elif "CONSOLIDADO CARTERA POR CLIENTE" in title:
            _remove_shape(pic)
            _slide_consolidado(slide, box, model)
        elif "SUBPRODUCTOS" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "SUBPRODUCTOS")
        elif "OTRAS CUENTAS POR COBRAR" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "OTRAS CxC")
        elif "CARTERA LOCAL" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "LOCAL")
        elif "CARTERA REGIONAL" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "REGIONAL")
        elif "CARTERA NACIONAL" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "NACIONAL")
        elif "INSTITUCIONAL" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "INSTITUCIONAL")
        elif title.strip().endswith("DESPOSTE") or "3.8 DESPOSTE" in title:
            _remove_shape(pic)
            _slide_unidad_detalle(slide, box, model, "DESPOSTE")
        elif "RECORD DE CARTERA VENCIDA" in title:
            _remove_shape(pic)
            _slide_record(slide, box, model)
        elif "ROTACIÓN" in title or "ROTACION" in title:
            hist = model.get("historico") or {}
            if k.get("dias_rotacion") or hist.get("edades_hist") or hist.get("record_series"):
                _remove_shape(pic)
                _slide_rotacion(slide, box, model)

    # Insertar slide ejecutivo extra después de "2. INFORME DE CARTERA" si existe
    _try_insert_executive_summary(prs, model, corte_label)

    prs.save(str(output_path))
    return output_path


# ── helpers base ─────────────────────────────────────────────

def _slide_title(slide) -> str:
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        try:
            if sh.is_placeholder and sh.placeholder_format.idx == 0:
                return sh.text_frame.text.replace("\x0b", " ").replace("\n", " ").strip()
        except Exception:
            continue
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            return sh.text_frame.text.replace("\x0b", " ").replace("\n", " ").strip()
    return ""


def _set_slide_body_text(slide, text: str):
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        try:
            if sh.is_placeholder and sh.placeholder_format.idx == 1:
                sh.text_frame.paragraphs[0].text = text
                return
        except Exception:
            continue


def _first_picture(slide):
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return sh
    return None


def _remove_shape(shape):
    el = shape._element
    el.getparent().remove(el)


def _textbox(slide, left, top, width, height, text, size=12, bold=False, color=BLACK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Century Gothic"
    run.font.color.rgb = _rgb(color)
    return box


def _card(slide, left, top, width, height, label, value, sub=None, danger=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(BG_CARD)
    shape.line.color.rgb = _rgb(GREEN_LIGHT)
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), height)
    accent.fill.solid()
    accent.fill.fore_color.rgb = _rgb(RED if danger else GREEN_HDR)
    accent.line.fill.background()
    _textbox(slide, left + Inches(0.18), top + Inches(0.1), width - Inches(0.25), Inches(0.28), label, size=10, color=MUTED)
    _textbox(
        slide,
        left + Inches(0.18),
        top + Inches(0.38),
        width - Inches(0.25),
        Inches(0.4),
        value,
        size=18,
        bold=True,
        color=RED if danger else GREEN,
    )
    if sub:
        _textbox(slide, left + Inches(0.18), top + Inches(0.82), width - Inches(0.25), Inches(0.25), sub, size=9, color=MUTED)


def _style_header_cell(cell):
    cell.fill.solid()
    cell.fill.fore_color.rgb = _rgb(GREEN_HDR)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(9)
            run.font.color.rgb = _rgb(WHITE)
            run.font.name = "Century Gothic"


def _style_body_cell(cell, alt=False, bold=False, align_right=True):
    cell.fill.solid()
    cell.fill.fore_color.rgb = _rgb(ROW_ALT if alt else WHITE)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.RIGHT if align_right else PP_ALIGN.LEFT
        for run in p.runs:
            run.font.bold = bold
            run.font.size = Pt(8)
            run.font.color.rgb = _rgb(BLACK)
            run.font.name = "Century Gothic"


def _add_table(slide, left, top, width, height, headers, rows, first_col_left=True):
    table_shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, height)
    table = table_shape.table
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = str(h)
        _style_header_cell(cell)
    for i, row in enumerate(rows):
        is_total = any("TOTAL" in str(x).upper() for x in row[:3])
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            left_align = first_col_left and j == 0 and not str(headers[0]).startswith("#")
            if j == 1 and headers[0] in ("#", "NIT"):
                left_align = True
            _style_body_cell(cell, alt=(i % 2 == 1), bold=is_total, align_right=not left_align)
    return table_shape


def _add_chart(slide, chart_type, left, top, width, height, categories, series_map, title=None):
    data = CategoryChartData()
    data.categories = categories
    for name, values in series_map.items():
        data.add_series(name, values)
    chart_shape = slide.shapes.add_chart(chart_type, left, top, width, height, data)
    chart = chart_shape.chart
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = title
        for run in chart.chart_title.text_frame.paragraphs[0].runs:
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.name = "Century Gothic"
            run.font.color.rgb = _rgb(GREEN)
    chart.has_legend = len(series_map) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    return chart_shape


def _split_box(box):
    """Divide el área del pantallazo en franja superior (KPIs) + cuerpo."""
    left, top, width, height = box
    kpi_h = min(Inches(1.15), int(height * 0.22))
    body_top = top + kpi_h + Inches(0.08)
    body_h = height - kpi_h - Inches(0.08)
    return (left, top, width, kpi_h), (left, body_top, width, body_h)


# ── slides de junta ──────────────────────────────────────────

def _slide_edades_dashboard(slide, box, model):
    k = model["kpis"]
    kpi_box, body = _split_box(box)
    l, t, w, h = kpi_box
    card_w = int(w / 4) - Inches(0.08)
    _card(slide, l, t, card_w, h, "Total cartera", money_mm(k["total_cartera"]))
    _card(slide, l + card_w + Inches(0.1), t, card_w, h, "Corriente", money_mm(k["cartera_corriente"]), pct(k["pct_corriente"]))
    _card(slide, l + 2 * (card_w + Inches(0.1)), t, card_w, h, "Vencida", money_mm(k["cartera_vencida"]), pct(k["pct_vencida"]), danger=True)
    _card(slide, l + 3 * (card_w + Inches(0.1)), t, card_w, h, "Clientes", str(k["n_clientes"]))

    bl, bt, bw, bh = body
    # gráfico izq: composición edades (millones)
    cats = []
    vals = []
    for b in AGE_BUCKETS:
        v = float(model["edades"].get(b, 0) or 0) / 1e6
        if v > 0.05:
            cats.append(b.replace(" AL ", "-"))
            vals.append(round(v, 1))
    chart_w = int(bw * 0.48)
    _add_chart(
        slide,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        bl,
        bt,
        chart_w,
        bh,
        cats,
        {"Saldo (MM)": vals},
        title="Composición por edades (MM)",
    )
    # tabla derecha resumen
    headers = ["Concepto", "Saldo", "Corriente", "Vencido", "% Venc."]
    rows = []
    for label, tot, pv in [
        ("Con Olímpica", model["totales"], k["pct_vencida"]),
        ("Sin Olímpica", model["totales_sin_olimpica"], k["pct_vencida_sin_olimpica"]),
        ("Olímpica", model["totales_olimpica"], None),
    ]:
        s = tot.get("saldo") or 0
        v = tot.get("vencido") or 0
        if pv is None:
            pv = (v / s) if s else 0
        rows.append([label, money_mm(s), money_mm(tot.get("CORRIENTE", 0)), money_mm(v), pct(pv)])
    _add_table(slide, bl + chart_w + Inches(0.15), bt, bw - chart_w - Inches(0.15), bh, headers, rows)


def _slide_indicador(slide, box, model):
    k = model["kpis"]
    l, t, w, h = box
    # 3 KPI grandes
    cw = int(w / 3) - Inches(0.1)
    ch = Inches(1.2)
    _card(slide, l, t, cw, ch, "% Cartera vencida", pct(k["pct_vencida"]), money_mm(k["cartera_vencida"]), danger=True)
    _card(slide, l + cw + Inches(0.12), t, cw, ch, "% Sin Olímpica", pct(k["pct_vencida_sin_olimpica"]), money_mm(k["vencida_sin_olimpica"]), danger=True)
    _card(slide, l + 2 * (cw + Inches(0.12)), t, cw, ch, "% Corriente", pct(k["pct_corriente"]), money_mm(k["cartera_corriente"]))

    # gráfico comparativo
    chart_top = t + ch + Inches(0.2)
    chart_h = h - ch - Inches(0.25)
    _add_chart(
        slide,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        l,
        chart_top,
        w,
        chart_h,
        ["Con Olímpica", "Sin Olímpica", "Solo Olímpica"],
        {
            "Total (MM)": [
                round(k["total_cartera"] / 1e6, 1),
                round(k["total_sin_olimpica"] / 1e6, 1),
                round(k["olimpica_saldo"] / 1e6, 1),
            ],
            "Vencida (MM)": [
                round(k["cartera_vencida"] / 1e6, 1),
                round(k["vencida_sin_olimpica"] / 1e6, 1),
                round(k["olimpica_vencido"] / 1e6, 1),
            ],
        },
        title="Total vs vencida (millones de pesos)",
    )


def _slide_sin_olimpica(slide, box, model):
    k = model["kpis"]
    kpi_box, body = _split_box(box)
    l, t, w, h = kpi_box
    cw = int(w / 3) - Inches(0.1)
    _card(slide, l, t, cw, h, "Total sin Olímpica", money_mm(k["total_sin_olimpica"]))
    _card(slide, l + cw + Inches(0.1), t, cw, h, "Vencida", money_mm(k["vencida_sin_olimpica"]), pct(k["pct_vencida_sin_olimpica"]), danger=True)
    _card(slide, l + 2 * (cw + Inches(0.1)), t, cw, h, "Olímpica (ref.)", money_mm(k["olimpica_saldo"]), f"Venc. {money_mm(k['olimpica_vencido'])}")

    bl, bt, bw, bh = body
    df = model["por_cliente_sin_olimpica"].head(10)
    # pie top 6 + otros
    top = df.head(6)
    cats = [str(r["nombre"])[:18] for _, r in top.iterrows()]
    vals = [round(float(r["saldo"]) / 1e6, 1) for _, r in top.iterrows()]
    resto = float(model["totales_sin_olimpica"].get("saldo", 0)) - sum(float(r["saldo"]) for _, r in top.iterrows())
    if resto > 0:
        cats.append("Otros")
        vals.append(round(resto / 1e6, 1))
    chart_w = int(bw * 0.42)
    _add_chart(slide, XL_CHART_TYPE.PIE, bl, bt, chart_w, bh, cats, {"Saldo MM": vals}, title="Participación top clientes")

    headers = ["Cliente", "Saldo", "Vencido", "%"]
    rows = []
    for _, r in df.iterrows():
        s = r["saldo"]
        v = r.get("vencido", 0)
        rows.append([str(r["nombre"])[:34], money_mm(s), money_mm(v), pct((v / s) if s else 0)])
    rows.append(["TOTAL", money_mm(k["total_sin_olimpica"]), money_mm(k["vencida_sin_olimpica"]), pct(k["pct_vencida_sin_olimpica"])])
    _add_table(slide, bl + chart_w + Inches(0.12), bt, bw - chart_w - Inches(0.12), bh, headers, rows)


def _slide_unidades(slide, box, model):
    k = model["kpis"]
    kpi_box, body = _split_box(box)
    l, t, w, h = kpi_box
    cw = int(w / 3) - Inches(0.1)
    _card(slide, l, t, cw, h, "Total", money_mm(k["total_cartera"]))
    _card(slide, l + cw + Inches(0.1), t, cw, h, "Vencida", money_mm(k["cartera_vencida"]), pct(k["pct_vencida"]), danger=True)
    _card(slide, l + 2 * (cw + Inches(0.1)), t, cw, h, "Unidades", str(len(model["por_unidad"])))

    bl, bt, bw, bh = body
    udf = model["por_unidad"]
    cats = [str(r["unidad"])[:14] for _, r in udf.iterrows()]
    saldos = [round(float(r["saldo"]) / 1e6, 1) for _, r in udf.iterrows()]
    vencs = [round(float(r.get("vencido", 0)) / 1e6, 1) for _, r in udf.iterrows()]
    chart_w = int(bw * 0.48)
    _add_chart(
        slide,
        XL_CHART_TYPE.BAR_CLUSTERED,
        bl,
        bt,
        chart_w,
        bh,
        cats,
        {"Saldo MM": saldos, "Vencida MM": vencs},
        title="Cartera por unidad (MM)",
    )
    headers = ["Unidad", "Saldo", "Vencido", "%", "Corriente"]
    rows = []
    for _, r in udf.iterrows():
        s = r["saldo"]
        v = r.get("vencido", 0)
        rows.append([r["unidad"], money_mm(s), money_mm(v), pct((v / s) if s else 0), money_mm(r.get("CORRIENTE", 0))])
    rows.append(["TOTAL", money_mm(k["total_cartera"]), money_mm(k["cartera_vencida"]), pct(k["pct_vencida"]), money_mm(k["cartera_corriente"])])
    _add_table(slide, bl + chart_w + Inches(0.12), bt, bw - chart_w - Inches(0.12), bh, headers, rows)


def _slide_consolidado(slide, box, model):
    k = model["kpis"]
    kpi_box, body = _split_box(box)
    l, t, w, h = kpi_box
    cw = int(w / 3) - Inches(0.1)
    top3 = model["por_cliente"].head(3)["saldo"].sum()
    _card(slide, l, t, cw, h, "Top 3 clientes", money_mm(top3), f"{(top3/k['total_cartera']*100) if k['total_cartera'] else 0:.1f}% del total")
    _card(slide, l + cw + Inches(0.1), t, cw, h, "Total cartera", money_mm(k["total_cartera"]))
    _card(slide, l + 2 * (cw + Inches(0.1)), t, cw, h, "N° clientes", str(k["n_clientes"]))

    bl, bt, bw, bh = body
    df = model["por_cliente"].head(12)
    cats = [str(r["nombre"])[:16] for _, r in df.head(8).iterrows()]
    vals = [round(float(r["saldo"]) / 1e6, 1) for _, r in df.head(8).iterrows()]
    chart_w = int(bw * 0.42)
    _add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, bl, bt, chart_w, bh, cats, {"Saldo MM": vals}, title="Top 8 clientes (MM)")

    headers = ["#", "Cliente", "Saldo", "Corriente", "Vencido", "%"]
    rows = []
    for i, (_, r) in enumerate(df.iterrows(), 1):
        s = r["saldo"]
        v = r.get("vencido", 0)
        rows.append([str(i), str(r["nombre"])[:32], money_mm(s), money_mm(r.get("CORRIENTE", 0)), money_mm(v), pct((v / s) if s else 0)])
    rows.append(["", "TOTAL", money_mm(k["total_cartera"]), money_mm(k["cartera_corriente"]), money_mm(k["cartera_vencida"]), pct(k["pct_vencida"])])
    _add_table(slide, bl + chart_w + Inches(0.1), bt, bw - chart_w - Inches(0.1), bh, headers, rows)


def _slide_unidad_detalle(slide, box, model, unit):
    df = model["detalle_por_unidad"].get(unit)
    l, t, w, h = box
    if df is None or df.empty:
        _textbox(slide, l, t, w, Inches(0.5), f"Sin datos SIIMED para {unit}", size=14, bold=True, color=MUTED)
        return

    total = float(df["saldo"].sum())
    venc = float(df["vencido"].sum()) if "vencido" in df else 0
    corr = float(df["CORRIENTE"].sum()) if "CORRIENTE" in df else max(total - venc, 0)
    kpi_h = Inches(1.05)
    cw = int(w / 3) - Inches(0.1)
    _card(slide, l, t, cw, kpi_h, unit, money_mm(total))
    _card(slide, l + cw + Inches(0.1), t, cw, kpi_h, "Corriente", money_mm(corr))
    _card(slide, l + 2 * (cw + Inches(0.1)), t, cw, kpi_h, "Vencida", money_mm(venc), pct((venc / total) if total else 0), danger=True)

    body_top = t + kpi_h + Inches(0.1)
    body_h = h - kpi_h - Inches(0.1)
    top = df.head(8)
    cats = [str(r["nombre"])[:16] for _, r in top.iterrows()]
    vals = [round(float(r["saldo"]) / 1e6, 1) for _, r in top.iterrows()]
    chart_w = int(w * 0.40)
    if cats:
        _add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, l, body_top, chart_w, body_h, cats, {"Saldo MM": vals}, title=f"Top clientes · {unit}")

    headers = ["#", "Cliente", "Saldo", "Corriente", "Vencido"]
    rows = []
    for i, (_, r) in enumerate(df.head(12).iterrows(), 1):
        rows.append([
            str(i),
            str(r["nombre"])[:34],
            money_mm(r["saldo"]),
            money_mm(r.get("CORRIENTE", 0)),
            money_mm(r.get("vencido", 0)),
        ])
    rows.append(["", "TOTAL", money_mm(total), money_mm(corr), money_mm(venc)])
    _add_table(slide, l + chart_w + Inches(0.1), body_top, w - chart_w - Inches(0.1), body_h, headers, rows)


def _slide_record(slide, box, model):
    hist = model.get("historico") or {}
    series = hist.get("record_series") or []
    l, t, w, h = box
    k = model["kpis"]

    _card(slide, l, t, int(w * 0.32), Inches(1.05), "Mes actual", hist.get("mes_actual", "—"), pct(k["pct_vencida"]), danger=True)
    _card(slide, l + int(w * 0.34), t, int(w * 0.32), Inches(1.05), "Vencida actual", money_mm(k["cartera_vencida"]))
    _card(slide, l + int(w * 0.68), t, int(w * 0.32), Inches(1.05), "Total actual", money_mm(k["total_cartera"]))

    body_top = t + Inches(1.2)
    body_h = h - Inches(1.25)

    if series:
        use = series[-12:]
        cats = [s["label"].replace(" 202", "/").replace(" ", "")[:8] for s in use]
        pcts = [round(float(s["pct_vencida"]) * 100, 1) for s in use]
        totals = [round(float(s["total_mm"]), 0) for s in use]
        chart_w = int(w * 0.55)
        _add_chart(
            slide,
            XL_CHART_TYPE.LINE_MARKERS,
            l,
            body_top,
            chart_w,
            body_h,
            cats,
            {"% Vencida": pcts},
            title="Evolución % cartera vencida",
        )
        # mini tabla últimos meses
        headers = ["Mes", "Vencida MM", "%", "Total MM"]
        rows = [[s["label"], f"{s['vencida_mm']:,.0f}", pct(s["pct_vencida"]), f"{s['total_mm']:,.0f}"] for s in use[-8:]]
        _add_table(slide, l + chart_w + Inches(0.1), body_top, w - chart_w - Inches(0.1), body_h, headers, rows)
    else:
        headers = ["Mes", "Vencida MM", "%", "Corriente MM", "Total MM"]
        rows = [[
            hist.get("mes_actual", "Actual"),
            f"{k['cartera_vencida']/1e6:,.0f}",
            pct(k["pct_vencida"]),
            f"{k['cartera_corriente']/1e6:,.0f}",
            f"{k['total_cartera']/1e6:,.0f}",
        ]]
        _add_table(slide, l, body_top, w, body_h, headers, rows)


def _slide_rotacion(slide, box, model):
    k = model["kpis"]
    hist = model.get("historico") or {}
    l, t, w, h = box
    cw = int(w / 3) - Inches(0.1)
    _card(slide, l, t, cw, Inches(1.05), "Rotación aprox.", f"{k['dias_rotacion']:.1f} días" if k.get("dias_rotacion") else "N/D")
    _card(slide, l + cw + Inches(0.1), t, cw, Inches(1.05), "Ingresos mes", money_mm(k["ingresos_mes"]) if k.get("ingresos_mes") else "N/D")
    _card(slide, l + 2 * (cw + Inches(0.1)), t, cw, Inches(1.05), "Total cartera", money_mm(k["total_cartera"]))

    body_top = t + Inches(1.2)
    body_h = h - Inches(1.25)
    edades_hist = hist.get("edades_hist") or []
    if edades_hist:
        cats = [e["mes"] for e in edades_hist]
        totals = [round(e["total_mm"], 0) for e in edades_hist]
        venc_pct = [round(e["pct"] * 100, 1) for e in edades_hist]
        _add_chart(
            slide,
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            l,
            body_top,
            int(w * 0.55),
            body_h,
            cats,
            {"Total MM": totals, "% Vencida": venc_pct},
            title="Histórico reciente cartera",
        )
        headers = ["Mes", "Total MM", "Vencida MM", "%"]
        rows = [[e["mes"], f"{e['total_mm']:,.0f}", f"{e['vencida_mm']:,.0f}", pct(e["pct"])] for e in edades_hist]
        _add_table(slide, l + int(w * 0.55) + Inches(0.1), body_top, int(w * 0.45) - Inches(0.1), body_h, headers, rows)
    else:
        _textbox(
            slide,
            l,
            body_top,
            w,
            Inches(0.8),
            "Nota: rotación aproximada = Cartera / (Ingresos del balance ÷ 30). Validar con contabilidad si se presenta a junta.",
            size=12,
            color=MUTED,
        )


def _try_insert_executive_summary(prs, model, corte_label):
    """
    Si hay layout 'Solo el título' o 'En blanco', agrega una slide de resumen
    ejecutivo al final de la sección Cartera (antes de GRACIAS si existe).
    No falla si no se puede insertar en posición exacta.
    """
    # python-pptx no inserta fácilmente en medio; agregamos al final antes de gracias
    # buscando slide GRACIAS y dejando el resumen como penúltima visualmente
    # Por simplicidad: añadimos una slide nueva al final con layout Solo título / En blanco
    layout = None
    for ly in prs.slide_layouts:
        if ly.name in ("Solo el título", "En blanco", "Título y objetos", "Diseño personalizado"):
            layout = ly
            break
    if layout is None:
        return

    # Evitar duplicar si ya generamos muchas veces: solo una vez por run
    for s in prs.slides:
        if "RESUMEN EJECUTIVO CARTERA" in _slide_title(s).upper():
            return

    slide = prs.slides.add_slide(layout)
    # título
    try:
        for sh in slide.shapes:
            if sh.has_text_frame and sh.is_placeholder and sh.placeholder_format.idx == 0:
                sh.text_frame.paragraphs[0].text = "RESUMEN EJECUTIVO CARTERA"
                break
    except Exception:
        _textbox(slide, Inches(0.5), Inches(0.2), Inches(12), Inches(0.5), "RESUMEN EJECUTIVO CARTERA", size=22, bold=True, color=GREEN)

    k = model["kpis"]
    _textbox(slide, Inches(0.5), Inches(0.85), Inches(12), Inches(0.35), f"Corte {corte_label} · Lectura rápida para junta", size=12, color=MUTED)

    _card(slide, Inches(0.5), Inches(1.4), Inches(3.0), Inches(1.15), "Total cartera", money_mm(k["total_cartera"]))
    _card(slide, Inches(3.7), Inches(1.4), Inches(3.0), Inches(1.15), "Vencida", money_mm(k["cartera_vencida"]), pct(k["pct_vencida"]), danger=True)
    _card(slide, Inches(6.9), Inches(1.4), Inches(3.0), Inches(1.15), "Sin Olímpica", money_mm(k["total_sin_olimpica"]), pct(k["pct_vencida_sin_olimpica"]))
    _card(slide, Inches(10.1), Inches(1.4), Inches(2.7), Inches(1.15), "Clientes", str(k["n_clientes"]))

    # gráfico unidades
    udf = model["por_unidad"]
    if not udf.empty:
        cats = [str(r["unidad"])[:12] for _, r in udf.iterrows()]
        vals = [round(float(r["saldo"]) / 1e6, 1) for _, r in udf.iterrows()]
        _add_chart(
            slide,
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(0.5),
            Inches(2.8),
            Inches(7.2),
            Inches(4.0),
            cats,
            {"Saldo MM": vals},
            title="Distribución por unidad de negocio",
        )

    # mensajes clave
    msgs = [
        f"• La cartera vencida representa {pct(k['pct_vencida'])} del total.",
        f"• Sin Olímpica, el indicador de vencida es {pct(k['pct_vencida_sin_olimpica'])}.",
        f"• Vencida > 15 días: {money_mm(k['vencida_mayor_15'])}.",
        "• Detalle por canal y clientes en las siguientes diapositivas.",
    ]
    y = Inches(2.9)
    for m in msgs:
        _textbox(slide, Inches(8.0), y, Inches(4.8), Inches(0.55), m, size=13, color=BLACK)
        y += Inches(0.7)
