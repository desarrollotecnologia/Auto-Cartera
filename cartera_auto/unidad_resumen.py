# -*- coding: utf-8 -*-
"""Actualiza bloque resumen + EDADES (filas 14-27) en hoja unidad."""

from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .client_master import UNIDAD_VENCIDA_EXTRA
from .historico import MONTH_NAMES

FONT = "Century Gothic"
MM_FORMAT = "#,##0,,"  # muestra cifras en millones
THIN = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
GREEN_FILL = PatternFill("solid", fgColor="548235")
RED_FILL = PatternFill("solid", fgColor="C00000")
WHITE = "FFFFFF"

def _font(size=11, bold=False, color=None):
    kw = {"name": FONT, "size": size, "bold": bold}
    if color:
        kw["color"] = color
    return Font(**kw)


def _style_header(cell, fill):
    cell.font = _font(11, True, WHITE)
    cell.fill = fill
    cell.alignment = Alignment(horizontal="center")
    cell.border = THIN


def _style_mm(cell, formula: str):
    cell.value = formula
    cell.number_format = MM_FORMAT
    cell.font = _font(11)
    cell.border = THIN
    cell.alignment = Alignment(horizontal="center")


# Difícil cobro se reparte por la unidad de origen del cliente (informe manual:
# julio RUEDA → SUBPRODUCTOS; UMBARILA y EL SEÑOR DE LAS CARNES → CARNES).
DIFICIL_COBRO_ORIGEN = {nit: unit.upper() for nit, unit in UNIDAD_VENCIDA_EXTRA.items()}
DIFICIL_COBRO_DEFAULT = "CARNES"
BUCKET_LETTERS = {
    "vencido": "E", "1 AL 7": "F", "8 AL 15": "G", "16 AL 30": "H", "31 AL 60": "I",
    "61 AL 90": "J", "91 AL 180": "K", "181 AL 360": "L", "361+": "M",
}


def _dificil_cobro_split(model: dict) -> dict[str, dict[str, float]]:
    """{campo: {'SUBPRODUCTOS': x, 'CARNES': y}} para la fila DIFICIL COBRO."""
    split = {k: {"SUBPRODUCTOS": 0.0, "CARNES": 0.0} for k in BUCKET_LETTERS}
    rows = (model.get("detalle_por_unidad") or {}).get("DIFICIL COBRO")
    if rows is None or getattr(rows, "empty", True):
        return split
    for _, row in rows.iterrows():
        nit = str(row.get("nit") or "").strip()
        dest = DIFICIL_COBRO_ORIGEN.get(nit, DIFICIL_COBRO_DEFAULT)
        for field in BUCKET_LETTERS:
            split[field][dest] += float(row.get(field) or 0)
    return split


def _with_dificil_cobro(base: str, letter: str, own: float, other: float) -> str:
    if abs(own) < 0.5:
        return base
    if abs(other) < 0.5:
        return f"{base}+{letter}11"
    return f"{base}+{round(own, 2)}"


def update_unidad_resumen(wb, model: dict, consolidada_name: str, month: int) -> None:
    """
    Replica bloque del informe manual (filas 14-27):
    resumen TOTAL/CORRIENTE/VENCIDA + EDADES con fórmulas y formato millones.
    """
    if "unidad" not in wb.sheetnames:
        return
    ws = wb["unidad"]

    mes_label = MONTH_NAMES.get(month, "MES")
    corte = model["meta"].get("corte")
    if corte:
        try:
            mes_label = MONTH_NAMES[int(corte[5:7])]
        except Exception:
            pass

    ws.cell(14, 2, "DESCRIPCION")
    _style_header(ws.cell(14, 2), GREEN_FILL)
    ws.cell(14, 3, mes_label)
    _style_header(ws.cell(14, 3), GREEN_FILL)
    for c, label in ((4, "SUBPRODUCTOS"), (5, "SERVICIOS "), (6, "CARNES ")):
        ws.cell(14, c, label)
        _style_header(ws.cell(14, c), RED_FILL)

    split = _dificil_cobro_split(model)

    def sub_f(field: str) -> str:
        letter = BUCKET_LETTERS[field]
        s = split[field]
        return _with_dificil_cobro(f"=+{letter}3", letter, s["SUBPRODUCTOS"], s["CARNES"])

    def car_f(field: str) -> str:
        letter = BUCKET_LETTERS[field]
        s = split[field]
        return _with_dificil_cobro(f"=+{letter}10", letter, s["CARNES"], s["SUBPRODUCTOS"])

    def srv_f(field: str) -> str:
        letter = BUCKET_LETTERS[field]
        return "=+" + "+".join(f"{letter}{r}" for r in range(4, 10))

    rows_resumen = [
        (15, "TOTAL CARTERA (MM)", "=+C12", "=+D16+D17", "=+E16+E17", "=+F16+F17"),
        (16, "CARTERA CORRIENTE", "=SUM(D16:F16)", "=+D3", "=+D4+D5+D6+D7+D8+D9", "=+D10"),
        (17, "CARTERA VENCIDA", "=+E12", sub_f("vencido"), srv_f("vencido"), car_f("vencido")),
    ]
    for r, desc, c_mes, c_sub, c_srv, c_car in rows_resumen:
        ws.cell(r, 2, desc).font = _font(11, True)
        ws.cell(r, 2).border = THIN
        _style_mm(ws.cell(r, 3), c_mes)
        _style_mm(ws.cell(r, 4), c_sub)
        _style_mm(ws.cell(r, 5), c_srv)
        _style_mm(ws.cell(r, 6), c_car)

    ws.cell(18, 2, "EDADES ")
    _style_header(ws.cell(18, 2), RED_FILL)
    ws.cell(18, 3, "TOTAL ")
    _style_header(ws.cell(18, 3), RED_FILL)
    for c in range(4, 7):
        ws.cell(18, c, None)
        _style_header(ws.cell(18, c), RED_FILL)

    edades = [
        (r, label, sub_f(field), srv_f(field), car_f(field))
        for r, label, field in (
            (19, "1-7 DIAS", "1 AL 7"),
            (20, "8 - 15 DIAS", "8 AL 15"),
            (21, "16 - 30 DIAS", "16 AL 30"),
            (22, "31 - 60 DIAS", "31 AL 60"),
            (23, "61 - 90 DIAS", "61 AL 90"),
            (24, "91 - 180 DIAS", "91 AL 180"),
            (25, "181 - 360 DIAS", "181 AL 360"),
            (26, "> 361 DIAS", "361+"),
        )
    ]
    for r, label, f_sub, f_srv, f_car in edades:
        ws.cell(r, 2, label).font = _font(11)
        ws.cell(r, 2).border = THIN
        _style_mm(ws.cell(r, 3), f"=SUM(D{r}:F{r})")
        _style_mm(ws.cell(r, 4), f_sub)
        _style_mm(ws.cell(r, 5), f_srv)
        _style_mm(ws.cell(r, 6), f_car)

    ws.cell(27, 2, "TOTAL VENCIDA").font = _font(11, True)
    ws.cell(27, 2).border = THIN
    _style_mm(ws.cell(27, 3), "=SUM(C19:C26)")
    _style_mm(ws.cell(27, 4), "=SUM(D21:D26)")
    _style_mm(ws.cell(27, 5), "=SUM(E21:E26)")
    _style_mm(ws.cell(27, 6), "=SUM(F21:F26)")
