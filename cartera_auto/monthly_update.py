# -*- coding: utf-8 -*-
"""Actualización mensual de hojas secundarias del Informe de Cartera."""

from __future__ import annotations

import re
from copy import copy
from datetime import datetime

from openpyxl.utils import get_column_letter

from .historico import MONTH_NAMES, MONTH_SHORT

# Formato contable del informe original Cartera (sin decimales)
ACCOUNTING_FMT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
COMPORTAMIENTO_FMT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
NIT_FMT = "##,##0_);[Red]\\(##,##0\\)"
MM_FORMAT = "#,##0,,"

MONTH_SHEET = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

# informe jf fila → celda unidad (columna C fila unidad)
INFORME_JF_UNIDAD = {
    7: 3,   # SUBPRODUCTOS
    8: 4,   # OTRAS C*C
    9: 5,   # FAENADO LOCAL
    10: 6,  # FAENADO REGIONAL
    11: 7,  # FAENADO NACIONAL
    12: 8,  # FAENADO INSTITUCIONAL
    13: 9,  # DESPOSTE
    14: None,  # GANADO EN PIE
    15: 10,  # CARNES COLBEEF
    16: 11,  # CUENTAS DIFICIL RECAUDO
}

FORMULA_RECAUDO = {
    "MAY": "BK22",
    "JUN": "BL22",
    "JUL": "BM22",
    "AGO": "BN22",
    "SEP": "BO22",
    "OCT": "BP22",
    "NOV": "BQ22",
    "DIC": "BR22",
}

# Base auditada del original agosto V1. Reemplaza vínculos externos BI:BM
# (marzo-julio) para que la rotación conserve exactamente el histórico manual.
FORMULA_HISTORY_2026 = {
    "BI": {
        4: 8842292732, 5: 1139338537.1, 6: 4040351034.77, 7: 0,
        8: 45068850, 12: 1618889017.09, 13: 739161661.97,
        14: 1735998831.18, 15: 0, 16: 48429726,
    },
    "BJ": {
        4: 11596635460, 5: 1388629880.69, 6: 5332728377.46, 7: 0,
        8: 58722850, 12: 1683447646.73, 13: 830803817.97,
        14: 1831329001.5, 15: 0, 16: 50281857.24,
    },
    "BK": {
        4: 14886980132, 5: 1696237301.86, 6: 6767740156.41, 7: 0,
        8: 73834050, 12: 1698754468.5, 13: 437602663.52,
        14: 1512889220.33, 15: 0, 16: 49749657.24,
    },
    "BL": {
        4: 18193141417, 5: 1881999507.84, 6: 8286160165.62, 7: 0,
        8: 88946050, 12: 1598963278.52, 13: 323908373.32,
        14: 1418011053.76, 15: 0, 16: 89578921.16,
    },
    "BM": {
        4: 21584493056, 5: 2053084812.85, 6: 9381298258.75, 7: 0,
        8: 103874350, 12: 1444968890.4, 13: 163520051.17,
        14: 1387476316.1, 15: 0, 16: 82733992.56,
    },
}


def norm_nit(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^\d]", "", s)


def apply_monthly_updates(wb, model: dict) -> str | None:
    """Renombra hojas y actualiza comportamiento / informe jf del mes."""
    corte = model["meta"].get("corte")
    if not corte:
        return None
    year = int(corte[:4])
    month = int(corte[5:7])

    consolidada = rename_month_sheets(wb, month, year)
    update_comportamiento(wb, model, consolidada, month, year)
    update_informe_jf(wb, month, year)
    update_formula_rotacion(wb, model, month)
    return consolidada


def update_formula_rotacion(wb, model: dict, month: int) -> None:
    """Agrega el mes al cálculo anual de rotación usando balance y cartera."""
    summary = model.get("rotacion_balance") or {}
    if "Formula" not in wb.sheetnames or not summary:
        return

    ws = wb["Formula"]
    for col_letter, values in FORMULA_HISTORY_2026.items():
        for row, value in values.items():
            ws[f"{col_letter}{row}"] = value
        ws[f"{col_letter}9"] = f"=SUM({col_letter}4:{col_letter}8)"
        ws[f"{col_letter}17"] = f"=SUM({col_letter}12:{col_letter}16)"

    col = 58 + month  # BG=ENE, BH=FEB, ..., BN=AGO
    prev_col = col - 1
    letter = get_column_letter(col)
    prior_year = get_column_letter(col - 12)
    avg_start = get_column_letter(col - 11)

    for r in range(1, 35):
        src, dst = ws.cell(r, prev_col), ws.cell(r, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format

    corte = model.get("meta", {}).get("corte")
    ws.cell(2, col, datetime.strptime(corte, "%Y-%m-%d")) if corte else None

    ingresos = summary["ingresos_ytd"]
    cartera = summary["cartera"]
    for row, key in ((4, "faenado"), (5, "desposte"), (6, "carnes"), (8, "fomento")):
        ws.cell(row, col, float(ingresos.get(key, 0.0)))
    ws.cell(7, col, 0)
    ws.cell(9, col, f"=SUM({letter}4:{letter}8)")

    for row, key in ((12, "faenado"), (13, "desposte"), (14, "carnes"), (16, "fomento")):
        ws.cell(row, col, float(cartera.get(key, 0.0)))
    ws.cell(15, col, 0)
    ws.cell(17, col, f"=SUM({letter}12:{letter}16)")

    ws.cell(19, col, f"=+$BF$9-{prior_year}9+{letter}9")
    ws.cell(20, col, f"=AVERAGE({avg_start}17:{letter}17)")
    ws.cell(21, col, f"=+({letter}19/{letter}20)")
    ws.cell(22, col, f"=+$AF$22/{letter}21")

    for sales_row, avg_row, result_row, source_row in (
        (24, 25, 26, 6),
        (28, 29, 30, 4),
        (32, 33, 34, 5),
    ):
        ws.cell(
            sales_row, col,
            f"=+$BF${source_row}-{prior_year}{source_row}+{letter}{source_row}",
        )
        portfolio_row = {6: 14, 4: 12, 5: 13}[source_row]
        ws.cell(
            avg_row, col,
            f"=AVERAGE({avg_start}{portfolio_row}:{letter}{portfolio_row})",
        )
        ws.cell(
            result_row, col,
            f"=(+{letter}{avg_row}/{letter}{sales_row})*$AF${result_row}",
        )


def rename_month_sheets(wb, month: int, year: int) -> str:
    """Renombra consolidada/comportamiento al mes del corte. Retorna nombre consolidada."""
    mes = MONTH_SHEET[month]
    consolidada_new = f"consolidada {mes}"

    for old in list(wb.sheetnames):
        low = old.lower()
        if low.startswith("consolidada "):
            wb[old].title = consolidada_new
        elif low.startswith("comportamiento "):
            wb[old].title = f"comportamiento {mes}"

    return consolidada_new


def _month_header_names() -> set[str]:
    return {name.upper() for name in MONTH_NAMES.values()}


def _find_comportamiento_sheet(wb):
    for name in wb.sheetnames:
        if name.lower().startswith("comportamiento"):
            return wb[name]
    return None


def update_comportamiento(wb, model: dict, consolidada_name: str, month: int, year: int) -> None:
    ws = _find_comportamiento_sheet(wb)
    if ws is None:
        return

    mes_upper = MONTH_NAMES[month]
    ws.cell(2, 2, f"COMPORTAMIENTO DE CARTERA {mes_upper} {year}")

    hist = _read_comportamiento_history(ws)
    # Mismo plazo que la consolidada (maestro); el histórico solo completa faltantes.
    plazos: dict = {}
    if consolidada_name in wb.sheetnames:
        cws = wb[consolidada_name]
        for r in range(8, 200):
            nit = norm_nit(cws.cell(r, 2).value)
            p = cws.cell(r, 4).value
            if nit and p is not None:
                plazos[nit] = p
    for nit, data in hist.items():
        if data.get("PLAZOS") is not None:
            plazos.setdefault(nit, data["PLAZOS"])

    month_cols, pct_col = _find_comportamiento_month_cols(ws)
    target = mes_upper.upper()
    month_col = _ensure_comportamiento_month_col(ws, month_cols, pct_col, target)
    month_cols, pct_col = _find_comportamiento_month_cols(ws)

    pct_col = pct_col or month_col + 1
    for c in range(month_col + 1, 40):
        if str(ws.cell(4, c).value or "").strip().upper() == "%":
            pct_col = c
            break

    plazo_col = pct_col + 1
    var_col = pct_col + 3
    arrow_col = pct_col + 2
    for c in range(pct_col, pct_col + 6):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h == "PLAZOS":
            plazo_col = c
        elif "VARIACI" in h:
            var_col = c

    df = model["por_cliente"].sort_values("saldo", ascending=False)
    nit_saldo = {norm_nit(row["nit"]): float(row["saldo"] or 0) for _, row in df.iterrows()}

    # Unir clientes activos + históricos (conservar filas con saldo 0 en el mes)
    all_nits: list[str] = []
    seen: set[str] = set()
    for _, row in df.iterrows():
        n = norm_nit(row["nit"])
        if n and n not in seen:
            all_nits.append(n)
            seen.add(n)
    for n in hist:
        if n and n not in seen:
            all_nits.append(n)
            seen.add(n)

    _clear_block_comportamiento(ws, 5, 200, 3, var_col)

    n_clients = len(model["por_cliente"])
    # Misma fila Grand Total que consolidada (fija en 82 del V1)
    total_row = max(82, 8 + n_clients)
    total_ref_row = 11
    # El original corre la referencia con cada mes: agosto L→R11, septiembre M→S11.
    total_ref_col = month_col + 6
    total_ref = f"${get_column_letter(total_ref_col)}${total_ref_row}"
    for c in range(var_col + 1, total_ref_col + 4):
        if c != total_ref_col:
            ws.cell(total_ref_row, c).value = None
    ws.cell(total_ref_row, total_ref_col, f"=+'{consolidada_name}'!E{total_row}")

    month_letter = get_column_letter(month_col)
    prev_col = _prev_month_col(ws, month_col)
    prev_letter = get_column_letter(prev_col)

    nit_to_name = {norm_nit(row["nit"]): row["nombre"] for _, row in df.iterrows()}

    # El original lista solo los clientes con saldo en el mes, con subtotal TOP 20 en la fila 25.
    all_nits = [n for n in all_nits if n in nit_saldo]
    top_n = 20
    subtotal_row = 5 + top_n if len(all_nits) > top_n else None

    def _row_for(idx: int) -> int:
        return 5 + idx + (1 if subtotal_row and idx >= top_n else 0)

    for i, nit_s in enumerate(all_nits):
        r = _row_for(i)
        ws.cell(r, 3, int(nit_s) if nit_s.isdigit() else nit_s)
        ws.cell(r, 3).number_format = NIT_FMT
        old = hist.get(nit_s, {})
        name = nit_to_name.get(nit_s) or old.get("NOMBRE", "")
        ws.cell(r, 4, name)
        ws.cell(r, 4).number_format = "General"

        for c, h in month_cols.items():
            if c == month_col:
                continue
            if h in old:
                ws.cell(r, c, old[h])
                ws.cell(r, c).number_format = COMPORTAMIENTO_FMT

        ws.cell(r, month_col, float(nit_saldo.get(nit_s, 0.0)))
        ws.cell(r, month_col).number_format = COMPORTAMIENTO_FMT

        plazo = plazos.get(nit_s, old.get("PLAZOS"))
        if plazo is not None:
            ws.cell(r, plazo_col, plazo)
            ws.cell(r, plazo_col).number_format = ACCOUNTING_FMT

        ws.cell(r, pct_col, f"={month_letter}{r}/{total_ref}")
        ws.cell(r, arrow_col, f'=IF(({month_letter}{r}-{prev_letter}{r})>0,"⬆\n","⬇\n")')
        ws.cell(r, var_col, f"={month_letter}{r}-{prev_letter}{r}")

    if subtotal_row:
        ws.cell(subtotal_row, 4, "TOTAL")
        for c in [*month_cols.keys(), month_col, pct_col]:
            letter = get_column_letter(c)
            cell = ws.cell(subtotal_row, c, f"=SUM({letter}5:{letter}{subtotal_row - 1})")
            cell.number_format = "0%" if c == pct_col else COMPORTAMIENTO_FMT

    # Estructura visual y filtros como el original.
    _restore_comportamiento_merges(ws, plazo_col, arrow_col, var_col)
    ws.cell(4, plazo_col, "PLAZOS")
    last_row = max(5, _row_for(len(all_nits) - 1) if all_nits else 5)
    ws.auto_filter.ref = f"B4:{get_column_letter(var_col)}{last_row}"
    ws.print_area = f"B2:{get_column_letter(var_col)}{last_row}"


def _clear_block_comportamiento(ws, start_row, end_row, start_col, end_col):
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            try:
                ws.cell(r, c).value = None
            except Exception:
                pass


def _find_comportamiento_month_cols(ws) -> tuple[dict[int, str], int | None]:
    month_cols: dict[int, str] = {}
    pct_col = None
    for c in range(3, 40):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h in _month_header_names():
            month_cols[c] = h
        elif h == "%":
            pct_col = c
            break
    return month_cols, pct_col


def _ensure_comportamiento_month_col(ws, month_cols, pct_col, target: str) -> int:
    month_col = next((c for c, h in month_cols.items() if h == target), None)
    if month_col is not None:
        return month_col
    if not month_cols:
        ws.cell(4, 5, target)
        return 5
    last_col = max(month_cols)
    month_col = last_col + 1
    if pct_col and month_col >= pct_col:
        _insert_comportamiento_month_col(ws, pct_col, last_col)
        month_col = pct_col
    ws.cell(4, month_col, target)
    return month_col


def _insert_comportamiento_month_col(ws, idx: int, style_source: int) -> None:
    """Inserta mes conservando anchos, estilos y merges del diseño original."""
    merges = list(ws.merged_cells.ranges)
    for merged in merges:
        ws.unmerge_cells(str(merged))

    ws.insert_cols(idx)

    # Copiar ancho y estilo del mes anterior, sin copiar sus valores.
    src_letter = get_column_letter(style_source)
    dst_letter = get_column_letter(idx)
    ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width
    for r in range(1, max(ws.max_row, 200) + 1):
        src = ws.cell(r, style_source)
        dst = ws.cell(r, idx)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format

    # Desplazar/expandir combinaciones que cruzaban la nueva columna.
    for merged in merges:
        min_col, min_row, max_col, max_row = (
            merged.min_col,
            merged.min_row,
            merged.max_col,
            merged.max_row,
        )
        if min_col >= idx:
            min_col += 1
            max_col += 1
        elif min_col < idx <= max_col:
            max_col += 1
        ws.merge_cells(
            start_row=min_row,
            start_column=min_col,
            end_row=max_row,
            end_column=max_col,
        )


def _restore_comportamiento_merges(
    ws, plazo_col: int, arrow_col: int, var_col: int
) -> None:
    """Restaura B2:VAR y PLAZOS sobre sus dos columnas, como el original."""
    for merged in list(ws.merged_cells.ranges):
        # Título fila 2 o encabezado PLAZOS fila 4.
        if (
            merged.min_row <= 2 <= merged.max_row
            and merged.min_col <= 2 <= merged.max_col
        ) or (
            merged.min_row <= 4 <= merged.max_row
            and merged.max_col >= plazo_col - 1
            and merged.min_col <= arrow_col
        ):
            ws.unmerge_cells(str(merged))

    ws.merge_cells(
        start_row=2,
        start_column=2,
        end_row=2,
        end_column=var_col,
    )
    ws.merge_cells(
        start_row=4,
        start_column=plazo_col,
        end_row=4,
        end_column=arrow_col,
    )


def _prev_month_col(ws, month_col: int) -> int:
    prev_col = month_col - 1
    while prev_col >= 3:
        if str(ws.cell(4, prev_col).value or "").strip().upper() in _month_header_names():
            return prev_col
        prev_col -= 1
    return max(3, month_col - 1)


def _read_comportamiento_history(ws) -> dict[str, dict]:
    month_cols: dict[int, str] = {}
    plazo_col = None
    past_pct = False
    for c in range(3, 40):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h == "%":
            past_pct = True
        elif h == "PLAZOS":
            plazo_col = c
            break
        elif h in _month_header_names() and not past_pct:
            month_cols[c] = h
    out: dict[str, dict] = {}
    empty_run = 0
    for r in range(5, 200):
        nit = norm_nit(ws.cell(r, 3).value)
        if not nit:
            # La fila de subtotal TOP 20 no tiene NIT; cortar solo tras varias vacías
            empty_run += 1
            if empty_run >= 3:
                break
            continue
        empty_run = 0
        entry: dict = {}
        nombre = ws.cell(r, 4).value
        if nombre:
            entry["NOMBRE"] = nombre
        for c, h in month_cols.items():
            v = ws.cell(r, c).value
            if v is not None:
                entry[h] = v
        if plazo_col:
            p = ws.cell(r, plazo_col).value
            if p is not None:
                entry["PLAZOS"] = p
        out[nit] = entry
    return out


def update_informe_jf(wb, month: int, year: int) -> None:
    if "informe jf" not in wb.sheetnames:
        return
    ws = wb["informe jf"]
    short = MONTH_SHORT[month]

    # Bloque calendario reciente: columnas >= 18 (ENE… del año en curso)
    block_cols: list[tuple[int, str]] = []
    for c in range(18, 40):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h in MONTH_SHORT.values() or h in (
            "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"
        ):
            block_cols.append((c, h))
        elif block_cols and not h:
            break

    col = next((c for c, h in block_cols if h == short), None)
    if col is None:
        if block_cols:
            col = block_cols[-1][0] + 1
        else:
            col = 18
        ws.cell(4, col, short)
    else:
        ws.cell(4, col, short)

    for row, unidad_row in INFORME_JF_UNIDAD.items():
        if unidad_row is None:
            cell = ws.cell(row, col, 0)
        else:
            # Réplica original: vínculo al mes corriente
            cell = ws.cell(row, col, f"=+unidad!C{unidad_row}")
        cell.number_format = MM_FORMAT

    # TOTAL del mes (fila 17)
    from openpyxl.utils import get_column_letter as _gcl

    letter = _gcl(col)
    ws.cell(17, col, f"=SUM({letter}6:{letter}16)")
    ws.cell(17, col).number_format = MM_FORMAT

    # Quitar meses posteriores al corte (evita MAY/JUN basura y VARIACION=0)
    _clear_future_informe_jf_months(ws, col)
    _fix_informe_jf_variacion(ws, col)
    # Merge encabezado mes (filas 4-5) como el original
    try:
        merge_ref = f"{letter}4:{letter}5"
        already = any(str(m) == merge_ref for m in ws.merged_cells.ranges)
        if not already:
            ws.merge_cells(merge_ref)
    except Exception:
        pass
    _retarget_sheet_charts(ws, col, keep_cols={"D", "E", "AF"})
    _restore_informe_jf_history_chart(ws, col)
    _show_only_year_informe_jf(ws, col, year)


def _show_only_year_informe_jf(ws, month_col: int, year: int) -> None:
    """Oculta columnas de años anteriores; deja solo el año del corte."""
    # El bloque reciente suele empezar cerca de la col 18 (ENE del año en curso).
    # Detectar el ENE más cercano a la izquierda del mes actual.
    year_start = month_col
    for c in range(month_col, 4, -1):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h in ("ENE", "ENERO"):
            year_start = c
            break
        if h.startswith("VARIACION") or h in ("% PART.", "PART."):
            break

    keep = set(range(year_start, month_col + 1))
    for c in range(25, 40):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h.startswith("VARIACION") or h in ("% PART.", "PART.", "%PART.", "% PART", "VAR"):
            keep.add(c)
            if h.startswith("VARIACION"):
                keep.add(c + 1)

    for c in range(5, 40):
        letter = get_column_letter(c)
        ws.column_dimensions[letter].hidden = c not in keep

    # Desocultar columnas bajo las gráficas (si no, "no salen")
    for ch in getattr(ws, "_charts", []) or []:
        anchor = getattr(ch, "anchor", None)
        if anchor is None or not hasattr(anchor, "_from") or not hasattr(anchor, "to"):
            continue
        for c in range(anchor._from.col + 1, anchor.to.col + 2):
            letter = get_column_letter(c)
            ws.column_dimensions[letter].hidden = False
            if float(ws.column_dimensions[letter].width or 0) < 3:
                ws.column_dimensions[letter].width = 4


def _restore_informe_jf_history_chart(ws, month_col: int | None = None) -> None:
    """Restaura la gráfica histórica limitada al año visible."""
    charts = list(getattr(ws, "_charts", []) or [])
    if len(charts) < 2:
        return
    chart = charts[1]
    # Año visible: desde el ENE previo al mes hasta el mes actual.
    start_col = 18
    end_col = month_col or 25
    if month_col:
        for c in range(month_col, 4, -1):
            h = str(ws.cell(4, c).value or "").strip().upper()
            if h in ("ENE", "ENERO"):
                start_col = c
                break
    start_letter = get_column_letter(start_col)
    end_letter = get_column_letter(end_col)
    for index, series in enumerate(chart.series):
        row = 6 + index
        if series.val is not None and getattr(series.val, "numRef", None):
            series.val.numRef.f = f"'informe jf'!${start_letter}${row}:${end_letter}${row}"
        if series.cat is not None and getattr(series.cat, "strRef", None):
            series.cat.strRef.f = f"'informe jf'!${start_letter}$4:${end_letter}$4"


def _clear_future_informe_jf_months(ws, col: int) -> None:
    """Borra encabezados y datos de meses a la derecha del mes actual (antes de VARIACION)."""
    for c in range(col + 1, 30):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h.startswith("VARIACION") or h in ("VAR", "% PART.", "PART.", "%PART."):
            break
        ws.cell(4, c).value = None
        for r in range(6, 18):
            ws.cell(r, c).value = None


def _fix_informe_jf_variacion(ws, col: int) -> None:
    """VARIACION $ / % y % PART. respecto al mes anterior (como informe manual)."""
    from openpyxl.utils import get_column_letter

    cur = get_column_letter(col)
    prev = get_column_letter(col - 1)
    var_col = pct_col = part_col = None
    for c in range(25, 40):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h.startswith("VARIACION"):
            var_col = c
        elif h in ("% PART.", "PART.", "%PART.", "% PART"):
            part_col = c
    if var_col:
        pct_col = var_col + 1
    if not var_col:
        return

    total_row = 17
    for r in range(6, 18):
        label = str(ws.cell(r, 4).value or "").strip().upper()
        if not label and r != total_row:
            continue
        if r == total_row:
            # TOTAL: VARIACION = mes-ant; % PART = suma de participaciones (V1)
            ws.cell(r, var_col, f"={cur}{r}-{prev}{r}")
            if pct_col:
                ws.cell(r, pct_col, f"=+IFERROR({get_column_letter(var_col)}{r}/{prev}{r},0)")
                ws.cell(r, pct_col).number_format = "0%"
            if part_col:
                letter_p = get_column_letter(part_col)
                ws.cell(r, part_col, f"=SUM({letter_p}6:{letter_p}16)")
                ws.cell(r, part_col).number_format = "0%"
            continue
        ws.cell(r, var_col, f"={cur}{r}-{prev}{r}")
        if pct_col:
            ws.cell(r, pct_col, f"=+IFERROR({get_column_letter(var_col)}{r}/{prev}{r},0)")
            ws.cell(r, pct_col).number_format = "0%"
        if part_col:
            ws.cell(r, part_col, f"={cur}{r}/${cur}${total_row}")
            ws.cell(r, part_col).number_format = "0%"


def _retarget_sheet_charts(ws, col: int, keep_cols: set[str] | None = None) -> None:
    """
    Ajusta gráficos al mes del corte:
    - Serie de 1 columna ($W$7:$W$13) → columna del mes actual
    - Serie multi-mes ($E$7:$AC$7) → recorta hasta la columna del mes
    - También reescribe títulos de serie ($W$4:$W$5 → mes actual)
    """
    import re

    letter = get_column_letter(col)
    keep = keep_cols or {"B", "D"}

    def _rewire(formula: str | None) -> str | None:
        if not formula or "$" not in formula:
            return formula

        def repl_range(m: re.Match) -> str:
            c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
            if c1 in keep and c2 not in keep:
                return f"${c1}${r1}:${letter}${r2}"
            if c1 not in keep and c2 not in keep:
                if c1 == c2:
                    return f"${letter}${r1}:${letter}${r2}"
                return f"${c1}${r1}:${letter}${r2}"
            return m.group(0)

        out = re.sub(
            r"\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)",
            repl_range,
            formula,
        )

        def repl_cell(m: re.Match) -> str:
            c = m.group(1)
            if c in keep:
                return m.group(0)
            return f"${letter}${m.group(2)}"

        out = re.sub(r"\$([A-Z]{1,3})\$(\d+)", repl_cell, out)
        return out

    def _rewire_title(s) -> None:
        title = getattr(s, "title", None)
        if title is None:
            return
        str_ref = getattr(title, "strRef", None)
        if str_ref is not None and getattr(str_ref, "f", None):
            str_ref.f = _rewire(str_ref.f)
            return
        # Algunos charts usan tx.strRef
        tx = getattr(s, "tx", None)
        if tx is not None:
            str_ref = getattr(tx, "strRef", None)
            if str_ref is not None and getattr(str_ref, "f", None):
                str_ref.f = _rewire(str_ref.f)

    for ch in getattr(ws, "_charts", []) or []:
        for s in ch.series:
            if s.val is not None and getattr(s.val, "numRef", None) is not None:
                s.val.numRef.f = _rewire(s.val.numRef.f)
            if s.cat is not None and getattr(s.cat, "strRef", None) is not None:
                s.cat.strRef.f = _rewire(s.cat.strRef.f)
            _rewire_title(s)


def update_periodo_recaudo_formula(wb, month: int) -> None:
    """
    Restaura fórmulas de rotación (=+Formula!B*22) en PERIODO DE RECAUDO.
    Orden como el original: … JUN | JUL | PROM (mes antes de PROM).
    Busca el mes solo en el bloque del año en curso (tras último PROM 20XX).
    """
    import re as _re

    name = "PERIODO DE  RECAUDO"
    if name not in wb.sheetnames or "Formula" not in wb.sheetnames:
        return

    ws = wb[name]
    short = MONTH_SHORT[month]

    def _hdr(c: int) -> str:
        return str(ws.cell(4, c).value or "").strip().upper()

    def _swap_cols(c1: int, c2: int, max_row: int = 14) -> None:
        for r in range(1, max_row + 1):
            a, b = ws.cell(r, c1), ws.cell(r, c2)
            a.value, b.value = b.value, a.value
            fa, fb = a.number_format, b.number_format
            a.number_format, b.number_format = fb, fa

    # Bloque año en curso = después del último "PROM 2024/2025/…"
    block_start = 5
    for c in range(5, 50):
        h = _hdr(c)
        if _re.match(r"PROM\.?\s*20\d{2}", h) or (
            h.startswith("PROM ") and _re.search(r"20\d{2}", h)
        ):
            block_start = c + 1

    # Restaurar fórmulas solo en meses del año en curso; los años previos son valores fijos.
    for c in range(block_start, 50):
        h = _hdr(c)
        ref = FORMULA_RECAUDO.get(h)
        if not ref:
            continue
        cur = ws.cell(10, c).value
        if not (isinstance(cur, str) and "Formula!" in cur):
            ws.cell(10, c, f"=+Formula!{ref}")

    month_col = next((c for c in range(block_start, 50) if _hdr(c) == short), None)
    prom_col = next((c for c in range(block_start, 50) if _hdr(c) == "PROM"), None)

    if month_col is not None and prom_col is not None and month_col == prom_col + 1:
        _swap_cols(prom_col, month_col)
        month_col = prom_col
    elif month_col is not None and prom_col is not None and month_col > prom_col:
        ws.insert_cols(prom_col)
        src = month_col + 1
        for r in range(1, 15):
            ws.cell(r, prom_col).value = ws.cell(r, src).value
            ws.cell(r, prom_col).number_format = ws.cell(r, src).number_format
        ws.delete_cols(src)
        month_col = prom_col
    elif month_col is None:
        insert_at = prom_col if prom_col else block_start + 6
        if prom_col:
            ws.insert_cols(prom_col)
            insert_at = prom_col
        ws.cell(4, insert_at, short)
        ref = FORMULA_RECAUDO.get(short)
        if ref:
            ws.cell(10, insert_at, f"=+Formula!{ref}")
        month_col = insert_at
    else:
        ref = FORMULA_RECAUDO.get(short)
        if ref:
            ws.cell(10, month_col, f"=+Formula!{ref}")
        ws.cell(4, month_col, short)

    if month_col:
        formula_col = re.sub(r"\d", "", FORMULA_RECAUDO.get(short, ""))
        if formula_col:
            for row, formula_row in ((5, 30), (6, 34), (7, 26), (10, 22)):
                ws.cell(row, month_col, f"=+Formula!{formula_col}{formula_row}")
            ws.cell(8, month_col, 0)
            ws.cell(9, month_col, 0)

        prom_col = next(
            (c for c in range(month_col + 1, 50) if _hdr(c) == "PROM"),
            None,
        )
        if prom_col:
            month_letter = get_column_letter(month_col)
            year_proms = [
                c for c in range(5, block_start)
                if _re.match(r"PROM\.?\s*20\d{2}", _hdr(c))
            ]
            for row in range(5, 11):
                # V1: AVERAGE(PROM año-2, PROM año-1 : mes actual)
                if len(year_proms) >= 2:
                    older = get_column_letter(year_proms[-2])
                    last = get_column_letter(year_proms[-1])
                    formula = f"=AVERAGE({older}{row},{last}{row}:{month_letter}{row})"
                else:
                    formula = f"=AVERAGE({get_column_letter(block_start)}{row}:{month_letter}{row})"
                ws.cell(row, prom_col, formula)
        _retarget_sheet_charts(ws, month_col, keep_cols={"B", "D", get_column_letter(block_start)})
        _show_only_year_periodo(ws, block_start, month_col, prom_col)


def _show_only_year_periodo(ws, block_start: int, month_col: int, prom_col: int | None) -> None:
    """Oculta años anteriores en PERIODO DE RECAUDO y apunta gráficas a 2026."""
    keep = set(range(block_start, month_col + 1))
    if prom_col:
        keep.add(prom_col)
    keep.update({2, 3, 4})  # etiquetas
    for c in range(5, 45):
        letter = get_column_letter(c)
        ws.column_dimensions[letter].hidden = c not in keep

    start_l = get_column_letter(block_start)
    end_l = get_column_letter(month_col)
    charts = list(getattr(ws, "_charts", []) or [])
    if charts:
        # Área de evolución mensual del año
        ch0 = charts[0]
        for s in ch0.series:
            if s.val is not None and getattr(s.val, "numRef", None):
                # Mantener la fila de la serie (normalmente 10)
                old = s.val.numRef.f or ""
                import re as _re
                m = _re.search(r"\$(\d+)", old)
                row = m.group(1) if m else "10"
                s.val.numRef.f = f"'PERIODO DE  RECAUDO'!${start_l}${row}:${end_l}${row}"
            if s.cat is not None and getattr(s.cat, "strRef", None):
                s.cat.strRef.f = f"'PERIODO DE  RECAUDO'!${start_l}$4:${end_l}$4"
        # Reubicar ancla junto al bloque visible
        anchor = getattr(ch0, "anchor", None)
        if anchor is not None and hasattr(anchor, "_from"):
            anchor._from.col = max(block_start - 1, 2)
            anchor._from.row = 11
            anchor.to.col = (prom_col or month_col) + 6
            anchor.to.row = 30
            for c in range(anchor._from.col + 1, anchor.to.col + 2):
                letter = get_column_letter(c)
                ws.column_dimensions[letter].hidden = False
                if float(ws.column_dimensions[letter].width or 0) < 3:
                    ws.column_dimensions[letter].width = 4

    if len(charts) > 1:
        ch1 = charts[1]
        anchor = getattr(ch1, "anchor", None)
        if anchor is not None and hasattr(anchor, "_from"):
            # Barras del mes a la derecha del PROM
            anchor._from.col = (prom_col or month_col)
            anchor._from.row = 0
            anchor.to.col = (prom_col or month_col) + 6
            anchor.to.row = 16
            for c in range(anchor._from.col + 1, anchor.to.col + 2):
                letter = get_column_letter(c)
                ws.column_dimensions[letter].hidden = False
                if float(ws.column_dimensions[letter].width or 0) < 3:
                    ws.column_dimensions[letter].width = 4


def read_plazos_from_comportamiento(wb) -> dict[str, int | float | str]:
    """Lee PLAZOS por NIT desde hoja comportamiento (col PLAZOS)."""
    ws = _find_comportamiento_sheet(wb)
    if ws is None:
        return {}
    plazo_col = None
    for c in range(3, 25):
        h = str(ws.cell(4, c).value or "").strip().upper()
        if h == "PLAZOS":
            plazo_col = c
            break
    if not plazo_col:
        return {}
    out: dict[str, int | float | str] = {}
    for r in range(5, 200):
        nit = norm_nit(ws.cell(r, 3).value)
        plazo = ws.cell(r, plazo_col).value
        if nit and plazo is not None and str(plazo).strip() != "":
            out[nit] = plazo
    return out
