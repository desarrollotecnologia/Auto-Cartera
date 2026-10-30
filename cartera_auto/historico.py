# -*- coding: utf-8 -*-
"""Actualiza la 'base histórica' del Informe de Cartera (mes a mes)."""

from __future__ import annotations

import re
from copy import copy

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from .config import AGE_BUCKETS

MONTH_NAMES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}
MONTH_SHORT = {
    1: "ENE",
    2: "FEB",
    3: "MAR",
    4: "ABR",
    5: "MAY",
    6: "JUN",
    7: "JUL",
    8: "AGO",
    9: "SEP",
    10: "OCT",
    11: "NOV",
    12: "DIC",
}

# Filas en 'inf cartera x edades' (con Olímpica)
ROWS_CON = {
    "total": 4,
    "corriente": 5,
    "vencida": 6,
    "1 AL 7": 8,
    "8 AL 15": 9,
    "16 AL 30": 10,
    "31 AL 60": 11,
    "61 AL 90": 12,
    "91 AL 180": 13,
    "181 AL 360": 14,
    "361+": 15,
}
# Bloque sin Olímpica del original (encabezado fila 25, datos 26:38)
ROWS_SIN = {
    "total": 26,
    "corriente": 27,
    "vencida": 28,
    "1 AL 7": 30,
    "8 AL 15": 31,
    "16 AL 30": 32,
    "31 AL 60": 33,
    "61 AL 90": 34,
    "91 AL 180": 35,
    "181 AL 360": 36,
    "361+": 37,
}

FONT = "Century Gothic"
GREEN_FILL = PatternFill("solid", fgColor=Color(theme=4))
RED_FILL = PatternFill("solid", fgColor=Color(theme=5))
WHITE_FONT = Font(name=FONT, size=12, bold=True, color=Color(theme=0))
DATA_FONT = Font(name=FONT, size=12)
MED = Border(
    left=Side(style="medium", color="000000"),
    right=Side(style="medium", color="000000"),
    top=Side(style="medium", color="000000"),
    bottom=Side(style="medium", color="000000"),
)
THIN = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
CENTER = Alignment(horizontal="center")


def update_historico(wb, model: dict) -> dict:
    """
    Escribe el mes del corte en las hojas históricas.
    Retorna serie histórica para el PPT.
    """
    corte = model["meta"].get("corte")  # YYYY-MM-DD
    if not corte:
        return {"updated": False, "reason": "sin corte"}

    year = int(corte[:4])
    month = int(corte[5:7])

    edades_ws = wb["inf cartera x edades"]
    _ensure_edades_layout(edades_ws)
    col = _ensure_month_column(edades_ws, year, month)
    _write_edades_column(edades_ws, col, model)
    _repair_edades_history(edades_ws, col)
    _fix_diff_formulas(edades_ws, col)
    _show_only_year_columns(edades_ws, year, col)
    _apply_edades_column_widths(edades_ws, col)
    _apply_edades_design(edades_ws, col, year)
    _reposition_edades_chart(edades_ws, col)

    _update_record_cvda(wb["Record cvda"], year, month, model, col)
    _hide_prior_record_years(wb["Record cvda"], year)
    _fix_record_cvda_chart(wb["Record cvda"])

    from .monthly_update import update_periodo_recaudo_formula

    update_periodo_recaudo_formula(wb, month)

    series = _extract_series_for_ppt(wb, model, year, month, col)
    series["updated"] = True
    series["column"] = get_column_letter(col)
    series["year"] = year
    series["month"] = month
    return series


def _ensure_edades_layout(ws) -> None:
    """Migra plantillas antiguas al layout exacto del original de agosto."""
    if str(ws.cell(24, 2).value or "").strip().upper() == "CARTERA SIN OLIMPICA":
        ws.insert_rows(24, 1)
    # Si el bloque sin Olímpica quedó en filas viejas (25/26/27), bajar datos a 26/27/28.
    if (
        str(ws.cell(25, 2).value or "").strip().upper().startswith("TOTAL CARTERA")
        and str(ws.cell(26, 2).value or "").strip().upper() == "CORRIENTE"
    ):
        for c in range(3, 40):
            for src, dst in ((27, 28), (26, 27), (25, 26)):
                ws.cell(dst, c).value = ws.cell(src, c).value
                ws.cell(dst, c).number_format = ws.cell(src, c).number_format
            ws.cell(25, c).value = ws.cell(3, c).value
        for src, dst in ((36, 38), (35, 37), (34, 36), (33, 35), (32, 34), (31, 33), (30, 32), (29, 31), (28, 30)):
            for c in range(3, 40):
                ws.cell(dst, c).value = ws.cell(src, c).value
                ws.cell(dst, c).number_format = ws.cell(src, c).number_format

    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == merged.max_row and merged.min_row in (7, 29):
            ws.unmerge_cells(str(merged))
        elif merged.min_row == merged.max_row and 24 <= merged.min_row <= 30:
            ws.unmerge_cells(str(merged))
    ws.merge_cells("B7:AC7")
    ws.merge_cells("B29:AC29")
    labels = {
        18: "VENCIDA >15 DIAS",
        19: "TOTAL CARTERA",
        20: "CARTERA OLIMPICA",
        21: "CARTERA VENCIDA CON OLIMPICA",
        22: "CARTERA VENCIDA SIN OLIMPICA",
        25: "CARTERA SIN OLIMPICA",
        26: "TOTAL CARTERA (MM)",
        27: "CORRIENTE",
        28: "VENCIDA",
        29: "EDADES",
        30: "1-7",
        31: "8-15",
        32: "16 - 30 ",
        33: "31 - 60 ",
        34: "61 - 90 ",
        35: "91 - 180 ",
        36: "181 - 360 ",
        37: "> 361 ",
        38: "TOTAL VENCIDA",
    }
    for row, label in labels.items():
        ws.cell(row, 2, label)


def _normalize_month(name: str | None) -> str:
    if not name:
        return ""
    return str(name).strip().upper().replace(" ", "")


def _ensure_month_column(ws, year: int, month: int) -> int:
    """
    Localiza (o crea) la columna del mes/año en inf cartera x edades.
    Layout conocido: ... DIC(2024) + 2025 + DIC(2025) + 2026 ENE..
    """
    target = MONTH_NAMES[month]
    header_row = 3
    # Mapear columnas existentes con año estimado
    cols = []  # (col_idx, month_num, year_est)
    # Empezamos asumiendo que la primera DICIEMBRE es dic del (primer año-1)
    # Heurística: contar ENERO para incrementar año; primera col C=3 es DIC => año base
    # Mejor: leer hacia atrás desde la derecha el último mes con dato y proyectar.

    last_data_col = None
    for c in range(3, 40):
        h = _normalize_month(ws.cell(header_row, c).value)
        if h in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART"):
            break
        if not h:
            # hueco: candidato a nueva columna
            if last_data_col is not None:
                # si el mes pedido es el siguiente al último, usar este hueco
                break
            continue
        last_data_col = c
        cols.append((c, h))

    # Estimar años recorriendo desde el final
    # Última columna con header de mes: asumir que es el mes más reciente de la plantilla.
    # Plantilla junio 2026: última = JUNIO en col 21.
    # Construimos timeline andando de izquierda a derecha:
    timeline = []
    # Detectar año de la última columna por el valor del corte esperado: usamos
    # conteo de ENERO: al ver ENERO después de DICIEMBRE, año++.
    # Año inicial: year_target - n_eneros_hasta_target... más simple:
    # Buscar de derecha a izquierda la secuencia de meses y asignar year al último = year
    # si el último mes_num <= month, else year-1 etc.

    parsed = []
    for c, h in cols:
        mnum = None
        for num, name in MONTH_NAMES.items():
            if _normalize_month(name) == h:
                mnum = num
                break
        if mnum:
            parsed.append((c, mnum, h))

    if not parsed:
        # crear en col 3
        ws.cell(header_row, 3, target)
        return 3

    # Asignar años: el ÚLTIMO mes de la timeline corresponde al mes más reciente
    # de la plantilla. Anclar el último al year/month del corte si coincide el nombre,
    # si no, anclar el último y proyectar.
    last_c, last_m, _ = parsed[-1]
    # Construir años hacia atrás desde last
    years = [0] * len(parsed)
    # Si el último mes == month del corte, year_last = year; si last_m < month, year_last = year
    # si last_m > month, year_last = year - 1 (plantilla adelantada?)
    if last_m == month:
        y_last = year
    elif last_m < month:
        # plantilla todavía en un mes anterior del mismo año
        y_last = year
    else:
        # último mes en plantilla es posterior al corte (raro) o estamos en año nuevo
        y_last = year - 1 if month == 12 else year

    # Recalcular: anclar buscando coincidencia exacta month+aproximación
    # Más robusto: caminar desde la derecha asignando y_last al último,
    # decrementando año cuando pasamos de ENE a DIC hacia la izquierda.
    years[-1] = y_last
    for i in range(len(parsed) - 2, -1, -1):
        m_cur = parsed[i][1]
        m_next = parsed[i + 1][1]
        y = years[i + 1]
        # si vamos hacia atrás y el mes actual es >= mes siguiente, cruzamos año
        # Ej: ... DIC(15), ENE(16) hacia atrás desde ENE: DIC es año-1
        if m_cur >= m_next:
            y = y - 1
        years[i] = y

    # Buscar match
    for i, (c, mnum, h) in enumerate(parsed):
        if mnum == month and years[i] == year:
            return c

    # No existe: crear columna nueva a la derecha del último mes
    new_col = last_c + 1
    # Si new_col choca con DIF, insertar columna
    dif_header = _normalize_month(ws.cell(header_row, new_col).value)
    if dif_header in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART") or dif_header == "":
        if dif_header in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART"):
            ws.insert_cols(new_col)
        ws.cell(header_row, new_col, target)
        # copiar formato del header anterior
        src = ws.cell(header_row, last_c)
        dst = ws.cell(header_row, new_col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.alignment = copy(src.alignment)
        # también header del bloque sin olímpica fila 25
        ws.cell(25, new_col, target)
        src25 = ws.cell(25, last_c)
        dst25 = ws.cell(25, new_col)
        if src25.has_style:
            dst25.font = copy(src25.font)
            dst25.fill = copy(src25.fill)
        # fórmulas SUM de total vencida
        for r in (16,):
            prev = get_column_letter(last_c)
            letter = get_column_letter(new_col)
            ws.cell(r, new_col, f"=SUM({letter}8:{letter}15)")
        # y filas 19-21 indicadores si existen en prev
        for r in range(19, 22):
            f = ws.cell(r, last_c).value
            if isinstance(f, str) and f.startswith("="):
                ws.cell(r, new_col, f.replace(get_column_letter(last_c), get_column_letter(new_col)))
    else:
        ws.cell(header_row, new_col, target)
        ws.cell(25, new_col, target)

    return new_col


# Vínculos a hoja unidad (como informe manual Cartera)
_UNIDAD_LINKS_CON = {
    "total": "C15",
    "corriente": "C16",
    "vencida": "C17",
    "1 AL 7": "C19",
    "8 AL 15": "C20",
    "16 AL 30": "C21",
    "31 AL 60": "C22",
    "61 AL 90": "C23",
    "91 AL 180": "C24",
    "181 AL 360": "C25",
    "361+": "C26",
}
MM_FORMAT = "#,##0,,"

_UNIDAD_LINK_RE = re.compile(r"^=\+?unidad!([A-Za-z]+\d+)\s*$", re.I)


def _num(v, default: float = 0.0) -> float:
    try:
        if v is None:
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def _values_con_from_model(model: dict) -> dict[str, float]:
    """Valores bloque con Olímpica (pesos exactos, como el informe original)."""
    tot = model.get("totales") or {}
    k = model.get("kpis") or {}
    return {
        "total": _num(tot.get("saldo"), _num(k.get("total_cartera"))),
        "corriente": _num(tot.get("CORRIENTE"), _num(k.get("cartera_corriente"))),
        "vencida": _num(tot.get("vencido"), _num(k.get("cartera_vencida"))),
        "1 AL 7": _num(tot.get("1 AL 7")),
        "8 AL 15": _num(tot.get("8 AL 15")),
        "16 AL 30": _num(tot.get("16 AL 30")),
        "31 AL 60": _num(tot.get("31 AL 60")),
        "61 AL 90": _num(tot.get("61 AL 90")),
        "91 AL 180": _num(tot.get("91 AL 180")),
        "181 AL 360": _num(tot.get("181 AL 360")),
        "361+": _num(tot.get("361+")),
    }


def _resolve_unidad_link(uws, ref: str) -> float:
    """
    Resuelve unidad!ref con la misma lógica del original.
    C15/C16/C17 del resumen dependen de consolidada; usar tabla superior (C3:C11).
    """
    ref = ref.upper()
    if ref == "C15":
        # TOTAL CARTERA = Grand Total fila 12, o suma C3:C11
        gt = uws.cell(12, 3).value
        if isinstance(gt, (int, float)):
            return float(gt)
        return sum(_num(uws.cell(r, 3).value) for r in range(3, 12))
    if ref == "C16":
        return sum(_num(uws.cell(r, 4).value) for r in range(3, 12))
    if ref == "C17":
        return sum(_num(uws.cell(r, 5).value) for r in range(3, 12))
    # Edades C19..C26 → cols F..M de la tabla superior (filas 3-11)
    age_col = {
        "C19": 6, "C20": 7, "C21": 8, "C22": 9,
        "C23": 10, "C24": 11, "C25": 12, "C26": 13,
    }.get(ref)
    if age_col:
        return sum(_num(uws.cell(r, age_col).value) for r in range(3, 12))
    # Filas de cuenta C3..C11 (informe jf)
    m = re.fullmatch(r"C(\d+)", ref)
    if m:
        return _num(uws.cell(int(m.group(1)), 3).value)
    col_letter, row = coordinate_from_string(ref)
    return _resolve_cell(uws, row, column_index_from_string(col_letter))


def _eval_sheet_formula(ws, formula: str, depth: int = 0) -> float:
    """Evalúa fórmulas simples de la hoja unidad (SUM, +, refs) para congelar histórico."""
    if depth > 25:
        return 0.0
    f = (formula or "").strip()
    if not f.startswith("="):
        return _num(f)
    expr = f[1:].lstrip("+").strip()

    m = re.fullmatch(r"SUM\(([A-Za-z]+\d+):([A-Za-z]+\d+)\)", expr, re.I)
    if m:
        c1, r1 = coordinate_from_string(m.group(1).upper())
        c2, r2 = coordinate_from_string(m.group(2).upper())
        i1, i2 = column_index_from_string(c1), column_index_from_string(c2)
        total = 0.0
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(i1, i2), max(i1, i2) + 1):
                total += _resolve_cell(ws, r, c, depth + 1)
        return total

    # A1+B2+C3…
    if re.fullmatch(r"[A-Za-z]+\d+(?:\s*\+\s*[A-Za-z]+\d+)+", expr):
        total = 0.0
        for tok in re.findall(r"[A-Za-z]+\d+", expr):
            col_letter, row = coordinate_from_string(tok.upper())
            total += _resolve_cell(ws, row, column_index_from_string(col_letter), depth + 1)
        return total

    m = re.fullmatch(r"([A-Za-z]+\d+)", expr)
    if m:
        col_letter, row = coordinate_from_string(m.group(1).upper())
        return _resolve_cell(ws, row, column_index_from_string(col_letter), depth + 1)

    return 0.0


def _resolve_cell(ws, row: int, col: int, depth: int = 0) -> float:
    v = ws.cell(row, col).value
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and v.startswith("="):
        return _eval_sheet_formula(ws, v, depth)
    return 0.0


def freeze_unidad_linked_historico(wb) -> int:
    """
    Convierte =+unidad!Xn en valores numéricos usando la hoja unidad actual.
    Llamar ANTES de sobrescribir 'unidad' con el mes nuevo, para que Abr/May/Jun
    queden fijos como en el informe manual (solo el mes corriente usaba fórmulas).
    """
    if "unidad" not in wb.sheetnames:
        return 0
    uws = wb["unidad"]
    frozen = 0

    targets: list[tuple] = []
    if "inf cartera x edades" in wb.sheetnames:
        targets.append((wb["inf cartera x edades"], range(4, 40), range(3, 35)))
    if "informe jf" in wb.sheetnames:
        targets.append((wb["informe jf"], range(6, 18), range(5, 30)))
    # Record cvda: congelar meses previos con =+unidad! antes de pisar el mes nuevo
    if "Record cvda" in wb.sheetnames:
        targets.append((wb["Record cvda"], range(5, 90), range(3, 8)))

    for ws, rows, cols in targets:
        for r in rows:
            for c in cols:
                cell = ws.cell(r, c)
                v = cell.value
                if not isinstance(v, str) or "unidad!" not in v.lower():
                    continue
                m = _UNIDAD_LINK_RE.match(v.strip())
                if not m:
                    continue
                ref = m.group(1).upper()
                val = _resolve_unidad_link(uws, ref)
                cell.value = float(val)
                if not cell.number_format or cell.number_format == "General":
                    cell.number_format = MM_FORMAT
                frozen += 1
    return frozen


def freeze_historico_with_model(wb, model: dict) -> int:
    """
    Congela =+unidad! del mes corriente a pesos exactos del model (para plantilla del mes siguiente).
    El Excel entregado conserva las fórmulas; la plantilla guarda valores como el original histórico.
    """
    values = _values_con_from_model(model)
    frozen = 0
    if "inf cartera x edades" in wb.sheetnames:
        ws = wb["inf cartera x edades"]
        for c in range(3, 35):
            for key, row in ROWS_CON.items():
                cell = ws.cell(row, c)
                v = cell.value
                if isinstance(v, str) and "unidad!" in v.lower():
                    cell.value = float(values.get(key, 0) or 0)
                    cell.number_format = MM_FORMAT
                    frozen += 1
    if "informe jf" in wb.sheetnames and "unidad" in wb.sheetnames:
        from .monthly_update import INFORME_JF_UNIDAD

        ws = wb["informe jf"]
        uws = wb["unidad"]
        for c in range(5, 30):
            for row, unidad_row in INFORME_JF_UNIDAD.items():
                cell = ws.cell(row, c)
                v = cell.value
                if not isinstance(v, str) or "unidad!" not in v.lower():
                    continue
                if unidad_row is None:
                    cell.value = 0
                else:
                    cell.value = _num(uws.cell(unidad_row, 3).value)
                cell.number_format = MM_FORMAT
                frozen += 1
    return frozen


def _write_edades_column(ws, col: int, model: dict):
    """
    Réplica del original: mes corriente con =+unidad!C15…C26 y formato #,##0,,.
    Bloque sin Olímpica: pesos exactos (float), mismo formato.
    """
    sin = model["totales_sin_olimpica"]
    k = model["kpis"]

    for key, row in ROWS_CON.items():
        ref = _UNIDAD_LINKS_CON.get(key)
        if ref:
            cell = ws.cell(row, col, f"=+unidad!{ref}")
        else:
            cell = ws.cell(row, col, float(_values_con_from_model(model).get(key, 0) or 0))
        cell.number_format = MM_FORMAT
        cell.font = DATA_FONT
        cell.alignment = CENTER

    ws.cell(18, col, ws.cell(3, col).value)

    values_sin = {
        "total": k["total_sin_olimpica"],
        "corriente": float(sin.get("CORRIENTE", 0) or 0),
        "vencida": k["vencida_sin_olimpica"],
        "1 AL 7": sin.get("1 AL 7", 0),
        "8 AL 15": sin.get("8 AL 15", 0),
        "16 AL 30": sin.get("16 AL 30", 0),
        "31 AL 60": sin.get("31 AL 60", 0),
        "61 AL 90": sin.get("61 AL 90", 0),
        "91 AL 180": sin.get("91 AL 180", 0),
        "181 AL 360": sin.get("181 AL 360", 0),
        "361+": sin.get("361+", 0),
    }
    if not values_sin["corriente"] and values_sin["total"]:
        values_sin["corriente"] = values_sin["total"] - values_sin["vencida"]

    for key, row in ROWS_SIN.items():
        cell = ws.cell(row, col, float(values_sin[key] or 0))
        cell.number_format = MM_FORMAT
        cell.font = DATA_FONT
        cell.alignment = CENTER

    letter = get_column_letter(col)
    ws.cell(16, col, f"=SUM({letter}8:{letter}15)")
    ws.cell(16, col).number_format = MM_FORMAT
    ws.cell(38, col, f"=SUM({letter}30:{letter}37)")
    ws.cell(38, col).number_format = MM_FORMAT

    # Indicadores en el mismo orden del informe manual.
    ws.cell(19, col, f"=+{letter}4")
    ws.cell(19, col).number_format = MM_FORMAT
    ws.cell(20, col, float(model["kpis"].get("olimpica_saldo", 0) or 0))
    ws.cell(20, col).number_format = MM_FORMAT
    ws.cell(21, col, f"=SUM({letter}10:{letter}15)/{letter}4")
    ws.cell(21, col).number_format = "0.0%"
    ws.cell(22, col, f"=SUM({letter}32:{letter}37)/{letter}26")
    ws.cell(22, col).number_format = "0.0%"
    ws.cell(25, col, ws.cell(3, col).value)

    # Limpiar meses futuros a la derecha (evita JUNIO/JULIO basura en corte MAYO)
    _clear_future_month_cols(ws, col)
    # Reapuntar gráfico CARTERA POR EDADES a la columna del mes
    from .monthly_update import _retarget_sheet_charts

    _retarget_sheet_charts(ws, col, keep_cols={"B"})


def _iter_month_cols(ws, upto_col: int | None = None) -> list[int]:
    cols = []
    for c in range(3, 45):
        h = _normalize_month(ws.cell(3, c).value)
        if h in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART"):
            break
        if not h:
            continue
        if any(_normalize_month(name) == h for name in MONTH_NAMES.values()):
            if upto_col is None or c <= upto_col:
                cols.append(c)
    return cols


def _repair_edades_history(ws, current_col: int) -> None:
    """Repara VENCIDA >15, VENCIDA sin Olímpica y totales desalinhados."""
    for c in _iter_month_cols(ws, current_col):
        letter = get_column_letter(c)
        month_name = ws.cell(3, c).value
        ws.cell(18, c, month_name)
        ws.cell(25, c, month_name)

        ws.cell(19, c, f"=+{letter}4")
        ws.cell(19, c).number_format = MM_FORMAT

        if c != current_col:
            ol = ws.cell(20, c).value
            fmt = str(ws.cell(20, c).number_format or "")
            if isinstance(ol, str) or "%" in fmt:
                tot = ws.cell(4, c).value
                sin = ws.cell(26, c).value
                if isinstance(tot, (int, float)) and isinstance(sin, (int, float)):
                    ws.cell(20, c, float(tot) - float(sin))
                    ws.cell(20, c).number_format = MM_FORMAT
            elif isinstance(ol, (int, float)):
                ws.cell(20, c).number_format = MM_FORMAT

        ws.cell(21, c, f"=SUM({letter}10:{letter}15)/{letter}4")
        ws.cell(21, c).number_format = "0.0%"
        ws.cell(22, c, f"=SUM({letter}32:{letter}37)/{letter}26")
        ws.cell(22, c).number_format = "0.0%"

        venc = ws.cell(28, c).value
        tot_sin = ws.cell(26, c).value
        corr_sin = ws.cell(27, c).value
        if venc in (None, "") and isinstance(tot_sin, (int, float)) and isinstance(
            corr_sin, (int, float)
        ):
            ws.cell(28, c, float(tot_sin) - float(corr_sin))
            ws.cell(28, c).number_format = MM_FORMAT
        elif isinstance(venc, (int, float)):
            ws.cell(28, c).number_format = MM_FORMAT

        ws.cell(16, c, f"=SUM({letter}8:{letter}15)")
        ws.cell(16, c).number_format = MM_FORMAT
        ws.cell(38, c, f"=SUM({letter}30:{letter}37)")
        ws.cell(38, c).number_format = MM_FORMAT


def _place_dif_var_part(ws, month_col: int) -> tuple[int, int, int]:
    """DIF/VAR/PART fijas en AA:AC mientras quepan (agosto W→AA, septiembre X→AA)."""
    target_dif = max(27, month_col + 3)
    target_var = target_dif + 1
    target_part = target_dif + 2

    def _safe_clear(row: int, col: int) -> None:
        try:
            ws.cell(row, col).value = None
        except AttributeError:
            # MergedCell de solo lectura: deshacer merge que la cubre
            for rng in list(ws.merged_cells.ranges):
                if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                    try:
                        ws.unmerge_cells(str(rng))
                    except Exception:
                        pass
            try:
                ws.cell(row, col).value = None
            except Exception:
                pass

    for c in range(month_col + 1, 45):
        h = _normalize_month(ws.cell(3, c).value)
        if h in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART"):
            if c not in (target_dif, target_var, target_part):
                _safe_clear(3, c)
                _safe_clear(25, c)
                for r in range(4, 39):
                    _safe_clear(r, c)

    for gap_col in range(month_col + 1, target_dif):
        _safe_clear(3, gap_col)
        _safe_clear(25, gap_col)

    ws.cell(3, target_dif, "DIF.")
    ws.cell(3, target_var, "VAR.")
    ws.cell(3, target_part, "PART.")
    ws.cell(25, target_dif, "DIF.")
    ws.cell(25, target_var, "VAR.")
    ws.cell(25, target_part, "PART.")
    return target_dif, target_var, target_part


def _apply_edades_design(ws, month_col: int, year: int | None = None) -> None:
    """Aplica el diseño visual del original (verdes/rojos, fuentes, bordes)."""
    if year is None:
        month_cols = [c for c in _iter_month_cols(ws, month_col) if not ws.column_dimensions[get_column_letter(c)].hidden]
    else:
        month_num = None
        for num, name in MONTH_NAMES.items():
            if _normalize_month(ws.cell(3, month_col).value) == _normalize_month(name):
                month_num = num
                break
        timeline = _month_timeline(ws, year, month_num or 1)
        month_cols = [c for c, _m, y in timeline if y == year and c <= month_col]
    dif_c, var_c, part_c = _find_dif_var_part_cols(ws)
    value_cols = month_cols + [c for c in (dif_c, var_c, part_c) if c]

    def paint(row: int, fill, font=WHITE_FONT, border=MED, cols=None):
        targets = cols if cols is not None else [2, *value_cols]
        for c in targets:
            cell = ws.cell(row, c)
            cell.fill = fill
            cell.font = font
            cell.border = border
            if c != 2:
                cell.alignment = CENTER

    # Encabezados verdes
    paint(3, GREEN_FILL)
    paint(18, GREEN_FILL, cols=[2, *month_cols])
    paint(25, GREEN_FILL)
    paint(16, GREEN_FILL)
    paint(38, GREEN_FILL)

    # Filas rojas: EDADES + TOTAL/OLÍMPICA del bloque medio
    paint(7, RED_FILL, cols=[2])
    paint(29, RED_FILL, cols=[2])
    paint(19, RED_FILL, cols=[2, *month_cols])
    paint(20, RED_FILL, cols=[2, *month_cols])

    # Datos del mes corriente y porcentajes con tipografía del original
    for r in [4, 5, 6, *range(8, 16), 21, 22, 26, 27, 28, *range(30, 38)]:
        for c in value_cols:
            cell = ws.cell(r, c)
            if r in (21, 22):
                cell.font = Font(name="Verdana", size=10)
            else:
                cell.font = DATA_FONT
            cell.alignment = CENTER
            if r in (19, 20, 21, 22):
                cell.border = MED
            else:
                cell.border = THIN

    # Etiquetas del bloque medio
    for r, bold in ((21, False), (22, False)):
        cell = ws.cell(r, 2)
        cell.font = Font(name="Verdana", size=10, bold=bold)
        cell.border = MED


def _month_timeline(ws, year: int, month: int) -> list[tuple[int, int, int]]:
    """[(col, month_num, year_est), ...] anclado al mes/año del corte."""
    parsed: list[tuple[int, int]] = []
    for c in range(3, 45):
        h = _normalize_month(ws.cell(3, c).value)
        if h in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART"):
            break
        mnum = None
        for num, name in MONTH_NAMES.items():
            if _normalize_month(name) == h:
                mnum = num
                break
        if mnum:
            parsed.append((c, mnum))
    if not parsed:
        return []

    last_c, last_m = parsed[-1]
    if last_m == month or last_m < month:
        y_last = year
    else:
        y_last = year - 1 if month == 12 else year

    years = [0] * len(parsed)
    years[-1] = y_last
    for i in range(len(parsed) - 2, -1, -1):
        m_cur = parsed[i][1]
        m_next = parsed[i + 1][1]
        y = years[i + 1]
        if m_cur >= m_next:
            y -= 1
        years[i] = y
    return [(c, mnum, years[i]) for i, (c, mnum) in enumerate(parsed)]


def _show_only_year_columns(ws, year: int, month_col: int) -> None:
    """
    Deja visibles solo columnas del año del corte (+ DIF/VAR/PART).
    Oculta años anteriores y meses futuros sin uso.
    """
    month_num = None
    for num, name in MONTH_NAMES.items():
        if _normalize_month(ws.cell(3, month_col).value) == _normalize_month(name):
            month_num = num
            break
    if month_num is None:
        return

    timeline = _month_timeline(ws, year, month_num)
    keep = {c for c, _m, y in timeline if y == year and c <= month_col}
    dif_c, var_c, part_c = _find_dif_var_part_cols(ws)
    keep.update(c for c in (dif_c, var_c, part_c, (dif_c or month_col + 2) - 1) if c)

    for c in range(3, 45):
        letter = get_column_letter(c)
        dim = ws.column_dimensions[letter]
        if c in keep:
            dim.hidden = False
        else:
            dim.hidden = True

    # En VENCIDA >15 solo etiquetar meses del año visible.
    for c, _m, y in timeline:
        if y == year and c <= month_col:
            ws.cell(18, c, ws.cell(3, c).value)
        else:
            ws.cell(18, c).value = None


def _reposition_edades_chart(ws, month_col: int) -> None:
    """Coloca la gráfica junto al bloque visible de 2026 (a la derecha de PART.)."""
    charts = list(getattr(ws, "_charts", []) or [])
    if not charts:
        return
    dif_c, var_c, part_c = _find_dif_var_part_cols(ws)
    # openpyxl usa índice 0-based. PART es 1-based; la gráfica inicia en la col siguiente.
    part_idx = (part_c or (month_col + 4)) - 1  # 0-based index of PART
    from_col = part_idx + 1  # primera col libre a la derecha de PART
    to_col = from_col + 8
    for ch in charts:
        anchor = getattr(ch, "anchor", None)
        if anchor is None or not hasattr(anchor, "_from") or not hasattr(anchor, "to"):
            continue
        anchor._from.col = from_col
        anchor._from.row = 1
        anchor._from.colOff = 120000
        anchor._from.rowOff = 80000
        anchor.to.col = to_col
        anchor.to.row = 17
        anchor.to.colOff = 0
        anchor.to.rowOff = 0

    # Crítico: esas columnas no pueden quedar hidden o la gráfica "no sale".
    for c in range(from_col + 1, to_col + 2):  # 1-based
        letter = get_column_letter(c)
        ws.column_dimensions[letter].hidden = False
        width = float(ws.column_dimensions[letter].width or 0)
        if width < 4:
            ws.column_dimensions[letter].width = 4.5


def _apply_edades_column_widths(ws, month_col: int) -> None:
    """Anchos suficientes para que no aparezcan ##### ni columnas en 0."""
    ws.column_dimensions["B"].width = 31.375
    ws.column_dimensions["B"].hidden = False
    # Solo ensanchar columnas visibles del año actual.
    for c in range(3, month_col + 1):
        letter = get_column_letter(c)
        dim = ws.column_dimensions[letter]
        if dim.hidden:
            continue
        dim.hidden = False
        width = float(dim.width or 0)
        dim.width = 11.0 if c == month_col else max(width, 11.0)
    dif_c, var_c, part_c = _find_dif_var_part_cols(ws)
    gap = get_column_letter((dif_c or month_col + 2) - 1)
    ws.column_dimensions[gap].width = max(float(ws.column_dimensions[gap].width or 0), 3.0)
    ws.column_dimensions[gap].hidden = False
    for c, w in ((dif_c, 8.0), (var_c, 9.5), (part_c, 8.0)):
        if not c:
            continue
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = w
        ws.column_dimensions[letter].hidden = False


def _clear_future_month_cols(ws, col: int) -> None:
    """Borra columnas de mes posteriores al mes escrito hasta antes de DIF/VAR/PART."""
    for c in range(col + 1, 40):
        h = _normalize_month(ws.cell(3, c).value)
        if h in ("DIF.", "DIF", "VAR.", "VAR", "PART.", "PART"):
            break
        for r in list(ROWS_CON.values()) + list(ROWS_SIN.values()) + [
            16, 18, 19, 20, 21, 22, 38
        ]:
            ws.cell(r, c).value = None
        ws.cell(3, c).value = None
        ws.cell(25, c).value = None


def _find_dif_var_part_cols(ws) -> tuple[int | None, int | None, int | None]:
    dif = var = part = None
    for c in range(3, 45):
        h = _normalize_month(ws.cell(3, c).value)
        if h in ("DIF.", "DIF"):
            dif = c
        elif h in ("VAR.", "VAR"):
            var = c
        elif h in ("PART.", "PART"):
            part = c
    return dif, var, part


def _fix_diff_formulas(ws, col: int):
    """DIF / VAR / PART respecto al mes anterior (columnas detectadas por encabezado)."""
    if col <= 3:
        return
    dif_c, var_c, part_c = _place_dif_var_part(ws, col)
    prev = get_column_letter(col - 1)
    cur = get_column_letter(col)
    dif_l = get_column_letter(dif_c)
    rows = [
        4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16,
        26, 27, 28, 30, 31, 32, 33, 34, 35, 36, 37, 38,
    ]
    for r in rows:
        ws.cell(r, dif_c, f"={cur}{r}-{prev}{r}")
        ws.cell(r, dif_c).number_format = MM_FORMAT
        if var_c:
            ws.cell(r, var_c, f"=IFERROR({dif_l}{r}/{prev}{r},0)")
            ws.cell(r, var_c).number_format = "0%"
    if not part_c:
        return
    for r in (4, 5, 6, *range(8, 17)):
        ws.cell(r, part_c, f"={cur}{r}/${cur}$4")
        ws.cell(r, part_c).number_format = "0%"
    for r in (26, 27, 28, *range(30, 39)):
        ws.cell(r, part_c, f"={cur}{r}/${cur}$26")
        ws.cell(r, part_c).number_format = "0%"


def _update_record_cvda(ws, year: int, month: int, model: dict, edades_col: int):
    """
    Actualiza la fila del mes en Record cvda (bloque con Olímpica y sin Olímpica).
    Réplica original: vencida =+unidad!C17; corriente/total a inf cartera; % con formato 0%.
    """
    k = model["kpis"]
    pct_v = k["pct_vencida"]
    pct_c = k["pct_corriente"]

    # Reparar meses previos del año que aún tengan =+unidad! (plantilla sucia)
    _repair_record_prior_months(ws, year, month, edades_col)

    # Bloque principal: filas 5-36 (años apilados)
    row = _find_record_row(ws, year, month, start=5, end=40)
    if row:
        letter = get_column_letter(edades_col)
        ws.cell(row, 3, "=+unidad!C17")  # vencida (como original)
        ws.cell(row, 3).number_format = MM_FORMAT
        ws.cell(row, 4, pct_v)
        ws.cell(row, 4).number_format = "0%"
        ws.cell(row, 5, f"='inf cartera x edades'!{letter}5")  # corriente
        ws.cell(row, 5).number_format = MM_FORMAT
        ws.cell(row, 6, pct_c)
        ws.cell(row, 6).number_format = "0%"
        ws.cell(row, 7, f"='inf cartera x edades'!{letter}4")  # total
        ws.cell(row, 7).number_format = MM_FORMAT
        ws.cell(row, 2, MONTH_SHORT[month])

    # Sin Olímpica bloque ~72-77
    # El bloque 2026 sin Olímpica llega hasta ~85; no limitarlo a JUN (fila 78).
    row_sin = _find_record_row(ws, year, month, start=60, end=90)
    if row_sin:
        letter = get_column_letter(edades_col)
        ws.cell(row_sin, 3, f"='inf cartera x edades'!{letter}28")
        ws.cell(row_sin, 3).number_format = MM_FORMAT
        ws.cell(row_sin, 4, k["pct_vencida_sin_olimpica"])
        ws.cell(row_sin, 4).number_format = "0%"
        ws.cell(row_sin, 5, f"='inf cartera x edades'!{letter}27")
        ws.cell(row_sin, 5).number_format = MM_FORMAT
        ws.cell(row_sin, 6, 1 - k["pct_vencida_sin_olimpica"])
        ws.cell(row_sin, 6).number_format = "0%"
        ws.cell(row_sin, 7, f"='inf cartera x edades'!{letter}26")
        ws.cell(row_sin, 7).number_format = MM_FORMAT

    normalize_record_row_totals(ws)


def _repair_record_prior_months(ws, year: int, month: int, edades_col: int) -> None:
    """Fija meses anteriores del año con los valores ya congelados en edades."""
    wb = ws.parent
    if wb is None or "inf cartera x edades" not in wb.sheetnames:
        return
    edades = wb["inf cartera x edades"]
    month_to_col: dict[str, int] = {}
    # Timeline anclada al mes actual para no mezclar AGO 2025 con AGO 2026
    timeline = _month_timeline(edades, year, month)
    for c, mnum, y in timeline:
        if y == year and c < edades_col:
            month_to_col[MONTH_SHORT[mnum]] = c

    for start, end, rows_edades in (
        (5, 42, {3: 6, 5: 5, 7: 4}),       # con Olímpica: venc/corr/total
        (60, 90, {3: 28, 5: 27, 7: 26}),   # sin Olímpica
    ):
        targets: dict[int, int] = {}
        for mnum in range(1, month):
            short = MONTH_SHORT[mnum]
            if short not in month_to_col:
                continue
            r = _find_record_row(ws, year, mnum, start=start, end=end)
            if r:
                targets[r] = month_to_col[short]
        for r, col_e in targets.items():
            for c, erow in rows_edades.items():
                src = edades.cell(erow, col_e).value
                if isinstance(src, (int, float)):
                    ws.cell(r, c, float(src))
                    ws.cell(r, c).number_format = MM_FORMAT
            # % desde total si hay números
            venc = ws.cell(r, 3).value
            total = ws.cell(r, 7).value
            if isinstance(venc, (int, float)) and isinstance(total, (int, float)) and total:
                ws.cell(r, 4, float(venc) / float(total))
                ws.cell(r, 4).number_format = "0%"
                corr = ws.cell(r, 5).value
                if isinstance(corr, (int, float)):
                    ws.cell(r, 6, float(corr) / float(total))
                    ws.cell(r, 6).number_format = "0%"


def _fix_record_cvda_chart(ws) -> None:
    """Gráfico Record cvda solo con meses del año en curso."""
    def _current_year_span(start: int) -> tuple[int, int]:
        """Último tramo ENE..mes de la sección."""
        blocks = _record_year_blocks(ws, start)
        if not blocks:
            return (start, start)
        _, r0, r1 = blocks[-1]
        return (r0, r1)

    header = _record_sin_olimpica_header(ws)
    main = _current_year_span(5)
    sin = _current_year_span(header + 1)
    main_val = f"'Record cvda'!$D${main[0]}:$D${main[1]}"
    main_cat = f"'Record cvda'!$B${main[0]}:$B${main[1]}"
    sin_val = f"'Record cvda'!$D${sin[0]}:$D${sin[1]}"

    for ch in getattr(ws, "_charts", []) or []:
        series = list(ch.series or [])
        if not series:
            continue
        if len(series) >= 2:
            s0, s1 = series[0], series[1]
            if s0.val is not None and getattr(s0.val, "numRef", None):
                s0.val.numRef.f = sin_val
            if s0.cat is not None and getattr(s0.cat, "strRef", None):
                s0.cat.strRef.f = main_cat
            if s1.val is not None and getattr(s1.val, "numRef", None):
                s1.val.numRef.f = main_val
            if s1.cat is not None and getattr(s1.cat, "strRef", None):
                s1.cat.strRef.f = main_cat
        else:
            s = series[0]
            if s.val is not None and getattr(s.val, "numRef", None):
                s.val.numRef.f = main_val
            if s.cat is not None and getattr(s.cat, "strRef", None):
                s.cat.strRef.f = main_cat


def _hide_prior_record_years(ws, year: int) -> None:
    """Oculta filas de años anteriores; deja solo el año del corte."""
    header = _record_sin_olimpica_header(ws)

    def _block_start(start: int, fallback: int) -> int:
        rows = [r0 for y, r0, _ in _record_year_blocks(ws, start) if y == year]
        return rows[0] if rows else fallback

    con_start = _block_start(5, 31)
    sin_start = _block_start(header + 1, header + 29)
    for r in range(4, con_start):
        ws.row_dimensions[r].hidden = True
    for r in range(con_start, header + 1):
        ws.row_dimensions[r].hidden = False
    for r in range(header + 1, sin_start):
        ws.row_dimensions[r].hidden = True
    for r in range(sin_start, sin_start + 15):
        ws.row_dimensions[r].hidden = False
    ws.row_dimensions[2].hidden = False
    ws.row_dimensions[3].hidden = False
    _reposition_record_chart(ws)


def _reposition_record_chart(ws) -> None:
    """Mantiene la gráfica visible sobre el bloque 2026 (sin filas ocultas)."""
    for ch in getattr(ws, "_charts", []) or []:
        anchor = getattr(ch, "anchor", None)
        if anchor is None or not hasattr(anchor, "_from") or not hasattr(anchor, "to"):
            continue
        anchor._from.col = 7
        anchor._from.row = 1
        anchor._from.colOff = 200000
        anchor._from.rowOff = 50000
        anchor.to.col = 16
        anchor.to.row = 40
        anchor.to.colOff = 0
        anchor.to.rowOff = 0
        # Asegurar columnas del ancla visibles
        for c in range(8, 18):
            letter = get_column_letter(c)
            ws.column_dimensions[letter].hidden = False
            if float(ws.column_dimensions[letter].width or 0) < 3:
                ws.column_dimensions[letter].width = 4


def _find_record_row(ws, year: int, month: int, start: int, end: int) -> int | None:
    """Encuentra la fila ENE/FEB/... del año objetivo dentro de un bloque."""
    short = MONTH_SHORT[month]
    # Detectar bloques por PROM./años
    # Recorremos y asignamos año según encabezados PROM. 202X o secuencias ENE tras PROM
    current_year = None
    # Buscar indicios de año cerca
    for r in range(start, end + 1):
        label = str(ws.cell(r, 2).value or "").strip().upper()
        if "2023" in label:
            current_year = 2023
        elif "2024" in label and "PROM" in label:
            current_year = 2024
        elif "2025" in label and "PROM" in label:
            current_year = 2025
        elif "2026" in label:
            current_year = 2026
        elif label in MONTH_SHORT.values() or label in ("MZO", "SEP", "SET"):
            # Si estamos en secuencia de meses, inferir año por posición
            # Heurística mejor: contar bloques ENE
            pass

    for y, r0, r1 in _record_year_blocks(ws, start):
        if y != year:
            continue
        for r in range(r0, r1 + 1):
            label = str(ws.cell(r, 2).value or "").strip().upper().replace(" ", "")
            if label == short or (month == 3 and label in ("MAR", "MZO")):
                return r
        # Si no existe la fila del mes, crear debajo del último mes del bloque
        new_r = r1 + 1
        # insertar fila si la siguiente no está vacía / es PROMEDIO
        nxt = str(ws.cell(new_r, 2).value or "").upper()
        if nxt.startswith("PROM") or nxt in MONTH_SHORT.values():
            insert_rows_keep_refs(ws, new_r)
        ws.cell(new_r, 2, short)
        ws.cell(new_r, 7, f"=+C{new_r}+E{new_r}")
        return new_r
    return None


def _record_sin_olimpica_header(ws) -> int:
    for r in range(1, (ws.max_row or 0) + 1):
        label = str(ws.cell(r, 2).value or "").upper()
        if label.startswith("HISTORICO") and "SIN OLIMPICA" in label:
            return r
    return 45


def _record_year_blocks(ws, start: int) -> list[tuple[int, int, int]]:
    """
    Bloques (año, fila ENE, último mes) de la sección con o sin Olímpica.

    Cada sección arranca en 'PROM. 2023' y cada ENE abre el año siguiente
    (2024, 2025, 2026…); las filas se corren al insertar meses, así que no
    pueden ser fijas.
    """
    header = _record_sin_olimpica_header(ws)
    r_from, r_to = (1, header - 1) if start < header else (header, ws.max_row or header)
    months = set(MONTH_SHORT.values()) | {"MZO", "SET"}
    blocks: list[list[int]] = []
    current: list[int] | None = None
    year = 2023
    for r in range(r_from, r_to + 1):
        label = str(ws.cell(r, 2).value or "").strip().upper().replace(" ", "")
        if label == "ENE":
            year += 1
            current = [year, r, r]
            blocks.append(current)
        elif current is not None and label in months:
            current[2] = r
        elif label:
            current = None
    return [tuple(b) for b in blocks]


def _shift_formula_rows(formula: str, sheet_title: str, at_row: int, n: int) -> str:
    """Desplaza n filas las referencias a la propia hoja con fila >= at_row."""
    from openpyxl.formula.tokenizer import Token, Tokenizer

    def shift_ref(ref: str) -> str:
        def repl(m: re.Match) -> str:
            row = int(m.group(4))
            if row >= at_row:
                row += n
            return f"{m.group(1)}{m.group(2)}{m.group(3)}{row}"
        return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", repl, ref)

    try:
        tok = Tokenizer(formula)
    except Exception:
        return formula
    own = {sheet_title.lower(), f"'{sheet_title.lower()}'"}
    parts = []
    for item in tok.items:
        value = item.value
        if item.type == Token.OPERAND and item.subtype == Token.RANGE:
            if "!" in value:
                sheet, ref = value.rsplit("!", 1)
                if sheet.lower() in own:
                    value = f"{sheet}!{shift_ref(ref)}"
            else:
                value = shift_ref(value)
        parts.append(value)
    return "=" + "".join(parts)


def insert_rows_keep_refs(ws, idx: int, amount: int = 1) -> None:
    """insert_rows de openpyxl no ajusta fórmulas; Excel sí lo hace."""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = _shift_formula_rows(cell.value, ws.title, idx, amount)
    ws.insert_rows(idx, amount)


def normalize_record_row_totals(ws) -> None:
    """La columna G de cada mes suma su propia fila (=+C{r}+E{r} / =SUM(C{r},E{r}))."""
    plus = re.compile(r"^=\+C(\d+)\+E(\d+)$")
    summ = re.compile(r"^=SUM\(C(\d+),E(\d+)\)$")
    for r in range(1, (ws.max_row or 0) + 1):
        value = ws.cell(r, 7).value
        if not isinstance(value, str):
            continue
        if plus.match(value):
            ws.cell(r, 7).value = f"=+C{r}+E{r}"
        elif summ.match(value):
            ws.cell(r, 7).value = f"=SUM(C{r},E{r})"


def _extract_series_for_ppt(wb, model, year, month, col) -> dict:
    """Arma series históricas (millones) para slides de record."""
    ws = wb["Record cvda"]
    series = []
    # Leer año anterior + año actual
    current_row = _find_record_row(ws, year, month, 5, 40)
    rows_by_year = [
        (y, r)
        for y, r0, r1 in _record_year_blocks(ws, 5)
        if y in (year - 1, year)
        for r in range(r0, r1 + 1)
    ]
    for y, r in rows_by_year:
        label = str(ws.cell(r, 2).value or "").strip()
        if not label or label.upper().startswith("PROM"):
            continue
        venc = ws.cell(r, 3).value
        corr = ws.cell(r, 5).value
        pct_v = ws.cell(r, 4).value
        # Si es fórmula del mes actual, usar model
        if r == current_row:
            series.append({
                "label": f"{label} {y}",
                "vencida_mm": model["kpis"]["cartera_vencida"] / 1e6,
                "corriente_mm": model["kpis"]["cartera_corriente"] / 1e6,
                "total_mm": model["kpis"]["total_cartera"] / 1e6,
                "pct_vencida": model["kpis"]["pct_vencida"],
            })
            continue
        if isinstance(venc, str) or isinstance(corr, str):
            continue
        if venc is None:
            continue
        # Normalizar a millones si viene en pesos
        venc_f = float(venc)
        corr_f = float(corr or 0)
        if venc_f > 10000:  # pesos
            venc_f /= 1e6
            corr_f /= 1e6
        tot = venc_f + corr_f
        pv = float(pct_v) if isinstance(pct_v, (int, float)) else (venc_f / tot if tot else 0)
        series.append({
            "label": f"{label} {y}",
            "vencida_mm": venc_f,
            "corriente_mm": corr_f,
            "total_mm": tot,
            "pct_vencida": pv,
        })

    # Edades del mes (para slide indicador histórico corto)
    edades_hist = []
    ws_e = wb["inf cartera x edades"]
    # últimas 6 columnas de meses antes de DIF
    month_cols = []
    for c in range(3, 40):
        h = str(ws_e.cell(3, c).value or "").strip().upper()
        if h in ("DIF.", "VAR.", "PART.") or h.startswith("DIF") or h.startswith("VAR") or h.startswith("PART"):
            break
        if h in [MONTH_NAMES[i] for i in range(1, 13)]:
            month_cols.append((c, h))
    for c, h in month_cols[-6:]:
        total = ws_e.cell(4, c).value
        venc = ws_e.cell(6, c).value
        if isinstance(total, (int, float)) and isinstance(venc, (int, float)):
            edades_hist.append({
                "mes": h[:3],
                "total_mm": float(total) / 1e6,
                "vencida_mm": float(venc) / 1e6,
                "pct": float(venc) / float(total) if total else 0,
            })

    return {
        "record_series": series[-18:],  # últimos ~18 meses
        "edades_hist": edades_hist,
        "mes_actual": f"{MONTH_SHORT[month]} {year}",
    }
