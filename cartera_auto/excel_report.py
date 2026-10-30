# -*- coding: utf-8 -*-
"""Genera Excel clonando la plantilla real del Informe de Cartera COLBEEF."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .client_master import (
    carnes_extra_for,
    load_maestro,
    merge_from_run,
    plazo_for,
    principal_for,
)
from .config import AGE_BUCKETS, COMPANY_NAME, ROOT, UNIT_ORDER
from .historico import freeze_unidad_linked_historico, update_historico
from .monthly_update import apply_monthly_updates, read_plazos_from_comportamiento
from .parsers import format_corte_label
from .unidad_resumen import update_unidad_resumen

TEMPLATE = ROOT / "templates" / "plantilla_informe_cartera.xlsx"
FONT = "Century Gothic"
GREEN = "92D050"
HEADER_FILL = PatternFill("solid", fgColor="4472C4")  # fallback; plantilla usa theme
AGE_FILL = PatternFill("solid", fgColor=GREEN)
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL_RED = PatternFill("solid", fgColor="FF0000")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
# Formato contable del original Cartera (enteros, ceros como guión)
MONEY_FMT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
MONEY_FMT_INT = MONEY_FMT

# Buckets como en SIIMED / plantilla
AGE_COLS = [
    "CORRIENTE",
    "1 AL 7",
    "8 AL 15",
    "16 AL 30",
    "31 AL 60",
    "61 AL 90",
    "91 AL 180",
    "181 AL 360",
    "361+",
]

UNIT_SHEET_MAP = {
    "SUBPRODUCTOS": "subproductos",
    "OTRAS CxC": "otras cxc",
    "LOCAL": "local",
    "REGIONAL": "regional",
    "NACIONAL": "nacional",
    "INSTITUCIONAL": "institucional",
    "DESPOSTE": "desposte",
    "CARNES COLBEEF": "carnes",
}

# Ancho del rótulo TOTAL (filas total + porcentaje) en los originales.
TOTAL_MERGE_END = {
    "carnes": 5,
    "subproductos": 3,
    "otras cxc": 4,
    "local": 4,
    "regional": 3,
    "nacional": 4,
    "institucional": 4,
    "desposte": 3,
}

UNIT_LABELS = {
    "OTRAS CxC": "OTRAS C*C CLIENTES",
}

ACCOUNT_LABELS = {
    "SUBPRODUCTOS": "130503  SUBPRODUCTOS",
    "OTRAS CxC": "130505  OTRAS C*C CLIENTES",
    "LOCAL": "13050901  LOCAL",
    "REGIONAL": "13050902  REGIONAL",
    "NACIONAL": "13050903  NACIONAL",
    "INSTITUCIONAL": "13050904  INSTITUCIONAL",
    "DESPOSTE": "13050920  DESPOSTE",
    "CARNES COLBEEF": "13051401  CARNES COLBEEF",
}

# Filas de datos fijas en los originales (TOTAL siempre en la misma fila);
# si hay más clientes la tabla crece.
UNIT_MIN_ROWS = {
    "subproductos": 10,
    "otras cxc": 20,
    "local": 40,
    "regional": 14,
    "institucional": 10,
    "desposte": 13,
    "carnes": 25,
}

# Unidades extra que se muestran en la hoja carnes. Vacío: DIFICIL COBRO no se lista
# (criterio de Cartera desde agosto 2026; antes variaba mes a mes).
CARNES_EXTRA_UNITS: tuple[str, ...] = ()


def generate_excel(model: dict, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not TEMPLATE.exists():
        raise FileNotFoundError(f"No está la plantilla Excel: {TEMPLATE}")

    shutil.copy2(TEMPLATE, output_path)
    wb = load_workbook(output_path)

    corte = model["meta"].get("corte")
    corte_label = format_corte_label(corte)
    corte_siimed = _siimed_corte_text(corte)
    procesado = model["meta"].get("procesado")
    maestro = load_maestro()

    plazos = _read_plazos_from_template(wb)
    plazos.update(read_plazos_from_comportamiento(wb))
    for nit, entry in maestro.items():
        if entry.get("plazo") is not None:
            plazos[nit] = entry["plazo"]

    # Congelar meses previos (fórmulas unidad! → valores) antes de pisar la hoja unidad
    freeze_unidad_linked_historico(wb)

    _fill_consolidada(wb, model, corte_siimed, procesado, plazos, maestro)
    _fill_unidad(wb, model)
    _fill_unit_sheets(wb, model, maestro)
    _touch_titles(wb, corte_siimed, corte_label, procesado)

    consolidada_name = apply_monthly_updates(wb, model)
    if consolidada_name and corte:
        update_unidad_resumen(wb, model, consolidada_name, int(corte[5:7]))

    # Actualizar base histórica (edades, record, rotación)
    model["historico"] = update_historico(wb, model)

    # Réplica del original: no agregar hojas auxiliares al entregable.
    if "AUTO_RESUMEN" in wb.sheetnames:
        del wb["AUTO_RESUMEN"]

    # Quitar vínculos externos heredados de la plantilla (evita "Reparar" en Excel)
    _sanitize_external_formulas(wb)

    wb.save(output_path)

    # NO restaurar style/colors vía zip: corrompe drawings y Excel quita las gráficas.
    # openpyxl ya conserva los 7 charts; los colores de tema bastan para uso operativo.

    merge_from_run(model, maestro)
    return output_path


def _sanitize_external_formulas(wb) -> int:
    """
    Convierte =+[N]Hoja!ref / =+'[N]Hoja'!ref en referencias locales.
    Si la hoja no existe (ej. Hoja1 de otro libro), deja 0 para no romper el xlsx.
    Evita el aviso 'Reparar' de Excel por externalLinks.
    """
    names = {n.lower(): n for n in wb.sheetnames}
    fixed = 0
    for ws in wb.worksheets:
        max_r = min(ws.max_row or 1, 400)
        # Formula usa columnas BI/BJ/BK (61–63) para rotación histórica.
        max_c = min(ws.max_column or 1, 80)
        for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c):
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or "[" not in v or not v.startswith("="):
                    continue
                new_v = _localize_external_formula(v, names)
                if new_v != v:
                    cell.value = new_v
                    fixed += 1
    try:
        if getattr(wb, "_external_links", None):
            wb._external_links = []
    except Exception:
        pass
    return fixed


def _localize_external_formula(formula: str, names: dict[str, str]) -> str:
    """Reemplaza refs [N]Hoja! por hoja local o 0 si no existe."""

    def sheet_prefix(sheet: str) -> str | None:
        real = names.get(sheet.strip().lower())
        if not real:
            return None
        if " " in real or any(ch in real for ch in ("*", "/", "?", ":")):
            return f"='{real}'!"
        return f"={real}!"

    # =+'[N]Nombre hoja'!celda
    def repl_quoted(m: re.Match) -> str:
        pref = sheet_prefix(m.group(1))
        return (pref or "=0*") + m.group(2)

    out = re.sub(
        r"=\+?'\[(\d+)\]([^']+)'!(.+)$",
        repl_quoted,
        formula,
        flags=re.IGNORECASE,
    )
    if out != formula:
        if out.startswith("=0*"):
            return "0"
        return out

    # =+[N]Hoja!celda
    def repl_plain(m: re.Match) -> str:
        pref = sheet_prefix(m.group(1))
        return (pref or "=0*") + m.group(2)

    out = re.sub(
        r"=\+?\[(\d+)\]([^'!]+)!(.+)$",
        repl_plain,
        formula,
        flags=re.IGNORECASE,
    )
    if out.startswith("=0*"):
        return "0"
    return out


def _siimed_corte_text(corte: str | None) -> str:
    if not corte:
        return "Con corte a: —"
    try:
        y, m, d = corte.split("-")
        months = [
            "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
        ]
        return f"Con corte a: {months[int(m)]} {int(d)} Del {y} - Fecha Documento"
    except Exception:
        return f"Con corte a: {corte}"


def _font(size=11, bold=False, color=None):
    kw = {"name": FONT, "size": size, "bold": bold}
    if color:
        kw["color"] = color
    return Font(**kw)


def _money(cell, value):
    cell.value = float(value or 0)
    cell.number_format = MONEY_FMT
    cell.font = _font(10)
    cell.alignment = Alignment(horizontal="right")
    cell.border = THIN


def _money_formula(cell, formula: str, *, bold: bool = False):
    """Fórmula monetaria con el formato contable del original."""
    cell.value = formula
    cell.number_format = MONEY_FMT
    cell.font = _font(10 if not bold else 11, bold, "FFFFFF" if bold else None)
    cell.alignment = Alignment(horizontal="right")
    cell.border = THIN


def _clear_block(ws, start_row, end_row, start_col, end_col, clear_style: bool = False):
    # deshacer merges que intersecten el bloque
    to_unmerge = []
    for rng in list(ws.merged_cells.ranges):
        if rng.max_row >= start_row and rng.min_row <= end_row and rng.max_col >= start_col and rng.min_col <= end_col:
            to_unmerge.append(str(rng))
    for rng in to_unmerge:
        try:
            ws.unmerge_cells(rng)
        except Exception:
            pass
    empty_fill = PatternFill(fill_type=None)
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            try:
                cell.value = None
            except AttributeError:
                continue
            if clear_style:
                cell.fill = empty_fill
                cell.font = Font(name=FONT, size=10)
                cell.border = Border()
                cell.number_format = "General"


def _read_plazos_from_template(wb) -> dict[str, int | float | str]:
    """Conserva PLAZO (col D) del informe anterior en plantilla, por NIT."""
    plazos: dict[str, int | float | str] = {}
    plazos.update(_read_plazos_master())
    for name in wb.sheetnames:
        if not name.lower().startswith("consolidada"):
            continue
        ws = wb[name]
        for r in range(8, 200):
            nit = ws.cell(r, 2).value
            plazo = ws.cell(r, 4).value
            if nit and plazo is not None and str(plazo).strip() != "":
                plazos[str(nit).strip().replace(".0", "")] = plazo
        break
    return plazos


def _read_plazos_master() -> dict[str, int | float | str]:
    """Maestro NIT→PLAZO (data/plazos_nit.json), actualizable mes a mes."""
    import json

    path = ROOT / "data" / "plazos_nit.json"
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return {str(k): v for k, v in data.items()}
    except Exception:
        return {}


def _fill_consolidada(wb, model, corte_siimed, procesado=None, plazos=None, maestro=None):
    name = "consolidada junio" if "consolidada junio" in wb.sheetnames else None
    if not name:
        # buscar hoja consolidada*
        for n in wb.sheetnames:
            if n.lower().startswith("consolidada"):
                name = n
                break
    if not name:
        return
    ws = wb[name]
    ws["B1"] = COMPANY_NAME
    ws["B1"].font = _font(10, True)
    ws["B2"] = "EDADES DE CARTERA A 10 COLUMNAS GENERAL POR IDENTIFICACION DEL TERCERO"
    ws["B2"].font = _font(10, True)
    ws["B3"] = corte_siimed
    ws["B3"].font = _font(10, True)
    # Fecha corta esquina superior derecha (AGO/31/2026), como el original
    corte = model.get("meta", {}).get("corte")
    if corte:
        try:
            y, m, d = corte.split("-")
            short = [
                "", "ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                "JUL", "AGO", "SEP", "OCT", "NOV", "DIC",
            ][int(m)]
            ws.cell(1, 16, f"{short}/{int(d):02d}/{y}").font = _font(10, True)
        except Exception:
            pass
    if procesado:
        ws["B5"] = procesado if str(procesado).lower().startswith("procesado") else f"Procesado: {procesado}"
        ws["B5"].font = _font(10)

    plazos = plazos or {}

    # cabeceras fila 7
    headers = ["NIT", "NOMBRE", "PLAZO", "SALDO DEL NIT", "CORRIENTE", "VALOR VENCIDO"] + [
        "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60", "61 AL 90", "91 AL 180", "181 AL 360", "<== 361"
    ]
    for i, h in enumerate(headers):
        cell = ws.cell(7, 2 + i, h)
        cell.font = _font(11, True, "FFFFFF")
        cell.fill = AGE_FILL if i >= 6 else PatternFill("solid", fgColor="5B9BD5")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN

    df = model["por_cliente"]
    # limpiar filas viejas (hasta ~200)
    _clear_block(ws, 8, 200, 2, 16, clear_style=True)

    for i in range(len(df)):
        r = 8 + i
        row = df.iloc[i]
        nit_s = str(row["nit"]).strip()
        nit_val: int | str = int(nit_s) if nit_s.isdigit() else nit_s
        ws.cell(r, 2, nit_val).font = _font(10)
        ws.cell(r, 2).number_format = "##,##0_);[Red]\\(##,##0\\)"
        ws.cell(r, 2).border = THIN
        ws.cell(r, 3, row["nombre"]).font = _font(10)
        ws.cell(r, 3).number_format = "General"
        ws.cell(r, 3).border = THIN
        plazo = plazos.get(nit_s) if plazos else None
        if plazo is None and maestro:
            plazo = plazo_for(nit_s, maestro)
        if plazo is not None:
            ws.cell(r, 4, int(plazo) if str(plazo).isdigit() else plazo).font = _font(10)
        ws.cell(r, 4).number_format = "General"
        ws.cell(r, 4).border = THIN
        _money(ws.cell(r, 5), row["saldo"])
        _money(ws.cell(r, 6), row.get("CORRIENTE", 0))
        _money(ws.cell(r, 7), row.get("vencido", 0))
        for j, b in enumerate([x for x in AGE_COLS if x != "CORRIENTE"]):
            key = b if b != "361+" else "361+"
            _money(ws.cell(r, 8 + j), row.get(key, 0))

    # Grand Total en fila fija 82 (hueco vacío debajo de clientes), como el V1.
    last_data = 7 + len(df)
    r = max(82, last_data + 1)
    ws.cell(r, 2, "Grand Total").font = _font(11, True)
    for c in range(2, 16):
        ws.cell(r, c).fill = TOTAL_FILL_RED
        ws.cell(r, c).font = _font(11, True, "FFFFFF")
        ws.cell(r, c).border = THIN
    for c in range(5, 16):
        letter = get_column_letter(c)
        _money_formula(ws.cell(r, c), f"=SUM({letter}8:{letter}{r - 1})", bold=True)
        ws.cell(r, c).fill = TOTAL_FILL_RED

    # Quitar filas rojas/vacías heredadas de la plantilla debajo del total.
    _clear_block(ws, r + 1, 200, 2, 16, clear_style=True)

    ws.auto_filter.ref = f"B7:O{r}"
    ws.print_area = f"B1:O{r}"


def _fill_unidad(wb, model):
    if "unidad" not in wb.sheetnames:
        return
    ws = wb["unidad"]
    headers = ["CUENTA", "SALDO DEL NIT", "CORRIENTE", "VALOR VENCIDO"] + [
        "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60", "61 AL 90", "91 AL 180", "181 AL 360", "<== 361"
    ]
    for i, h in enumerate(headers):
        cell = ws.cell(2, 2 + i, h)
        cell.font = _font(11, True, "FFFFFF")
        cell.fill = AGE_FILL if i >= 3 else PatternFill("solid", fgColor="5B9BD5")
        cell.border = THIN
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    _clear_block(ws, 3, 12, 2, 14)
    df = model["por_unidad"]
    for i in range(len(df)):
        r = 3 + i
        row = df.iloc[i]
        unit = row["unidad"]
        label = UNIT_LABELS.get(unit, unit)
        existing = ws.cell(r, 2).value
        if existing and str(existing).strip():
            label = str(existing).strip()
        ws.cell(r, 2, label).font = _font(11, True)
        ws.cell(r, 2).border = THIN
        _money(ws.cell(r, 3), row["saldo"])
        _money(ws.cell(r, 4), row.get("CORRIENTE", 0))
        _money(ws.cell(r, 5), row.get("vencido", 0))
        for j, b in enumerate([x for x in AGE_COLS if x != "CORRIENTE"]):
            _money(ws.cell(r, 6 + j), row.get(b, 0))

    r = 3 + len(df)
    ws.cell(r, 2, "Grand Total").font = _font(11, True)
    for c in range(2, 15):
        ws.cell(r, c).fill = TOTAL_FILL_RED
        ws.cell(r, c).font = _font(11, True, "FFFFFF")
        ws.cell(r, c).border = THIN
    for c in range(3, 15):
        letter = get_column_letter(c)
        _money_formula(ws.cell(r, c), f"=SUM({letter}3:{letter}{r - 1})", bold=True)
        ws.cell(r, c).fill = TOTAL_FILL_RED
    ws.print_area = f"B1:N{r}"


def _header_col_map(ws, header_row: int, start_col: int = 2, end_col: int = 20) -> dict[str, int]:
    """Mapea encabezados de la plantilla → índice de columna."""
    mapping: dict[str, int] = {}
    aliases = {
        "NO": "NO",
        "NIT": "NIT",
        "NOMBRE": "NOMBRE",
        "CUENTA": "CUENTA",
        "PRINCIPAL": "PRINCIPAL",
        "CONDICION DE PAGO": "CONDICION",
        "CUPO": "CUPO",
        "REGIÓN": "REGION",
        "REGION": "REGION",
        "SALDO DEL NIT": "SALDO",
        "CORRIENTE": "CORRIENTE",
        "VALOR VENCIDO": "VENCIDO",
        "1 AL 7": "1 AL 7",
        "8 AL 15": "8 AL 15",
        "16 AL 30": "16 AL 30",
        "31 AL 60": "31 AL 60",
        "61 AL 90": "61 AL 90",
        "91 AL 180": "91 AL 180",
        "181 AL 360": "181 AL 360",
        "<== 361": "361+",
        "361+": "361+",
    }
    # Ordenar aliases: claves más largas primero (evita que NIT matchee SALDO DEL NIT)
    ordered = sorted(aliases.items(), key=lambda x: len(x[0]), reverse=True)
    for c in range(start_col, end_col + 1):
        raw = str(ws.cell(header_row, c).value or "").strip().upper()
        if not raw:
            continue
        for key, tag in ordered:
            if raw == key or (key in raw and tag not in mapping):
                if tag == "NIT" and "SALDO" in raw:
                    continue
                mapping[tag] = c
                break
    return mapping


def _unit_detail_df(model: dict, unit: str):
    import pandas as pd

    df = model["detalle_por_unidad"].get(unit)
    if df is None or df.empty:
        return None
    if unit != "CARNES COLBEEF":
        return df
    frames = [df]
    for extra in CARNES_EXTRA_UNITS:
        extra_df = model["detalle_por_unidad"].get(extra)
        if extra_df is not None and not extra_df.empty:
            frames.append(extra_df)
    if len(frames) == 1:
        return df
    merged = pd.concat(frames, ignore_index=True)
    return merged.sort_values("saldo", ascending=False).reset_index(drop=True)


def _fill_unit_sheets(wb, model, maestro=None):
    maestro = maestro or {}
    age_keys = [x for x in AGE_COLS if x != "CORRIENTE"]
    for unit, sheet_name in UNIT_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue
        df = _unit_detail_df(model, unit)
        if df is None or df.empty:
            continue
        ws = wb[sheet_name]
        header_row = 2
        for r in range(1, 6):
            vals = [str(ws.cell(r, c).value or "").upper() for c in range(1, 12)]
            if any("NOMBRE" in v or "SALDO" in v for v in vals):
                header_row = r
                break

        colmap = _header_col_map(ws, header_row)
        if not colmap:
            continue

        _clear_block(ws, header_row + 1, header_row + 120, 2, 16)

        for i in range(len(df)):
            r = header_row + 1 + i
            row = df.iloc[i]
            if "NO" in colmap:
                cell = ws.cell(r, colmap["NO"], i + 1)
                cell.font = _font(11)
                cell.border = THIN
                cell.fill = PatternFill(fill_type=None)
            if "NIT" in colmap:
                # El original numera 1, 2, 3… aunque el encabezado diga NIT
                cell = ws.cell(r, colmap["NIT"], i + 1)
                cell.font = _font(10)
                cell.border = THIN
                cell.fill = PatternFill(fill_type=None)
            if "NOMBRE" in colmap:
                cell = ws.cell(r, colmap["NOMBRE"], row["nombre"])
                cell.font = _font(10)
                cell.border = THIN
                cell.fill = PatternFill(fill_type=None)
            if sheet_name == "carnes":
                extra = carnes_extra_for(row["nombre"], maestro)
                if "CONDICION" in colmap and extra.get("condicion_pago"):
                    ws.cell(r, colmap["CONDICION"], extra["condicion_pago"]).font = _font(10)
                    ws.cell(r, colmap["CONDICION"]).border = THIN
                if "CUPO" in colmap and extra.get("cupo") is not None:
                    _money(ws.cell(r, colmap["CUPO"]), extra["cupo"])
            if sheet_name == "local" and "PRINCIPAL" in colmap:
                pr = principal_for(row["nombre"], maestro)
                if pr:
                    ws.cell(r, colmap["PRINCIPAL"], pr).font = _font(10)
                    ws.cell(r, colmap["PRINCIPAL"]).border = THIN
            # CUENTA / REGION: usar cuenta del detalle o etiqueta de unidad
            cuenta_val = ACCOUNT_LABELS.get(unit, row.get("cuenta") or unit)
            if "CUENTA" in colmap:
                cell = ws.cell(r, colmap["CUENTA"], cuenta_val)
                cell.font = _font(10)
                cell.border = THIN
                cell.fill = PatternFill(fill_type=None)
            if "REGION" in colmap:
                cell = ws.cell(r, colmap["REGION"], cuenta_val)
                cell.font = _font(10)
                cell.border = THIN
                cell.fill = PatternFill(fill_type=None)
            if "SALDO" in colmap:
                _money(ws.cell(r, colmap["SALDO"]), row["saldo"])
            if "CORRIENTE" in colmap:
                _money(ws.cell(r, colmap["CORRIENTE"]), row.get("CORRIENTE", 0))
            if "VENCIDO" in colmap:
                _money(ws.cell(r, colmap["VENCIDO"]), row.get("vencido", 0))
            for b in age_keys:
                if b in colmap:
                    _money(ws.cell(r, colmap[b]), row.get(b, 0))

        n_rows = max(len(df), UNIT_MIN_ROWS.get(sheet_name, 0))
        num_col = colmap.get("NO") or colmap.get("NIT")
        for i in range(len(df), n_rows):
            rr = header_row + 1 + i
            if num_col:
                ws.cell(rr, num_col, i + 1).font = _font(11)
            for c in sorted(set(colmap.values())):
                ws.cell(rr, c).border = THIN
                ws.cell(rr, c).fill = PatternFill(fill_type=None)

        # Fila TOTAL (roja) + fila % (como informe manual Cartera)
        r = header_row + 1 + n_rows
        white = _font(11, True, "FFFFFF")
        money_cols = [k for k in ("SALDO", "CORRIENTE", "VENCIDO", *age_keys) if k in colmap]
        saldo_col = colmap.get("SALDO")
        first_col = min(colmap.values())
        last_col = max(colmap.values())
        rp = r + 1

        # El original combina el rótulo TOTAL sobre las dos filas, hasta SALDO.
        if saldo_col and saldo_col > first_col:
            merge_end = min(
                saldo_col - 1,
                TOTAL_MERGE_END.get(sheet_name, saldo_col - 1),
            )
            for rr in (r, rp):
                for c in range(first_col, merge_end + 1):
                    ws.cell(rr, c).fill = TOTAL_FILL_RED
                    ws.cell(rr, c).border = THIN
            try:
                ws.merge_cells(
                    start_row=r,
                    start_column=first_col,
                    end_row=rp,
                    end_column=merge_end,
                )
            except ValueError:
                pass
            label = ws.cell(r, first_col, "TOTAL")
        else:
            label_col = colmap.get("NOMBRE", first_col)
            label = ws.cell(r, label_col, "TOTAL")
        label.font = white
        label.fill = TOTAL_FILL_RED
        label.border = THIN

        for key in money_cols:
            c = colmap[key]
            letter = get_column_letter(c)
            _money_formula(
                ws.cell(r, c),
                f"=SUM({letter}{header_row + 1}:{letter}{r - 1})",
                bold=True,
            )
            ws.cell(r, c).fill = TOTAL_FILL_RED

        # Fila de participación % bajo TOTAL
        if "SALDO" in colmap and money_cols:
            saldo_letter = get_column_letter(colmap["SALDO"])
            share_cols = [colmap[k] for k in money_cols if k != "SALDO"]
            for k in money_cols:
                c = colmap[k]
                letter = get_column_letter(c)
                if k == "SALDO" and share_cols:
                    first_l = get_column_letter(min(share_cols))
                    last_l = get_column_letter(max(share_cols))
                    formula = f"=SUM({first_l}{rp}:{last_l}{rp})"
                else:
                    formula = f"=+{letter}{r}/${saldo_letter}${r}"
                cell = ws.cell(rp, c, formula)
                cell.number_format = "0%"
                cell.fill = TOTAL_FILL_RED
                cell.font = white
                cell.border = THIN
        # Quitar filas rojas vacías heredadas debajo del bloque TOTAL/%.
        _clear_block(ws, rp + 1, header_row + 120, 2, 16, clear_style=True)
        ws.print_area = (
            f"{get_column_letter(first_col)}1:"
            f"{get_column_letter(last_col)}{rp}"
        )


def _touch_titles(wb, corte_siimed, corte_label, procesado=None):
    for name in wb.sheetnames:
        ws = wb[name]
        # actualizar textos de corte si existen en primeras filas
        for r in range(1, 6):
            for c in range(1, 5):
                v = ws.cell(r, c).value
                if isinstance(v, str) and "corte a" in v.lower():
                    ws.cell(r, c).value = corte_siimed
                    ws.cell(r, c).font = _font(10, True)
                if procesado and isinstance(v, str) and v.lower().startswith("procesado"):
                    ws.cell(r, c).value = (
                        procesado if str(procesado).lower().startswith("procesado")
                        else f"Procesado: {procesado}"
                    )
                    ws.cell(r, c).font = _font(10)


def _fill_auto_resumen(ws, model, corte_label):
    ws["A1"] = COMPANY_NAME
    ws["A1"].font = _font(14, True, "185E22")
    ws["A2"] = "RESUMEN AUTOMÁTICO DE CARTERA (generado por Cartera Auto)"
    ws["A2"].font = _font(12, True)
    ws["A3"] = f"Corte: {corte_label}"
    ws["A3"].font = _font(11)

    k = model["kpis"]
    rows = [
        ("Total cartera", k["total_cartera"]),
        ("Cartera corriente", k["cartera_corriente"]),
        ("Cartera vencida", k["cartera_vencida"]),
        ("% Vencida", k["pct_vencida"]),
        ("Total sin Olímpica", k["total_sin_olimpica"]),
        ("% Vencida sin Olímpica", k["pct_vencida_sin_olimpica"]),
        ("Saldo Olímpica", k["olimpica_saldo"]),
        ("N° clientes", k["n_clientes"]),
    ]
    ws["A5"] = "Indicador"
    ws["B5"] = "Valor"
    for cell in (ws["A5"], ws["B5"]):
        cell.font = _font(11, True, "FFFFFF")
        cell.fill = PatternFill("solid", fgColor="185E22")
        cell.border = THIN

    for i, (lab, val) in enumerate(rows):
        r = 6 + i
        ws.cell(r, 1, lab).font = _font(11)
        ws.cell(r, 1).border = THIN
        if " %" in lab or lab.startswith("%"):
            ws.cell(r, 2, val).number_format = "0.0%"
            ws.cell(r, 2).font = _font(11, True)
        elif lab.startswith("N°"):
            ws.cell(r, 2, int(val)).font = _font(11, True)
        else:
            _money(ws.cell(r, 2), val)
        ws.cell(r, 2).border = THIN

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
