# -*- coding: utf-8 -*-
"""Compara estructura de exportes SIIMED entre meses (columnas, metadatos, señales de edición)."""

from __future__ import annotations

import sys
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
DL = Path(r"C:\Users\CAMPUSLANDS\Downloads")

GROUPS = {
    "CARTERA POR EDADES (general)": [
        "CARTERA GENERAL 30 ABRIL.xlsx",
        "CARTERA A CORTE 31 MAYO.xlsx",
        "cartera a 30 junio.xlsx",
        "Cartera a corte 31 julio.xlsx",
        "cartera general 31 agosto 2026.xlsx",
        "CARTERA A 30 SEPT.xlsx",
    ],
    "CARTERA POR CUENTA": [
        "CARTERA POR CUENTA 30 ABRIL.xlsx",
        "cartera por cuenta 31 mayo.xlsx",
        "cartera por unidad 30 junio.xlsx",
        "Cartera por cuenta 31 julio.xlsx",
        "cartera por cuenta agosto 2026.xlsx",
        "cartera por cuenta 30 sept.xlsx",
    ],
    "BALANCE": [
        "BALANCE ABRIL.xlsx",
        "datos balance mayo.xlsx",
        "balance ingresos Cartera junio 06 2026.xlsx",
        "Balence Julio cartera 2026.xlsx",
        "Balance agosto cartera.xlsx",
        "BALANCE DE PRUEBA CTA 13 (1).xlsx",
        "BALANCE DE PRUEBA CTA 4 (1).xlsx",
    ],
}


def core_props(path: Path) -> dict:
    import re

    out = {}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "docProps/core.xml" in names:
            xml = z.read("docProps/core.xml").decode("utf-8", "replace")
            for tag in ("creator", "lastModifiedBy", "created", "modified"):
                m = re.search(rf"<[^>]*{tag}[^>]*>([^<]*)<", xml)
                out[tag] = m.group(1) if m else None
        if "docProps/app.xml" in names:
            xml = z.read("docProps/app.xml").decode("utf-8", "replace")
            m = re.search(r"<Application>([^<]*)<", xml)
            out["app"] = m.group(1) if m else None
            m = re.search(r"<AppVersion>([^<]*)<", xml)
            out["appver"] = m.group(1) if m else None
        out["vba"] = any("vbaProject" in n for n in names)
        out["comments"] = sum(1 for n in names if "comments" in n)
        out["calcChain"] = "xl/calcChain.xml" in names
        out["pivot"] = any("pivot" in n for n in names)
    return out


def find_header_row(ws, max_r=15):
    best, best_n = None, 0
    for r in range(1, max_r + 1):
        vals = [ws.cell(r, c).value for c in range(1, (ws.max_column or 1) + 1)]
        texts = [v for v in vals if isinstance(v, str) and v.strip()]
        if len(texts) > best_n:
            best, best_n = r, len(texts)
    return best


def analyze(path: Path) -> None:
    print(f"\n### {path.name}  ({path.stat().st_size:,} bytes)")
    if not path.exists():
        print("   NO EXISTE")
        return
    p = core_props(path)
    print(
        f"   creado={p.get('created')} por={p.get('creator')!r} | modificado={p.get('modified')} "
        f"por={p.get('lastModifiedBy')!r} | app={p.get('app')} {p.get('appver')} | "
        f"calcChain={p.get('calcChain')} comentarios={p.get('comments')} vba={p.get('vba')} pivot={p.get('pivot')}"
    )
    wb = load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        hr = find_header_row(ws)
        print(
            f"   [hoja {ws.title!r}] dim={ws.dimensions} max_row={ws.max_row} max_col={ws.max_column} "
            f"estado={ws.sheet_state} header_row={hr}"
        )
        for r in range(1, (hr or 1)):
            vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 25) + 1)]
            vals = [v for v in vals if v is not None]
            if vals:
                print(f"      titulo r{r}: {vals}")
        if hr:
            headers = [ws.cell(hr, c).value for c in range(1, ws.max_column + 1)]
            print(f"      encabezados: {headers}")
        formulas = 0
        fonts, fills, fmts = Counter(), Counter(), Counter()
        data_rows = 0
        manual_fill_cells = []
        for row in ws.iter_rows(min_row=(hr or 1) + 1, max_row=ws.max_row):
            if any(c.value is not None for c in row):
                data_rows += 1
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas += 1
                if c.value is not None:
                    fonts[(c.font.name, c.font.b, c.font.color.rgb if c.font.color else None)] += 1
                    fmts[c.number_format] += 1
                fg = c.fill.fgColor.rgb if c.fill and c.fill.fill_type == "solid" else None
                if fg:
                    fills[fg] += 1
                    if len(manual_fill_cells) < 5:
                        manual_fill_cells.append(c.coordinate)
        hidden_rows = [r for r, d in ws.row_dimensions.items() if d.hidden]
        hidden_cols = [k for k, d in ws.column_dimensions.items() if d.hidden]
        print(
            f"      filas_datos={data_rows} formulas={formulas} merges={len(ws.merged_cells.ranges)} "
            f"autofiltro={ws.auto_filter.ref} filas_ocultas={hidden_rows[:10]} cols_ocultas={hidden_cols}"
        )
        print(f"      fuentes={fonts.most_common(4)}")
        print(f"      rellenos={fills.most_common(4)} ej={manual_fill_cells}")
        print(f"      formatos={fmts.most_common(4)}")
        # Últimas filas (totales / pie de reporte)
        tail = []
        for r in range(max(1, ws.max_row - 3), ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 20) + 1)]
            vals = [v for v in vals if v is not None]
            if vals:
                tail.append((r, vals[:10]))
        for r, v in tail:
            print(f"      pie r{r}: {v}")


if __name__ == "__main__":
    for group, files in GROUPS.items():
        print("\n" + "=" * 100)
        print(group)
        print("=" * 100)
        for f in files:
            analyze(DL / f)
