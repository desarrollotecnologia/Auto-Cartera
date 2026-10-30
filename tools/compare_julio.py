# -*- coding: utf-8 -*-
"""Comparación profunda Informe original vs generado (Julio 2026)."""
import sys
import io
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import openpyxl
import pandas as pd

ORIG = Path(r"c:\Users\CAMPUSLANDS\Downloads\1. Informe de Cartera julio V1.xlsx")
GEN = Path(r"c:\Users\CAMPUSLANDS\Downloads\Informe_Cartera_2026-07-31_20260813_0803.xlsx")
SRC_CARTERA = Path(r"c:\Users\CAMPUSLANDS\Downloads\Cartera a corte 31 julio.xlsx")
SRC_UNIDAD = Path(r"c:\Users\CAMPUSLANDS\Downloads\Cartera por cuenta 31 julio.xlsx")
SRC_BALANCE = Path(r"c:\Users\CAMPUSLANDS\Downloads\Balence Julio cartera 2026.xlsx")

TOL = 1.0  # pesos


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def find_total_row(ws, col=2, labels=("TOTAL", "GRAND TOTAL"), start=1, end=400):
    for r in range(start, end + 1):
        v = ws.cell(r, col).value
        if v and str(v).strip().upper() in labels:
            return r
    return None


def find_month_col(ws, month_name="JULIO", header_rows=(3, 24)):
    for hr in header_rows:
        for c in range(3, 40):
            v = ws.cell(hr, c).value
            if v and month_name.upper() in str(v).upper():
                return c, hr
    return None, None


AGE_COLS = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15]  # corriente, vencido, buckets...


def compare_consolidada(wb_o, wb_g):
    ws_o = wb_o["consolidada julio"]
    ws_g = wb_g["consolidada junio"]
    tr_o = find_total_row(ws_o)
    tr_g = find_total_row(ws_g)
    diffs = []
    cols = {
        "saldo": 5,
        "corriente": 6,
        "vencido": 7,
        "1 AL 7": 8,
        "8 AL 15": 9,
        "16 AL 30": 10,
        "31 AL 60": 11,
        "61 AL 90": 12,
        "91 AL 180": 13,
        "181 AL 360": 14,
        "361+": 15,
    }
    for k, c in cols.items():
        vo, vg = num(ws_o.cell(tr_o, c).value), num(ws_g.cell(tr_g, c).value)
        if vo is not None and vg is not None and abs(vo - vg) > TOL:
            diffs.append(("consolidada TOTAL", k, vo, vg, vg - vo))

    # clientes por NIT
    clients_o = {}
    for r in range(8, tr_o):
        nit = str(ws_o.cell(r, 2).value or "").strip()
        if not nit:
            continue
        clients_o[nit] = {k: num(ws_o.cell(r, c).value) for k, c in cols.items()}

    clients_g = {}
    for r in range(8, tr_g):
        nit = str(ws_g.cell(r, 2).value or "").strip()
        if not nit:
            continue
        clients_g[nit] = {k: num(ws_g.cell(r, c).value) for k, c in cols.items()}

    only_o = set(clients_o) - set(clients_g)
    only_g = set(clients_g) - set(clients_o)
    if only_o:
        diffs.append(("consolidada clientes", f"solo en original ({len(only_o)})", list(only_o)[:5], None, None))
    if only_g:
        diffs.append(("consolidada clientes", f"solo en generado ({len(only_g)})", list(only_g)[:5], None, None))

    nit_diffs = []
    for nit in set(clients_o) & set(clients_g):
        for k in cols:
            vo, vg = clients_o[nit][k], clients_g[nit][k]
            if vo is not None and vg is not None and abs(vo - vg) > TOL:
                nit_diffs.append((nit, k, vo, vg, vg - vo))
    nit_diffs.sort(key=lambda x: abs(x[4]), reverse=True)
    return diffs, nit_diffs[:20], len(clients_o), len(clients_g)


def compare_unidad(wb_o, wb_g):
    ws_o, ws_g = wb_o["unidad"], wb_g["unidad"]
    tr_o = find_total_row(ws_o, labels=("GRAND TOTAL", "TOTAL"))
    tr_g = find_total_row(ws_g, labels=("GRAND TOTAL", "TOTAL"))
    cols = {"saldo": 3, "corriente": 4, "vencido": 5}
    diffs = []
    for k, c in cols.items():
        vo, vg = num(ws_o.cell(tr_o, c).value), num(ws_g.cell(tr_g, c).value)
        if vo is not None and vg is not None and abs(vo - vg) > TOL:
            diffs.append(("unidad TOTAL", k, vo, vg, vg - vo))

    units_o, units_g = {}, {}
    for r in range(3, tr_o):
        u = str(ws_o.cell(r, 2).value or "").strip()
        if u:
            units_o[u] = {k: num(ws_o.cell(r, c).value) for k, c in cols.items()}
    for r in range(3, tr_g):
        u = str(ws_g.cell(r, 2).value or "").strip()
        if u:
            units_g[u] = {k: num(ws_g.cell(r, c).value) for k, c in cols.items()}

    for u in set(units_o) | set(units_g):
        if u not in units_o:
            diffs.append(("unidad fila", f"falta en original: {u}", None, units_g[u]["saldo"], None))
        elif u not in units_g:
            diffs.append(("unidad fila", f"falta en generado: {u}", units_o[u]["saldo"], None, None))
        else:
            for k in cols:
                vo, vg = units_o[u][k], units_g[u][k]
                if vo is not None and vg is not None and abs(vo - vg) > TOL:
                    diffs.append((f"unidad {u}", k, vo, vg, vg - vo))
    return diffs


def compare_unit_sheet(wb_o, wb_g, sheet):
    if sheet not in wb_o.sheetnames or sheet not in wb_g.sheetnames:
        return [("hoja", sheet, "no existe en uno de los archivos", None, None)]
    ws_o, ws_g = wb_o[sheet], wb_g[sheet]

    def total_row(ws):
        for r in range(1, 120):
            v = ws.cell(r, 3).value
            if v and str(v).strip().upper() == "TOTAL":
                return r
        return None

    tr_o, tr_g = total_row(ws_o), total_row(ws_g)
    cols = {"saldo": 4, "corriente": 5, "vencido": 6}
    diffs = []
    if tr_o and tr_g:
        for k, c in cols.items():
            vo, vg = num(ws_o.cell(tr_o, c).value), num(ws_g.cell(tr_g, c).value)
            if vo is not None and vg is not None and abs(vo - vg) > TOL:
                diffs.append((f"{sheet} TOTAL", k, vo, vg, vg - vo))
    return diffs


def compare_inf_edades(wb_o, wb_g):
    ws_o = wb_o["inf cartera x edades"]
    ws_g = wb_g["inf cartera x edades"]
    col_o, _ = find_month_col(ws_o, "JULIO")
    col_g, _ = find_month_col(ws_g, "JULIO")
    if not col_g:
        col_g, _ = find_month_col(ws_g, "JUL")
    diffs = []
    if col_o is None:
        diffs.append(("inf edades", "columna JULIO original", "no encontrada", None, None))
    if col_g is None:
        diffs.append(("inf edades", "columna JULIO generado", "no encontrada", None, None))
    if not col_o or not col_g:
        return diffs, col_o, col_g

    rows = {
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
        "total_sin": 25,
        "corriente_sin": 26,
        "vencida_sin": 27,
    }
    for k, r in rows.items():
        vo, vg = num(ws_o.cell(r, col_o).value), num(ws_g.cell(r, col_g).value)
        if vo is not None and vg is not None and abs(vo - vg) > TOL:
            diffs.append((f"inf edades {k}", f"col {col_o} vs {col_g}", vo, vg, vg - vo))
    return diffs, col_o, col_g


def compare_record(wb_o, wb_g):
    ws_o = wb_o["Record cvda"]
    ws_g = wb_g["Record cvda"]
    diffs = []
    for r in range(1, 90):
        mo = str(ws_o.cell(r, 2).value or "").strip().upper()
        mg = str(ws_g.cell(r, 2).value or "").strip().upper()
        if mo == "JUL" and 30 <= r <= 45:
            for c, name in [(3, "vencida"), (4, "pct_v"), (5, "corriente"), (6, "pct_c"), (7, "total")]:
                vo, vg = num(ws_o.cell(r, c).value), num(ws_g.cell(r, c).value)
                if vo is not None and vg is not None and abs(vo - vg) > (0.0001 if c in (4, 6) else TOL):
                    diffs.append((f"Record cvda JUL row {r}", name, vo, vg, vg - vo))
        if mo == "JUL" and r >= 70:
            for c, name in [(3, "vencida_sin"), (4, "pct_v_sin"), (5, "corriente_sin"), (7, "total_sin")]:
                vo, vg = num(ws_o.cell(r, c).value), num(ws_g.cell(r, c).value)
                if vo is not None and vg is not None and abs(vo - vg) > (0.0001 if c == 4 else TOL):
                    diffs.append((f"Record cvda sin OL JUL row {r}", name, vo, vg, vg - vo))
    return diffs


def parse_siimed_totals():
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from cartera_auto.parsers import parse_cartera_file, parse_unidad_file, parse_balance_file
    from cartera_auto.transform import build_report_model

    cartera = parse_cartera_file(SRC_CARTERA)
    unidad = parse_unidad_file(SRC_UNIDAD)
    balance = parse_balance_file(SRC_BALANCE)
    model = build_report_model(cartera, unidad, balance)
    return model, cartera, unidad


def main():
    wb_o = openpyxl.load_workbook(ORIG, data_only=True)
    wb_g = openpyxl.load_workbook(GEN, data_only=True)

    print("=" * 70)
    print("COMPARACIÓN JULIO 2026: ORIGINAL vs GENERADO")
    print("=" * 70)

    d_cons, nit_top, n_o, n_g = compare_consolidada(wb_o, wb_g)
    print(f"\n[1] CONSOLIDADA — clientes orig={n_o} gen={n_g}")
    for d in d_cons:
        print(f"  {d[0]} | {d[1]}: orig={d[2]:,.2f} gen={d[3]:,.2f} delta={d[4]:,.2f}" if d[3] is not None else f"  {d}")
    if nit_top:
        print("  Top diferencias por cliente:")
        for nit, k, vo, vg, delta in nit_top[:10]:
            print(f"    NIT {nit} {k}: orig={vo:,.2f} gen={vg:,.2f} delta={delta:,.2f}")

    d_uni = compare_unidad(wb_o, wb_g)
    print(f"\n[2] UNIDAD — {len(d_uni)} diferencias")
    for d in d_uni:
        if d[3] is not None and d[2] is not None:
            print(f"  {d[0]} | {d[1]}: orig={d[2]:,.2f} gen={d[3]:,.2f} delta={d[4]:,.2f}")
        else:
            print(f"  {d}")

    unit_sheets = ["carnes", "subproductos", "otras cxc", "local", "regional", "nacional", "institucional", "desposte"]
    print("\n[3] HOJAS POR UNIDAD")
    for sh in unit_sheets:
        ds = compare_unit_sheet(wb_o, wb_g, sh)
        if ds:
            print(f"  {sh}:")
            for d in ds:
                if d[3] is not None:
                    print(f"    {d[1]}: orig={d[2]:,.2f} gen={d[3]:,.2f} delta={d[4]:,.2f}")

    d_ed, col_o, col_g = compare_inf_edades(wb_o, wb_g)
    print(f"\n[4] INF CARTERA X EDADES — col orig={col_o} gen={col_g}")
    for d in d_ed:
        print(f"  {d[0]}: orig={d[2]:,.2f} gen={d[3]:,.2f} delta={d[4]:,.2f}")

    d_rec = compare_record(wb_o, wb_g)
    print(f"\n[5] RECORD CVDA — {len(d_rec)} diferencias JUL")
    for d in d_rec:
        print(f"  {d[0]} {d[1]}: orig={d[2]} gen={d[3]} delta={d[4]}")

    # SIIMED vs original vs model
    print("\n[6] SIIMED / MODELO vs ORIGINAL (totales)")
    model, cartera, unidad = parse_siimed_totals()
    k = model["kpis"]
    tot = model["totales"]
    ws_o = wb_o["consolidada julio"]
    tr_o = find_total_row(ws_o)
    orig_total = num(ws_o.cell(tr_o, 5).value)
    orig_venc = num(ws_o.cell(tr_o, 7).value)
    print(f"  SIIMED modelo total:     {k['total_cartera']:,.2f}")
    print(f"  SIIMED modelo vencida:   {k['cartera_vencida']:,.2f}")
    print(f"  Original total:          {orig_total:,.2f}")
    print(f"  Original vencida:        {orig_venc:,.2f}")
    print(f"  Generado total:          {num(wb_g['consolidada junio'].cell(find_total_row(wb_g['consolidada junio']), 5).value):,.2f}")
    print(f"  Clientes SIIMED cartera: {len(cartera['por_cliente'])}")
    print(f"  Clientes SIIMED unidad det: {len(unidad.get('detalle', []))}")
    print(f"  Unidades SIIMED: {list(unidad['por_unidad']['unidad']) if not unidad['por_unidad'].empty else []}")

    # TRABAJAR sheet
    if "TRABAJAR" in wb_o.sheetnames:
        ws = wb_o["TRABAJAR"]
        print("\n[7] HOJA TRABAJAR (original) — primeras filas con datos")
        for r in range(1, 15):
            row = [ws.cell(r, c).value for c in range(1, 8)]
            if any(row):
                print(f"  R{r}: {row}")

    wb_o.close()
    wb_g.close()


if __name__ == "__main__":
    main()
