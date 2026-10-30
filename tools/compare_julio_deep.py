# -*- coding: utf-8 -*-
"""Comparación exhaustiva hoja por hoja - Julio 2026."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import openpyxl
from openpyxl.utils import get_column_letter

ORIG = r"c:\Users\CAMPUSLANDS\Downloads\1. Informe de Cartera julio V1.xlsx"
GEN = r"c:\Users\CAMPUSLANDS\Downloads\Informe_Cartera_2026-07-31_20260813_0803.xlsx"
TOL = 1.0

SHEET_MAP = {
    "consolidada julio": "consolidada junio",
    "comportamiento julio": "comportamiento junio",
}


def num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",", ""))
    except: return None


def scan_sheet(ws_o, ws_g, name_o, name_g):
    """Compare all numeric cells in used range."""
    diffs = []
    max_r = max(ws_o.max_row or 1, ws_g.max_row or 1)
    max_c = max(ws_o.max_column or 1, ws_g.max_column or 1)
    max_r = min(max_r, 200)
    max_c = min(max_c, 30)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            vo, vg = ws_o.cell(r, c).value, ws_g.cell(r, c).value
            no, ng = num(vo), num(vg)
            if no is not None and ng is not None and abs(no - ng) > TOL:
                diffs.append((r, c, vo, vg, ng - no))
            elif (vo != vg) and not (no is not None and ng is not None):
                # text diff (ignore None vs empty)
                so = str(vo or "").strip()
                sg = str(vg or "").strip()
                if so != sg and (so or sg):
                    # skip formula strings in gen if orig has value
                    if isinstance(vg, str) and vg.startswith("="):
                        continue
                    if isinstance(vo, str) and vo.startswith("="):
                        continue
                    diffs.append((r, c, vo, vg, "TEXT"))
    return diffs


def main():
    wb_o = openpyxl.load_workbook(ORIG, data_only=True)
    wb_g = openpyxl.load_workbook(GEN, data_only=True)
    wb_o_f = openpyxl.load_workbook(ORIG, data_only=False)
    wb_g_f = openpyxl.load_workbook(GEN, data_only=False)

    skip = {"AUTO_RESUMEN", "DATOS", "TRABAJAR", "Formula", "2023"}
    all_sheets_o = [s for s in wb_o.sheetnames if s not in skip]

    print("=" * 80)
    print("ESCANEO EXHAUSTIVO POR HOJA")
    print("=" * 80)

    summary = []
    for sn_o in all_sheets_o:
        sn_g = SHEET_MAP.get(sn_o, sn_o)
        if sn_g not in wb_g.sheetnames:
            summary.append((sn_o, "HOJA FALTANTE EN GENERADO", 0, []))
            continue
        diffs = scan_sheet(wb_o[sn_o], wb_g[sn_g], sn_o, sn_g)
        num_diffs = [d for d in diffs if d[4] != "TEXT"]
        text_diffs = [d for d in diffs if d[4] == "TEXT"]
        if num_diffs or text_diffs:
            summary.append((sn_o, sn_g, len(num_diffs), num_diffs[:15]))
            print(f"\n--- {sn_o} vs {sn_g} ---")
            print(f"  Diferencias numéricas: {len(num_diffs)}, texto: {len(text_diffs)}")
            for d in num_diffs[:12]:
                print(f"  R{d[0]}C{d[1]}({get_column_letter(d[1])}{d[0]}): orig={d[2]} gen={d[3]} delta={d[4]:,.2f}")
            for d in text_diffs[:8]:
                print(f"  TEXT R{d[0]}C{d[1]}: orig={d[2]!r} gen={d[3]!r}")

    print("\n" + "=" * 80)
    print("RESUMEN")
    for item in summary:
        print(f"  {item[0]}: {item[2]} dif numéricas")

    # Formula differences in key sheets
    print("\n" + "=" * 80)
    print("FÓRMULAS CLAVE (inf cartera x edades, Record cvda)")
    for sn in ["inf cartera x edades", "Record cvda", "informe jf", "PERIODO DE  RECAUDO"]:
        if sn not in wb_o_f.sheetnames:
            continue
        ws_o, ws_g = wb_o_f[sn], wb_g_f[sn]
        print(f"\n--- {sn} ---")
        # JUL col = 10
        col = 10
        for r in [4,5,6,16,19,20,21,25,26,27,37]:
            fo = ws_o.cell(r, col).value
            fg = ws_g.cell(r, col).value
            if fo != fg:
                print(f"  R{r}C{col}: orig={fo!r} gen={fg!r}")
        # diff/var/part cols 25-27
        for c in [25,26,27]:
            for r in [4,6,16,25,27]:
                fo = ws_o.cell(r, c).value
                fg = ws_g.cell(r, c).value
                if fo != fg:
                    print(f"  R{r}C{c}: orig={fo!r} gen={fg!r}")

    # informe jf full scan
    print("\n--- informe jf numeric ---")
    if "informe jf" in wb_o.sheetnames:
        diffs = scan_sheet(wb_o["informe jf"], wb_g["informe jf"], "informe jf", "informe jf")
        nd = [d for d in diffs if d[4] != "TEXT"]
        print(f"  {len(nd)} dif num, {len(diffs)-len(nd)} text")
        for d in nd[:20]:
            print(f"  R{d[0]}C{d[1]}: orig={d[2]} gen={d[3]} delta={d[4]:,.2f}")

    # Unit sheet client-level
    print("\n--- DETALLE CLIENTES POR UNIDAD ---")
    for sh in ["carnes", "local", "regional", "subproductos"]:
        if sh not in wb_o.sheetnames:
            continue
        ws_o, ws_g = wb_o[sh], wb_g[sh]
        # find totals
        def clients(ws):
            out = {}
            for r in range(1, 100):
                name = ws.cell(r, 3).value
                saldo = num(ws.cell(r, 4).value)
                if name and saldo and str(name).upper() != "TOTAL" and str(name).upper() != "NOMBRE":
                    out[str(name).strip()] = saldo
            return out
        co, cg = clients(ws_o), clients(ws_g)
        saldo_diffs = []
        for n in set(co) | set(cg):
            if n not in co:
                saldo_diffs.append((n, None, cg[n], "solo gen"))
            elif n not in cg:
                saldo_diffs.append((n, co[n], None, "solo orig"))
            elif abs(co[n] - cg[n]) > TOL:
                saldo_diffs.append((n, co[n], cg[n], cg[n]-co[n]))
        if saldo_diffs:
            print(f"  {sh}: {len(saldo_diffs)} difs")
            for d in sorted(saldo_diffs, key=lambda x: abs(x[3] if isinstance(x[3], float) else 0), reverse=True)[:5]:
                print(f"    {d[0][:40]}: orig={d[1]} gen={d[2]} ({d[3]})")

    wb_o.close(); wb_g.close(); wb_o_f.close(); wb_g_f.close()


if __name__ == "__main__":
    main()
