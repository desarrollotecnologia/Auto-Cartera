# -*- coding: utf-8 -*-
"""Validación Julio corregido vs original."""
import sys, io
sys.path.insert(0, r"c:\Users\CAMPUSLANDS\Desktop\presentacion")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import openpyxl
from pathlib import Path

ORIG = Path(r"c:\Users\CAMPUSLANDS\Downloads\1. Informe de Cartera julio V1.xlsx")
GEN = Path(r"c:\Users\CAMPUSLANDS\Desktop\presentacion\output\JULIO_CORREGIDO\Informe_Cartera_2026-07-31_20260818_1418.xlsx")
TOL = 1.0


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def main():
    wb_o = openpyxl.load_workbook(ORIG, data_only=True)
    wb_g = openpyxl.load_workbook(GEN, data_only=True)
    wb_gf = openpyxl.load_workbook(GEN, data_only=False)

    print("=== HOJAS ===")
    print("Gen:", [s for s in wb_g.sheetnames if "consolid" in s.lower() or "comport" in s.lower()])

    # Totales
    ws_o = wb_o["consolidada julio"]
    ws_g = wb_g["consolidada julio"]
    def find_total(ws):
        for r in range(8, 120):
            v = ws.cell(r, 2).value
            if v and str(v).strip().upper() == "TOTAL":
                return r
        return 75

    tr_o = find_total(ws_o)
    tr_g = find_total(ws_g)
    print(f"\nConsolidada TOTAL orig={num(ws_o.cell(tr_o, 5).value):,.2f} gen={num(ws_g.cell(tr_g, 5).value):,.2f}")

    # PLAZO
    missing = 0
    for r in range(8, tr_o):
        if ws_o.cell(r, 4).value and not ws_g.cell(r, 4).value:
            missing += 1
    print(f"PLAZO faltantes en gen: {missing}")

    # Comportamiento JULIO
    ws_o = wb_o["comportamiento julio"]
    ws_g = wb_g["comportamiento julio"]
    difs = 0
    for r in range(5, 73):
        vo = num(ws_o.cell(r, 11).value)
        vg = num(ws_g.cell(r, 11).value)
        if vo is not None and vg is not None and abs(vo - vg) > TOL:
            difs += 1
    print(f"\nComportamiento JULIO difs: {difs} (title gen: {ws_g.cell(2,2).value})")
    print(f"  Olímpica orig={num(ws_o.cell(5,11).value):,.2f} gen={num(ws_g.cell(5,11).value):,.2f}")

    # informe jf JUL col 24
    ws_o = wb_o["informe jf"]
    ws_gf = wb_gf["informe jf"]
    print(f"\ninforme jf C24 header gen: {ws_gf.cell(4,24).value}")
    print(f"  R7 SUBPROD orig={num(ws_o.cell(7,24).value):,.2f} formula gen={ws_gf.cell(7,24).value}")

    # PERIODO JUL
    ws_o = wb_o["PERIODO DE  RECAUDO"]
    ws_gf = wb_gf["PERIODO DE  RECAUDO"]
    print(f"\nPERIODO JUL orig={ws_o.cell(10,31).value} gen formula={ws_gf.cell(10,31).value}")

    # carnes count
    ws_o = wb_o["carnes"]
    ws_g = wb_g["carnes"]
    co = sum(1 for r in range(3, 50) if ws_o.cell(r, 3).value and "TOTAL" not in str(ws_o.cell(r, 3).value).upper())
    cg = sum(1 for r in range(3, 50) if ws_g.cell(r, 3).value and str(ws_g.cell(r, 3).value).upper() != "TOTAL")
    print(f"\nCarnes clientes orig={co} gen={cg}")

    # unidad label
    print(f"Unidad OTRAS orig={wb_o['unidad'].cell(4,2).value} gen={wb_g['unidad'].cell(4,2).value}")

    wb_o.close()
    wb_g.close()
    wb_gf.close()


if __name__ == "__main__":
    main()
