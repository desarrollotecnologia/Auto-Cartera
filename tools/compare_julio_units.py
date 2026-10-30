# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import openpyxl

ORIG = r"c:\Users\CAMPUSLANDS\Downloads\1. Informe de Cartera julio V1.xlsx"
GEN = r"c:\Users\CAMPUSLANDS\Downloads\Informe_Cartera_2026-07-31_20260813_0803.xlsx"

def num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",", ""))
    except: return 0.0

wb_o = openpyxl.load_workbook(ORIG, data_only=True)
wb_g = openpyxl.load_workbook(GEN, data_only=True)

# CARNES
ws = wb_o["carnes"]
s = sum(num(ws.cell(r, 6).value) for r in range(3, 200) if ws.cell(r, 3).value and "TOTAL" not in str(ws.cell(r, 3).value).upper())
ws2 = wb_g["carnes"]
s2 = sum(num(ws2.cell(r, 4).value) for r in range(3, 200) if ws2.cell(r, 3).value and str(ws2.cell(r, 3).value).upper() != "TOTAL")
tot_g = 0
for r in range(3, 200):
    if str(ws2.cell(r, 3).value or "").upper() == "TOTAL":
        tot_g = num(ws2.cell(r, 4).value)
print(f"CARNES orig sum={s:,.2f} gen sum={s2:,.2f} gen TOTAL row={tot_g:,.2f}")

# UNIDAD
print("\nUNIDAD comparison:")
ws_o, ws_g = wb_o["unidad"], wb_g["unidad"]
for r in range(3, 13):
    uo, vo = ws_o.cell(r, 2).value, num(ws_o.cell(r, 3).value)
    ug, vg = ws_g.cell(r, 2).value, num(ws_g.cell(r, 3).value)
    if uo or ug:
        d = vg - vo
        print(f"  {str(uo):22} {vo:>18,.2f} | {str(ug):15} {vg:>18,.2f} | delta={d:,.2f}")

# Each unit sheet total vs unidad
units = {
    "subproductos": (5, 4),
    "local": (5, 4),
    "regional": (5, 4),
    "nacional": (5, 4),
    "institucional": (5, 4),
    "desposte": (5, 4),
    "otras cxc": (5, 4),
}
print("\nUnit sheet TOTAL vs unidad:")
unidad_map = {}
for r in range(3, 12):
    u = str(ws_o.cell(r, 2).value or "")
    unidad_map[u.upper()] = num(ws_o.cell(r, 3).value)

for sh, (saldo_col_orig, saldo_col_gen) in units.items():
    ws = wb_o[sh]
    # find total in orig - last row with data in saldo col
    orig_total = None
    for r in range(ws.max_row, 1, -1):
        v = ws.cell(r, saldo_col_orig).value
        if num(v) > 0 and ws.cell(r, 3).value and "TOTAL" in str(ws.cell(r, 3).value).upper():
            orig_total = num(v)
            break
    if orig_total is None:
        orig_total = sum(num(ws.cell(r, saldo_col_orig).value) for r in range(3, 200) if ws.cell(r, 3).value)

    ws = wb_g[sh]
    gen_total = None
    for r in range(3, 200):
        if str(ws.cell(r, 3).value or "").upper() == "TOTAL":
            gen_total = num(ws.cell(r, saldo_col_gen).value)
    print(f"  {sh:15} orig={orig_total:,.2f} gen={gen_total:,.2f}")

# Carnes vs unidad
carnes_u = unidad_map.get("CARNES COLBEEF", 0)
print(f"\nCARNES unidad sheet: {carnes_u:,.2f}")

# Client-level carnes comparison
print("\nCARNES client diffs (by name, saldo):")
co, cg = {}, {}
ws = wb_o["carnes"]
for r in range(3, 200):
    n = ws.cell(r, 3).value
    if n and "TOTAL" not in str(n).upper():
        co[str(n).strip()] = num(ws.cell(r, 6).value)
ws = wb_g["carnes"]
for r in range(3, 200):
    n = ws.cell(r, 3).value
    if n and str(n).upper() != "TOTAL":
        cg[str(n).strip()] = num(ws.cell(r, 4).value)
diffs = []
for n in set(co) | set(cg):
    if n not in co: diffs.append((n, 0, cg[n], cg[n]))
    elif n not in cg: diffs.append((n, co[n], 0, -co[n]))
    elif abs(co[n] - cg[n]) > 1: diffs.append((n, co[n], cg[n], cg[n] - co[n]))
print(f"  {len(diffs)} differences")
for d in sorted(diffs, key=lambda x: abs(x[3]), reverse=True)[:12]:
    print(f"    {d[0][:42]:42} orig={d[1]:,.2f} gen={d[2]:,.2f} delta={d[3]:,.2f}")

wb_o.close()
wb_g.close()
