# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import openpyxl

ORIG = r"c:\Users\CAMPUSLANDS\Downloads\1. Informe de Cartera julio V1.xlsx"
GEN = r"c:\Users\CAMPUSLANDS\Downloads\Informe_Cartera_2026-07-31_20260813_0803.xlsx"

def num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",", ""))
    except: return None

wb_o = openpyxl.load_workbook(ORIG, data_only=True)
wb_g = openpyxl.load_workbook(GEN, data_only=True)

# Unit totals
print("=== TOTALES HOJAS UNIDAD ===")
for sh in ["carnes", "subproductos", "otras cxc", "local", "regional", "nacional", "institucional", "desposte"]:
    ws_o, ws_g = wb_o[sh], wb_g[sh]
    tr_o = tr_g = None
    saldo_o = saldo_g = None
    for r in range(1, 120):
        if str(ws_o.cell(r, 3).value or "").upper() == "TOTAL":
            tr_o = r
            saldo_o = num(ws_o.cell(r, 5).value) or num(ws_o.cell(r, 4).value)
        if str(ws_g.cell(r, 3).value or "").upper() == "TOTAL":
            tr_g = r
            saldo_g = num(ws_g.cell(r, 4).value)
    delta = (saldo_g or 0) - (saldo_o or 0)
    print(f"  {sh:15} orig={saldo_o:,.2f} gen={saldo_g:,.2f} delta={delta:,.2f}")

# inf edades col 10
print("\n=== inf cartera x edades JULIO (col 10) ===")
ws_o = wb_o["inf cartera x edades"]
ws_g = wb_g["inf cartera x edades"]
for r, k in [(4, "total"), (5, "corriente"), (6, "vencida"), (25, "total sin OL"), (26, "corr sin"), (27, "venc sin")]:
    vo, vg = num(ws_o.cell(r, 10).value), num(ws_g.cell(r, 10).value)
    print(f"  {k}: orig={vo:,.2f} gen={vg:,.2f} delta={(vg or 0)-(vo or 0):,.2f}")

# Compare carnes clients by name
print("\n=== CARNES: clientes orig vs gen (saldo) ===")
ws_o, ws_g = wb_o["carnes"], wb_g["carnes"]
co, cg = {}, {}
for r in range(3, 100):
    n = ws_o.cell(r, 3).value
    s = num(ws_o.cell(r, 6).value)  # SALDO DEL NIT col F
    if n and s and str(n).upper() != "TOTAL":
        co[str(n).strip()] = s
for r in range(3, 100):
    n = ws_g.cell(r, 3).value
    s = num(ws_g.cell(r, 4).value)  # gen SALDO col D
    if n and s and str(n).upper() != "TOTAL":
        cg[str(n).strip()] = s

diffs = []
for n in set(co) | set(cg):
    if n not in co:
        diffs.append((n, None, cg[n], "solo gen"))
    elif n not in cg:
        diffs.append((n, co[n], None, "solo orig"))
    elif abs(co[n] - cg[n]) > 1:
        diffs.append((n, co[n], cg[n], cg[n] - co[n]))
print(f"  Diferencias: {len(diffs)}")
for d in sorted(diffs, key=lambda x: abs(x[3] if isinstance(x[3], float) else 0), reverse=True)[:15]:
    print(f"    {d[0][:45]}: orig={d[1]} gen={d[2]} ({d[3]})")

# subproductos by name
print("\n=== SUBPRODUCTOS: clientes ===")
ws_o, ws_g = wb_o["subproductos"], wb_g["subproductos"]
co, cg = {}, {}
for r in range(3, 50):
    n = ws_o.cell(r, 3).value
    s = num(ws_o.cell(r, 5).value)
    if n and s: co[str(n).strip()] = s
for r in range(3, 50):
    n = ws_g.cell(r, 3).value
    s = num(ws_g.cell(r, 4).value)
    if n and s: cg[str(n).strip()] = s
diffs = []
for n in set(co) | set(cg):
    if n not in co: diffs.append((n, None, cg[n]))
    elif n not in cg: diffs.append((n, co[n], None))
    elif abs(co[n]-cg[n])>1: diffs.append((n, co[n], cg[n]))
print(f"  Diferencias saldo: {len(diffs)}")
for d in diffs[:10]:
    print(f"    {d}")

# PERIODO RECAUDO
print("\n=== PERIODO RECAUDO ===")
ws_o = wb_o["PERIODO DE  RECAUDO"]
ws_g = wb_g["PERIODO DE  RECAUDO"]
for c in range(5, 35):
    h = ws_o.cell(4, c).value
    if h and str(h).strip().upper() in ("JUL", "JUN", "JULIO", "JUNIO"):
        vo, vg = num(ws_o.cell(10, c).value), num(ws_g.cell(10, c).value)
        print(f"  Col {c} ({h}): orig={vo} gen={vg}")

# comportamiento - check if JULIO column exists
print("\n=== COMPORTAMIENTO ===")
ws_o = wb_o["comportamiento julio"]
ws_g = wb_g["comportamiento junio"]
# find JULIO col in orig row 4
jul_col = None
for c in range(3, 20):
    if ws_o.cell(4, c).value and "JULIO" in str(ws_o.cell(4, c).value).upper():
        jul_col = c
        break
print(f"  JULIO col orig: {jul_col}")
if jul_col:
    # compare Olímpica row 5
    vo = num(ws_o.cell(5, jul_col).value)
    vg_col = None
    for c in range(3, 20):
        if ws_g.cell(4, c).value and "JULIO" in str(ws_g.cell(4, c).value).upper():
            vg_col = c
    print(f"  JULIO col gen: {vg_col}")
    if vg_col:
        vg = num(ws_g.cell(5, vg_col).value)
        print(f"  Olímpica JULIO: orig={vo:,.2f} gen={vg:,.2f}")

wb_o.close()
wb_g.close()
