# -*- coding: utf-8 -*-
import sys, io
sys.path.insert(0, r"c:\Users\CAMPUSLANDS\Desktop\presentacion")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from cartera_auto.parsers import parse_cartera_file, parse_unidad_file

c = parse_cartera_file(r"c:\Users\CAMPUSLANDS\Downloads\Cartera a corte 31 julio.xlsx")
u = parse_unidad_file(r"c:\Users\CAMPUSLANDS\Downloads\Cartera por cuenta 31 julio.xlsx")

names = ["DISTRIBUIDORA EL SEÑOR", "UMBARILA MALAGON"]
for n in names:
    print(f"\n=== {n} ===")
    df = c["detalle"]
    m = df[df["nombre"].str.contains(n.split()[0], case=False, na=False)]
    if m.empty:
        print("  NOT in cartera detalle")
    else:
        print(m[["nit", "nombre", "cuenta", "unidad", "saldo"]].to_string())

    m2 = c["por_cliente"]
    m2 = m2[m2["nombre"].str.contains(n.split()[0], case=False, na=False)]
    if not m2.empty:
        print("por_cliente:", m2[["nit", "nombre", "saldo"]].to_string())

    ud = u["detalle"]
    if not ud.empty:
        m3 = ud[ud["nombre"].str.contains(n.split()[0], case=False, na=False)]
        if not m3.empty:
            print("unidad detalle:", m3[["nit", "nombre", "cuenta", "unidad", "saldo"]].to_string())

# carnes unit clients count
from cartera_auto.transform import build_report_model
model = build_report_model(c, u, None)
carnes = model["detalle_por_unidad"].get("CARNES COLBEEF")
print(f"\nCARNES COLBEEF in model: {len(carnes)} clients, sum={carnes['saldo'].sum():,.2f}")

# DIFICIL COBRO
dc = model["detalle_por_unidad"].get("DIFICIL COBRO")
if dc is not None and not dc.empty:
    print(f"DIFICIL COBRO: {len(dc)} clients, sum={dc['saldo'].sum():,.2f}")
else:
    ud = u["detalle"]
    dc2 = ud[ud["unidad"] == "DIFICIL COBRO"] if not ud.empty else None
    if dc2 is not None and not dc2.empty:
        print(f"DIFICIL COBRO in unidad detalle: {len(dc2)} rows, sum={dc2['saldo'].sum():,.2f}")
