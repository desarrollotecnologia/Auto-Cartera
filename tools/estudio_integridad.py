# -*- coding: utf-8 -*-
"""Cruza totales e integridad fila a fila de los exportes SIIMED (agosto vs septiembre)."""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
DL = Path(r"C:\Users\CAMPUSLANDS\Downloads")

BUCKETS = [
    "<== -361", "DE -360 A -181", "DE -180 A -121", "DE -120 A -91",
    "DE -90 A -61", "DE -60 A -31", "DE -30 A -16", "DE -15 A -8", "DE -7 A -1",
]


def cartera_general(path: Path) -> None:
    print(f"\n=== {path.name} — hoja DATOS (detalle crudo) ===")
    ws = load_workbook(path, data_only=False)["DATOS"]
    hdr = {str(ws.cell(7, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
    c_nit, c_cta, c_nom = 1, 3, 2
    c_saldo, c_venc, c_corr = hdr["SALDO DEL NIT"], hdr["VALOR VENCIDO"], hdr["0 ==>"]
    b_cols = [hdr[b] for b in BUCKETS]
    tot = defaultdict(float)
    por_cuenta = defaultdict(float)
    bad_rows, typed = [], []
    nits = set()
    dificil = []
    for r in range(8, ws.max_row + 1):
        nit = ws.cell(r, c_nit).value
        if not isinstance(nit, (int, float)):
            continue
        saldo = float(ws.cell(r, c_saldo).value or 0)
        venc = float(ws.cell(r, c_venc).value or 0)
        corr = float(ws.cell(r, c_corr).value or 0)
        bsum = sum(float(ws.cell(r, c).value or 0) for c in b_cols)
        cta = str(ws.cell(r, c_cta).value or "")
        nits.add(int(nit))
        tot["saldo"] += saldo
        tot["vencido"] += venc
        tot["corriente"] += corr
        por_cuenta[cta] += saldo
        if "DIFICIL" in cta.upper():
            dificil.append((int(nit), ws.cell(r, c_nom).value, saldo))
        if abs(venc - bsum) > 1 or abs(saldo - venc - corr) > 1:
            bad_rows.append((r, int(nit), saldo, venc, corr, bsum))
        for c in [c_saldo, c_venc, c_corr, *b_cols]:
            if isinstance(ws.cell(r, c).value, str):
                typed.append(ws.cell(r, c).coordinate)
    print(f"   clientes={len(nits)} saldo={tot['saldo']:,.2f} vencido={tot['vencido']:,.2f} corriente={tot['corriente']:,.2f}")
    print(f"   filas incoherentes (vencido≠suma edades o saldo≠venc+corr): {len(bad_rows)} {bad_rows[:5]}")
    print(f"   celdas numéricas guardadas como texto: {len(typed)} {typed[:5]}")
    print("   saldo por cuenta:")
    for k in sorted(por_cuenta):
        print(f"      {k:<45} {por_cuenta[k]:>20,.2f}")
    print(f"   difícil cobro: {dificil}")
    return tot, por_cuenta, nits


def por_cuenta_resumen(path: Path, sheet: str, label_col=1, saldo_col=None):
    ws = load_workbook(path, data_only=True)[sheet]
    hdr = {str(ws.cell(7, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
    sc = saldo_col or hdr["SALDO DEL NIT"]
    out = {}
    for r in range(8, ws.max_row + 1):
        lab = str(ws.cell(r, label_col).value or "")
        if lab.endswith("Total"):
            out[lab.replace(" Total", "")] = float(ws.cell(r, sc).value or 0)
    return out


if __name__ == "__main__":
    ago = cartera_general(DL / "cartera general 31 agosto 2026.xlsx")
    sep = cartera_general(DL / "CARTERA A 30 SEPT.xlsx")

    print("\n=== cartera por cuenta: totales por cuenta ===")
    pa = por_cuenta_resumen(DL / "cartera por cuenta agosto 2026.xlsx", "por unidad")
    ps = por_cuenta_resumen(DL / "cartera por cuenta 30 sept.xlsx", "DATOS")
    for k in sorted(set(pa) | set(ps)):
        print(f"   {k:<45} ago={pa.get(k, 0):>18,.2f}  sep={ps.get(k, 0):>18,.2f}")

    print("\n=== septiembre: por cuenta (archivo) vs cartera general (DATOS) ===")
    _, pc_sep, _ = sep
    for k, v in sorted(ps.items()):
        if k == "Grand":
            continue
        g = pc_sep.get(k, 0.0)
        flag = "OK" if abs(g - v) < 1 else f"DIF {v - g:,.2f}"
        print(f"   {k:<45} archivo={v:>18,.2f} general={g:>18,.2f} {flag}")

    print("\n=== clientes que salen / entran (agosto → septiembre) ===")
    ga = load_workbook(DL / "cartera general 31 agosto 2026.xlsx", data_only=True)["DATOS"]
    names = {}
    for r in range(8, ga.max_row + 1):
        if isinstance(ga.cell(r, 1).value, (int, float)):
            names[int(ga.cell(r, 1).value)] = ga.cell(r, 2).value
    salen = ago[2] - sep[2]
    entran = sep[2] - ago[2]
    print(f"   salen {len(salen)}: {[names.get(n) for n in list(salen)[:15]]}")
    print(f"   entran {len(entran)}")
