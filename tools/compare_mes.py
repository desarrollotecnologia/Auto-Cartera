# -*- coding: utf-8 -*-
"""Compara el informe generado (recalculado en Excel) contra el manual del mes."""
import io
import sys
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import openpyxl
from openpyxl.utils import get_column_letter as L

ORIG = sys.argv[1]
GEN = sys.argv[2]
TOL = 1.0
SKIP = {"DATOS", "TRABAJAR", "2023"}


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def match_sheet(name, names):
    if name in names:
        return name
    pref = name.split(" ")[0].lower()
    cands = [n for n in names if n.lower().startswith(pref)]
    return cands[0] if len(cands) == 1 else None


def main():
    vo = openpyxl.load_workbook(ORIG, data_only=True)
    vg = openpyxl.load_workbook(GEN, data_only=True)
    fo = openpyxl.load_workbook(ORIG)
    fg = openpyxl.load_workbook(GEN)

    for sn in vo.sheetnames:
        if sn in SKIP:
            continue
        gn = match_sheet(sn, vg.sheetnames)
        print("\n" + "=" * 90)
        print(f"{sn}  vs  {gn}")
        if gn is None:
            print("  HOJA FALTANTE")
            continue
        a, b, af, bf = vo[sn], vg[gn], fo[sn], fg[gn]
        max_r = min(max(a.max_row, b.max_row), 220)
        max_c = min(max(a.max_column, b.max_column), 70)
        nums, texts, forms, fmts = [], [], [], []
        for r in range(1, max_r + 1):
            if a.row_dimensions[r].hidden and b.row_dimensions[r].hidden:
                hidden_both = True
            else:
                hidden_both = False
            for c in range(1, max_c + 1):
                x, y = a.cell(r, c).value, b.cell(r, c).value
                nx, ny = num(x), num(y)
                coord = f"{L(c)}{r}"
                if nx is not None or ny is not None:
                    if (nx or 0) != (ny or 0) and abs((nx or 0) - (ny or 0)) > TOL:
                        if nx is None and ny is None:
                            pass
                        else:
                            nums.append((coord, x, y, hidden_both))
                else:
                    sx, sy = str(x or "").strip(), str(y or "").strip()
                    if sx != sy:
                        texts.append((coord, x, y, hidden_both))
                fx, fy = af.cell(r, c).value, bf.cell(r, c).value
                isfx = isinstance(fx, str) and fx.startswith("=")
                isfy = isinstance(fy, str) and fy.startswith("=")
                if (isfx or isfy) and fx != fy:
                    forms.append((coord, fx, fy))
                if x not in (None, "") and af.cell(r, c).number_format != bf.cell(r, c).number_format:
                    fmts.append((coord, af.cell(r, c).number_format, bf.cell(r, c).number_format))
        hid_a = [r for r in range(1, max_r + 1) if a.row_dimensions[r].hidden]
        hid_b = [r for r in range(1, max_r + 1) if b.row_dimensions[r].hidden]
        hc_a = sorted(k for k, d in a.column_dimensions.items() if d.hidden)
        hc_b = sorted(k for k, d in b.column_dimensions.items() if d.hidden)
        vis_nums = [n for n in nums if not n[3]]
        print(f"  valores numéricos distintos: {len(nums)} (visibles {len(vis_nums)})")
        for n in vis_nums[:25]:
            d = (num(n[2]) or 0) - (num(n[1]) or 0)
            print(f"    {n[0]}: manual={n[1]!r} programa={n[2]!r} dif={d:,.2f}")
        vis_t = [t for t in texts if not t[3]]
        print(f"  textos distintos: {len(texts)} (visibles {len(vis_t)})")
        for t in vis_t[:15]:
            print(f"    {t[0]}: manual={str(t[1])[:45]!r} programa={str(t[2])[:45]!r}")
        print(f"  fórmulas distintas: {len(forms)}")
        for f in forms[:12]:
            print(f"    {f[0]}: manual={str(f[1])[:55]!r} programa={str(f[2])[:55]!r}")
        print(f"  formatos de número distintos: {len(fmts)}")
        grp = defaultdict(list)
        for f in fmts:
            grp[(f[1], f[2])].append(f[0])
        for (m, p), cells in list(grp.items())[:6]:
            print(f"    manual={m!r} programa={p!r} en {len(cells)} celdas ej {cells[:4]}")
        if hid_a != hid_b:
            print(f"  filas ocultas: manual={hid_a[:40]} programa={hid_b[:40]}")
        if hc_a != hc_b:
            print(f"  columnas ocultas: manual={hc_a} programa={hc_b}")
        ca, cb = af._charts, bf._charts
        if len(ca) != len(cb):
            print(f"  gráficas: manual={len(ca)} programa={len(cb)}")
        for i, (x, y) in enumerate(zip(ca, cb)):
            sx = [(s.val.numRef.f if s.val and s.val.numRef else None) for s in x.series]
            sy = [(s.val.numRef.f if s.val and s.val.numRef else None) for s in y.series]
            if sx != sy:
                print(f"  gráfica {i+1} series: manual={sx} programa={sy}")


if __name__ == "__main__":
    main()
