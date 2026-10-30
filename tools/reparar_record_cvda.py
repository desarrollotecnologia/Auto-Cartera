# -*- coding: utf-8 -*-
"""Restaura fórmulas internas de 'Record cvda' en la plantilla desde el original manual."""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from cartera_auto.historico import normalize_record_row_totals  # noqa: E402

SRC = Path.home() / "Downloads" / "1. Informe de Cartera agosto V1.xlsx"
TEMPLATE = ROOT / "templates" / "plantilla_informe_cartera.xlsx"


def main(template: Path = TEMPLATE, source: Path = SRC) -> None:
    ref = load_workbook(source)["Record cvda"]
    wb = load_workbook(template)
    ws = wb["Record cvda"]
    changed = []
    for row in ref.iter_rows():
        for cell in row:
            v1 = cell.value
            cur = ws.cell(cell.row, cell.column).value
            if not (isinstance(v1, str) and v1.startswith("=")) or "[" in v1:
                continue
            if isinstance(cur, str) and cur.startswith("=") and cur != v1:
                ws.cell(cell.row, cell.column).value = v1
                changed.append((cell.coordinate, cur, v1))
    normalize_record_row_totals(ws)
    wb.save(template)
    for coord, old, new in changed:
        print(f"{coord}: {old} -> {new}")
    print("celdas reparadas:", len(changed))


if __name__ == "__main__":
    main()
