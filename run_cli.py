# -*- coding: utf-8 -*-
"""CLI: python run_cli.py --cartera archivo.xlsx [--unidad ...] [--balance ...]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from cartera_auto.pipeline import run_pipeline
from cartera_auto.parsers import format_corte_label
from cartera_auto.transform import money_mm, pct


def main():
    ap = argparse.ArgumentParser(description="Genera Excel + PPT de Cartera desde SIIMED")
    ap.add_argument("--cartera", required=True, help="Excel cartera por edades")
    ap.add_argument("--unidad", default=None, help="Excel cartera por unidad")
    ap.add_argument(
        "--balance",
        nargs="+",
        default=None,
        help="Excel balance (uno solo, o CTA 13 y CTA 4 por separado)",
    )
    ap.add_argument("--out", default=None, help="Carpeta de salida")
    args = ap.parse_args()

    result = run_pipeline(
        cartera_path=args.cartera,
        unidad_path=args.unidad,
        balance_path=args.balance,
        output_dir=args.out or (ROOT / "output"),
    )
    k = result["kpis"]
    print("OK")
    print("Corte:", format_corte_label(result["meta"].get("corte")))
    print("Total cartera:", money_mm(k["total_cartera"]))
    print("Vencida:", money_mm(k["cartera_vencida"]), pct(k["pct_vencida"]))
    print("Sin Olímpica:", money_mm(k["total_sin_olimpica"]))
    print("Excel:", result["excel"])
    print("PPTX:", result["pptx"])
    hist = result.get("historico") or {}
    if hist.get("updated"):
        print("Histórico:", hist.get("mes_actual"), "col", hist.get("column"))
        print("Series record:", len(hist.get("record_series") or []))
    else:
        print("Histórico: no actualizado", hist.get("reason"))
    validation = result.get("validation") or {}
    if validation.get("ok"):
        print(
            "Validación: OK -",
            validation.get("charts", 0),
            "gráficas /",
            validation.get("drawings", 0),
            "drawings",
        )
    for warning in result.get("warnings") or []:
        print("ADVERTENCIA:", warning)


if __name__ == "__main__":
    main()
