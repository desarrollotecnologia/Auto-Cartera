# -*- coding: utf-8 -*-
"""
Actualiza data/maestro_clientes.json con lo que el gestor digitó en el informe manual.

- carnes: CONDICION DE PAGO y CUPO (si el manual los trae)
- local: PRINCIPAL / vendedor (si el manual lo trae)
- consolidada: PLAZO solo para clientes que aún no lo tienen
- comportamiento: PLAZOS (oficial; reemplaza el del maestro)
Uso: python tools/aprender_maestro.py "1. Informe de Cartera septiembre V1.xlsx"
"""
from __future__ import annotations

import io
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import load_workbook  # noqa: E402

from cartera_auto.client_master import MASTER_PATH, load_maestro, norm_nit, save_maestro  # noqa: E402


def _sheet(wb, prefix: str):
    return next((wb[n] for n in wb.sheetnames if n.lower().startswith(prefix)), None)


def main(path: str) -> None:
    wb = load_workbook(path, data_only=True)
    maestro = load_maestro()
    changes: list[str] = []

    cons = _sheet(wb, "consolidada")
    name_to_nit: dict[str, str] = {}
    for r in range(8, 200):
        nit = norm_nit(cons.cell(r, 2).value)
        name = str(cons.cell(r, 3).value or "").strip()
        if not nit or not name:
            continue
        name_to_nit[name.upper()] = nit
        entry = maestro.setdefault(nit, {"nombre": name})
        plazo = cons.cell(r, 4).value
        if plazo not in (None, "") and entry.get("plazo") in (None, ""):
            entry["plazo"] = plazo
            changes.append(f"plazo {name}: {plazo}")

    comp = _sheet(wb, "comportamiento")
    if comp is not None:
        plazo_col = next(
            (c for c in range(3, 40) if str(comp.cell(4, c).value or "").strip().upper() == "PLAZOS"),
            None,
        )
        for r in range(5, 200) if plazo_col else ():
            nit = norm_nit(comp.cell(r, 3).value)
            plazo = comp.cell(r, plazo_col).value
            if not nit or plazo in (None, ""):
                continue
            entry = maestro.setdefault(nit, {"nombre": comp.cell(r, 4).value})
            if entry.get("plazo") != plazo:
                changes.append(f"plazo (comportamiento) {entry.get('nombre')}: {entry.get('plazo')!r} -> {plazo!r}")
                entry["plazo"] = plazo

    def by_name(name: str) -> dict | None:
        nit = name_to_nit.get(name.strip().upper())
        if nit:
            return maestro.setdefault(nit, {"nombre": name})
        for entry in maestro.values():
            if str(entry.get("nombre", "")).strip().upper() == name.strip().upper():
                return entry
        return None

    carnes = wb["carnes"]
    for r in range(3, 60):
        name = str(carnes.cell(r, 3).value or "").strip()
        if not name or name.upper() in ("NOMBRE", "TOTAL"):
            continue
        entry = by_name(name)
        if entry is None:
            continue
        for field, col in (("condicion_pago", 4), ("cupo", 5)):
            value = carnes.cell(r, col).value
            if value not in (None, "") and entry.get(field) != value:
                changes.append(f"{field} {name}: {entry.get(field)!r} -> {value!r}")
                entry[field] = value

    local = wb["local"]
    for r in range(3, 60):
        name = str(local.cell(r, 3).value or "").strip()
        if not name or name.upper() in ("NOMBRE", "TOTAL"):
            continue
        entry = by_name(name)
        value = local.cell(r, 4).value
        if entry is not None and value not in (None, "") and entry.get("principal") != value:
            changes.append(f"principal {name}: {entry.get('principal')!r} -> {value!r}")
            entry["principal"] = value

    if changes:
        backup = MASTER_PATH.with_name(f"maestro_clientes_backup_{datetime.now():%Y%m%d_%H%M%S}.json")
        shutil.copy2(MASTER_PATH, backup)
        save_maestro(maestro)
    for c in changes:
        print(c)
    print("cambios:", len(changes))


if __name__ == "__main__":
    main(sys.argv[1])
