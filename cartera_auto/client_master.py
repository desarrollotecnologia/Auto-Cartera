# -*- coding: utf-8 -*-
"""Maestro de clientes (PLAZO, CUPO, condición de pago, principal)."""

from __future__ import annotations

import json
import re
from pathlib import Path

from .config import ROOT

MASTER_PATH = ROOT / "data" / "maestro_clientes.json"
LEGACY_PLAZO_PATH = ROOT / "data" / "plazos_nit.json"

# NITs con ajuste 361+ opcional (legacy; resumen unidad usa E3+E11 / E10)
UNIDAD_VENCIDA_EXTRA = {
    "91244390": "subproductos",
    "901406008": "carnes",
}


def norm_nit(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^\d]", "", s)


def load_maestro() -> dict[str, dict]:
    """Carga maestro NIT → {plazo, cupo, condicion_pago, principal, nombre}."""
    data: dict[str, dict] = {}
    if MASTER_PATH.exists():
        try:
            raw = json.loads(MASTER_PATH.read_text(encoding="utf-8"))
            for k, v in raw.items():
                nit = norm_nit(k)
                if nit and isinstance(v, dict):
                    data[nit] = dict(v)
        except Exception:
            pass
    # Compatibilidad con plazos_nit.json anterior
    if LEGACY_PLAZO_PATH.exists():
        try:
            legacy = json.loads(LEGACY_PLAZO_PATH.read_text(encoding="utf-8"))
            for k, plazo in legacy.items():
                nit = norm_nit(k)
                if nit:
                    data.setdefault(nit, {})
                    data[nit].setdefault("plazo", plazo)
        except Exception:
            pass
    return data


def save_maestro(data: dict[str, dict]) -> None:
    MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    MASTER_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def merge_from_run(model: dict, maestro: dict | None = None) -> dict[str, dict]:
    """Actualiza maestro con clientes del mes (conserva datos manuales existentes)."""
    maestro = dict(maestro or load_maestro())
    for _, row in model["por_cliente"].iterrows():
        nit = norm_nit(row["nit"])
        if not nit:
            continue
        entry = maestro.setdefault(nit, {})
        if not entry.get("nombre"):
            entry["nombre"] = row["nombre"]
    save_maestro(maestro)
    return maestro


def plazo_for(nit: str, maestro: dict[str, dict]) -> int | float | str | None:
    return maestro.get(norm_nit(nit), {}).get("plazo")


def carnes_extra_for(nombre: str, maestro: dict[str, dict]) -> dict:
    """Busca cupo y condición por nombre (hoja carnes)."""
    nom = (nombre or "").strip().upper()
    for entry in maestro.values():
        if str(entry.get("nombre", "")).strip().upper() == nom:
            return {
                "cupo": entry.get("cupo"),
                "condicion_pago": entry.get("condicion_pago"),
            }
    return {}


def principal_for(nombre: str, maestro: dict[str, dict]) -> str | None:
    nom = (nombre or "").strip().upper()
    for entry in maestro.values():
        if str(entry.get("nombre", "")).strip().upper() == nom:
            p = entry.get("principal")
            return str(p) if p else None
    return None


def vencida_extra_361(model: dict) -> tuple[float, float]:
    """Valores 361+ para ajuste resumen unidad (RUEDA, DISTRIBUIDORA)."""
    sub_extra = 0.0
    carnes_extra = 0.0
    df = model["por_cliente"]
    for nit, bucket in UNIDAD_VENCIDA_EXTRA.items():
        rows = df[df["nit"].astype(str) == nit]
        if rows.empty:
            continue
        val = float(rows.iloc[0].get("361+", 0) or 0)
        if bucket == "subproductos":
            sub_extra += val
        else:
            carnes_extra += val
    return sub_extra, carnes_extra
