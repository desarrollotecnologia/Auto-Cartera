# -*- coding: utf-8 -*-
"""Parsers de exportes SIIMED (Excel)."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook

from .config import ACCOUNT_UNIT_MAP, OLIMPICA_NIT


def _num(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _nit(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return str(v)
    s = str(v).strip()
    s = s.replace(",", "").replace(".0", "")
    # filas de total tipo "890,107,487  Total"
    if "total" in s.lower():
        return ""
    digits = re.sub(r"[^\d]", "", s)
    return digits


def _is_total_row(values: list[Any]) -> bool:
    """Detecta filas de total SIIMED sin confundir razones sociales (ej. TOTAL FOOD)."""
    parts = [str(v).strip() for v in values if v is not None and str(v).strip()]
    if not parts:
        return True
    first = parts[0].lower()
    # Totales de cuenta / grand total
    if first.endswith(" total") or first.endswith("total"):
        if first in ("total", "grand total") or first.endswith(" total"):
            return True
    if "grand total" in first:
        return True
    # "0 Total" / "NIT Total" en columnas de nit
    for p in parts[:4]:
        pl = p.lower()
        if pl.endswith(" total") and "food" not in pl:
            # "890,107,487  Total" o "0 Total"
            return True
    return False


def _account_code(cuenta: str) -> str:
    if not cuenta:
        return ""
    m = re.match(r"(\d+)", str(cuenta).strip())
    return m.group(1) if m else ""


def unit_from_account(cuenta: str) -> str:
    code = _account_code(cuenta)
    if code in ACCOUNT_UNIT_MAP:
        return ACCOUNT_UNIT_MAP[code]
    # probar prefijos largos primero
    for key in sorted(ACCOUNT_UNIT_MAP.keys(), key=len, reverse=True):
        if code.startswith(key):
            return ACCOUNT_UNIT_MAP[key]
    return "OTROS"


def parse_corte_from_text(text: str) -> str | None:
    if not text:
        return None
    # "Con corte a: Junio 30 Del 2026"
    m = re.search(
        r"corte a:\s*([A-Za-zÁÉÍÓÚáéíóú]+)\s+(\d{1,2})\s+Del\s+(\d{4})",
        text,
        re.I,
    )
    if not m:
        return None
    months = {
        "enero": 1,
        "febrero": 2,
        "marzo": 3,
        "abril": 4,
        "mayo": 5,
        "junio": 6,
        "julio": 7,
        "agosto": 8,
        "septiembre": 9,
        "octubre": 10,
        "noviembre": 11,
        "diciembre": 12,
    }
    mes = months.get(m.group(1).lower())
    if not mes:
        return None
    return f"{int(m.group(3)):04d}-{mes:02d}-{int(m.group(2)):02d}"


def _find_header_row(ws, keywords=("NIT", "NOMBRE", "SALDO")) -> int:
    for r in range(1, min(20, (ws.max_row or 1) + 1)):
        vals = [str(ws.cell(r, c).value or "").upper() for c in range(1, 20)]
        joined = "|".join(vals)
        if all(k.upper() in joined for k in keywords[:2]):
            return r
    return 7


def _map_siimed_age_headers(headers: list[str]) -> dict[str, int]:
    """Mapea encabezados SIIMED (días negativos) a buckets de presentación."""
    mapping: dict[str, int] = {}
    for idx, h in enumerate(headers):
        hu = (h or "").upper().replace(" ", "")
        if "SALDO" in hu and "NIT" in hu:
            mapping["SALDO"] = idx
        elif "VALORVENCIDO" in hu or hu == "VENCIDO":
            mapping["VENCIDO"] = idx
        elif "0==>" in hu or hu.startswith("0=="):
            mapping.setdefault("CORRIENTE", idx)
        elif "-7A-1" in hu or "DE-7A-1" in hu:
            mapping["1 AL 7"] = idx
        elif "-15A-8" in hu:
            mapping["8 AL 15"] = idx
        elif "-30A-16" in hu:
            mapping["16 AL 30"] = idx
        elif "-60A-31" in hu:
            mapping["31 AL 60"] = idx
        elif "-90A-61" in hu:
            mapping["61 AL 90"] = idx
        elif "-120A-91" in hu or "-180A-121" in hu:
            # se acumulan en 91-180
            key = "91 AL 180"
            mapping.setdefault(key, idx)
            if "-120A-91" in hu:
                mapping["91_120"] = idx
            if "-180A-121" in hu:
                mapping["121_180"] = idx
        elif "-360A-181" in hu:
            mapping["181 AL 360"] = idx
        elif "<=-361" in hu or "<==-361" in hu or "=-361" in hu:
            mapping["361+"] = idx
        # formato ya remapeado (hoja unidad)
        elif hu == "CORRIENTE":
            mapping.setdefault("CORRIENTE", idx)
        elif "1AL7" in hu:
            mapping["1 AL 7"] = idx
        elif "8AL15" in hu:
            mapping["8 AL 15"] = idx
        elif "16AL30" in hu:
            mapping["16 AL 30"] = idx
        elif "31AL60" in hu:
            mapping["31 AL 60"] = idx
        elif "61AL90" in hu:
            mapping["61 AL 90"] = idx
        elif "91AL180" in hu or hu == "91A180":
            mapping["91 AL 180"] = idx
        elif "181AL360" in hu:
            mapping["181 AL 360"] = idx
        elif "361" in hu:
            mapping["361+"] = idx
    return mapping


def parse_cartera_file(path: str | Path) -> dict:
    """Parsea 'cartera a XX.xlsx' (edades por cliente / unidad)."""
    path = Path(path)
    wb = load_workbook(path, data_only=True)

    meta = {"archivo": path.name, "corte": None, "procesado": None, "empresa": None}
    # meta desde primera hoja
    ws0 = wb[wb.sheetnames[0]]
    for r in range(1, 6):
        for c in range(1, 5):
            v = ws0.cell(r, c).value
            if not v:
                continue
            s = str(v)
            if "SIIMED" in s or "COLBEEF" in s:
                meta["empresa"] = s
            if "corte a" in s.lower():
                meta["corte"] = parse_corte_from_text(s)
            if s.lower().startswith("procesado"):
                meta["procesado"] = s

    # Preferir hoja 'unidad' si trae buckets ya remapeados
    detail_rows: list[dict] = []
    sheet_name = "unidad" if "unidad" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    header_row = _find_header_row(ws)
    headers = [str(ws.cell(header_row, c).value or "").strip() for c in range(1, 20)]
    hmap = {h.upper(): i for i, h in enumerate(headers) if h}
    age_map = _map_siimed_age_headers(headers)

    nit_i = next((hmap[k] for k in hmap if k == "NIT"), 0)
    nom_i = next((hmap[k] for k in hmap if "NOMBRE" in k), 1)
    cta_i = next((hmap[k] for k in hmap if k == "CUENTA"), None)

    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        vals = [ws.cell(r, c).value for c in range(1, 20)]
        if _is_total_row(vals):
            continue
        nit = _nit(vals[nit_i] if nit_i < len(vals) else None)
        if not nit:
            continue
        nombre = str(vals[nom_i] or "").strip()
        cuenta = str(vals[cta_i] or "").strip() if cta_i is not None else ""
        row = {
            "nit": nit,
            "nombre": nombre,
            "cuenta": cuenta,
            "unidad": unit_from_account(cuenta) if cuenta else "",
            "es_olimpica": nit == OLIMPICA_NIT,
        }
        # saldos
        if "SALDO" in age_map:
            row["saldo"] = _num(vals[age_map["SALDO"]])
        else:
            # col D típica en hoja unidad
            row["saldo"] = _num(vals[3] if len(vals) > 3 else 0)

        if "VENCIDO" in age_map:
            row["vencido"] = _num(vals[age_map["VENCIDO"]])
        else:
            row["vencido"] = 0.0

        for bucket in [
            "CORRIENTE",
            "1 AL 7",
            "8 AL 15",
            "16 AL 30",
            "31 AL 60",
            "61 AL 90",
            "91 AL 180",
            "181 AL 360",
            "361+",
        ]:
            if bucket == "91 AL 180" and ("91_120" in age_map or "121_180" in age_map):
                row[bucket] = _num(vals[age_map["91_120"]]) if "91_120" in age_map else 0.0
                row[bucket] += _num(vals[age_map["121_180"]]) if "121_180" in age_map else 0.0
            elif bucket in age_map:
                row[bucket] = _num(vals[age_map[bucket]])
            else:
                row[bucket] = 0.0

        # si no había CORRIENTE mapeado pero hay saldo, inferir
        if "CORRIENTE" not in age_map and row["saldo"]:
            overdue = sum(row.get(b, 0) for b in [
                "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
                "61 AL 90", "91 AL 180", "181 AL 360", "361+",
            ])
            row["CORRIENTE"] = max(row["saldo"] - overdue, 0)
            row["vencido"] = overdue if not row["vencido"] else row["vencido"]

        detail_rows.append(row)

    # Si unidad no tenía cuenta, intentar DATOS
    if detail_rows and not any(r["cuenta"] for r in detail_rows) and "DATOS" in wb.sheetnames:
        detail_rows = _parse_datos_sheet(wb["DATOS"])

    # Agregado por cliente (sumar cuentas)
    by_client = _aggregate_clients(detail_rows)

    # Preferir hoja 'cartera' / 'CARTERA' (totales por NIT del SIIMED)
    cartera_ws = _find_sheet(wb, "cartera")
    if cartera_ws is not None:
        sheet_clients = _parse_client_sheet(cartera_ws)
        if sheet_clients:
            # El gestor a veces deja la hoja CARTERA por NIT+cuenta: consolidar
            # por NIT y descartarla si no cuadra con el detalle crudo.
            sheet_clients = _aggregate_clients(sheet_clients)
            detail_total = sum(r.get("saldo", 0) for r in detail_rows)
            sheet_total = sum(r.get("saldo", 0) for r in sheet_clients)
            if not detail_rows or abs(detail_total - sheet_total) < 1:
                by_client = sheet_clients

    anticipos = _parse_anticipos(wb)

    wb.close()
    return {
        "meta": meta,
        "detalle": pd.DataFrame(detail_rows),
        "por_cliente": pd.DataFrame(by_client),
        "anticipos": anticipos,
    }


def _find_sheet(wb, name: str):
    """Busca hoja por nombre exacto o case-insensitive / prefijo."""
    target = name.strip().lower()
    for sn in wb.sheetnames:
        if sn.strip().lower() == target:
            return wb[sn]
    for sn in wb.sheetnames:
        if sn.strip().lower().startswith(target) and "anticipo" not in sn.lower():
            return wb[sn]
    return None


def _parse_anticipos(wb) -> dict[str, float]:
    """
    Lee hoja CARTERA -ANTICIPOS → {nit: monto_anticipo}.
    Los anticipos suelen ser negativos (reducen saldo).
    """
    ws = None
    for sn in wb.sheetnames:
        if "anticipo" in sn.lower():
            ws = wb[sn]
            break
    if ws is None:
        return {}
    header_row = _find_header_row(ws)
    headers = [str(ws.cell(header_row, c).value or "").strip().upper() for c in range(1, 12)]
    try:
        nit_i = headers.index("NIT")
    except ValueError:
        nit_i = 0
    ant_i = next((i for i, h in enumerate(headers) if "ANTICIPO" in h), 3)
    out: dict[str, float] = {}
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        nit = _nit(ws.cell(r, nit_i + 1).value)
        if not nit:
            continue
        ant = _num(ws.cell(r, ant_i + 1).value)
        if abs(ant) > 0.001:
            out[nit] = ant
    return out


def _parse_client_sheet(ws) -> list[dict]:
    """Parsea hoja agregada por cliente (cartera / Cartera sin Olimpica)."""
    header_row = _find_header_row(ws)
    headers = [str(ws.cell(header_row, c).value or "").strip() for c in range(1, 20)]
    age_map = _map_siimed_age_headers(headers)
    hmap = {h.upper(): i for i, h in enumerate(headers) if h}
    nit_i = next((hmap[k] for k in hmap if k == "NIT"), 0)
    nom_i = next((hmap[k] for k in hmap if "NOMBRE" in k), 1)
    rows = []
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        vals = [ws.cell(r, c).value for c in range(1, 20)]
        if _is_total_row(vals):
            continue
        nit = _nit(vals[nit_i] if nit_i < len(vals) else None)
        if not nit:
            continue
        row = {
            "nit": nit,
            "nombre": str(vals[nom_i] or "").strip(),
            "es_olimpica": nit == OLIMPICA_NIT,
            "saldo": _num(vals[age_map["SALDO"]]) if "SALDO" in age_map else 0.0,
            "vencido": _num(vals[age_map["VENCIDO"]]) if "VENCIDO" in age_map else 0.0,
        }
        for bucket in [
            "CORRIENTE", "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
            "61 AL 90", "91 AL 180", "181 AL 360", "361+",
        ]:
            if bucket == "91 AL 180" and ("91_120" in age_map or "121_180" in age_map):
                row[bucket] = (_num(vals[age_map["91_120"]]) if "91_120" in age_map else 0.0) + (
                    _num(vals[age_map["121_180"]]) if "121_180" in age_map else 0.0
                )
            else:
                row[bucket] = _num(vals[age_map[bucket]]) if bucket in age_map else 0.0
        if not row.get("CORRIENTE"):
            overdue = sum(row.get(b, 0) for b in [
                "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
                "61 AL 90", "91 AL 180", "181 AL 360", "361+",
            ])
            row["CORRIENTE"] = max(row["saldo"] - overdue, 0)
            if not row["vencido"]:
                row["vencido"] = overdue
        rows.append(row)
    rows.sort(key=lambda x: x["saldo"], reverse=True)
    return rows


def _parse_datos_sheet(ws) -> list[dict]:
    header_row = _find_header_row(ws)
    headers = [str(ws.cell(header_row, c).value or "").strip() for c in range(1, 22)]
    age_map = _map_siimed_age_headers(headers)
    hmap = {h.upper(): i for i, h in enumerate(headers) if h}
    nit_i = next((hmap[k] for k in hmap if k == "NIT"), 0)
    nom_i = next((hmap[k] for k in hmap if "NOMBRE" in k and "NIT" not in k), 1)
    cta_i = next((hmap[k] for k in hmap if k == "CUENTA"), 2)

    rows = []
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        vals = [ws.cell(r, c).value for c in range(1, 22)]
        if _is_total_row(vals):
            continue
        nit = _nit(vals[nit_i])
        if not nit:
            continue
        cuenta = str(vals[cta_i] or "").strip()
        row = {
            "nit": nit,
            "nombre": str(vals[nom_i] or "").strip(),
            "cuenta": cuenta,
            "unidad": unit_from_account(cuenta),
            "es_olimpica": nit == OLIMPICA_NIT,
            "saldo": _num(vals[age_map["SALDO"]]) if "SALDO" in age_map else 0.0,
            "vencido": _num(vals[age_map["VENCIDO"]]) if "VENCIDO" in age_map else 0.0,
        }
        for bucket in [
            "CORRIENTE", "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
            "61 AL 90", "91 AL 180", "181 AL 360", "361+",
        ]:
            if bucket == "91 AL 180" and ("91_120" in age_map or "121_180" in age_map):
                row[bucket] = (_num(vals[age_map["91_120"]]) if "91_120" in age_map else 0.0) + (
                    _num(vals[age_map["121_180"]]) if "121_180" in age_map else 0.0
                )
            else:
                row[bucket] = _num(vals[age_map[bucket]]) if bucket in age_map else 0.0
        if not row.get("CORRIENTE"):
            overdue = sum(row.get(b, 0) for b in [
                "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
                "61 AL 90", "91 AL 180", "181 AL 360", "361+",
            ])
            row["CORRIENTE"] = max(row["saldo"] - overdue, 0)
            if not row["vencido"]:
                row["vencido"] = overdue
        rows.append(row)
    return rows


def _aggregate_clients(detail_rows: list[dict]) -> list[dict]:
    buckets = [
        "CORRIENTE", "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
        "61 AL 90", "91 AL 180", "181 AL 360", "361+",
    ]
    agg: dict[str, dict] = {}
    for r in detail_rows:
        nit = r["nit"]
        if nit not in agg:
            agg[nit] = {
                "nit": nit,
                "nombre": r["nombre"],
                "es_olimpica": r["es_olimpica"],
                "saldo": 0.0,
                "vencido": 0.0,
                **{b: 0.0 for b in buckets},
            }
        agg[nit]["saldo"] += r.get("saldo", 0)
        agg[nit]["vencido"] += r.get("vencido", 0)
        for b in buckets:
            agg[nit][b] += r.get(b, 0)
        if not agg[nit]["nombre"] and r["nombre"]:
            agg[nit]["nombre"] = r["nombre"]
    out = list(agg.values())
    out.sort(key=lambda x: x["saldo"], reverse=True)
    return out


def parse_unidad_file(path: str | Path) -> dict:
    """Parsea 'cartera por unidad XX.xlsx'."""
    path = Path(path)
    wb = load_workbook(path, data_only=True)
    meta = {"archivo": path.name, "corte": None}
    ws0 = wb[wb.sheetnames[0]]
    for r in range(1, 6):
        v = ws0.cell(r, 1).value
        if v and "corte a" in str(v).lower():
            meta["corte"] = parse_corte_from_text(str(v))

    rows = []
    # Preferir hoja CARTERA (totales)
    sheet = "CARTERA" if "CARTERA" in wb.sheetnames else wb.sheetnames[-1]
    ws = wb[sheet]
    header_row = _find_header_row(ws, keywords=("CUENTA", "SALDO"))
    headers = [str(ws.cell(header_row, c).value or "").strip() for c in range(1, 25)]
    age_map = _map_siimed_age_headers(headers)

    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        vals = [ws.cell(r, c).value for c in range(1, 25)]
        cuenta_raw = str(vals[0] or "").strip()
        if not cuenta_raw or cuenta_raw.lower().startswith("grand"):
            continue
        cuenta_clean = cuenta_raw.replace(" Total", "").strip()
        row = {
            "cuenta": cuenta_clean,
            "unidad": unit_from_account(cuenta_clean),
            "saldo": _num(vals[age_map["SALDO"]]) if "SALDO" in age_map else _num(vals[1]),
            "vencido": _num(vals[age_map["VENCIDO"]]) if "VENCIDO" in age_map else _num(vals[2]),
        }
        for bucket in [
            "CORRIENTE", "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
            "61 AL 90", "91 AL 180", "181 AL 360", "361+",
        ]:
            if bucket == "91 AL 180" and ("91_120" in age_map or "121_180" in age_map):
                row[bucket] = (_num(vals[age_map["91_120"]]) if "91_120" in age_map else 0.0) + (
                    _num(vals[age_map["121_180"]]) if "121_180" in age_map else 0.0
                )
            else:
                row[bucket] = _num(vals[age_map[bucket]]) if bucket in age_map else 0.0
        if not row.get("CORRIENTE") and "CORRIENTE" not in age_map:
            overdue = sum(row.get(b, 0) for b in [
                "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
                "61 AL 90", "91 AL 180", "181 AL 360", "361+",
            ])
            row["CORRIENTE"] = max(row["saldo"] - overdue, 0)
        rows.append(row)

    # detalle clientes por unidad desde DATOS si existe
    detalle = []
    if "DATOS" in wb.sheetnames:
        detalle = _parse_unidad_datos(wb["DATOS"])

    wb.close()
    return {
        "meta": meta,
        "por_unidad": pd.DataFrame(rows),
        "detalle": pd.DataFrame(detalle) if detalle else pd.DataFrame(),
    }


def _parse_unidad_datos(ws) -> list[dict]:
    header_row = _find_header_row(ws, keywords=("CUENTA", "NIT", "SALDO"))
    headers = [str(ws.cell(header_row, c).value or "").strip() for c in range(1, 20)]
    age_map = _map_siimed_age_headers(headers)
    rows = []
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        vals = [ws.cell(r, c).value for c in range(1, 20)]
        if _is_total_row(vals):
            continue
        cuenta = str(vals[0] or "").strip()
        nit = _nit(vals[3] if len(vals) > 3 else None)
        if not nit or not cuenta:
            continue
        nombre = str(vals[5] or "").strip() if len(vals) > 5 else ""
        row = {
            "nit": nit,
            "nombre": nombre,
            "cuenta": cuenta,
            "unidad": unit_from_account(cuenta),
            "es_olimpica": nit == OLIMPICA_NIT,
            "saldo": _num(vals[age_map["SALDO"]]) if "SALDO" in age_map else _num(vals[6]),
            "vencido": _num(vals[age_map["VENCIDO"]]) if "VENCIDO" in age_map else _num(vals[7]),
        }
        for bucket in [
            "CORRIENTE", "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
            "61 AL 90", "91 AL 180", "181 AL 360", "361+",
        ]:
            if bucket == "91 AL 180" and ("91_120" in age_map or "121_180" in age_map):
                row[bucket] = (_num(vals[age_map["91_120"]]) if "91_120" in age_map else 0.0) + (
                    _num(vals[age_map["121_180"]]) if "121_180" in age_map else 0.0
                )
            else:
                row[bucket] = _num(vals[age_map[bucket]]) if bucket in age_map else 0.0
        if "CORRIENTE" not in age_map:
            overdue = sum(row.get(b, 0) for b in [
                "1 AL 7", "8 AL 15", "16 AL 30", "31 AL 60",
                "61 AL 90", "91 AL 180", "181 AL 360", "361+",
            ])
            row["CORRIENTE"] = max(row["saldo"] - overdue, 0)
        rows.append(row)
    return rows


def parse_balance_file(path: str | Path) -> dict:
    """Parsea balance de prueba / ingresos (opcional para rotación)."""
    path = Path(path)
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    meta = {"archivo": path.name, "periodo": None}
    for r in range(1, 6):
        v = ws.cell(r, 1).value
        if v and "período" in str(v).lower() or (v and "periodo" in str(v).lower()):
            meta["periodo"] = str(v)

    # SIIMED varía el balance: algunos exportes terminan en L y otros agregan
    # CUENTA completa, desplazando DESCRIPCION a I y NUEVOSALDO a M.
    header_row = _find_balance_header_row(ws)
    columns = _map_balance_columns(ws, header_row)
    isolated = _balance_isolated_values(ws, columns["nuevo_saldo"])
    rotation_summary = _rotation_summary_from_values(isolated)

    rows = []
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        desc = ws.cell(r, columns["descripcion"]).value
        saldo = ws.cell(r, columns["nuevo_saldo"]).value
        if desc is None and saldo is None:
            continue
        if isinstance(saldo, str) and saldo.startswith("="):
            continue
        rows.append({
            "clase": ws.cell(r, columns["clase"]).value,
            "grupo": ws.cell(r, columns["grupo"]).value,
            "cuenta": ws.cell(r, columns["cuenta"]).value,
            "codigo_cuenta": ws.cell(r, columns["codigo_cuenta"]).value,
            "subcuenta": ws.cell(r, columns["subcuenta"]).value,
            "auxiliar": ws.cell(r, columns["auxiliar"]).value,
            "descripcion": str(desc or "").strip(),
            "saldo_anterior": _num(ws.cell(r, columns["saldo_anterior"]).value),
            "debito": _num(ws.cell(r, columns["debito"]).value),
            "credito": _num(ws.cell(r, columns["credito"]).value),
            "nuevo_saldo": _num(saldo),
        })
    wb.close()
    return {
        "meta": meta,
        "cuentas": pd.DataFrame(rows),
        "rotacion_resumen": rotation_summary,
        "valores_aislados": isolated,
    }


def parse_balance_files(paths) -> dict:
    """
    Acepta un balance único (cuentas 1 y 4) o el balance dividido por clase
    (CTA 13 con los 5 saldos de cartera y CTA 4 con los 4 acumulados de ingresos).
    """
    if isinstance(paths, (str, Path)):
        paths = [paths]
    parsed = [parse_balance_file(p) for p in paths]
    if len(parsed) == 1:
        return parsed[0]

    cuentas = pd.concat(
        [p["cuentas"] for p in parsed if not p["cuentas"].empty], ignore_index=True
    )
    summary = next((p["rotacion_resumen"] for p in parsed if p["rotacion_resumen"]), {})
    if not summary:
        ingresos: list[float] = []
        cartera: list[float] = []
        for p in parsed:
            clases = set(
                pd.to_numeric(p["cuentas"].get("clase"), errors="coerce")
                .dropna()
                .astype(int)
            ) if "clase" in p["cuentas"].columns else set()
            values = p["valores_aislados"]
            if 4 in clases and len(values) >= 4:
                ingresos = values[-4:]
            elif 1 in clases and len(values) >= 5:
                cartera = values[-5:]
        if ingresos and cartera:
            summary = _rotation_summary_from_values(ingresos + cartera)

    return {
        "meta": {
            "archivo": " + ".join(p["meta"]["archivo"] for p in parsed),
            "periodo": next((p["meta"]["periodo"] for p in parsed if p["meta"]["periodo"]), None),
        },
        "cuentas": cuentas,
        "rotacion_resumen": summary,
        "valores_aislados": [],
    }


def _find_balance_header_row(ws) -> int:
    for r in range(1, min(ws.max_row or 1, 20) + 1):
        headers = {
            str(ws.cell(r, c).value or "").strip().upper().replace(" ", "")
            for c in range(1, (ws.max_column or 1) + 1)
        }
        if "DESCRIPCION" in headers and (
            "SALDOANTERIOR" in headers or "NUEVOSALDO" in headers
        ):
            return r
    return 7


def _map_balance_columns(ws, header_row: int) -> dict[str, int]:
    positions: dict[str, list[int]] = {}
    for c in range(1, (ws.max_column or 1) + 1):
        key = (
            str(ws.cell(header_row, c).value or "")
            .strip()
            .upper()
            .replace(" ", "")
            .replace("Ó", "O")
            .replace("Í", "I")
        )
        if key:
            positions.setdefault(key, []).append(c)

    def first(name: str, fallback: int) -> int:
        return positions.get(name, [fallback])[0]

    # La primera CUENTA es el componente contable; la última suele ser el
    # código completo y no reemplaza CUENTA/SUBCUENTA/AUXILIAR.
    nuevo_saldo = first("NUEVOSALDO", ws.max_column or 12)
    if "NUEVOSALDO" not in positions:
        # En algunos archivos el encabezado viene vacío, pero es la columna
        # inmediatamente posterior a CREDITO.
        nuevo_saldo = first("CREDITO", 11) + 1

    return {
        "clase": first("CLASE", 2),
        "grupo": first("GRUPO", 3),
        "cuenta": first("CUENTA", 4),
        "codigo_cuenta": positions.get("CUENTA", [4])[-1],
        "subcuenta": first("SUBCUENTA", 5),
        "auxiliar": first("AUXILIAR", 6),
        "descripcion": first("DESCRIPCION", 8),
        "saldo_anterior": first("SALDOANTERIOR", 9),
        "debito": first("DEBITO", 10),
        "credito": first("CREDITO", 11),
        "nuevo_saldo": nuevo_saldo,
    }


def _balance_isolated_values(ws, value_col: int) -> list[float]:
    """
    Lee los bloques auxiliares que SIIMED deja al final del balance.

    Son la fuente exacta usada por el informe manual para Formula!BN4:BN16:
    cuatro acumulados de ingresos y cinco saldos de cartera.
    """
    isolated: list[float] = []
    for r in range(8, (ws.max_row or 0) + 1):
        value = ws.cell(r, value_col).value
        if not isinstance(value, (int, float)):
            continue
        if all(ws.cell(r, c).value in (None, "") for c in range(1, value_col)):
            isolated.append(float(value))
    return isolated


def _rotation_summary_from_values(isolated: list[float]) -> dict:
    if len(isolated) < 9:
        return {}
    values = isolated[-9:]
    return {
        "ingresos_ytd": {
            "faenado": values[0],
            "desposte": values[1],
            "carnes": values[2],
            "fomento": values[3],
            "total": sum(values[:4]),
        },
        "cartera": {
            "faenado": values[4],
            "desposte": values[5],
            "carnes": values[6],
            "fomento": values[7],
            "total": values[8],
        },
    }


def format_corte_label(corte: str | None) -> str:
    if not corte:
        return datetime.now().strftime("%d de %B de %Y")
    try:
        dt = datetime.strptime(corte, "%Y-%m-%d")
    except ValueError:
        return corte
    months = [
        "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]
    return f"{dt.day} de {months[dt.month]} de {dt.year}"
