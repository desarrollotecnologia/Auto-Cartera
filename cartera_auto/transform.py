# -*- coding: utf-8 -*-
"""Transformaciones y KPIs de cartera."""

from __future__ import annotations

import pandas as pd

from .config import AGE_BUCKETS, OLIMPICA_NIT, UNIT_ORDER

# Anticipos materiales que Cartera descuenta del saldo (mayo 2026: TECNOLOGIAS -783.296)
# Umbral evita ruido de centavos; Olímpica/cadenas grandes suelen quedar fuera.
ANTICIPO_MATERIAL_MIN = 500_000.0


def _safe_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.copy() if df is not None and not df.empty else pd.DataFrame()


def _balance_amount(
    cuentas: pd.DataFrame, codes: str | list[str], field: str = "nuevo_saldo"
) -> float:
    if cuentas.empty or "codigo_cuenta" not in cuentas.columns:
        return 0.0
    wanted = {str(code) for code in ([codes] if isinstance(codes, str) else codes)}
    normalized = cuentas["codigo_cuenta"].apply(
        lambda value: str(int(value)) if pd.notna(value) else ""
    )
    rows = cuentas[normalized.isin(wanted)]
    return abs(float(rows[field].sum())) if not rows.empty else 0.0


def _rotation_balance_summary(
    balance: dict | None, por_unidad: pd.DataFrame
) -> dict:
    exact = (balance or {}).get("rotacion_resumen") or {}
    if exact:
        return exact

    cuentas = _safe_df((balance or {}).get("cuentas"))
    if cuentas.empty:
        return {}

    operational = _balance_amount(cuentas, "41")
    carnes = _balance_amount(cuentas, "4120")
    desposte = _balance_amount(
        cuentas, ["41057210", "41057306", "41358510", "413985"]
    )
    fomento = _balance_amount(cuentas, ["413518", "4145"])
    faenado = max(operational - carnes - desposte - fomento, 0.0)

    units = {
        str(row["unidad"]).upper(): float(row["saldo"])
        for _, row in por_unidad.iterrows()
    } if not por_unidad.empty else {}
    cartera = {
        "carnes": units.get("CARNES COLBEEF", 0.0),
        "desposte": units.get("DESPOSTE", 0.0),
        "fomento": units.get("DIFICIL COBRO", 0.0),
    }
    cartera["faenado"] = max(
        float(por_unidad["saldo"].sum()) - sum(cartera.values()), 0.0
    ) if not por_unidad.empty else 0.0

    return {
        "ingresos_ytd": {
            "faenado": faenado,
            "desposte": desposte,
            "carnes": carnes,
            "fomento": fomento,
            "total": operational,
        },
        "cartera": cartera,
    }


def _apply_anticipos(detalle: pd.DataFrame, por_cliente: pd.DataFrame, anticipos: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Descuenta anticipos materiales del saldo/corriente (lógica informe Cartera mayo).
    Solo clientes con cartera en unidad LOCAL y |anticipo| >= umbral.
    Cadenas (Olímpica, etc.) mantienen saldo bruto; el anticipo queda en hoja aparte.
    """
    if not anticipos or detalle.empty:
        return detalle, por_cliente

    local_nits = set()
    if "unidad" in detalle.columns:
        local_nits = set(
            detalle.loc[detalle["unidad"].astype(str).str.upper() == "LOCAL", "nit"].astype(str)
        )

    detalle = detalle.copy()
    for nit, ant in anticipos.items():
        nit_s = str(nit)
        if abs(float(ant)) < ANTICIPO_MATERIAL_MIN:
            continue
        if nit_s == OLIMPICA_NIT:
            continue
        if local_nits and nit_s not in local_nits:
            continue
        mask = detalle["nit"].astype(str) == nit_s
        if not mask.any():
            continue
        sub = detalle.loc[mask]
        total = float(sub["saldo"].sum()) or 1.0
        for idx in sub.index:
            share = float(detalle.at[idx, "saldo"]) / total
            adj = float(ant) * share
            detalle.at[idx, "saldo"] = float(detalle.at[idx, "saldo"]) + adj
            if "CORRIENTE" in detalle.columns:
                detalle.at[idx, "CORRIENTE"] = max(float(detalle.at[idx, "CORRIENTE"]) + adj, 0)

    if not por_cliente.empty:
        por_cliente = por_cliente.copy()
        for nit, ant in anticipos.items():
            nit_s = str(nit)
            if abs(float(ant)) < ANTICIPO_MATERIAL_MIN:
                continue
            if nit_s == OLIMPICA_NIT:
                continue
            if local_nits and nit_s not in local_nits:
                continue
            mask = por_cliente["nit"].astype(str) == nit_s
            if not mask.any():
                continue
            por_cliente.loc[mask, "saldo"] = por_cliente.loc[mask, "saldo"].astype(float) + float(ant)
            if "CORRIENTE" in por_cliente.columns:
                por_cliente.loc[mask, "CORRIENTE"] = (
                    por_cliente.loc[mask, "CORRIENTE"].astype(float) + float(ant)
                ).clip(lower=0)
    return detalle, por_cliente


def build_report_model(
    cartera: dict,
    unidad: dict | None = None,
    balance: dict | None = None,
) -> dict:
    """Construye el modelo de datos listo para Excel/PPT."""
    detalle = _safe_df(cartera.get("detalle"))
    por_cliente = _safe_df(cartera.get("por_cliente"))
    anticipos = dict(cartera.get("anticipos") or {})

    # Preferir detalle del archivo por unidad (fuente coherente con totales SIIMED por cuenta)
    if unidad and not _safe_df(unidad.get("detalle")).empty:
        detalle = _ensure_vencido(unidad["detalle"].copy())
        # Consolidada desde la misma fuente (evita mezclar cartera vs cuenta de fechas distintas)
        por_cliente = _aggregate_from_detalle(detalle)
    else:
        if not detalle.empty:
            detalle = _ensure_vencido(detalle)
        if por_cliente.empty and not detalle.empty:
            por_cliente = _aggregate_from_detalle(detalle)
        elif not por_cliente.empty:
            por_cliente = _ensure_vencido(por_cliente)

    # Anticipos materiales (hoja CARTERA -ANTICIPOS)
    detalle, por_cliente = _apply_anticipos(detalle, por_cliente, anticipos)

    # Por unidad: recalcular desde detalle ajustado si hay unidad
    if not detalle.empty and "unidad" in detalle.columns:
        cols = ["saldo", "vencido"] + [b for b in AGE_BUCKETS if b in detalle.columns]
        por_unidad = detalle.groupby("unidad", as_index=False)[cols].sum(numeric_only=True)
    elif unidad and not _safe_df(unidad.get("por_unidad")).empty:
        por_unidad = unidad["por_unidad"].copy()
        por_unidad = (
            por_unidad.groupby("unidad", as_index=False)[
                ["saldo", "vencido"] + [b for b in AGE_BUCKETS if b in por_unidad.columns]
            ].sum(numeric_only=True)
        )
    else:
        por_unidad = pd.DataFrame()

    if not por_unidad.empty:
        por_unidad = _ensure_vencido(por_unidad)
        por_unidad["_ord"] = por_unidad["unidad"].apply(
            lambda u: UNIT_ORDER.index(u) if u in UNIT_ORDER else 99
        )
        por_unidad = por_unidad.sort_values("_ord").drop(columns="_ord")

    # Totales oficiales: preferir suma por unidad SIIMED (Grand Total exacto)
    if not por_unidad.empty:
        totales = _sum_row(por_unidad, "TOTAL CARTERA")
    else:
        totales = _sum_row(por_cliente, "TOTAL CARTERA")

    # Sin Olímpica / Olímpica a partir de clientes
    sin_ol = por_cliente[por_cliente["nit"].astype(str) != OLIMPICA_NIT].copy()
    olimpica = por_cliente[por_cliente["nit"].astype(str) == OLIMPICA_NIT]
    tot_ol = _sum_row(olimpica, "OLÍMPICA") if not olimpica.empty else _empty_sum("OLÍMPICA")
    # Asegurar vencido Olímpica desde buckets
    if tot_ol.get("vencido", 0) == 0:
        tot_ol["vencido"] = sum(float(tot_ol.get(b, 0) or 0) for b in AGE_BUCKETS if b != "CORRIENTE")

    # Si tenemos total oficial, sin Olímpica = total - Olímpica (más exacto)
    if not por_unidad.empty and tot_ol.get("saldo", 0):
        totales_sin = {
            "label": "TOTAL SIN OLÍMPICA",
            "saldo": float(totales.get("saldo", 0)) - float(tot_ol.get("saldo", 0)),
            "vencido": float(totales.get("vencido", 0)) - float(tot_ol.get("vencido", 0)),
        }
        for b in AGE_BUCKETS:
            totales_sin[b] = float(totales.get(b, 0)) - float(tot_ol.get(b, 0))
    else:
        totales_sin = _sum_row(sin_ol, "TOTAL SIN OLÍMPICA")

    # Detalle por unidad (top clientes)
    detalle_por_unidad = {}
    if not detalle.empty and "unidad" in detalle.columns:
        for unit in UNIT_ORDER:
            sub = detalle[detalle["unidad"] == unit].copy()
            if sub.empty:
                continue
            # agregar por nit dentro de la unidad
            gcols = ["saldo", "vencido"] + [b for b in AGE_BUCKETS if b in sub.columns]
            g = sub.groupby(["nit", "nombre"], as_index=False)[gcols].sum(numeric_only=True)
            g = g.sort_values("saldo", ascending=False)
            detalle_por_unidad[unit] = g

    # KPIs
    total_saldo = float(totales.get("saldo", 0) or 0)
    total_vencido = float(totales.get("vencido", 0) or 0)
    total_corriente = float(totales.get("CORRIENTE", 0) or 0)
    if not total_corriente and total_saldo:
        total_corriente = total_saldo - total_vencido

    pct_vencida = (total_vencido / total_saldo) if total_saldo else 0.0
    pct_corriente = (total_corriente / total_saldo) if total_saldo else 0.0

    sin_saldo = float(totales_sin.get("saldo", 0) or 0)
    sin_vencido = float(totales_sin.get("vencido", 0) or 0)
    pct_vencida_sin = (sin_vencido / sin_saldo) if sin_saldo else 0.0

    # Edades consolidado (para slide)
    edades = {b: float(totales.get(b, 0) or 0) for b in AGE_BUCKETS}
    edades_sin = {b: float(totales_sin.get(b, 0) or 0) for b in AGE_BUCKETS}

    # Vencida > 15 días
    vencida_15 = sum(
        float(totales.get(b, 0) or 0)
        for b in ["16 AL 30", "31 AL 60", "61 AL 90", "91 AL 180", "181 AL 360", "361+"]
    )
    vencida_15_sin = sum(
        float(totales_sin.get(b, 0) or 0)
        for b in ["16 AL 30", "31 AL 60", "61 AL 90", "91 AL 180", "181 AL 360", "361+"]
    )

    # Ingresos del balance (opcional, para rotación simplificada)
    ingresos_mes = None
    if balance and not _safe_df(balance.get("cuentas")).empty:
        cta = balance["cuentas"]
        # Clase 4 INGRESO total si existe
        ing = cta[cta["descripcion"].str.upper() == "INGRESO"]
        if not ing.empty:
            # Movimiento neto del mes; NUEVOSALDO es acumulado del año.
            ingresos_mes = abs(
                float(ing.iloc[0]["credito"]) - float(ing.iloc[0]["debito"])
            )

    dias_rotacion = None
    if ingresos_mes and ingresos_mes > 0 and total_saldo:
        # aproximación: cartera / (ingresos/30)
        dias_rotacion = total_saldo / (ingresos_mes / 30.0)

    meta = dict(cartera.get("meta") or {})
    if unidad and unidad.get("meta", {}).get("corte") and not meta.get("corte"):
        meta["corte"] = unidad["meta"]["corte"]

    return {
        "meta": meta,
        "por_cliente": por_cliente,
        "por_cliente_sin_olimpica": sin_ol,
        "por_unidad": por_unidad,
        "detalle_por_unidad": detalle_por_unidad,
        "totales": totales,
        "totales_sin_olimpica": totales_sin,
        "totales_olimpica": tot_ol,
        "edades": edades,
        "edades_sin_olimpica": edades_sin,
        "rotacion_balance": _rotation_balance_summary(balance, por_unidad),
        "kpis": {
            "total_cartera": total_saldo,
            "cartera_corriente": total_corriente,
            "cartera_vencida": total_vencido,
            "pct_vencida": pct_vencida,
            "pct_corriente": pct_corriente,
            "total_sin_olimpica": sin_saldo,
            "vencida_sin_olimpica": sin_vencido,
            "pct_vencida_sin_olimpica": pct_vencida_sin,
            "olimpica_saldo": float(tot_ol.get("saldo", 0) or 0),
            "olimpica_vencido": float(tot_ol.get("vencido", 0) or 0),
            "vencida_mayor_15": vencida_15,
            "vencida_mayor_15_sin": vencida_15_sin,
            "ingresos_mes": ingresos_mes,
            "dias_rotacion": dias_rotacion,
            "n_clientes": int(len(por_cliente)),
            "n_clientes_sin_olimpica": int(len(sin_ol)),
        },
    }


def _ensure_vencido(df: pd.DataFrame) -> pd.DataFrame:
    """Recalcula vencido = suma de buckets vencidos si falta o está en cero."""
    out = df.copy()
    overdue_cols = [b for b in AGE_BUCKETS if b != "CORRIENTE" and b in out.columns]
    if not overdue_cols:
        return out
    computed = out[overdue_cols].sum(axis=1)
    if "vencido" not in out.columns:
        out["vencido"] = computed
    else:
        # reemplazar filas en cero cuando sí hay vencido en buckets
        mask = (out["vencido"].fillna(0) == 0) & (computed > 0)
        out.loc[mask, "vencido"] = computed[mask]
    if "CORRIENTE" not in out.columns and "saldo" in out.columns:
        out["CORRIENTE"] = (out["saldo"] - out["vencido"]).clip(lower=0)
    elif "CORRIENTE" in out.columns and "saldo" in out.columns:
        mask = (out["CORRIENTE"].fillna(0) == 0) & (out["saldo"] > out["vencido"])
        out.loc[mask, "CORRIENTE"] = (out.loc[mask, "saldo"] - out.loc[mask, "vencido"]).clip(lower=0)
    return out


def _aggregate_from_detalle(detalle: pd.DataFrame) -> pd.DataFrame:
    buckets = [b for b in AGE_BUCKETS if b in detalle.columns]
    cols = ["saldo", "vencido"] + buckets
    g = detalle.groupby(["nit", "nombre"], as_index=False)[cols].sum(numeric_only=True)
    g["es_olimpica"] = g["nit"].astype(str) == OLIMPICA_NIT
    return g.sort_values("saldo", ascending=False).reset_index(drop=True)


def _sum_row(df: pd.DataFrame, label: str) -> dict:
    if df is None or df.empty:
        return _empty_sum(label)
    out = {"label": label}
    for col in ["saldo", "vencido"] + AGE_BUCKETS:
        if col in df.columns:
            out[col] = float(df[col].sum())
        else:
            out[col] = 0.0
    return out


def _empty_sum(label: str) -> dict:
    return {"label": label, "saldo": 0.0, "vencido": 0.0, **{b: 0.0 for b in AGE_BUCKETS}}


def money_mm(value: float, decimals: int = 1) -> str:
    """Formatea a millones de pesos (como se lee en el informe original)."""
    return f"${value / 1_000_000:,.{decimals}f} MM"


def snap_pesos_mm(value: float) -> float:
    """Redondea pesos al millón entero (solo usos de visualización; no usar en Excel)."""
    return round(float(value or 0) / 1_000_000) * 1_000_000


def money_cop(value: float) -> str:
    return f"${value:,.0f}"


def pct(value: float) -> str:
    return f"{value * 100:.1f}%"
