# -*- coding: utf-8 -*-
"""Validación / detección del tipo de export SIIMED."""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook


class FileValidationError(Exception):
    """Archivo no corresponde al tipo esperado."""


def _sheet_texts(ws, max_rows=8, max_cols=12) -> str:
    parts = []
    for r in range(1, max_rows + 1):
        for c in range(1, max_cols + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip():
                parts.append(v.strip().upper())
    return " | ".join(parts)


def detect_siimed_type(path: str | Path) -> str:
    """
    Detecta el tipo de archivo SIIMED.
    Retorna: 'cartera_edades' | 'cartera_unidad' | 'balance' | 'desconocido'
    """
    path = Path(path)
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = [n.upper().strip() for n in wb.sheetnames]
        # Leer textos de la primera hoja y, si existe, DATOS
        texts = []
        for sn in wb.sheetnames[:3]:
            texts.append(_sheet_texts(wb[sn]))
        text = " || ".join(texts)

        # 1) Balance de prueba (más específico)
        if "BALANCE DE PRUEBA" in text:
            return "balance"
        if "SALDOANTERIOR" in text and "DEBITO" in text and "CREDITO" in text:
            return "balance"
        if "NUEVOSALDO" in text and "CLASE" in text and "DESCRIPCION" in text:
            return "balance"

        # 2) Cartera por unidad / por CUENTA
        if "POR CUENTA" in text or "EDADES DE PROVEEDORES" in text:
            return "cartera_unidad"
        if "CENTRO COSTO" in text and "SUBCENTRO" in text:
            return "cartera_unidad"

        # 3) Cartera por edades / por identificación del tercero (clientes)
        if "IDENTIFICACION DEL TERCERO" in text or "IDENTIFICACIÓN DEL TERCERO" in text:
            return "cartera_edades"
        if "EDADES DE CARTERA" in text and "POR CUENTA" not in text:
            return "cartera_edades"
        if "CARTERA SIN OLIMPICA" in names:
            return "cartera_edades"

        # 4) Heurística por combinación de hojas
        # archivo unidad típico: DATOS + CARTERA (sin "sin olimpica", sin hoja unidad de clientes)
        if "CARTERA" in names and "DATOS" in names and "UNIDAD" not in names:
            # puede ser unidad (totales por cuenta) si título lo dice
            if "PROVEEDORES" in text or "POR CUENTA" in text:
                return "cartera_unidad"
        if "UNIDAD" in names and ("CARTERA" in names or "DATOS" in names):
            return "cartera_edades"

        return "desconocido"
    finally:
        wb.close()


LABELS = {
    "cartera_edades": "Cartera por edades (clientes / NIT)",
    "cartera_unidad": "Cartera por unidad / cuenta",
    "balance": "Balance de prueba / ingresos",
    "desconocido": "No reconocido como export SIIMED de cartera",
}


def validate_inputs(
    cartera_path: str | Path,
    unidad_path: str | Path | None = None,
    balance_path: str | Path | list[str | Path] | None = None,
) -> list[str]:
    """
    Valida que cada archivo sea del tipo correcto.
    Retorna lista de advertencias (vacía = OK).
    Lanza FileValidationError si el obligatorio está mal o hay cruce claro.
    """
    warnings: list[str] = []
    errors: list[str] = []

    t_cartera = detect_siimed_type(cartera_path)
    if t_cartera != "cartera_edades":
        errors.append(
            f"El archivo 1 no parece 'Cartera por edades'. "
            f"Detectado: {LABELS.get(t_cartera, t_cartera)}. "
            f"Sube el Excel tipo “cartera a 30 …xlsx” (edades por cliente/NIT)."
        )

    if unidad_path:
        t_u = detect_siimed_type(unidad_path)
        if t_u == "cartera_edades":
            errors.append(
                "El archivo 2 parece 'Cartera por edades', no 'por unidad'. "
                "Están intercambiados o elegiste el archivo equivocado."
            )
        elif t_u == "balance":
            errors.append(
                "El archivo 2 parece un Balance de ingresos, no cartera por unidad."
            )
        elif t_u != "cartera_unidad":
            errors.append(
                f"El archivo 2 no se reconoce como cartera por unidad. "
                f"Detectado: {LABELS.get(t_u, t_u)}."
            )

    balance_paths = (
        [balance_path] if isinstance(balance_path, (str, Path)) else list(balance_path or [])
    )
    for path in balance_paths:
        t_b = detect_siimed_type(path)
        if t_b != "balance":
            errors.append(
                f"El archivo 3 ({Path(path).name}) no parece Balance de ingresos. "
                f"Detectado: {LABELS.get(t_b, t_b)}. "
                f"Si no lo tienes, déjalo vacío (es opcional)."
            )

    if unidad_path and not errors:
        if Path(cartera_path).name.lower() == Path(unidad_path).name.lower():
            warnings.append("Subiste el mismo nombre de archivo en casilla 1 y 2.")

    if errors:
        raise FileValidationError("\n".join(errors))

    return warnings
