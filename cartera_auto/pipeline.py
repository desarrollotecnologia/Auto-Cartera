# -*- coding: utf-8 -*-
"""Orquestación: SIIMED → Excel + PPT."""

from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from .config import OUTPUT_DIR
from .excel_report import generate_excel
from .parsers import parse_balance_files, parse_cartera_file, parse_unidad_file
from .ppt_report import generate_pptx
from .report_validation import (
    validate_generated_excel,
    validate_with_excel_desktop,
)
from .transform import build_report_model
from .validate import FileValidationError, validate_inputs


def run_pipeline(
    cartera_path: str | Path,
    unidad_path: str | Path | None = None,
    balance_path: str | Path | list[str | Path] | None = None,
    output_dir: str | Path | None = None,
    prefix: str | None = None,
) -> dict:
    """
    Ejecuta el pipeline completo.

    Returns:
        dict con paths de salida y modelo/kpis.
    """
    # Candado: no procesar ni tocar plantilla histórica si los archivos no cuadran
    warnings = validate_inputs(cartera_path, unidad_path, balance_path)

    requested_output = Path(output_dir or OUTPUT_DIR)
    actual_dir, archive_dir, managed_output = _resolve_output_dirs(requested_output)
    actual_dir.mkdir(parents=True, exist_ok=True)
    if managed_output:
        archive_dir.mkdir(parents=True, exist_ok=True)

    cartera = parse_cartera_file(cartera_path)
    unidad = parse_unidad_file(unidad_path) if unidad_path else None
    balance = parse_balance_files(balance_path) if balance_path else None

    model = build_report_model(cartera, unidad, balance)

    # Candado extra: totales absurdos / vacíos
    if not model["kpis"]["total_cartera"] or model["kpis"]["total_cartera"] <= 0:
        raise FileValidationError(
            "No se pudo calcular un total de cartera válido. "
            "Revisa que el archivo 1 sea el export SIIMED de edades del mes correcto."
        )
    if model["kpis"]["n_clientes"] < 3:
        raise FileValidationError(
            "Se detectaron muy pocos clientes. "
            "Probablemente el archivo subido no es el reporte de cartera por edades."
        )

    corte = model["meta"].get("corte") or datetime.now().strftime("%Y-%m-%d")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = prefix or f"Informe_Cartera_{corte}"

    excel_name = f"{prefix}_{stamp}.xlsx"
    ppt_name = f"{prefix}_{stamp}.pptx"
    warnings = list(warnings)

    # Generar fuera de ACTUAL: un fallo nunca reemplaza el informe vigente.
    staging_parent = OUTPUT_DIR if managed_output else actual_dir
    staging_parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".cartera_staging_", dir=staging_parent) as tmp:
        staging = Path(tmp)
        staged_excel = staging / excel_name
        staged_ppt = staging / ppt_name

        generate_excel(model, staged_excel)
        validation = validate_generated_excel(staged_excel, corte)
        warnings.extend(validation.get("warnings") or [])
        excel_runtime = validate_with_excel_desktop(staged_excel)
        validation["excel_desktop"] = excel_runtime
        if excel_runtime.get("warning"):
            warnings.append(
                "No se pudo ejecutar la validación con Excel Desktop: "
                + str(excel_runtime["warning"])
            )
        generate_pptx(model, staged_ppt)

        if managed_output:
            warnings.extend(_archive_current(actual_dir, archive_dir))

        excel_path = actual_dir / excel_name
        ppt_path = actual_dir / ppt_name
        shutil.copy2(staged_excel, excel_path)
        shutil.copy2(staged_ppt, ppt_path)

    # Actualizar plantilla solo después de validar y publicar ambos entregables.
    if model.get("historico", {}).get("updated"):
        template_warning = _update_historical_template(excel_path, model, corte)
        if template_warning:
            warnings.append(template_warning)

    return {
        "excel": excel_path,
        "pptx": ppt_path,
        "model": model,
        "kpis": model["kpis"],
        "meta": model["meta"],
        "historico": model.get("historico"),
        "validation": validation,
        "warnings": warnings,
    }


def _resolve_output_dirs(requested: Path) -> tuple[Path, Path, bool]:
    """Salida operativa única en output/ACTUAL y aprobados en output/ARCHIVO."""
    try:
        is_root_output = requested.resolve() == OUTPUT_DIR.resolve()
    except OSError:
        is_root_output = requested == OUTPUT_DIR
    is_actual = requested.name.upper() == "ACTUAL"
    if is_root_output:
        return requested / "ACTUAL", requested / "ARCHIVO", True
    if is_actual:
        return requested, requested.parent / "ARCHIVO", True
    return requested, requested / "ARCHIVO", False


def _archive_current(actual_dir: Path, archive_dir: Path) -> list[str]:
    """Mueve versiones anteriores; si Excel las tiene abiertas, las reporta."""
    warnings: list[str] = []
    for path in actual_dir.iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        target = archive_dir / path.name
        if target.exists():
            target = archive_dir / (
                f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
            )
        try:
            shutil.move(str(path), str(target))
        except OSError:
            warnings.append(
                f"No se pudo archivar {path.name}; ciérralo en Excel y elimínalo."
            )
    return warnings


def _update_historical_template(
    excel_path: Path, model: dict, corte: str
) -> str | None:
    """Actualización atómica con validación y respaldo de la plantilla previa."""
    from openpyxl import load_workbook

    from .historico import freeze_historico_with_model

    root = Path(__file__).resolve().parent.parent
    template = root / "templates" / "plantilla_informe_cartera.xlsx"
    backup_dir = root / "templates" / "archivo"
    backup_dir.mkdir(parents=True, exist_ok=True)

    fd, tmp_name = tempfile.mkstemp(
        prefix=".plantilla_", suffix=".xlsx", dir=template.parent
    )
    os.close(fd)
    tmp_template = Path(tmp_name)
    try:
        shutil.copy2(excel_path, tmp_template)
        wb = load_workbook(tmp_template)
        # Réplica exacta: la hoja auxiliar no pertenece al original.
        if "AUTO_RESUMEN" in wb.sheetnames:
            del wb["AUTO_RESUMEN"]
        freeze_historico_with_model(wb, model)
        wb.save(tmp_template)
        wb.close()

        validate_generated_excel(tmp_template, corte)

        if template.exists():
            backup = backup_dir / (
                f"plantilla_informe_cartera_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            shutil.copy2(template, backup)
        try:
            os.replace(tmp_template, template)
        except PermissionError:
            return (
                "El informe se generó y validó, pero la plantilla histórica está "
                "abierta o bloqueada. Ciérrala en Excel y vuelve a generar antes "
                "de procesar el mes siguiente."
            )
        _trim_backups(backup_dir, keep=3)
        return None
    finally:
        if tmp_template.exists():
            tmp_template.unlink()


def _trim_backups(directory: Path, keep: int) -> None:
    backups = sorted(
        directory.glob("plantilla_informe_cartera_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for path in backups[keep:]:
        try:
            path.unlink()
        except OSError:
            pass
