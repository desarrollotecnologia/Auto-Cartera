# -*- coding: utf-8 -*-
"""Validación estructural del informe Excel antes de publicarlo."""

from __future__ import annotations

import os
import posixpath
import re
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook

from .historico import MONTH_NAMES


EXPECTED_CHARTS = {
    "inf cartera x edades": 1,
    "informe jf": 2,
    "Record cvda": 1,
    "PERIODO DE  RECAUDO": 2,
    "2023": 1,
}

REQUIRED_SHEETS = {
    "unidad",
    "inf cartera x edades",
    "informe jf",
    "Record cvda",
    "PERIODO DE  RECAUDO",
    "Formula",
    "carnes",
    "local",
    "regional",
    "nacional",
    "subproductos",
    "otras cxc",
    "institucional",
    "desposte",
    "2023",
}

REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}


class ReportValidationError(Exception):
    """El Excel generado no es seguro para publicar."""


def validate_generated_excel(path: str | Path, corte: str | None = None) -> dict:
    """
    Comprueba ZIP, relaciones, hojas, fórmulas, meses y las 7 gráficas.

    Lanza ReportValidationError ante un error que haría incompleto o reparable
    el archivo en Excel.
    """
    path = Path(path)
    errors: list[str] = []
    warnings: list[str] = []

    if not path.exists() or path.stat().st_size < 100_000:
        raise ReportValidationError(f"Excel ausente o incompleto: {path}")

    package = _validate_package(path)
    errors.extend(package["errors"])

    try:
        wb = load_workbook(path, data_only=False, read_only=False)
    except Exception as exc:
        errors.append(f"openpyxl no pudo abrir el libro: {exc}")
        wb = None

    chart_total = 0
    if wb is not None:
        try:
            missing = sorted(REQUIRED_SHEETS - set(wb.sheetnames))
            if missing:
                errors.append("Faltan hojas obligatorias: " + ", ".join(missing))
            if not any(name.lower().startswith("consolidada ") for name in wb.sheetnames):
                errors.append("Falta la hoja consolidada del mes")
            if not any(
                name.lower().startswith("comportamiento ") for name in wb.sheetnames
            ):
                errors.append("Falta la hoja comportamiento del mes")

            for sheet_name, expected in EXPECTED_CHARTS.items():
                if sheet_name not in wb.sheetnames:
                    continue
                found = len(getattr(wb[sheet_name], "_charts", []) or [])
                chart_total += found
                if found != expected:
                    errors.append(
                        f"{sheet_name}: se esperaban {expected} gráfica(s), hay {found}"
                    )

            if chart_total != 7:
                errors.append(f"El informe debe contener 7 gráficas; contiene {chart_total}")

            broken = _find_broken_formulas(wb)
            if broken:
                sample = ", ".join(broken[:8])
                errors.append(f"Hay fórmulas con #REF! ({sample})")

            external = _find_external_formulas(wb)
            if external:
                warnings.append(
                    f"Quedaron {len(external)} fórmula(s) con vínculo externo"
                )

            if corte and "inf cartera x edades" in wb.sheetnames:
                month = int(corte[5:7])
                expected_month = MONTH_NAMES[month]
                headers = [
                    str(wb["inf cartera x edades"].cell(3, c).value or "")
                    .strip()
                    .upper()
                    for c in range(3, 45)
                ]
                if expected_month not in headers:
                    errors.append(
                        f"No aparece {expected_month} en 'inf cartera x edades'"
                    )
        finally:
            wb.close()

    if errors:
        raise ReportValidationError("\n".join(errors))

    return {
        "ok": True,
        "charts": chart_total,
        "drawings": package["drawings"],
        "warnings": warnings,
    }


def validate_with_excel_desktop(path: str | Path) -> dict:
    """
    Abre el archivo con Excel real (Windows) y comprueba que sobrevivan 7 charts.

    Si Excel no está disponible, retorna unavailable sin bloquear la generación.
    """
    if os.name != "nt" or os.environ.get("CARTERA_SKIP_EXCEL_CHECK") == "1":
        return {"available": False, "ok": None}

    absolute = str(Path(path).resolve()).replace("'", "''")
    script = rf"""
$excel = $null
$wb = $null
try {{
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open('{absolute}', 0, $true)
  $charts = 0
  foreach ($ws in $wb.Worksheets) {{ $charts += $ws.ChartObjects().Count }}
  if ($charts -ne 7) {{
    Write-Output "CARTERA_EXCEL_INVALID charts=$charts"
    exit 2
  }}
  Write-Output "CARTERA_EXCEL_OK charts=$charts sheets=$($wb.Worksheets.Count)"
}} catch {{
  Write-Output "CARTERA_EXCEL_ERROR $($_.Exception.Message)"
  exit 3
}} finally {{
  if ($wb) {{ $wb.Close($false) }}
  if ($excel) {{ $excel.Quit() }}
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}}
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=75,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {"available": False, "ok": None, "warning": str(exc)}

    output = (result.stdout + "\n" + result.stderr).strip()
    if result.returncode != 0 or "CARTERA_EXCEL_OK" not in output:
        raise ReportValidationError(
            "Excel Desktop no pudo abrir el informe conservando las 7 gráficas. "
            + output[-500:]
        )
    return {"available": True, "ok": True, "message": output}


def _validate_package(path: Path) -> dict:
    errors: list[str] = []
    drawings = 0
    try:
        with zipfile.ZipFile(path) as archive:
            bad_member = archive.testzip()
            if bad_member:
                errors.append(f"ZIP corrupto en {bad_member}")

            names = set(archive.namelist())
            drawings = len(
                [
                    name
                    for name in names
                    if re.fullmatch(r"xl/drawings/drawing\d+\.xml", name)
                ]
            )
            charts = len(
                [
                    name
                    for name in names
                    if re.fullmatch(r"xl/charts/chart\d+\.xml", name)
                ]
            )
            if drawings != 5:
                errors.append(f"Paquete: se esperaban 5 drawings, hay {drawings}")
            if charts != 7:
                errors.append(f"Paquete: se esperaban 7 charts, hay {charts}")

            errors.extend(_broken_relationships(archive, names))
    except (OSError, zipfile.BadZipFile) as exc:
        errors.append(f"El archivo XLSX no es un ZIP válido: {exc}")

    return {"errors": errors, "drawings": drawings}


def _broken_relationships(
    archive: zipfile.ZipFile, names: set[str]
) -> list[str]:
    errors: list[str] = []
    rel_files = [
        name
        for name in names
        if name.endswith(".rels")
        and (
            name.startswith("xl/drawings/_rels/")
            or name.startswith("xl/charts/_rels/")
        )
    ]
    for rel_name in rel_files:
        try:
            root = ET.fromstring(archive.read(rel_name))
        except ET.ParseError:
            errors.append(f"Relación XML inválida: {rel_name}")
            continue

        owner_dir = posixpath.dirname(posixpath.dirname(rel_name))
        for rel in root.findall("r:Relationship", REL_NS):
            target = rel.attrib.get("Target", "")
            if not target or rel.attrib.get("TargetMode") == "External":
                continue
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(posixpath.join(owner_dir, target))
            if resolved not in names:
                errors.append(
                    f"Relación rota: {rel_name} → {target}"
                )
    return errors


def _find_broken_formulas(wb) -> list[str]:
    found: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if (
                    isinstance(value, str)
                    and value.startswith("=")
                    and "#REF!" in value.upper()
                ):
                    found.append(f"{ws.title}!{cell.coordinate}")
    return found


def _find_external_formulas(wb) -> list[str]:
    found: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if (
                    isinstance(value, str)
                    and value.startswith("=")
                    and re.search(r"\[\d+\]", value)
                ):
                    found.append(f"{ws.title}!{cell.coordinate}")
    return found
