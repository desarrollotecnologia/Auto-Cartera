# -*- coding: utf-8 -*-
"""Configuración del prototipo Cartera Auto."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "output"

# NIT Olímpica (cliente a excluir en vista "sin Olímpica")
OLIMPICA_NIT = "890107487"

# Mapeo cuenta SIIMED → unidad de negocio
ACCOUNT_UNIT_MAP = {
    "130503": "SUBPRODUCTOS",
    "130505": "OTRAS CxC",
    "13050901": "LOCAL",
    "13050902": "REGIONAL",
    "13050903": "NACIONAL",
    "13050904": "INSTITUCIONAL",
    "13050920": "DESPOSTE",
    "13051401": "CARNES COLBEEF",
    "130514": "CARNES COLBEEF",
    "13059805": "DIFICIL COBRO",
    "130598": "DIFICIL COBRO",
}

UNIT_ORDER = [
    "SUBPRODUCTOS",
    "OTRAS CxC",
    "LOCAL",
    "REGIONAL",
    "NACIONAL",
    "INSTITUCIONAL",
    "DESPOSTE",
    "CARNES COLBEEF",
    "DIFICIL COBRO",
]

# Buckets de edad presentables (orden junta)
AGE_BUCKETS = [
    "CORRIENTE",
    "1 AL 7",
    "8 AL 15",
    "16 AL 30",
    "31 AL 60",
    "61 AL 90",
    "91 AL 180",
    "181 AL 360",
    "361+",
]

# Colores corporativos Colbeef (rojo carne / carbón)
COLORS = {
    "primary": "185E22",
    "primary_dark": "0F3D16",
    "accent": "70AD47",
    "text": "000000",
    "muted": "555555",
    "bg_light": "E8F5E9",
    "bg_header": "289B38",
    "white": "FFFFFF",
    "row_alt": "E8F5E9",
    "danger": "C01E1E",
    "ok": "17AD22",
}

COMPANY_NAME = "SIIMED S.A.S. - COLBEEF S.A.S."
REPORT_TITLE = "INFORME DE CARTERA"
