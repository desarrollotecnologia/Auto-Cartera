# -*- coding: utf-8 -*-
"""
Cartera Auto — Prototipo para el área de Cartera COLBEEF
Sube exportes SIIMED y descarga Excel + PowerPoint listos para junta.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cartera_auto.pipeline import run_pipeline
from cartera_auto.parsers import format_corte_label
from cartera_auto.transform import money_mm, pct
from cartera_auto.validate import FileValidationError, detect_siimed_type, LABELS


st.set_page_config(
    page_title="Cartera Auto · COLBEEF",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
<style>
  .stApp { background: linear-gradient(160deg, #f4faf4 0%, #e8f5e9 50%, #f7fbf7 100%); }
  h1, h2, h3 { color: #185E22 !important; font-family: 'Century Gothic', Calibri, sans-serif !important; }
  div[data-testid="stMetric"] {
    background: #fff; border: 1px solid #c8e6c9; border-radius: 12px; padding: 0.6rem 0.8rem;
  }
  .stButton>button { background: #289B38 !important; color: white !important; border: none !important; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("Cartera Auto")
st.caption("Prototipo · De exportes SIIMED a Excel gerencial + presentación de junta")

with st.expander("¿Cómo usarlo?", expanded=True):
    st.markdown(
        """
1. En SIIMED descarga los reportes del mes (edades de cartera, cartera por unidad y, opcional, balance de ingresos).
2. Súbelos aquí.
3. Pulsa **Generar informe**.
4. Descarga el **Excel** y el **PowerPoint** listos para revisión y junta.

**Archivos mínimos:** `cartera a …xlsx` (edades por cliente).  
**Recomendados:** + `cartera por unidad…xlsx` + `balance ingresos…xlsx`.
"""
    )

col1, col2, col3 = st.columns(3)
with col1:
    f_cartera = st.file_uploader(
        "1. Cartera por edades (obligatorio)",
        type=["xlsx"],
        help="Ej: cartera a 30 junio.xlsx — edades por cliente/NIT",
    )
with col2:
    f_unidad = st.file_uploader(
        "2. Cartera por unidad (recomendado)",
        type=["xlsx"],
        help="Ej: cartera por unidad 30 junio.xlsx — totales por cuenta",
    )
with col3:
    f_balance = st.file_uploader(
        "3. Balance ingresos (opcional)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Un balance único, o los dos por separado (CTA 13 y CTA 4)",
    )

# Vista previa de detección al subir
if f_cartera or f_unidad or f_balance:
    st.markdown("**Verificación de archivos**")
    vc1, vc2, vc3 = st.columns(3)
    with tempfile.TemporaryDirectory() as tmp_preview:
        tmp_preview = Path(tmp_preview)

        def _preview(uploaded, expected_key, col):
            if uploaded is None:
                col.caption("—")
                return
            p = tmp_preview / uploaded.name
            p.write_bytes(uploaded.getvalue())
            try:
                detected = detect_siimed_type(p)
                ok = (
                    (expected_key == "cartera" and detected == "cartera_edades")
                    or (expected_key == "unidad" and detected == "cartera_unidad")
                    or (expected_key == "balance" and detected == "balance")
                )
                if ok:
                    col.success(f"OK · {LABELS[detected]}")
                else:
                    col.error(f"No coincide · {LABELS.get(detected, detected)}")
            except Exception as e:
                col.warning(f"No se pudo leer: {e}")

        _preview(f_cartera, "cartera", vc1)
        _preview(f_unidad, "unidad", vc2)
        for f in f_balance or []:
            _preview(f, "balance", vc3)
        if not f_balance:
            vc3.caption("—")

run = st.button("Generar informe", type="primary", disabled=f_cartera is None)

if run and f_cartera is not None:
    with st.spinner("Validando archivos SIIMED y armando Excel + PPT…"):
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp = Path(tmp)
                p_cartera = tmp / f_cartera.name
                p_cartera.write_bytes(f_cartera.getvalue())

                p_unidad = None
                if f_unidad is not None:
                    p_unidad = tmp / f_unidad.name
                    p_unidad.write_bytes(f_unidad.getvalue())

                p_balance = []
                for f in f_balance or []:
                    p = tmp / f.name
                    p.write_bytes(f.getvalue())
                    p_balance.append(p)

                out_dir = ROOT / "output"
                result = run_pipeline(
                    cartera_path=p_cartera,
                    unidad_path=p_unidad,
                    balance_path=p_balance or None,
                    output_dir=out_dir,
                )

            k = result["kpis"]
            corte = format_corte_label(result["meta"].get("corte"))
            st.success(f"Informe generado · Corte {corte}")
            validation = result.get("validation") or {}
            if validation.get("ok"):
                st.success(
                    f"Excel validado · {validation.get('charts', 0)} gráficas "
                    f"· {validation.get('drawings', 0)} dibujos · sin #REF!"
                )
            for w in result.get("warnings") or []:
                st.warning(w)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total cartera", money_mm(k["total_cartera"]))
            m2.metric("Vencida", money_mm(k["cartera_vencida"]), pct(k["pct_vencida"]))
            m3.metric("Sin Olímpica", money_mm(k["total_sin_olimpica"]))
            m4.metric("Clientes", k["n_clientes"])

            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "Descargar Excel gerencial",
                    data=result["excel"].read_bytes(),
                    file_name=result["excel"].name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.caption(str(result["excel"]))
            with c2:
                st.download_button(
                    "Descargar PowerPoint junta",
                    data=result["pptx"].read_bytes(),
                    file_name=result["pptx"].name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
                st.caption(str(result["pptx"]))

            st.info(
                "Revise en 10–15 min: comentarios cualitativos, acuerdos de pago y slides de Tesorería "
                "(aún no automatizadas en este prototipo)."
            )
        except FileValidationError as e:
            st.error("No se generó el informe (archivo incorrecto o en casilla equivocada).")
            st.markdown(str(e).replace("\n", "\n\n"))
            st.warning(
                "La plantilla histórica **no se modificó**. Corrige los archivos y vuelve a intentar."
            )
        except Exception as e:
            st.error(f"Error al procesar: {e}")
            st.exception(e)

st.divider()
st.markdown(
    "**Alcance del prototipo:** sección Cartera (edades, sin Olímpica, unidades, detalle por canal). "
    "Tesorería (flujo de caja, deuda, cupos, fiducias) se agrega cuando entreguen esos insumos SIIMED.\n\n"
    "**Seguridad:** si subes un archivo en la casilla equivocada, el sistema lo detecta y **bloquea** la generación "
    "para no contaminar el histórico."
)
