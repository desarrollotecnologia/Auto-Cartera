# Cartera Auto — Prototipo para el área de Cartera (COLBEEF / SIIMED)

## ¿Qué resuelve?

Cada mes el área descarga reportes de **SIIMED**, arma un Excel gerencial a mano y pega capturas en PowerPoint para junta. Eso toma mucho tiempo.

**Cartera Auto** recibe los Excel de SIIMED y genera automáticamente:

1. **Excel gerencial** (resumen, edades, sin Olímpica, por unidad, detalle por canal)
2. **PowerPoint** listo para revisión/junta (sección Cartera)

## Diseño

El prototipo **clona las plantillas reales COLBEEF**:

- Excel ← `templates/plantilla_informe_cartera.xlsx` (tu Informe de Cartera V1)
- PowerPoint ← `templates/plantilla_ig_tesoreria.pptx` (tu IG Tesorería)

En el PPT se conservan portada, Tesorería y maestros Colbeef; en las slides de Cartera se reemplazan las capturas por **tablas editables** con la tipografía/colores verdes de la marca (Century Gothic / verde COLBEEF).

## Cómo usarlo (área de Cartera)

### Opción A — Aplicación visual (recomendada)

1. Doble clic en `INICIAR_CARTERA_AUTO.bat`
2. Se abre el navegador
3. Sube los archivos SIIMED del mes
4. Pulsa **Generar informe**
5. Descarga Excel + PowerPoint

### Opción B — Línea de comandos

```bat
python run_cli.py --cartera "C:\ruta\cartera a 30 junio.xlsx" --unidad "C:\ruta\cartera por unidad 30 junio.xlsx" --balance "C:\ruta\balance ingresos Cartera junio 06 2026.xlsx"
```

Los archivos quedan en la carpeta `output\`.

## Archivos SIIMED que pide

| Archivo | Obligatorio | Ejemplo |
|---------|-------------|---------|
| Cartera por edades | Sí | `cartera a 30 junio.xlsx` |
| Cartera por unidad | Recomendado | `cartera por unidad 30 junio.xlsx` |
| Balance ingresos | Opcional (rotación) | `balance ingresos Cartera junio….xlsx` |

## Requisitos

- Windows
- Python 3.10+ instalado ([python.org](https://www.python.org/downloads/))
- Al primer uso el `.bat` instala dependencias solas

```bat
python -m pip install -r requirements.txt
```

## Qué incluye el PowerPoint (prototipo)

- Portada e indicadores
- Cartera por edades (con / sin Olímpica)
- Indicador % vencida
- Cartera por unidad de negocio
- Consolidado top clientes
- Detalle por canal: Subproductos, Otras CxC, Local, Regional, Nacional, Institucional, Desposte, Carnes Colbeef, Difícil cobro

## Qué NO incluye aún (fase 2)

- Slides de **Tesorería** (flujo de caja, deuda, créditos, cupos, fiducias)
- Comentarios cualitativos / acuerdos de pago / jurídica
- Conexión directa a SIIMED (sigue siendo descarga manual de Excel)

Tras generar, el gestor debería revisar **10–15 minutos** (comentarios y slides no automáticas) antes de junta.

## Estructura del proyecto

```
presentacion/
  INICIAR_CARTERA_AUTO.bat   ← doble clic para abrir
  app.py                     ← interfaz
  run_cli.py                 ← línea de comandos
  cartera_auto/              ← motor (parsers, Excel, PPT)
  output/                    ← archivos generados
  requirements.txt
```

## Notas

- Olímpica se identifica por NIT `890107487` (configurable en `cartera_auto/config.py`).
- Las tablas del PPT son **editables** (ya no son pantallazos).
- Si SIIMED cambia el formato de columnas, avisar para ajustar el parser.
