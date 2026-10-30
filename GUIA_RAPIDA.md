# Guía rápida — Cartera Auto (COLBEEF)

## En 3 pasos

1. **Doble clic** en `INICIAR_CARTERA_AUTO.bat`
2. En el navegador, **sube** los Excel de SIIMED:
   - `Cartera a corte …xlsx` (obligatorio)
   - `Cartera por cuenta …xlsx` (recomendado)
   - `Balance … cartera …xlsx` (opcional, rotación)
3. Pulsa **Generar informe** y descarga Excel + PowerPoint

## Qué genera automáticamente

| Hoja / archivo | Contenido |
|----------------|-----------|
| **consolidada** | Clientes, edades, PLAZO, totales |
| **unidad** | Totales por cuenta + resumen MM + **EDADES** (fórmulas) |
| **comportamiento** | Histórico mensual por cliente + columna del mes |
| **informe jf** | Record por tipo de cliente (enlace a unidad) |
| **inf cartera x edades** | Histórico con vínculos a `unidad` (formato millones) |
| **Record cvda** | Evolución vencida/corriente |
| **PERIODO DE RECAUDO** | Rotación (fórmulas Formula!) |
| **carnes / local / …** | Detalle por unidad (CUPO, PRINCIPAL desde maestro) |
| **PowerPoint** | KPIs, gráficos y tablas editables (Cartera) |

## Revisar antes de junta (10–15 min)

- [ ] Total cartera = SIIMED (± redondeo)
- [ ] Abrir Excel y dejar recalcular (fórmulas de resumen / edades / Record)
- [ ] Columna del mes en **comportamiento** e **informe jf**
- [ ] Slides Cartera (edades, unidades, record)
- [ ] Acuerdos de pago / jurídica (manual)
- [ ] Slides **Tesorería** (aún no automatizadas)

## Maestro de clientes (`data/maestro_clientes.json`)

Datos que **no vienen de SIIMED** y se conservan mes a mes:

- **plazo** → consolidada y comportamiento
- **cupo** y **condicion_pago** → hoja carnes
- **principal** (vendedor) → hoja local

**Cliente nuevo:** agregar entrada en el JSON (o copiar del informe manual del mes anterior):

```json
"900123456": {
  "nombre": "RAZON SOCIAL S.A.S.",
  "plazo": 30,
  "cupo": 50000000,
  "condicion_pago": "VENCIMIENTO 30 DIAS",
  "principal": "NOMBRE VENDEDOR"
}
```

Tras cada generación exitosa, el maestro se actualiza con NITs nuevos (conserva los campos manuales).

## Anticipos SIIMED

Si el Excel de cartera trae hoja `CARTERA -ANTICIPOS`, el programa descuenta anticipos **materiales** (≥ $500.000) de clientes **LOCAL** (ej. TECNOLOGIAS en mayo). Olímpica y cadenas mantienen saldo bruto.

## Difícil cobro en hoja carnes

Los clientes de la cuenta **DIFÍCIL COBRO** (130598) **no** se listan en la hoja `carnes` (criterio de agosto 2026). Siguen incluidos en consolidada, unidad y en el total general. Si un mes se quieren mostrar, agregarlos a mano en el Excel.

## Archivos base del mes

Usar exportes SIIMED **del mismo corte y misma fecha de proceso**. Si `cartera a corte` y `cartera por cuenta` se regeneran en fechas distintas, los saldos pueden no cuadrar entre sí.

## Salida actual

Carpeta operativa: `output/ACTUAL/`  
Originales de referencia: `data/originales_cartera/`  
Plantilla histórica: `templates/plantilla_informe_cartera.xlsx`

**Gráficas del Excel (como el original):**
- `inf cartera x edades` — CARTERA POR EDADES
- `informe jf` — por tipo de cliente + histórico
- `Record cvda` — 2 series (general y sin Olímpica)
- `PERIODO DE RECAUDO` — area + barras del mes
- hoja `2023` — record vencida

Si Excel muestra “Reparar”, **no guardes** ese archivo (borra las gráficas). Usa el de `output/ACTUAL/`.

## Base histórica

Cada mes exitoso actualiza la plantilla para apilar el siguiente mes (Abr→Jul listos para **agosto**).

El Excel generado replica el original en datos/fórmulas/formatos y conserva las 7 gráficas estadísticas.

## Validado — mayo 2026

Referencia oficial: `data/originales_cartera/mayo_V1.xlsx`

- Total: **$3.699,0 MM** | Vencida: **$290,2 MM (7,8%)** — **10/10** unidades = informe manual
- Anticipo LOCAL TECNOLOGIAS aplicado (−$783.296)
- Resumen + EDADES (filas 14–27) con mismas fórmulas y formato millones
- `inf cartera x edades` enlaza a `unidad` (sin ####)

## Validado — julio 2026

Referencia oficial: `data/originales_cartera/julio_V1.xlsx`  
Último informe generado: revisar `output/ACTUAL/`

- Total: **$3.078,7 MM** | Vencida: **$391,3 MM (12,7%)**
- Cuadra con informe manual Cartera y SIIMED
- Validación automática: **7 gráficas, 5 drawings, sin `#REF!` ni vínculos externos**

## Seguridad de generación

- Primero se genera en una carpeta temporal.
- Solo se publica en `output/ACTUAL/` si supera todas las validaciones.
- La versión anterior pasa a `output/ARCHIVO/`.
- La plantilla se actualiza al final, con respaldo y reemplazo atómico.
- Si la plantilla está abierta en Excel, el informe se entrega pero se muestra una advertencia.

## Pendiente (fuera de Cartera)

- Automatizar slides **Tesorería** en PPT
- Validar **agosto 2026** cuando estén los archivos SIIMED

## Soporte

Si SIIMED cambia columnas del exporte, avisar para ajustar el lector.
