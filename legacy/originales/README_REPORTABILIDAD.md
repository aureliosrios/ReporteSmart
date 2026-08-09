# Organización de la reportabilidad RO

## Flujo de la fase local

```text
HTML por rol
   ↓
app.js normaliza el registro
   ↓
server.py guarda base_datos_ro_diaria.json
   ↓
generar_reportabilidad_excel.mjs consolida
   ↓
reportabilidad_consolidada.json + Reportabilidad_RO_Local.xlsx
```

## Responsabilidad de cada HTML

| HTML | Responsable | Registro principal | Impacto EVM |
|---|---|---|---|
| `tareador.html` | Tareador | Personas, HH, categoría y WBS | `AC_MO` |
| `almacenero.html` | Almacenero | Insumos, cantidades, PU y WBS | `AC_MAT` |
| `administradora.html` | Administradora | HM, equipos, servicios y WBS | `AC_EQP` / `AC_SUB` |
| `ing_campo.html` | Ing. de Campo | Partida, metrado, tramo y WBS | `EV_PRODUCCION` |
| `dashboard_ro.html` | Residencia / OT | Lectura EVM y brechas | Consulta |

## Contrato JSON

El contrato se encuentra en `reportabilidad_schema.json`. Cada registro diario usa los mismos campos: fecha, rol, WBS, código, detalle, cantidad, unidad, PU, costo, tipo EVM, origen HTML y estado de validación.

## Generar la salida local

Con el servidor detenido o después de guardar registros desde los HTML, ejecutar desde la carpeta del proyecto:

```powershell
& 'C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' .\generar_reportabilidad_excel.mjs
```

La salida queda en `outputs/reportabilidad_ro/`:

- `reportabilidad_consolidada.json`: JSON organizado para auditoría e integración.
- `Reportabilidad_RO_Local.xlsx`: libro con CONTROL, WBS, APU, registros diarios, EVM y catálogos.
- `preview_evm_wbs.png`: verificación visual del resumen EVM.

## Reglas de control

- El BAC se calcula desde `presupuesto_con_apu.json`.
- El PV se toma de `base_datos_ro_diaria.json`.
- Solo los registros `VALIDO` alimentan AC y EV.
- Los duplicados se conservan para trazabilidad, pero se marcan como `DUPLICADO` y no afectan los indicadores.
- El vínculo entre recursos y WBS se mantiene en cada registro.

## Etapas posteriores

1. Validar con nuevos registros reales de los cuatro formularios.
2. Publicar HTML, JS, CSS y JSON maestro en GitHub Pages.
3. Crear un endpoint seguro o Google Apps Script para recibir `REGISTROS_DIARIOS`.
4. Construir el dashboard de Google Sheets sobre el mismo modelo EVM.
