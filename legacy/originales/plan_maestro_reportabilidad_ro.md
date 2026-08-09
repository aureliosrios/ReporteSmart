# Plan Maestro de Implementación de Reportabilidad RO

**Proyecto:** Redes Sanitarias de Agua Potable y Alcantarillado - Habilitación Urbana Los Cedros  
**Objetivo:** evolucionar la reportabilidad desde formularios HTML locales hasta un sistema publicado en GitHub, con almacenamiento en Google Sheets y dashboard EVM.

## 1. Alcance y principio de implementación

El sistema se implementará por etapas. Cada etapa debe validarse antes de iniciar la siguiente.

```text
HTML por rol
    ↓
JSON normalizado
    ↓
Excel local auditable
    ↓
GitHub Pages
    ↓
Google Apps Script Web App
    ↓
Google Sheets
    ↓
Dashboard EVM y análisis de brechas
```

La regla principal es mantener el mismo contrato de datos en todas las etapas. El registro no debe cambiar de significado al pasar de HTML a JSON, Excel, GitHub o Google Sheets.

## 2. Fuentes HTML y responsabilidades

| Archivo | Responsable | Información capturada | Tipo EVM |
|---|---|---|---|
| `tareador.html` | Tareador | Trabajador, categoría, personas, horas y WBS | `AC_MO` |
| `almacenero.html` | Almacenero | Material, cantidad, PU, destino WBS | `AC_MAT` |
| `administradora.html` | Administradora | Equipo, servicio, HM/unidades, PU y WBS | `AC_EQP` / `AC_SUB` |
| `ing_campo.html` | Ing. de Campo | Partida, tramo, metrado y WBS | `EV_PRODUCCION` |
| `dashboard_ro.html` | Residente / OT | PV, EV, AC, CPI, SPI, EAC y brechas | Consulta |

## 3. Contrato único de datos

El contrato se encuentra en `reportabilidad_schema.json`.

Cada registro debe contener como mínimo:

```json
{
  "id": "ID único",
  "fecha": "2026-08-02",
  "rol": "Almacenero",
  "wbs": "WBS-200",
  "codigoRecurso": "MAT_TUB_PVC_200",
  "detalle": "Tubería PVC DN 200 - Tramo Calle 1",
  "cantidad": 100,
  "unidad": "m",
  "pu": 42.00,
  "costo": 4200.00,
  "tipo": "AC_MAT",
  "origen_html": "almacenero.html",
  "estado_validacion": "VALIDO"
}
```

### Reglas de negocio

1. Todo registro debe tener fecha, WBS, código, cantidad, unidad y tipo EVM.
2. El costo se calcula como `cantidad × pu`.
3. El EV solo proviene de `ing_campo.html`.
4. El AC proviene de tareo, almacén, equipos y servicios.
5. Los registros duplicados se conservan para trazabilidad, pero no participan en los indicadores.
6. El BAC proviene del APU meta; el PV proviene de la línea base programada.
7. Cada registro debe conservar el HTML de origen.

## 4. Etapa 0: línea base y catálogos

### Entradas

- `presupuesto_con_apu.json`
- `base_datos_ro_diaria.json`
- `presupuesto_ro_sintetico.json`
- `memoria_descriptiva_sintetica.md`

### Actividades

1. Confirmar los cuatro WBS:
   - `WBS-100`: Obras preliminares.
   - `WBS-200`: Alcantarillado, zanjas y buzones.
   - `WBS-300`: Agua potable y conexiones.
   - `WBS-400`: Pruebas, desinfección y entrega.
2. Validar que el BAC del APU coincida con la suma de partidas.
3. Validar precios unitarios de mano de obra, materiales, equipos y subcontratos.
4. Mantener un único catálogo maestro para todos los HTML.

### Entregable

`presupuesto_con_apu.json` como fuente económica principal y `reportabilidad_schema.json` como contrato de integración.

## 5. Etapa 1: captura local en HTML

### Actividades

1. El usuario abre el portal correspondiente desde `index.html`.
2. El formulario permite seleccionar WBS y recurso/partida.
3. `app.js` calcula PU, costo y tipo EVM.
4. El usuario guarda el registro.
5. El backend local recibe el registro mediante `/api/save-log`.

### Validaciones

- La fila no puede quedar sin WBS.
- La cantidad debe ser mayor que cero.
- El PU debe provenir del catálogo maestro.
- El costo debe coincidir con cantidad por PU.
- El ID debe ser único para evitar doble registro.

### Entregable

`base_datos_ro_diaria.json` actualizado con registros normalizados.

## 6. Etapa 2: consolidación JSON y Excel local

### Script

`generar_reportabilidad_excel.mjs`

### Hojas generadas

1. `CONTROL`: estado general de la carga y KPI globales.
2. `WBS`: catálogo, BAC y PV por frente.
3. `APU_PRESUPUESTO`: partidas, metrados, PU y BAC.
4. `REGISTROS_DIARIOS`: registros provenientes de los HTML.
5. `EVM_WBS`: PV, EV, AC, CV, CPI, SPI, EAC y estado.
6. `CATALOGOS`: recursos maestros para los formularios.

### Comando de generación

```powershell
& 'C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' .\generar_reportabilidad_excel.mjs
```

### Criterio de aprobación

No se continúa a GitHub hasta que:

- el JSON sea válido;
- no existan errores de fórmula;
- los duplicados estén identificados;
- BAC, PV, EV y AC reconcilien con los registros;
- el Excel sea legible y auditable.

## 7. Etapa 3: publicación en GitHub

### Archivos que se publicarán

- `index.html`
- `tareador.html`
- `almacenero.html`
- `administradora.html`
- `ing_campo.html`
- `dashboard_ro.html`
- `app.js`
- `style.css`
- `presupuesto_con_apu.json`
- `reportabilidad_schema.json`

### Archivos que no deben publicarse sin revisión

- Contraseñas.
- Tokens.
- Claves privadas.
- Datos personales sensibles.
- Credenciales de Google.

### Actividades

1. Crear repositorio GitHub.
2. Subir los archivos del portal.
3. Activar GitHub Pages.
4. Probar cada formulario desde la URL pública.
5. Verificar que el catálogo JSON cargue correctamente.
6. Verificar que el navegador pueda comunicarse con el endpoint de Google Apps Script.

### Restricción importante

GitHub Pages es un hosting estático. No debe guardar directamente credenciales de Google ni conectarse con una cuenta de servicio desde el navegador. La escritura en Google Sheets debe pasar por un endpoint controlado de Google Apps Script.

## 8. Etapa 4: Google Sheets como base de reportabilidad

### Estructura recomendada del libro

Crear las siguientes hojas:

- `REGISTROS_DIARIOS`: tabla principal recibida desde los HTML.
- `WBS`: catálogo de frentes y línea base.
- `APU_PRESUPUESTO`: partidas y BAC.
- `EVM_WBS`: indicadores calculados.
- `CONTROL_CARGA`: bitácora de envíos, errores y duplicados.

### Columnas de `REGISTROS_DIARIOS`

```text
id | fecha | rol | wbs | codigoRecurso | detalle | cantidad | unidad | pu | costo | tipo | origen_html | estado_validacion | fecha_recepcion
```

### Flujo de escritura

1. El HTML crea un objeto JSON.
2. `app.js` envía el JSON mediante `fetch()` al Web App de Apps Script.
3. Apps Script valida campos y tipo EVM.
4. Apps Script busca el ID en `REGISTROS_DIARIOS`.
5. Si el ID ya existe, registra el evento como duplicado.
6. Si es válido, agrega una fila.
7. Apps Script devuelve una respuesta JSON al HTML.

## 9. Etapa 5: Dashboard en Google Sheets

El dashboard debe leer exclusivamente las hojas de Google Sheets y no recalcular valores manualmente en el HTML.

### Indicadores

```text
EV = Σ registros tipo EV_PRODUCCION
AC = Σ registros tipo AC_MO + AC_MAT + AC_EQP + AC_SUB
CPI = EV / AC
SPI = EV / PV
EAC = AC + (BAC - EV) / CPI
CV = EV - AC
SV = EV - PV
```

### Vistas recomendadas

- KPI globales del proyecto.
- Tabla EVM por WBS.
- Costo real por componente.
- Avance físico por partida.
- Tendencia semanal de PV, EV y AC.
- Alertas para CPI o SPI menores que 1.00.

## 10. Script final de Google Apps Script

Este script se instala en Google Sheets mediante **Extensiones → Apps Script**. Luego se publica como **Aplicación web** con acceso autorizado según la política de la organización.

El HTML deberá enviar un arreglo JSON al endpoint publicado.

```javascript
const CONFIG = {
  registrosSheet: 'REGISTROS_DIARIOS',
  controlSheet: 'CONTROL_CARGA',
  idColumn: 1
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const registros = Array.isArray(body) ? body : [body];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registrosSheet = ss.getSheetByName(CONFIG.registrosSheet);
    const controlSheet = ss.getSheetByName(CONFIG.controlSheet);

    if (!registrosSheet) {
      throw new Error('No existe la hoja REGISTROS_DIARIOS');
    }

    const existentes = obtenerIds_(registrosSheet);
    const filas = [];
    const eventos = [];

    registros.forEach((r) => {
      const validacion = validarRegistro_(r);
      const id = String(r.id || '').trim();

      if (!validacion.ok) {
        eventos.push([new Date(), id, 'RECHAZADO', validacion.mensaje]);
        return;
      }

      if (existentes.has(id)) {
        eventos.push([new Date(), id, 'DUPLICADO', 'El ID ya existe']);
        return;
      }

      filas.push([
        id,
        r.fecha,
        r.rol,
        r.wbs,
        r.codigoRecurso,
        r.detalle || '',
        Number(r.cantidad),
        r.unidad,
        Number(r.pu),
        Number(r.costo),
        r.tipo,
        r.origen_html || '',
        'VALIDO',
        new Date()
      ]);
      existentes.add(id);
    });

    if (filas.length > 0) {
      registrosSheet
        .getRange(registrosSheet.getLastRow() + 1, 1, filas.length, filas[0].length)
        .setValues(filas);
    }

    if (controlSheet && eventos.length > 0) {
      controlSheet
        .getRange(controlSheet.getLastRow() + 1, 1, eventos.length, eventos[0].length)
        .setValues(eventos);
    }

    return respuesta_({
      ok: true,
      recibidos: registros.length,
      guardados: filas.length,
      observados: eventos.length
    });
  } catch (error) {
    return respuesta_({ ok: false, error: error.message });
  } finally {
    lock.releaseLock();
  }
}

function validarRegistro_(r) {
  const obligatorios = ['id', 'fecha', 'rol', 'wbs', 'codigoRecurso', 'unidad', 'tipo'];
  const faltantes = obligatorios.filter((campo) => r[campo] === undefined || r[campo] === '');

  if (faltantes.length > 0) {
    return { ok: false, mensaje: 'Faltan campos: ' + faltantes.join(', ') };
  }

  if (!(Number(r.cantidad) > 0) || !(Number(r.pu) >= 0) || !(Number(r.costo) >= 0)) {
    return { ok: false, mensaje: 'Cantidad, PU o costo inválido' };
  }

  const tipos = ['AC_MO', 'AC_MAT', 'AC_EQP', 'AC_SUB', 'EV_PRODUCCION'];
  if (!tipos.includes(r.tipo)) {
    return { ok: false, mensaje: 'Tipo EVM no permitido: ' + r.tipo };
  }

  return { ok: true };
}

function obtenerIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  const values = sheet.getRange(2, CONFIG.idColumn, lastRow - 1, 1).getValues();
  return new Set(values.map((row) => String(row[0]).trim()).filter(Boolean));
}

function respuesta_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## 11. Orden de ejecución recomendado

1. Probar los cuatro HTML en local.
2. Generar y revisar `reportabilidad_consolidada.json`.
3. Revisar `Reportabilidad_RO_Local.xlsx`.
4. Corregir duplicados, catálogos y reglas de negocio.
5. Publicar el portal en GitHub Pages.
6. Crear el libro de Google Sheets.
7. Crear las hojas y encabezados definidos.
8. Publicar el Apps Script como Web App.
9. Configurar la URL del Web App en `app.js`.
10. Probar un registro por cada rol.
11. Validar la llegada a Google Sheets.
12. Crear el dashboard final sobre las hojas EVM.

El proyecto solo se considera completo cuando un registro creado desde cada HTML llega a Google Sheets, conserva su trazabilidad y actualiza correctamente AC, EV y los indicadores EVM.
