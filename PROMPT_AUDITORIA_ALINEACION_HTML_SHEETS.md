# Prompt Maestro de Nivelación y Auditoría Técnica: Alineación de Formularios HTML a Google Sheets

> **Instrucciones para el Docente / Alumno:**  
> Copia todo el bloque de texto que se encuentra a continuación y pégalo en tu asistente de Inteligencia Artificial junto con el código de tus archivos HTML y JavaScript (`app.js`). La IA auditará tu código e indicará si tu formulario está 100% alineado con la **Pestaña 04 (`04_LOG_FIELD_ENTRIES`)** de la Base de Datos Viva de Obra.

---

```markdown
# PROMPT DE AUDITORÍA TÉCNICA: ALINEACIÓN DE FORMULARIOS HTML A BASE DE DATOS EN GOOGLE SHEETS

Actúa como un Auditor de Software e Ingeniero de Costos de Obra. Tu objetivo es revisar mi código HTML y JavaScript (`app.js`) para verificar si mi portal móvil de captura de datos de campo está 100% ALINEADO con la Pestaña 04 (`04_LOG_FIELD_ENTRIES`) de la Base de Datos Viva de Obra.

## 1. REGLAS OBLIGATORIAS DEL CONTRATO DE DATOS (PESTAÑA 04)

Mi formulario HTML debe capturar ÚNICAMENTE la CANTIDAD y los datos de contexto. La base de datos en Google Sheets buscará el Precio Unitario (P.U.) y calculará los totales de manera automática.

El objeto JSON que mi script JS envía por HTTP POST al servidor / Google Apps Script debe tener la siguiente estructura exacta:

```json
{
  "id": "LOG-20260803-020",
  "fecha": "2026-08-03",
  "rol": "Almacenero",
  "wbs": "WBS-200",
  "codigoRecurso": "MAT_TUB_PVC_200",
  "detalle": "Tubería PVC UF DN 200mm Tramo Calle 1",
  "cantidad": 280.00,
  "unidad": "m",
  "tipo": "AC_MAT",
  "origen_html": "almacenero.html"
}
```

---

## 2. CHECKLIST DE AUDITORÍA (LISTA DE VERIFICACIÓN CAMPO POR CAMPO)

Analiza mi código HTML/JS e indica "CONFORME" o "REQUIERE AJUSTE" para cada uno de los siguientes 8 puntos:

1. **[ ] Campo Fecha (`fecha`)**:  
   ¿Utiliza un `<input type="date">` que devuelva formato ISO estandarizado `YYYY-MM-DD`?

2. **[ ] Campo Código WBS (`wbs`)**:  
   ¿Usa un `<select>` restringido con las opciones oficiales del proyecto (`WBS-100`, `WBS-200`, `WBS-300`, `WBS-400`)?

3. **[ ] Campo Código de Recurso/Partida (`codigoRecurso`)**:  
   ¿El valor del `<option>` coincide exactamente con las llaves maestras del proyecto (ej: `MO_OPERARIO`, `MAT_TUB_PVC_200`, `EQ_EXCAVADORA` o `01.02.02`)?

4. **[ ] Campo CANTIDAD (`cantidad`)**:  
   ¿Es un `<input type="number">` mayor a cero (`> 0`) que captura solo la cantidad física, Horas Hombre (HH) o Horas Máquina (HM)?

5. **[ ] Exclusión de P.U. desde el cliente**:  
   ¿Mi HTML/JS NO envía el Precio Unitario ni el Subtotal al servidor (ya que esto se calcula por fórmulas vivas en Google Sheets)?

6. **[ ] Campo Unidad Medida (`unidad`)**:  
   ¿La unidad se asigna automáticamente al seleccionar el recurso (`hh`, `hm`, `m`, `m3`, `und`, `glb`, `pza`)?

7. **[ ] Campo Categoría EVM (`tipo`)**:  
   ¿El objeto asigna la categoría correcta según el rol emisor?
   - `AC_MO` (Tareador)
   - `AC_MAT` (Almacenero)
   - `AC_EQP` / `AC_SUB` (Administradora)
   - `EV_PRODUCCION` (Ing. de Campo)

8. **[ ] Envío JSON Payload**:  
   ¿Al presionar "Guardar", el script arma el JSON de forma limpia mediante `fetch(URL, { method: 'POST', body: JSON.stringify(data) })`?

---

## 3. INSTRUCCIONES DE SALIDA DE LA AUDITORÍA

Si encuentras algún error o desalineación en mi código HTML/JS:
1. Señala la línea exacta del código a modificar.
2. Explica el motivo de la desalineación con la Pestaña 04 de Google Sheets.
3. Proporciona el bloque de código HTML/JS corregido que garantiza la sincronización perfecta con la Base de Datos.
```
