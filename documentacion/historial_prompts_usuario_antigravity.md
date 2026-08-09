# 📜 Guía y Catálogo de Prompts del Sistema de Control de Resultado Operativo (RO)
## Interacción Usuario ↔ Antigravity AI | Proyecto: Redes Sanitarias "Los Cedros"

> [!NOTE]
> Este documento cataloga de forma estructurada los **18 prompts corregidos gramaticalmente** utilizados durante el desarrollo del sistema de Control de Resultado Operativo (RO). Para cada prompt se detalla su **propósito funcional**, **objetivo técnico** y el **entregable tangible** generado en el proyecto.

---

## 📌 Fase 1: Alineamiento Estratégico, Definición de Roles y Reglas EVM

### 💬 Prompt 1: Planteamiento de la Estrategia de Control RO
> *"El cliente desea que tengamos un Control de Resultado Operativo (RO). ¿Qué me planteas para tener el control de Valor Ganado ($EV$), Valor Planificado ($PV$), Costo Real ($AC$) y demás indicadores, teniendo en cuenta la estructura de personal que tiene la obra?"*

* **¿Para qué sirve?**: Para establecer la estrategia general de Gestión de Valor Ganado (EVM) adaptada a la estructura de personal de la obra (7 profesionales).
* **Objetivo Técnico**: Definir la arquitectura de cálculo de los indicadores $EV, PV, AC, CPI, SPI, EAC$ y la segregación de fuentes de datos.
* **Entregable Generado**: Documento de Memoria Descriptiva Sintética (`memoria_descriptiva_sintetica.md`) y estructura base del presupuesto en JSON (`presupuesto_ro_sintetico.json`).

---

### 💬 Prompt 2: Definición del Canal de Reportabilidad Web y Almacenamiento
> *"Tenemos que dar solución desde la reportabilidad. Me imagino tener un HTML de reporte diario que alimente a una base de datos. El HTML puede estar alojado en GitHub y desde allí analizar el Costo Real ($AC$) y el Valor Ganado ($EV$) comparados con el Valor Planificado ($PV$) ya definido."*

* **¿Para qué sirve?**: Para fijar el modelo de captura de datos mediante interfaces ligeras HTML compatibles con GitHub Pages/servidores locales.
* **Objetivo Técnico**: Establecer la arquitectura web cliente-servidor descentralizada alimentada por formularios web y sincronizada a una base de datos.
* **Entregable Generado**: Propuesta arquitectónica del portal web de captura diaria y plantilla inicial de visualización.

---

### 💬 Prompt 3: Matriz de Responsabilidades y Roles de Campo
> *"No te vayas muy adelante. Primero tenemos que definir los roles de obra: quién reporta el personal y el frente de trabajo, quién reporta el consumo de materiales, quién mide el avance diario de producción y cómo se debe realizar dicha reportabilidad."*

* **¿Para qué sirve?**: Para asignar las responsabilidades operativas exactas de captura de información a los 4 roles clave del proyecto.
* **Objetivo Técnico**: Crear la matriz RACI de reportabilidad: Tareador ($MO$), Almacenero ($MAT$), Administradora ($EQP$) e Ingeniero de Campo ($EV$).
* **Entregable Generado**: Asignación formal de roles y responsabilidades en el Plan de Trabajo.

---

### 💬 Prompt 4: Análisis y Estructuración de la EDT / WBS
> *"¿Cuáles son las plantillas de reporte recomendadas? Para ello debemos definir si dividimos el proyecto por frentes de trabajo / WBS, o si el proyecto es lo suficientemente pequeño como para trabajarlo con un solo frente. Analiza la envergadura del proyecto antes de fraccionarlo en nodos WBS."*

* **¿Para qué sirve?**: Para evaluar la envergadura de la obra (60 días / 2 meses) y determinar el nivel óptimo de descomposición de la WBS.
* **Objetivo Técnico**: Definir los 4 nodos principales WBS (`WBS-100` Obras Preliminares, `WBS-200` Alcantarillado, `WBS-300` Agua Potable, `WBS-400` Pruebas Hidráulicas).
* **Entregable Generado**: Estructura WBS de 4 nodos integrada en los archivos JSON y Excel.

---

### 💬 Prompt 5: Regla Operativa de Imputación Directa al WBS
> *"El tareador tiene que definir el frente de trabajo al que está asignado el personal, el cual corresponderá al código WBS, al igual que el reporte de equipos y materiales. Por tanto, cada reportador debe conocer la estructura WBS. No es necesario asignar recursos a la actividad, sería muy engorroso; es suficiente asignar recursos al WBS para el AC."*

* **¿Para qué sirve?**: Para simplificar el trabajo de campo de los reportadores, evitando la asignación microscópica de recursos por partida.
* **Objetivo Técnico**: Regla de negocio EVM: Imputación directa de Costos Reales ($AC$) únicamente a nivel de Nodos WBS ($AC_{\text{Nodo}}$) y medición de $EV$ por partida.
* **Entregable Generado**: Regla de imputación directa configurada en la lógica de validación de los portales web.

---

### 💬 Prompt 6: Plan de Trabajo y Diagrama de Flujo del RO
> *"Teniendo las cosas claras, elaboremos un plan de trabajo en formato Markdown (`.md`) e incluye un diagrama de flujo visual sobre cómo proceder."*

* **¿Para qué sirve?**: Para consolidar el mapa de ruta de implementación del RO y visualizar el flujo de información de campo.
* **Objetivo Técnico**: Redactar el Plan de Trabajo integral en 4 fases (Pre-obra, Captura Diaria, Consolidación Semanal, Gobierno de Obra).
* **Entregable Generado**: Documento del Plan de Trabajo (`plan_de_trabajo_control_ro.md`).

---

## 🎨 Fase 2: Visualización del Flujograma y Desarrollo del Portal Web Multi-Frente

### 💬 Prompt 7: Corrección de Visualización del Diagrama
> *"No se visualiza la imagen del diagrama de flujo en el documento Markdown."*

* **¿Para qué sirve?**: Detectar que la representación gráfica del flujo requería renderizado en formato de imagen accesible.
* **Objetivo Técnico**: Identificar la necesidad de compilar código gráfico a una imagen estándar `.png`.
* **Entregable Generado**: Diagnóstico de conversión de sintaxis Mermaid a imagen rasterizada.

---

### 💬 Prompt 8: Generación Automatizada del Diagrama en Imagen PNG
> *"Solo se muestran los bloques de código en texto Mermaid (MMD), pero no el gráfico visual renderizado."*

* **¿Para qué sirve?**: Generar una imagen física visible del diagrama de flujo ejecutable en cualquier visor Markdown.
* **Objetivo Técnico**: Crear un script en Python (`generar_diagrama_flujo.py`) utilizando `matplotlib` para dibujar y guardar el diagrama de flujo en PNG.
* **Entregable Generado**: Imagen en alta resolución `diagrama_flujo_ro.png` enlazada en el plan de trabajo.

---

### 💬 Prompt 9: Desarrollo del Portal Web de Reportabilidad HTML/JS
> *"Después de analizar el plan de trabajo, genera una propuesta de aplicativo/reporte diario en HTML para el tareador, almacenero, ingeniero de campo y administradora, tomando como base los insumos y partidas del presupuesto en JSON."*

* **¿Para qué sirve?**: Crear la primera versión del aplicativo web interactivo para captura diaria de datos en obra.
* **Objetivo Técnico**: Construir la interfaz `index.html`, hoja de estilos `style.css` y lógica `app.js` con servidor local HTTP (`http://localhost:8080`).
* **Entregable Generado**: Aplicativo web inicial (`index.html`, `style.css`, `app.js`).

---

### 💬 Prompt 10: Flexibilidad Multi-Frente y Selección por Categorías/JSON
> *"Es importante que en el parte del tareador se pueda identificar la categoría de cada trabajador (Capataz, Operario, Oficial, Peón) y brinde la flexibilidad de reportar diferentes frentes de trabajo (WBS) en paralelo en un solo formulario. En el caso del almacenero, debe contar con el catálogo de insumos desde la base JSON para seleccionar y poder despachar a 2 o más frentes WBS. La administradora imputará los equipos y subcontratos desde el JSON a los frentes WBS. Finalmente, para el ingeniero de campo debe estar disponible todo el catálogo de partidas a reportar, permitiendo también trabajar en 2 o más frentes WBS en paralelo."*

* **¿Para qué sirve?**: Permitir la captura dinámica multi-fila y multi-frente en paralelo para los 4 roles profesionales de la obra.
* **Objetivo Técnico**: Implementar tablas dinámicas con dropdowns sincronizados desde `presupuesto_con_apu.json` y cálculo automático en vivo.
* **Entregable Generado**: Presupuesto enriquecido `presupuesto_con_apu.json` y motor dinámico de tablas en `app.js`.

---

### 💬 Prompt 11: Corrección de Carga Offline y Protocolo `file://`
> *"Estoy revisando el archivo `index.html`, sin embargo no veo que se active el botón de agregar WBS. Revisa el comportamiento del index."*

* **¿Para qué sirve?**: Resolver el bloqueo por políticas CORS cuando el usuario abre `index.html` mediante doble clic local (`file:///...`).
* **Objetivo Técnico**: Incorporar un diccionario de respaldo (fallback offline) dentro de `app.js` para asegurar funcionamiento incondicional sin servidor.
* **Entregable Generado**: Parche de compatibilidad offline en `app.js`.

---

### 💬 Prompt 12: Arquitectura Jerárquica por Bloques WBS
> *"Está bien incluir varios WBS por reporte diario; sin embargo, es importante que para cada bloque WBS se puedan reportar múltiples recursos ($MO, MAT, EQP$) o incluso múltiples actividades de Valor Ganado ($EV$)."*

* **¿Para qué sirve?**: Agrupar visualmente la información por tarjetas/bloques WBS conteniendo sub-tablas de recursos y partidas.
* **Objetivo Técnico**: Maquetación jerárquica: `Contenedor WBS` $\rightarrow$ `Sub-sección MO` + `Sub-sección MAT` + `Sub-sección EQP` + `Sub-sección EV`.
* **Entregable Generado**: Interfaz jerárquica de bloques WBS en `index.html` y `style.css`.

---

### 💬 Prompt 13: Ejecución y Validación de Resultados
> *"¿Puedes generar y ejecutar el resultado final con esta estructura?"*

* **¿Para qué sirve?**: Validar la compilación e integración fluida de los bloques WBS jerárquicos.
* **Objetivo Técnico**: Compilación del motor `app.js` con soporte para tarjetas de bloque WBS y resumen global de costos.
* **Entregable Generado**: Versión funcional validada de la interfaz jerárquica por WBS.

---

## 🗃️ Fase 3: Detalle de EV, Portales Celulares Independientes y Base de Datos Trazable en Excel

### 💬 Prompt 14: Unidades Explícitas en EV y Cuantificación de Costos por WBS
> *"En el formulario de Valor Ganado ($EV$), incluye de forma explícita la unidad de medida de cada actividad registrada ($m, m^3, und, glb$). Asimismo, para el control de costos, cuantifica numéricamente los gastos ($MO, MAT, EQP$) agrupados por cada nodo WBS."*

* **¿Para qué sirve?**: Garantizar la claridad física de los metrados reportados y presentar un cuadro de cuantificación de costos por componente.
* **Objetivo Técnico**: Agregar columna de `Unidad` en la sub-tabla de $EV$ y construir la tabla resumen de cuantificación ($AC_{\text{MO}}, AC_{\text{MAT}}, AC_{\text{EQP}}$ vs $EV$).
* **Entregable Generado**: Tabla de Cuantificación de Costos por WBS en la pestaña Dashboard de `index.html`.

---

### 💬 Prompt 15: Creación de Portales HTML Celulares Independientes por Rol
> *"La reportabilidad debe ser independiente para cada usuario; es decir, el tareador reportará desde su celular únicamente lo que le compete. Por lo tanto, requiero portales HTML independientes e individuales para cada uno de los roles de obra (`tareador.html`, `almacenero.html`, `administradora.html`, `ing_campo.html`, `dashboard_ro.html`)."*

* **¿Para qué sirve?**: Proporcionar a cada profesional de campo un aplicativo web exclusivo y ultraligero optimizado para smartphones.
* **Objetivo Técnico**: Crear 5 archivos HTML independientes conectados a la misma base de datos unificada (`localStorage` / JSON):
  - 📱 `tareador.html` (Mano de Obra)
  - 📦 `almacenero.html` (Materiales)
  - 🚜 `administradora.html` (Equipos y Servicios)
  - 👷‍♀️ `ing_campo.html` (Producción EV)
  - 📊 `dashboard_ro.html` (Sala de Control Residencia/OT)
  - 🏠 `index.html` (Hub Navegador de Acceso Directo)
* **Entregables Generados**: Los 5 archivos HTML independientes (`tareador.html`, `almacenero.html`, `administradora.html`, `ing_campo.html`, `dashboard_ro.html`).

---

### 💬 Prompt 16: Propuesta de Base de Datos en JSON y Excel
> *"Los archivos HTML están funcionando muy bien. Ahora necesitamos contar con una base de datos donde almacenar toda la información de los reportes capturados desde los HTML. Genera una propuesta de base de datos en JSON y también compila un libro de Excel que tome como insumo los datos de la reportabilidad."*

* **¿Para qué sirve?**: Almacenar de forma persistente y estructurada la totalidad de partes diarios capturados en obra.
* **Objetivo Técnico**: Crear el esquema JSON transaccional `base_datos_ro_diaria.json` y el script Python `generar_base_datos_excel.py` con `openpyxl`.
* **Entregables Generados**: Archivo de base de datos `base_datos_ro_diaria.json` y libro de Excel `Base_de_Datos_RO_Reportabilidad.xlsx`.

---

### 💬 Prompt 17: Cobertura Total de Recursos y Trazabilidad con Fórmulas de Excel
> *"En la base de datos debes contemplar la totalidad de los recursos del archivo JSON maestro ($MO, MAT, EQP, APU$), de modo que cuando se registre el reporte diario, la cuantificación sea 100% trazable y auditada mediante fórmulas dinámicas de Excel (`BUSCARV`, `SUMAR.SI.CONJUNTO`)."*

* **¿Para qué sirve?**: Permitir que cualquier auditoría de costos o revisión del Residente pueda trazar cada Sol gastado directamente hasta el catálogo maestro de precios y APUs.
* **Objetivo Técnico**: Configurar la hoja `MAESTRO_RECURSOS_Y_APU` con el 100% de insumos e implementar fórmulas vivas en Excel:
  - En `DB_REGISTROS_DIARIOS`: `=BUSCARV(código, MAESTRO_RECURSOS_Y_APU, 4, FALSO)` y `=Cantidad * Costo_Unitario`.
  - En `DASHBOARD_EVM_WBS`: `=SUMAR.SI.CONJUNTO(...)` para sumar $AC$ y $EV$ de forma dinámica.
* **Entregable Generado**: Libro de Excel auditado e interconectado con fórmulas dinámicas (`Base_de_Datos_RO_Reportabilidad.xlsx`).

---

### 💬 Prompt 18: Generación del Documento de Catálogo de Prompts
> *"¿Puedes generar un archivo Markdown (`.md`) con todos los prompts con los que hemos interactuado entre el Usuario y Antigravity, corrigiéndolos gramaticalmente sin perder el contexto?"*

* **¿Para qué sirve?**: Documentar la memoria técnica y guía de prompts de todo el proceso de desarrollo del sistema RO.
* **Objetivo Técnico**: Sintetizar y catalogar los 18 prompts con sus propósitos, objetivos técnicos y entregables tangibles asociados.
* **Entregable Generado**: Archivo de documentación `historial_prompts_usuario_antigravity.md`.
