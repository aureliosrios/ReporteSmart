# Plan de Trabajo: Implementación del Sistema Ágil de Control de Resultado Operativo (RO) con IA

**Proyecto:** Instalación de Redes Sanitarias de Agua Potable y Alcantarillado - Habilitación Urbana  
**Duración:** 60 Días Calendario (8 Semanas)  
**Objetivo:** Establecer el sistema de control ágil de avance (EV, PV) y costos (AC) a nivel de Resultado Operativo (RO) integrando Bildin, Almacén, Producción e Inteligencia Artificial para la detección temprana de brechas.

---

## 1. Estructura de Costos Simplificada en Obra (Catálogo WBS)

Para garantizar una alta adherencia operativa en campo sin sobrecargar de trabajo al personal, la captura de costos reales (AC) se realiza exclusivamente sobre **4 Nodos WBS**:

| Código WBS | Descripción del Frente / WBS de Costo | Partidas Incluidas (Para Medición de EV) |
| :---: | :--- | :--- |
| **`WBS-100`** | Obras Preliminares, Trazo y Movilización | Cartel de obra, caseta provisional, trazo y replanteo topográfico. |
| **`WBS-200`** | Red de Alcantarillado, Zanjas y Buzones | Excavación a máquina, cama de arena, tubería PVC 200mm, relleno compactado, buzones prefabricados y conexiones alcantarillado. |
| **`WBS-300`** | Red de Agua Potable y Conexiones | Tubería PVC 110mm, accesorios, abrazaderas y conexiones domiciliarias de agua. |
| **`WBS-400`** | Pruebas Hidráulicas y Desinfección | Prueba hidrostaticas, prueba de estanqueidad, desinfección con cloro y análisis de laboratorio. |

---

## 2. Diagrama de Flujo Operativo del Sistema

![Diagrama de Flujo Operativo del Resultado Operativo](diagrama_flujo_ro.png)

*Figura 1: Flujo de captura diaria simplificada en 4 Nodos WBS, consolidación de EVM y evaluación semanal de brechas con Inteligencia Artificial.*

---

## 3. Plan de Implementación Cronológico (Fase a Fase)

### Fase 1: Configuración de la Línea Base Pre-Obra (Días -3 a Día 0)
* **Paso 1.1**: Carga del `presupuesto_con_apu.json` comercial/meta en el sistema de control.
* **Paso 1.2**: Alta de los 4 códigos WBS (`WBS-100`, `WBS-200`, `WBS-300`, `WBS-400`) en el aplicativo **Bildin** para el tareo facial de la mano de obra.
* **Paso 1.3**: Socialización de 15 minutos con el Tareador, Almacenero e Ingeniero de Campo para alinear los códigos WBS.

### Fase 2: Ejecución del Registro Diario de Campo (Semanas 1 a 8)
* **Paso 2.1**: **Tareador (Bildin)** registra el tareo facial diario asignando cuadrillas a su respectivo código WBS.
* **Paso 2.2**: **Almacenero** registra la salida física de insumos asociándolos al código WBS de destino.
* **Paso 2.3**: **Ingeniero de Campo** registra el metrado ejecutado del día por partida en el formulario web/tablet.
* **Paso 2.4**: La **Administradora** registra las horas máquina (HM) de excavadoras/rodillos al código WBS.

### Fase 3: Consolidación y Generación de Brechas con IA (Todos los Lunes AM)
* **Paso 3.1**: Consolidación de los datos semanales en la base de datos JSON/Excel.
* **Paso 3.2**: Cálculo automático por código WBS de:
  $$\text{EV} = \sum (\text{Metrado Real} \times \text{P.U. Meta}), \quad \text{AC} = \text{AC}_{\text{MO}} + \text{AC}_{\text{MAT}} + \text{AC}_{\text{EQP}}, \quad \text{CPI} = \frac{\text{EV}}{\text{AC}}$$
* **Paso 3.3**: Ejecución del **Agente IA de Brechas** para evaluar partidas críticas con $\text{CPI} < 1.0$ y proyectar el Costo a Término ($\text{EAC}$).

### Fase 4: Gobierno de Obra y Toma de Acciones (Todos los Lunes 8:30 AM)
* **Paso 4.1**: Reunión de 20 minutos liderada por el **Residente de Obra** con el Ing. de Campo e Ing. de OT.
* **Paso 4.2**: Análisis del reporte de desvíos y aprobación de medidas correctivas para la semana en curso (ej: ajuste de rendimientos $HH/m$ o control de maquinarias en stand-by).

---

## 4. Entregables del Plan de Trabajo

1. 🖼️ **[diagrama_flujo_ro.png](diagrama_flujo_ro.png)**: Diagrama gráfico de flujo operativo del sistema.
2. 📄 **Catálogo de Códigos WBS y Guía Rápida de Campo**: Ficha digital para celular.
3. 🌐 **Formulario Web Diaria & Dashboard RO**: Aplicativo de captura e integración con base de datos JSON.
4. 📊 **Informe Semanal de Resultado Operativo (Excel/JSON)**: Matriz EVM con curva S, CPI, SPI y proyección EAC.
5. 🤖 **Agente IA Analyst**: Módulo de análisis de brechas e informe sintético de causa raíz.
