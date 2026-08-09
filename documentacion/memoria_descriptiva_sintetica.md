# Memoria Descriptiva & Plan de Control Ágile (RO con IA)

## 1. Datos Generales del Proyecto Sintético

* **Proyecto:** Instalación de Redes Sanitarias de Agua Potable y Alcantarillado - Habilitación Urbana "Los Cedros".
* **Ubicación:** Lima, Perú.
* **Plazo de Ejecución:** 60 días calendario (8 Semanas).
* **Modalidad de Contratación:** Suma Alzada / Precios Unitarios.
* **Costo Meta (RO Base):** S/ 398,200.00 (Sin IGV).
* **Venta Total:** S/ 485,000.00 (Sin IGV).
* **Margen Esperado:** 17.90% (S/ 86,800.00).

---

## 2. Alcance Técnico del Proyecto

El proyecto contempla la construcción integral de los sistemas de agua potable y alcantarillado sanitario para una habilitación urbana compuesta por 120 lotes:

1. **Red de Alcantarillado Sanitarios:**
   - Trazo, excavación a máquina y perfilado de zanja para 1,400 metros lineales.
   - Tendido de tubería PVC UF DN 200mm Serie S-20.
   - Construcción de 32 buzones de concreto prefabricado.
   - Instalación de 120 conexiones domiciliarias de alcantarillado con caja de registro.
2. **Red de Agua Potable:**
   - Excavación e instalación de 1,000 metros lineales de tubería PVC C-10 DN 110mm.
   - Instalación de accesorios, válvulas de compuerta y grifos contra incendio.
   - Instalación de 120 conexiones domiciliarias de agua potable.
3. **Pruebas y Puesta en Servicio:**
   - Prueba a zanja abierta de estanqueidad y deflexión en tuberías de alcantarillado.
   - Prueba de presión hidráulica y desinfección con cloro en red de agua potable.

---

## 3. Modelo de Gestión Ágil & Control de Resultado Operativo (RO)

Debido al **corto plazo de ejecución (8 semanas)**, cualquier desviación no detectada en la primera mitad del proyecto destruirá el margen operativo. Por ello, se establece un modelo de **Control Ágil Diario con Consolidación Semanal de RO e IA**.

```
                         [ FLUJO DE CONTROL ÁGIL DIARIO ]

  [ ENTRADAS DIARIAS ]               [ ENGINE / IA ]                 [ RESULTADOS ]
+----------------------+         +-----------------------+      +-----------------------+
| 1. BILDIN (MO)       | -------->                       | ---> | - Dashboard RO        |
|    Tareo Facial HH   |         |                       |      |   (PV vs EV vs AC)    |
|                      |         |  Motor ETL Python /   |      |                       |
| 2. ALMACÉN (MAT)     | -------->  Power BI + Agente IA | ---> | - Brechas Semanales   |
|    Vales de Salida   |         |  (Evaluación de       |      |   (CPI, SPI, Rend. HH)|
|                      |         |   Rendimientos)       |      |                       |
| 3. INGENIERO CAMPO   | -------->                       | ---> | - Alertamiento        |
|    Parte Diario      |         +-----------------------+      |   Temprano & EAC      |
+----------------------+                                        +-----------------------+
```

---

## 4. Estrategia de Captura de Datos y Asignación de Roles

| Rol en Obra | Instrumento de Entrada | Frecuencia | Dato Clave Capturado |
| :--- | :--- | :--- | :--- |
| **Tareador** | **Bildin** (App Celular / Rostro) | Diario (6:45 AM / 5:00 PM) | Horas Hombre (HH) por trabajador y código de costo/partida. |
| **Almacenero** | Formulario Digital / Vale Almacén | Diario (Tarde) | Despacho de insumos (Tuberías, bolsas cemento, agregado, arena). |
| **Maquinista / Adm.**| Parte Diario de Equipos | Diario | Horas máquina (HM) de excavadoras, retroexcavadoras y rodillos. |
| **Ing. de Campo** | Formulario de Producción Diaria | Diario (Cierre del día) | Metrados ejecutados ($EV$) por tramo, buzón o conexión. |
| **Ing. de OT** | Dashboard RO + IA Analyst | Semanal (Lunes AM) | Análisis del Resultado Operativo, proyección a origen ($EAC$) y mitigación. |

---

## 5. Algoritmo de Control e Indicadores Clave (KPIs)

Semanalmente, la IA procesa las bases de datos unificadas y calcula los siguientes indicadores por código de costo:

1. **Earned Value (Valor Ganado - EV):**
   $$\text{EV} = \text{Metrado Real Ejecutado} \times \text{Precio Unitario Meta}$$

2. **Actual Cost (Costo Real - AC):**
   $$\text{AC} = \text{Costo Real MO (Bildin)} + \text{Costo Real Mat (Almacén)} + \text{Costo Real Eq/Sub}$$

3. **Cost Performance Index (Índice de Rendimiento de Costo - CPI):**
   $$\text{CPI} = \frac{\text{EV}}{\text{AC}} \quad \begin{cases} > 1.0 & \text{Bajo presupuesto (Saludable)} \\ < 1.0 & \text{Sobrecosto (Alerta crítica)} \end{cases}$$

4. **Estimate at Completion (Costo Proyectado al Cierre - EAC):**
   $$\text{EAC} = \text{AC} + \frac{\text{Presupuesto Meta} - \text{EV}}{\text{CPI}}$$

---

## 6. Hoja de Ruta para Presentar la Propuesta al Cliente (Paso a Paso)

1. **Fase 1: Configuración de la Línea Base (Día 1 a 3 antes del inicio de obra)**
   - Carga del JSON de Presupuesto Meta en el sistema de control.
   - Definición del Catálogo de Códigos de Costo sincronizados con Bildin (Mano de Obra).

2. **Fase 2: Despliegue de Captura Digital (Semana 1)**
   - Capacitación rápida de 15 min al Tareador en el uso de Bildin.
   - Estandarización de partes diarios en Google Forms / Excel estructurado para el Ing. de Campo y Almacenero.

3. **Fase 3: Primer Reporte Semanal de RO con IA (Fin de Semana 1)**
   - Generación del primer reporte de brechas de rendimiento (HH/m lineal en excavación de zanjas).
   - Reunión de 20 minutos con el Residente e Ing. de OT para corregir desviaciones.
