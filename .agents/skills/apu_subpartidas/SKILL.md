---
name: apu_subpartidas
description: Skill especializado en la gestión, descomposicion y consolidación de APUs anidados y subpartidas (Taller/Obra).
---

# apu_subpartidas (Sub-Skill)

Este sub-skill define cómo se tratan las **subpartidas** (partidas anidadas) dentro de los Análisis de Precios Unitarios (APU) para mantener la trazabilidad exacta en Excel y JSON y evitar la "doble contabilización" en la hoja consolidada de insumos.

## 1. El Concepto de Subpartida en S10
Una subpartida es una actividad cuya tarifa o costo unitario se calcula mediante su propio APU completo en taller o planta, y luego se inserta como un **componente o insumo compuesto** dentro de una partida de nivel superior.

### Ejemplos Típicos en Ingeniería de Mina y Construcción:
1. **Fabricación de Cimbras Metálicas en Taller (`SUB-01`):**
   - **APU de la Subpartida:** Mano de obra de soldadores/oficiales, acero de perfil H, planchas de acople, equipo de soldar. Genera P.U. (ej: $4.29 USD/kg$).
   - **Partida Principal de Obra (`Sostenimiento Roca IV`):** Consume 65.0 kg de `SUB-01` a $4.29 USD/kg$.
2. **Preparación de Concreto en Planta (`SUB-02`):**
   - **APU de la Subpartida:** Cemento, agregados, aditivos acelerantes, mezcladora y operarios de planta. Genera P.U. (ej: $149.24 USD/m3$).
   - **Partida Principal de Obra (`Reparación de Losa`):** Consume 1.0 m3 de `SUB-02` a $149.24 USD/m3$.

---

## 2. Reglas de Exportación Trazable a Excel (Openpyxl)

### A. Catálogo Inicial de Sub-APUs en Excel
- En la hoja `02. PU y Subpartidas`, se debe renderizar un banner verde suave (`fill_subpartida_hdr`) denominado `--- CATÁLOGO DE SUBPARTIDAS COMPUESTAS DE TALLER Y PLANTA ---`.
- Cada subpartida cuenta con su propio APU estándar de 8 columnas con la celda de P.U. Directo (`='02. PU y Subpartidas'!H{row_total_sub}`).

### B. Vinculación Viva en el APU Principal
- En el APU de la partida principal se habilita la sección **`SUBPARTIDAS COMPUESTAS / COMPONENTES`**.
- **Fórmula de Precio (Col G):** Referencia dinámicamente a la celda del P.U. de la subpartida:
  $$\text{Precio Subpartida (Col G)} = \text{`='02. PU y Subpartidas'!H{row_total_subpartida}`}$$
- **Fórmula Parcial (Col H):** `=ROUND(F{row}*G{row}, 2)`.
- El P.U. Directo de la Partida Principal incluye el Subtotal de Subpartidas: `=SUM(H{sub_mo}, H{sub_subpartidas}, H{sub_mat}, H{sub_eq})`.

---

## 3. Regla Antiduplicidad S10 en Consolidado de Insumos (`INSUMOS`)
> [!CRITICAL]
> **REGLA DE ORO DE SUBPARTIDAS:**
> En la pestaña/sección de **Consolidado General de Insumos**, **NO SE DEBEN LISTAR LAS SUBPARTIDAS COMO INSUMOS**. Solo se explosionan los **insumos primarios elementales** (Mano de Obra de taller y obra, materiales base y maquinaria) acumulados en cascada.

### Algoritmo de Explosión en Cascadas:
$$\text{Metrado Efectivo Subpartida} = \text{Metrado Partida Principal} \times \text{Aporte Subpartida en APU}$$
$$\text{Cantidad Insumo Primario} = (\text{Metrado Partida} \times \text{Aporte Insumo}) + \sum (\text{Metrado Efectivo Subpartida} \times \text{Aporte Insumo Subpartida})$$

Esto garantiza en todo momento que:
$$\sum \text{Insumos Primarios Consolidados} \equiv \text{Costo Directo Total del Presupuesto}$$

---

## 4. Representación en JSON (`presupuesto_con_subpartidas.json`)
```json
{
  "subpartidas_catalogo": {
    "SUB_CIMBRA_FAB": {
      "codigo": "SUB-01",
      "descripcion": "FABRICACION DE CIMBRA METALICA EN TALLER DE MINA",
      "unidad": "kg",
      "rendimiento": 250.0,
      "costo_unitario_directo": 4.29
    }
  },
  "partidas": [
    {
      "item_code": "01.02.01.03.03",
      "description": "SOSTENIMIENTO ROCA TIPO IV",
      "apu": {
        "subpartidas": [
          {
            "subpartida_id": "SUB_CIMBRA_FAB",
            "aporte": 65.0,
            "precio": 4.29,
            "parcial": 278.85,
            "formula_parcial": "ROUND(cantidad * precio_unitario_subpartida, 2)"
          }
        ]
      }
    }
  ]
}
```

