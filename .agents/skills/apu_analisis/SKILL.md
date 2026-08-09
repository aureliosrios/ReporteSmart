---
name: apu_analisis
description: Skill especializado en la descomposición, formulación matemática y cálculo de Análisis de Precios Unitarios (APU).
---

# apu_analisis (Sub-Skill)

Este sub-skill gestiona la estructura interna de un APU (Mano de Obra, Materiales, Equipos y Herramientas Manuales), la aplicación de las normas de truncamiento S10 y su representación trazable en Excel y JSON.

## 1. Categorías del APU y Fórmulas Explicitas

### A. Mano de Obra (MO)
- **Fórmula de Aporte Unitario (Cantidad hh):**
  $$\text{Cantidad } (hh) = \text{Round}\left(\frac{\text{Cuadrilla} \times \text{Jornal}}{\text{Rendimiento}}, 4\right)$$
  *(En Excel:* `=ROUND(D{row}*E{row}/B${rend_row}, 4)` *donde Jornal = 8.00 hh)*
- **Parcial ($P_{mo}$):** $\text{Round}(\text{Cantidad} \times \text{Costo Hora Hombre}, 2)$
  *(En Excel:* `=ROUND(F{row}*G{row}, 2)`*)*

### B. Materiales
- **Cantidad:** Consumo unitario estimado de especificaciones técnicas.
- **Parcial ($P_{mat}$):** $\text{Round}(\text{Cantidad} \times \text{Precio Insumo Material}, 2)$

### C. Equipos y Maquinaria
- **Aporte Unitario ($hm$):** Similar a la mano de obra, calculado con cuadrilla de máquina, jornal (8.00 hh) y rendimiento.
- **Herramientas Manuales (% MO):**
  - Expresado como un porcentaje del subtotal de mano de obra (ej. $3\%$).
  - **Enlace de Fórmulas en Excel:**
    $$\text{Precio Herramienta (Col G)} = \text{Subtotal MO } (=H\{\text{row\_subtotal\_mo}\})$$
    $$\text{Cantidad (Col F)} = 0.03 \quad (3.00\%)$$
    $$\text{Parcial Herramienta (Col H)} = \text{Round}(F{row} \times G{row}, 2)$$

## 2. Costo Unitario Directo
$$PU_{directo} = \sum P_{mo} + \sum P_{mat} + \sum P_{eq} + P_{herr} + \sum P_{subpartida}$$
*(En Excel se calcula sumando los subtotales de cada sección:* `=SUM(H{sub_mo}, H{sub_mat}, H{sub_eq})`*)*

## 3. Ejemplo Estructura APU JSON Trazable
```json
{
  "codigo": "01.01.01.01",
  "descripcion": "MOVILIZACION Y DESMOVILIZACION DE EQUIPOS",
  "unidad": "glb",
  "rendimiento": 1.0,
  "mano_obra": [
    {
      "recurso_id": "MO_CAPATAZ",
      "cuadrilla": 0.5,
      "aporte": 4.0,
      "precio": 8.67,
      "parcial": 34.68,
      "jornal": 8.0,
      "formula_cantidad": "ROUND(cuadrilla * 8.0 / rendimiento, 4)",
      "formula_parcial": "ROUND(cantidad * precio, 2)"
    }
  ],
  "materiales": [],
  "equipos": [
    {
      "recurso_id": "EQ_HERRAMIENTAS",
      "cuadrilla": 0.0,
      "aporte": 0.03,
      "precio": 335.32,
      "parcial": 10.06,
      "formula_cantidad": "0.03 (3% MO)",
      "formula_precio": "Subtotal Mano de Obra",
      "formula_parcial": "ROUND(0.03 * subtotal_mo, 2)"
    }
  ],
  "costo_unitario_directo": 13508.48
}
```
