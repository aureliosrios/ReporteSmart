---
name: trazabilidad_presupuestos
description: Skill maestro para la arquitectura, cálculo y trazabilidad de presupuestos S10 en JSON y exportación dinámica a Excel.
---

# trazabilidad_presupuestos (Skill Maestro)

Este skill es el **núcleo principal de trazabilidad y coordinación** para el procesamiento de presupuestos de construcción civil compatibles con la metodología S10 / CAPECO / INEI.

## 1. Responsabilidades del Skill Maestro
1. **Coordinación del Pipeline:** Orquestar los sub-skills especializados de presupuestos, APUs, subpartidas, APU cerrados y fórmula polinómica.
2. **Garantía de Trazabilidad Total (Regla de Oro):**
   $$\text{Suma Total de Insumos Consolidados} \equiv \text{Costo Directo Total del Presupuesto}$$
3. **Control de Precisión y Truncamiento:**
   - Cantidad de aporte unitario ($C_u$): $\text{Round}\left(\frac{\text{Cuadrilla} \times \text{Jornal}}{\text{Rendimiento}}, 4\right)$ (Jornal estándar = 8.00 hh)
   - Parcial de Recurso ($P_r$): $\text{Round}(C_u \times \text{Precio Unitario}, 2)$
   - Costo Parcial de Partida: $\text{Round}(\text{Metrado} \times \text{PU Directo}, 2)$
4. **Reglas de Exportación Dinámica a Excel (Openpyxl) - Fórmulas Vivas Explicitas:**
   - **Vínculo Vivo P.U. Presupuesto Base:** La columna de P.U. en `01. Presupuesto Base` DEBE contener la fórmula directa que apunta a la celda del Total P.U. de la pestaña de APUs (`='02. Analisis de PU'!H{row_total}`).
   - **Estructura APU de 8 Columnas:** `[Código, Descripción Recurso, Unidad, Cuadrilla, Jornal, Cantidad / Aporte, Precio (USD), Parcial (USD)]`.
   - **Fórmula de Cantidad APU (MO y EQ):** `=ROUND(D{row}*E{row}/B${rend_row}, 4)` vinculada a las celdas de Cuadrilla, Jornal y Rendimiento.
   - **Herramientas Manuales (% MO):** Precio vinculado al subtotal de mano de obra (`=H{subtotal_mo_row}`), Cantidad `0.03` (3.00%) y Parcial `=ROUND(F{row}*G{row}, 2)`.
   - **Subtotales y Total PU:** Sumatorias dinámicas `=SUM(...)` por sección de MO, Materiales y Equipos, y Total PU `=SUM(H{sub_mo}, H{sub_mat}, H{sub_eq})`.
   - **Explosión Insumos con % Incidencia:** Fórmula viva `=G{row}/'01. Presupuesto Base'!F${row_costo_directo}`.
   - **Diseño Adaptativo Auto-Fit:** Ajuste de ancho de columnas dinámico `min(max(max_len + 4, 12), 65)` ignorando celdas combinadas de títulos.
5. **Enriquecimiento del JSON (`presupuesto_con_apu.json`):**
   - Incluir metadata explícita en los APU JSON: `"jornal": 8.0`, `"formula_cantidad"`, `"formula_parcial"`, `"formula_precio"`.

## 2. Mapa de Sub-Skills Coordinados
- **[presupuestos_tipos](file:///d:/Agentes%20de%20IA/Minera%20Huanchor/.agents/skills/presupuestos_tipos/SKILL.md):** Tipologías de presupuestos (Licitación, Venta/Metrado Cerrado, Taller/Manufactura, Gastos Generales y Adicionales).
- **[apu_analisis](file:///d:/Agentes%20de%20IA/Minera%20Huanchor/.agents/skills/apu_analisis/SKILL.md):** Estructura del APU, desglose por componentes (MO, Materiales, Equipos, % Herramientas Manuales).
- **[apu_subpartidas](file:///d:/Agentes%20de%20IA/Minera%20Huanchor/.agents/skills/apu_subpartidas/SKILL.md):** Tratamiento de APUs anidados en cascada (Obra -> Taller/Manufactura) y consolidación limpia en Insumos.
- **[apu_cerrado_ajustes](file:///d:/Agentes%20de%20IA/Minera%20Huanchor/.agents/skills/apu_cerrado_ajustes/SKILL.md):** Ajuste de PU meta cerrado mediante optimización de rendimientos y cuadrillas sin costos/rendimientos negativos.
- **[formula_polinomica](file:///d:/Agentes%20de%20IA/Minera%20Huanchor/.agents/skills/formula_polinomica/SKILL.md):** Agrupación normativa INEI (D.S. 011-79-VC), asignación de IUs y conformación de hasta 8 monomios.

## 3. Flujo de Ejecución Recomendado
```
[JSON Base / Estructura WBS]
       │
       ├──> Call: presupuestos_tipos (Determina regla del tipo de presupuesto)
       ├──> Call: apu_analisis & apu_subpartidas (Desglosa recursos y subpartidas)
       ├──> Call: apu_cerrado_ajustes (Ajusta PU prefijados/cerrados si aplica)
       ├──> Call: trazabilidad_presupuestos (Verifica Trazabilidad CD == Insumos & Genera Excel con Fórmulas Vivas)
       └──> Call: formula_polinomica (Genera monomios y reajuste K)
```

