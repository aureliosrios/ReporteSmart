---
name: presupuestos_tipos
description: Skill especializado en la clasificación, estructuración y reglas de negocio para los distintos tipos de presupuestos.
---

# presupuestos_tipos (Sub-Skill)

Este sub-skill gestiona las reglas específicas según la tipología del presupuesto que se esté formulando.

## 1. Tipos de Presupuestos y Sus Reglas

### A. Presupuesto de Licitación Pública (S10 Estándar)
- **Estructura:** Partidas desglosadas con código WBS/EDT (ej: 01.01.01), metrado exacto del expediente y APUs tradicionales.
- **Pie de Presupuesto:**
  - Costo Directo ($CD$)
  - Gastos Generales ($GG$) = $CD \times \%GG$ (ej. 10%)
  - Utilidad ($UT$) = $CD \times \%UT$ (ej. 8%)
  - Subtotal = $CD + GG + UT$
  - $IGV$ = $\text{Subtotal} \times 18\%$
  - Total Presupuesto = $\text{Subtotal} + IGV$

### B. Presupuesto a Suma Alzada / Venta (Metrado Cerrado)
- El PU ofertado se mantiene estricto. La variabilidad se absorbe en el margen o en el desglose interno del APU.

### C. Presupuesto Mixto Taller - Obra (Manufactura e Instalación)
- Divide el presupuesto en dos ámbitos de ejecución:
  1. **Fabricación / Taller (Subpartidas):** APUs internos con insumos de taller (planchas, soldadura, pintura taller, operario taller).
  2. **Instalación en Obra (Partida Principal):** APU principal que consume la subpartida de taller más recursos de campo (andamios, grúa, peón de montaje).

### D. Presupuesto de Adicionales o Mayores Metrados
- Utiliza la misma estructura de precios de la oferta base, pero aísla los recursos para no alterar el contrato principal.

## 2. Estructura JSON para Tipos de Presupuestos
```json
{
  "proyecto": {
    "titulo": "PROYECTO EJEMPLO",
    "tipo_presupuesto": "LICITACION_PUBLICA | TALLER_OBRA | SUMA_ALZADA",
    "porcentajes": {
      "gastos_generales": 10.0,
      "utilidad": 8.0,
      "igv": 18.0
    }
  }
}
```
