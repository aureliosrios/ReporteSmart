---
name: formula_polinomica
description: Skill especializado en la formulación, asignación de Índices Unificados INEI, agrupación de monomios y reajuste de valorizaciones K.
---

# formula_polinomica (Sub-Skill)

Este sub-skill gestiona la generación normativa de la **Fórmula Polinómica de Reajuste de Precios ($K$)** según el Decreto Supremo N° 011-79-VC y las normativas del INEI en el Perú.

## 1. Estructura General de la Fórmula Polinómica
$$K = a \frac{J_r}{J_o} + b \frac{M_r}{M_o} + c \frac{E_r}{E_o} + d \frac{V_r}{V_o} + e \frac{GU_r}{GU_o}$$

Donde:
- $a, b, c, d, e$: Coeficientes de incidencia (deben sumar exactamente $1.000$).
- $J_r, M_r, E_r, V_r, GU_r$: Índices de precio al momento de la valorización (mes de reajuste).
- $J_o, M_o, E_o, V_o, GU_o$: Índices de precio al momento del presupuesto base (mes del valor referencial).

## 2. Reglas de Negocio Normativas
1. **Número Máximo de Monomios:** Máximo 8 monomios.
2. **Número Máximo de Sub-índices por Monomio:** Máximo 3 sub-índices por monomio.
3. **Límite de Incidencia Mínima:** Ningún monomio ni sub-índice puede tener un coeficiente de incidencia menor a **0.050 (5%)**.
4. **Agrupación de Monomios Menores:** Si la incidencia de un IU es menor a 0.050, debe ser sumada/agrupada obligatoriamente con el IU afín de mayor peso.
5. **Gastos Generales y Utilidad:** Se asignan al **Índice Unificado 39 (Índice General de Precios al Consumidor)**.

## 3. Algoritmo de Normalización a 1.000
Para evitar discrepancias de decimales en Excel:
1. Redondear los coeficientes de los primeros $N-1$ monomios a 3 decimales: $c_i = \text{Round}(\text{incidencia}_i, 3)$.
2. Calcular el último monomio ($M_N$, usualmente Gastos Generales/Utilidad) por diferencia directa:
   $$c_N = \text{Round}\left(1.000 - \sum_{i=1}^{N-1} c_i, 3\right)$$

Esto garantiza matemáticamente que $\sum c = 1.000$ en todas las auditorías.
