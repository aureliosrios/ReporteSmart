---
name: apu_cerrado_ajustes
description: Skill especializado en la resolución de situaciones donde el Precio Unitario (PU) ya está prefijado o cerrado y se requiere ajustar rendimientos y cuadrillas sin generar valores o rendimientos negativos.
---

# apu_cerrado_ajustes (Sub-Skill)

Este sub-skill aborda la problemática crítica de ingeniería de costos: **¿Cómo estructurar o regenerar un APU cuando el Precio Unitario ($PU_{meta}$) ya ha sido prefijado/cerrado en un contrato o concurso, respetando la física de la construcción y evitando montos negativos?**

## 1. Declaración del Problema
En ofertas cerradas o auditorías contractuales, a menudo se exige presentar el desglose del APU para una partida cuyo precio unitario total ya está contractualizado (ejemplo: $PU_{meta} = S/. 125.00 / m2$).

Si se intenta ajustar arbitrariamente los insumos, se corre el riesgo de:
1. Generar **rendimientos negativos** (físicamente imposible).
2. Generar **cantidades o precios negativos** (ilegal e inauditable).
3. Provocar desfases en la trazabilidad por diferencias de redondeo.

## 2. Algoritmo de Resolver de APU Cerrado (Reverse Engineering)

Dado un $PU_{meta}$ objetivo y un conjunto de insumos conocidos (Mano de Obra $MO$, Materiales $MAT$, Equipos $EQ$):

### Paso 1: Fijar Costos Fijos Inamovibles ($C_{fijos}$)
Identificar los insumos con precio comercial y consumo fijo conocido (ejemplo: materiales principales como cemento, fierro o pintura epóxica).
$$C_{materiales} = \sum \text{Round}(\text{Cant}_i \times \text{Precio}_i, 2)$$

### Paso 2: Calcular el Margen Remanente para Mano de Obra y Equipos ($R_{mo\_eq}$)
$$R_{mo\_eq} = PU_{meta} - C_{materiales}$$

> [!WARNING]
> **Condición de Viabilidad Física:**
> Si $R_{mo\_eq} \le 0$, el $PU_{meta}$ es **inviable físicamente** porque los materiales solos superan el precio cerrado. El agente debe alertar al usuario antes de proceder.

### Paso 3: Ajuste Fisiológico de Rendimiento ($Rend$) y Cuadrilla ($Cuad$)
Sabiendo que el subtotal de Mano de Obra ($S_{mo}$) y Herramientas Manuales ($\%H$) absorben el remanente $R_{mo\_eq}$:

$$S_{mo} \times (1 + \%H/100) = R_{mo\_eq}$$
$$S_{mo} = \frac{R_{mo\_eq}}{1 + \%H/100}$$

Como la Mano de Obra es $S_{mo} = \sum \left[ \text{Round}\left(\frac{\text{Cuadrilla}_j \times 8}{Rend}, 4\right) \times \text{Tarifa}_j \right]$:

Podemos despejar el **Rendimiento Requerido Exacto ($Rend_{requerido}$)**:

$$Rend_{requerido} = \frac{8 \times \sum (\text{Cuadrilla}_j \times \text{Tarifa}_j)}{S_{mo}}$$

### Paso 4: Restricciones Sanitarias y De Seguridad Numérica
1. **$Rend_{requerido} > 0$:** Debe ser strictly positivo.
2. **Límites Normativos CAPECO:** El rendimiento ajustado debe estar dentro de un rango viable de mercado (ej: $5.0 \le Rend \le 120.0$). Si el rendimiento resultante es absurdamente alto o bajo, se ajusta la **Cuadrilla** (ej. de 2 peones a 1 peón) y se recalcula el $Rend$.
3. **Re-evaluación de Trazabilidad:** Al calcular la cantidad final a 4 decimales $C_u = \text{Round}(\text{Cuadrilla} \times 8 / Rend, 4)$, se aplica un **micro-ajuste de centavos en la tarifa o en el porcentaje de herramientas manuales** para que $\sum P_{mo} + \sum P_{mat} + \sum P_{eq}$ sume exactamente $PU_{meta}$ al centavo.

## 3. Ejemplo de Script Python Incorporado
```python
def resolver_apu_cerrado(pu_meta, materiales_list, cuadrilla_mo, tarifas_mo, pct_herramientas=3.0):
    costo_mat = sum(round(m['cant'] * m['precio'], 2) for m in materiales_list)
    remanente = pu_meta - costo_mat
    if remanente <= 0:
        raise ValueError(f"PU Meta S/. {pu_meta} es inferior al costo directo de materiales S/. {costo_mat}")
    
    subtotal_mo_target = remanente / (1 + pct_herramientas / 100.0)
    costo_cuadrilla_dia = sum(c * t for c, t in zip(cuadrilla_mo, tarifas_mo))
    
    # Rendimiento exacto continuo
    rendimiento_exacto = (costo_cuadrilla_dia * 8.0) / subtotal_mo_target
    return round(rendimiento_exacto, 2)
```
