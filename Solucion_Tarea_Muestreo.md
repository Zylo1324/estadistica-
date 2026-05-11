# Solución de la Tarea: Optimización de la Inversión en el Sector Agrario

**Datos Generales del Caso:**
*   **Población (N):** 12,500 productores de café.
*   **Costo por encuesta:** $25.00

---

## Tarea 1: Estimación de la Tasa de Acceso al Crédito

**Información disponible:**
*   Población ($N$): 12,500
*   Proporción esperada ($p$): 35% = 0.35
*   Probabilidad de fracaso ($q$ = 1 - p): 0.65
*   Nivel de Confianza (NC): 95% $\rightarrow$ Valor Z = 1.96
*   Margen de Error ($e$): 4% = 0.04

### a) Calcule el tamaño de muestra necesario
Para calcular el tamaño de muestra de una proporción en una **población finita**, usamos la siguiente fórmula:

$$n = \frac{N \cdot Z^2 \cdot p \cdot q}{(N-1) \cdot e^2 + Z^2 \cdot p \cdot q}$$

**Paso 1: Sustituir los valores**
$$n = \frac{12500 \cdot (1.96)^2 \cdot (0.35) \cdot (0.65)}{(12500-1) \cdot (0.04)^2 + (1.96)^2 \cdot (0.35) \cdot (0.65)}$$

**Paso 2: Calcular el numerador**
Numerador $= 12500 \cdot 3.8416 \cdot 0.2275 = 10,924.55$

**Paso 3: Calcular el denominador**
Denominador $= 12499 \cdot 0.0016 + 3.8416 \cdot 0.2275 = 19.9984 + 0.873964 = 20.872364$

**Paso 4: Dividir para obtener n**
$$n = \frac{10,924.55}{20.8724} = 523.40$$

*Al tratarse de personas (encuestas), siempre se debe redondear al entero superior.*
**Respuesta (a):** Se requiere un tamaño de muestra de **524 encuestas**.

### b) Costo total de la recolección de datos
Costo Total $= n \times \text{Costo por encuesta}$
Costo Total $= 524 \times \$25.00$
**Respuesta (b):** El costo total de recolección para este indicador será de **$13,100.00**.

---

## Tarea 2: Estimación del Ingreso Promedio Mensual

**Información disponible:**
*   Población ($N$): 12,500
*   Desviación estándar ($\sigma$): $180
*   Nivel de Confianza (NC): 99% $\rightarrow$ Valor Z = 2.576 (se puede usar 2.58 según el profesor, pero 2.576 es más exacto).
*   Error Máximo admisible ($E$): $15

### c) Calcule el tamaño de muestra requerido
Para calcular el tamaño de muestra para una media poblacional tratándose de una población finita:

$$n = \frac{N \cdot Z^2 \cdot \sigma^2}{(N-1) \cdot E^2 + Z^2 \cdot \sigma^2}$$

**Paso 1: Sustituir valores**
$$n = \frac{12500 \cdot (2.576)^2 \cdot (180)^2}{(12499) \cdot (15)^2 + (2.576)^2 \cdot (180)^2}$$

**Paso 2: Calcular Numerador y Denominador**
Numerador $= 12500 \cdot 6.6358 \cdot 32400 = 2,687,500,000 \text{ (aprox.)}$
Denominador $= (12499 \cdot 225) + (6.6358 \cdot 32400) = 2,812,275 + 215,000 = 3,027,275$

**Paso 3: Calcular n**
*Cálculo exacto:* $n = 2,687,489,280 / 3,027,274.14 = 887.76$

*Redondeando superior:*
**Respuesta (c):** El tamaño requerido es de **888 encuestas**.
*(Nota: si el profesor usa población infinita para este cálculo la muestra sería n = 956. Con N conocida, usar la fórmula finita es más correcto: n = 888).*

### d) Viabilidad del Estudio (Presupuesto $15,000)
Costo Real $= 888 \text{ encuestas} \times \$25.00 = \$22,200.00$

**Respuesta (d):** **No es viable**. El costo calculado (\$22,200) supera ampliamente el presupuesto máximo asignado de \$15,000. Para poder realizarlo habría que reducir parámetros estadísticos (como disminuir la confianza o aumentar el margen de error).

---

## Tarea 3: Análisis de Sensibilidad y Criterio Económico

### e) Reducir el Nivel de Confianza en la Tarea 2 (del 99% al 95%)
Si bajamos al 95%, el nuevo $Z = 1.96$. Recalculamos el n de la Tarea 2:
$$n_{nuevo} = \frac{12500 \cdot (1.96)^2 \cdot (180)^2}{(12499 \cdot 15^2) + (1.96^2 \cdot 180^2)}$$
$$n_{nuevo} = \frac{1,555,848,000}{2,812,275 + 124,467.84} = \frac{1,555,848,000}{2,936,742.84} = 529.78$$

El nuevo tamaño es de **530 encuestas**.

*   **Ahorro en encuestas:** $888 \text{ (muestra original)} - 530 \text{ (nueva muestra)} =$ **358 encuestas ahorradas**.
*   **Ahorro en dinero:** $358 \times \$25.00 =$ **$8,950 ahorrados**. 
*(Además, con esta nueva muestra, el costo sería $530 \times \$25 = \$13,250$, logrando entrar al presupuesto de \$15,000).*

### f) Asumir Máxima Variabilidad en Tarea 1
Si no supiéramos que la proporción es 35%, debemos castigar el estadístico usando máxima varianza: $p = 0.5$ y $q = 0.5$.
Recalculamos la Tarea 1 con $p=0.5$:
$$n_{p=0.5} = \frac{12500 \cdot (1.96)^2 \cdot 0.5 \cdot 0.5}{(12499)(0.04)^2 + (1.96)^2 \cdot 0.5 \cdot 0.5}$$
$$n_{p=0.5} = \frac{12,005}{19.9984 + 0.9604} = \frac{12,005}{20.9588} = 572.79$$

El nuevo tamaño sería de **573 encuestas**.

La muestra original era de 524 encuestas, por lo que el aumento es de **49 encuestas**.
**Porcentaje de aumento:** $\frac{49}{524} \times 100\% = \mathbf{9.35 \%}$.

**Conclusión:** Asumir máxima variabilidad aumentaría la muestra original casi un 9.4% respecto al cálculo original, incrementando los costos innecesariamente.
