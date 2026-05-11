<div style="font-family: 'Times New Roman', Times, serif;">

# Tarea: Optimización de la Inversión en el Sector Agrario

**Caso:** Usted ha sido contratado por el Ministerio de Desarrollo Agrario para diseñar la ficha técnica de una encuesta regional. El objetivo es diagnosticar la situación de los productores de café en una zona de alta productividad que cuenta con 12,500 productores registrados. Los resultados determinarán la cuantía de los fondos de apoyo para el próximo año. Se le solicita calcular los tamaños de muestra necesarios para dos indicadores críticos, considerando que el presupuesto para el trabajo de campo es limitado y cada encuesta aplicada cuesta $25.00.

<br>

**Tarea 1: Estimación de la Tasa de Acceso al Crédito**
El Ministerio necesita saber qué proporción de productores tiene deudas con la banca formal para diseñar un programa de refinanciamiento.
Información disponible: Un estudio piloto indica que aproximadamente el 35% de los productores tiene crédito activo.
Requerimiento: Trabaje con un Nivel de Confianza del 95% y un Margen de Error del 4%.

**a) Calcule el tamaño de muestra necesario considerando que la población es finita (N = 12,500)**

**Solución a):**
*   **Población (N):** 12,500
*   **Proporción esperada (p):** 0.35 (35%)
*   **Probabilidad de fracaso (q):** 0.65 (65%)
*   **Nivel de Confianza (NC):** 95%, cuyo valor Z correspondiente es **1.96**
*   **Margen de Error (e):** 0.04 (4%)

Para calcular el tamaño de muestra tratándose de la estimación de una proporción en una población finita, aplicamos la siguiente fórmula:

$$n = \frac{N \cdot Z^2 \cdot p \cdot q}{(N-1) \cdot e^2 + Z^2 \cdot p \cdot q}$$

Sustituimos los valores en la fórmula:
$$n = \frac{12500 \cdot (1.96)^2 \cdot 0.35 \cdot 0.65}{(12500-1) \cdot (0.04)^2 + (1.96)^2 \cdot 0.35 \cdot 0.65}$$

Calculamos el numerador:
Numerador $= 12500 \cdot 3.8416 \cdot 0.2275 = 10,924.55$

Calculamos el denominador:
Denominador $= 12499 \cdot 0.0016 + 3.8416 \cdot 0.2275 = 19.9984 + 0.873964 = 20.872364$

Dividimos para obtener $n$:
$$n = \frac{10,924.55}{20.872364} = 523.40$$

*Dado que hablamos de personas encuestadas, se redondea al entero superior.*
**Respuesta:** El tamaño de muestra necesario es de **524 encuestas**.

<br>

**b) ¿Cuál sería el costo total de la recolección de datos para este indicador?**

**Solución b):**
Multiplicamos el número de encuestas calculadas por el costo unitario de aplicar cada una.
Costo Total $= n \times \$25.00$
Costo Total $= 524 \times \$25.00 = \$13,100.00$

**Respuesta:** El costo total de recolección de datos será de **$13,100.00**.

<br><br>

**Tarea 2: Estimación del Ingreso Promedio Mensual**
Para calcular el impacto del subsidio, se requiere conocer el ingreso neto promedio mensual por hectárea.
Información disponible: En el censo agropecuario anterior, la desviación estándar del ingreso en esta zona fue de $180
Requerimiento: El Ministro solicita una precisión alta: Nivel de Confianza del 99% y un Error Máximo admisible de $15.

**c) Calcule el tamaño de muestra requerido.**

**Solución c):**
*   **Población (N):** 12,500
*   **Desviación Estándar ($\sigma$):** $180
*   **Nivel de Confianza (NC):** 99%, cuyo valor Z correspondiente es **2.576** (o 2.58 según profesor).
*   **Error Máximo admisible (E):** $15

Para estimar una media poblacional conociendo el tamaño total de la población (finita), se usa la siguiente fórmula:

$$n = \frac{N \cdot Z^2 \cdot \sigma^2}{(N-1) \cdot E^2 + Z^2 \cdot \sigma^2}$$

Sustituimos los valores:
$$n = \frac{12500 \cdot (2.576)^2 \cdot (180)^2}{(12500-1) \cdot (15)^2 + (2.576)^2 \cdot (180)^2}$$

Calculamos el numerador:
Numerador $= 12500 \cdot 6.6358 \cdot 32400 \approx 2,687,489,280$

Calculamos el denominador:
Denominador $= 12499 \cdot 225 + 6.6358 \cdot 32400 = 2,812,275 + 214,999.14 \approx 3,027,274.14$

Dividimos para obtener $n$:
$$n = \frac{2,687,489,280}{3,027,274.14} = 887.76$$

**Respuesta:** Redondeando al entero superior, el tamaño requerido es de **888 encuestas**.

<br>

**d) Si el presupuesto máximo asignado para este estudio es de $15,000, ¿es viable realizar esta investigación con los parámetros exigidos? Justifique su respuesta con el cálculo del costo.**

**Solución d):**
Calculamos el costo de realizar 888 encuestas.
Costo $= 888 \text{ encuestas} \times \$25.00/\text{encuesta} = \$22,200.00$

**Respuesta:** **No es viable**. Con un costo estimado de \$22,200, el proyecto supera el límite del presupuesto máximo asignado de \$15,000 en \$7,200. Sería necesario relajar los requerimientos estadísticos.

<br><br>

**Tarea 3: Análisis de Sensibilidad y Criterio Económico**
Como experto, usted nota que los costos están excediendo el presupuesto inicial:

**e) Si se decide reducir el Nivel de Confianza de la Tarea 2 del 99% al 95%, ¿cuántas encuestas se ahorraría el Ministerio y cuánto dinero representaría ese ahorro?**

**Solución e):**
Para el 95% de confianza, el nuevo valor $Z$ es **1.96**. Recalculamos $n$:
$$n_{nuevo} = \frac{12500 \cdot (1.96)^2 \cdot (180)^2}{12499 \cdot (15)^2 + (1.96)^2 \cdot (180)^2}$$
$$n_{nuevo} = \frac{12500 \cdot 3.8416 \cdot 32400}{12499 \cdot 225 + 3.8416 \cdot 32400}$$
$$n_{nuevo} = \frac{1,555,848,000}{2,812,275 + 124,467.84} = \frac{1,555,848,000}{2,936,742.84} = 529.78$$

El nuevo tamaño sería de **530 encuestas**.

*   **Ahorro en encuestas:** $888 \text{ (muestra anterior)} - 530 \text{ (muestra nueva)} =$ **358 encuestas ahorradas**.
*   **Ahorro en dinero:** $358 \text{ encuestas} \times \$25.00 =$ **$8,950** de ahorro. (Con esto, el nuevo costo baja a \$13,250, entrando perfectamente en el presupuesto original).

<br>

**f) Si no se tuviera el dato del 35% de acceso al crédito en la Tarea 1 y tuviera que asumir máxima variabilidad, ¿en qué porcentaje aumentaría el tamaño de la muestra respecto al cálculo original?**

**Solución f):**
Asumir máxima variabilidad significa asignar **$p = 0.5$** y **$q = 0.5$**. Recalculamos la Tarea 1 manteniendo los demás datos:

$$n_{p=0.5} = \frac{12500 \cdot (1.96)^2 \cdot 0.5 \cdot 0.5}{(12499) \cdot 0.04^2 + (1.96)^2 \cdot 0.5 \cdot 0.5}$$
$$n_{p=0.5} = \frac{12500 \cdot 3.8416 \cdot 0.25}{19.9984 + 3.8416 \cdot 0.25}$$
$$n_{p=0.5} = \frac{12,005}{19.9984 + 0.9604} = \frac{12,005}{20.9588} = 572.79$$

El nuevo tamaño es de **573 encuestas**.

El cálculo original arrojó **524 encuestas**.
Aumento en encuestas: $573 - 524 = 49 \text{ encuestas más.}$

Para hallar el porcentaje de aumento:
$\text{Porcentaje de aumento} = \left(\frac{\text{Aumento en muestra}}{\text{Muestra original}}\right) \times 100\%$
$\text{Porcentaje} = \left(\frac{49}{524}\right) \times 100\% = 9.35\%$

**Respuesta:** El tamaño de la muestra aumentaría en un **9.35%** respecto al cálculo original debido al incremento en la varianza estadística.

</div>
