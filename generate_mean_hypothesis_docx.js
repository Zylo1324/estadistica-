const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const baseDocxPath = path.join(__dirname, '_tmp_locked_copy.docx');
const outputDocxPath = path.join(
  __dirname,
  'UNIDAD 2',
  'Estimacion y prueba de hipotesis de una media_resuelto.docx'
);

const exercises = [
  {
    statement:
      'El gerente de la cadena de una tienda afirma que en promedio los clientes gastaron $400 el año pasado. Sin embargo, analizando el mercado, nosotros creemos, que dicho gerente ha exagerado. Para someter a prueba estas hipótesis se tomó una muestra aleatoria de 100 clientes que el año pasado habían comprado en dicha tienda, ésta reveló una media de $450 y una desviación estándar $100.',
    parts: [
      'a) En el nivel de significación de 0.05, ¿Es posible concluir que los clientes de esta tienda están gastando más?',
      'b) Calcule e interprete su intervalo de confianza al 95%.',
    ],
    blocks: [
      p('Datos: muestra de 100 clientes, media muestral de 450 dólares y desviación estándar muestral de 100 dólares.'),
      eq([text('n = 100; '), xbar(), text(' = 450; s = 100; '), mu0(), text(' = 400')]),
      p('Como se quiere comprobar si los clientes gastan más, la prueba es de cola derecha.'),
      eq([h0(), text(': μ ≤ 400')]),
      eq([h1(), text(': μ > 400')]),
      p('Se calcula el estadístico de prueba usando aproximación normal porque la muestra es grande.'),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('s'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('450 - 400')], frac(text('100'), rad(text('100')))),
        text(' = 5,00'),
      ]),
      eq([sub('Z', 'crítico'), text(' = 1,645')]),
      p('Como 5,00 es mayor que 1,645, el estadístico cae en la región de rechazo. Por tanto, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 95%, se usa el valor crítico 1,96.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('s'), rad(text('n')))]),
      eq([
        text('IC = 450 ± 1,96'),
        frac(text('100'), rad(text('100'))),
        text(' = 450 ± 19,60'),
      ]),
      eq([sub('IC', '95%'), text(' = (430,40; 469,60)')]),
    ],
    response:
      'Sí es posible concluir que los clientes están gastando más de 400 dólares. Con 95% de confianza, el gasto promedio está entre 430,40 y 469,60 dólares.',
  },
  {
    statement:
      'El gerente de ventas de una empresa que elabora cápsulas de uña de gato indica que la demanda semanal tiene distribución normal con una media de 1000 cápsulas y una desviación estándar de 360 cápsulas. Sin embargo, en un estudio reciente, una muestra aleatoria de 36 semanas dio una demanda promedio de 950 cápsulas.',
    parts: [
      'a) ¿Es posible concluir que la producción promedio semanal es menos de 1000 cápsulas al 1% de significación?',
      'b) Calcule e interprete su intervalo de confianza al 92%.',
    ],
    blocks: [
      p('Datos: la desviación estándar poblacional es conocida, por eso se trabaja con Z.'),
      eq([text('n = 36; '), xbar(), text(' = 950; σ = 360; '), mu0(), text(' = 1000')]),
      p('Como se quiere comprobar si el promedio es menor que 1000 cápsulas, la prueba es de cola izquierda.'),
      eq([h0(), text(': μ ≥ 1000')]),
      eq([h1(), text(': μ < 1000')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('σ'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('950 - 1000')], frac(text('360'), rad(text('36')))),
        text(' = -0,83'),
      ]),
      eq([sub('Z', 'crítico'), text(' = -2,326')]),
      p('Como -0,83 no es menor que -2,326, no se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 92%, se usa el valor crítico 1,751.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('σ'), rad(text('n')))]),
      eq([
        text('IC = 950 ± 1,751'),
        frac(text('360'), rad(text('36'))),
        text(' = 950 ± 105,04'),
      ]),
      eq([sub('IC', '92%'), text(' = (844,96; 1055,04)')]),
    ],
    response:
      'No hay evidencia suficiente al 1% para concluir que la demanda promedio semanal sea menor que 1000 cápsulas. Con 92% de confianza, la demanda promedio está entre 844,96 y 1055,04 cápsulas.',
  },
  {
    statement:
      'Un fabricante de aparatos de TV afirma que se necesita a lo mucho 250 micro amperes de corriente para alcanzar cierto grado de brillantez con un tipo de televisor en particular. Una muestra de 20 aparatos de TV produce un promedio muestral de corriente de 257.3 micro amperes. Denotemos por m el verdadero promedio de corriente necesaria para alcanzar la brillantez deseada con aparatos de este tipo, y supongamos que μ es la media de una población con σ = 15.',
    parts: [
      'a) Pruebe al nivel de significación del 5% la hipótesis nula de que la media es a lo sumo 250 micro amperes.',
      'b) Calcule e interprete su intervalo de confianza al 90%.',
    ],
    blocks: [
      p('Datos: se conoce la desviación estándar poblacional, por eso se utiliza Z.'),
      eq([text('n = 20; '), xbar(), text(' = 257,3; σ = 15; '), mu0(), text(' = 250')]),
      p('La afirmación indica que la media es a lo sumo 250 micro amperes; se contrasta si realmente se necesita más corriente.'),
      eq([h0(), text(': μ ≤ 250')]),
      eq([h1(), text(': μ > 250')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('σ'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('257,3 - 250')], frac(text('15'), rad(text('20')))),
        text(' = 2,18'),
      ]),
      eq([sub('Z', 'crítico'), text(' = 1,645')]),
      p('Como 2,18 es mayor que 1,645, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 90%, se usa el valor crítico 1,645.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('σ'), rad(text('n')))]),
      eq([
        text('IC = 257,3 ± 1,645'),
        frac(text('15'), rad(text('20'))),
        text(' = 257,3 ± 5,52'),
      ]),
      eq([sub('IC', '90%'), text(' = (251,78; 262,82)')]),
    ],
    response:
      'Se rechaza la afirmación de que la media sea a lo sumo 250 micro amperes. Con 90% de confianza, la corriente promedio necesaria está entre 251,78 y 262,82 micro amperes.',
  },
  {
    statement:
      'En una conocida prueba de autoimagen que da por resultado puntajes que se distribuyen de manera normal, se espera que el puntaje medio de los beneficiarios de la asistencia pública sea de 65. La prueba se aplica a una muestra aleatoria de 28 beneficiarios de la asistencia pública quienes logran un puntaje medio de 62.1 con una desviación estándar de 5.83.',
    parts: [
      'a) Al nivel de significación del 1%, ¿los beneficiarios de la asistencia pública tienen puntajes diferentes en promedio, a los que se esperaban?',
      'b) Calcule e interprete su intervalo de confianza 97%.',
    ],
    blocks: [
      p('Datos: la población es normal, la muestra es pequeña y se tiene desviación estándar muestral; por eso se usa t de Student con 27 grados de libertad.'),
      eq([text('n = 28; '), xbar(), text(' = 62,1; s = 5,83; '), mu0(), text(' = 65')]),
      p('Como se pregunta si los puntajes son diferentes, la prueba es bilateral.'),
      eq([h0(), text(': μ = 65')]),
      eq([h1(), text(': μ ≠ 65')]),
      eq([text('t = '), frac([xbar(), text(' - '), mu0()], frac(text('s'), rad(text('n'))))]),
      eq([
        text('t = '),
        frac([text('62,1 - 65')], frac(text('5,83'), rad(text('28')))),
        text(' = -2,63'),
      ]),
      eq([sub('t', 'crítico'), text(' = ±2,771')]),
      p('Como -2,63 se encuentra dentro del intervalo de aceptación, no se rechaza la hipótesis nula al 1%.'),
      p('Para el intervalo de confianza al 97%, se usa t crítico igual a 2,291.'),
      eq([text('IC = '), xbar(), text(' ± t'), frac(text('s'), rad(text('n')))]),
      eq([
        text('IC = 62,1 ± 2,291'),
        frac(text('5,83'), rad(text('28'))),
        text(' = 62,1 ± 2,52'),
      ]),
      eq([sub('IC', '97%'), text(' = (59,58; 64,62)')]),
    ],
    response:
      'Al 1% de significación, no hay evidencia suficiente para afirmar que el puntaje promedio sea diferente de 65. Con 97% de confianza, el puntaje promedio está entre 59,58 y 64,62 puntos.',
  },
  {
    statement:
      'Para tratar de estimar el consumo promedio por cliente, en un gran restaurante, se reunieron datos de una muestra de 49 clientes durante un periodo de tres semanas. Si la media de la muestra es de $22.60 dólares con una desviación estándar de $2.5,',
    parts: [
      'a) ¿Existe evidencia para decir que el consumo promedio de la población es menor a 25 dólares? Pruebe con α = 0.05',
      'b) Calcule e interprete su intervalo de confianza al 97.5%.',
    ],
    blocks: [
      p('Datos: la muestra es grande, por eso se usa aproximación normal con la desviación estándar muestral.'),
      eq([text('n = 49; '), xbar(), text(' = 22,60; s = 2,5; '), mu0(), text(' = 25')]),
      p('Como se quiere comprobar si el consumo promedio es menor que 25 dólares, la prueba es de cola izquierda.'),
      eq([h0(), text(': μ ≥ 25')]),
      eq([h1(), text(': μ < 25')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('s'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('22,60 - 25')], frac(text('2,5'), rad(text('49')))),
        text(' = -6,72'),
      ]),
      eq([sub('Z', 'crítico'), text(' = -1,645')]),
      p('Como -6,72 es menor que -1,645, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 97,5%, se usa el valor crítico 2,241.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('s'), rad(text('n')))]),
      eq([
        text('IC = 22,60 ± 2,241'),
        frac(text('2,5'), rad(text('49'))),
        text(' = 22,60 ± 0,80'),
      ]),
      eq([sub('IC', '97,5%'), text(' = (21,80; 23,40)')]),
    ],
    response:
      'Sí existe evidencia de que el consumo promedio sea menor a 25 dólares. Con 97,5% de confianza, el consumo promedio está entre 21,80 y 23,40 dólares.',
  },
  {
    statement:
      'Una muestra aleatoria de 12 alumnas graduadas de una escuela secretarial mecanografió un promedio de 79.5 palabras por minuto con una desviación estándar de 7.8 palabras por minuto.',
    parts: [
      'a) ¿Se tiene evidencia estadística para decir que el número promedio de palabras mecanografiadas por todas las graduadas de esa escuela es de 80 palabras por minuto con α = 0,01?',
      'b) Calcule e interprete su intervalo de confianza al 99%.',
    ],
    blocks: [
      p('Datos: la muestra es pequeña y se tiene desviación estándar muestral; por eso se usa t de Student con 11 grados de libertad.'),
      eq([text('n = 12; '), xbar(), text(' = 79,5; s = 7,8; '), mu0(), text(' = 80')]),
      p('Se contrasta si el promedio es igual o diferente de 80 palabras por minuto.'),
      eq([h0(), text(': μ = 80')]),
      eq([h1(), text(': μ ≠ 80')]),
      eq([text('t = '), frac([xbar(), text(' - '), mu0()], frac(text('s'), rad(text('n'))))]),
      eq([
        text('t = '),
        frac([text('79,5 - 80')], frac(text('7,8'), rad(text('12')))),
        text(' = -0,22'),
      ]),
      eq([sub('t', 'crítico'), text(' = ±3,106')]),
      p('Como -0,22 se encuentra dentro del intervalo de aceptación, no se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 99%, se usa t crítico igual a 3,106.'),
      eq([text('IC = '), xbar(), text(' ± t'), frac(text('s'), rad(text('n')))]),
      eq([
        text('IC = 79,5 ± 3,106'),
        frac(text('7,8'), rad(text('12'))),
        text(' = 79,5 ± 6,99'),
      ]),
      eq([sub('IC', '99%'), text(' = (72,51; 86,49)')]),
    ],
    response:
      'No se rechaza que el promedio sea de 80 palabras por minuto. Con 99% de confianza, el promedio real está entre 72,51 y 86,49 palabras por minuto.',
  },
  {
    statement:
      'Un empresario está considerando la posibilidad de ampliar su negocio mediante la adquisición de un pequeño bar. El dueño actual del bar afirma que el ingreso diario del establecimiento sigue una distribución normal de media 675 soles y una desviación estándar de 75 soles. Para comprobar si decía la verdad, tomó una muestra de treinta días y ésta reveló un ingreso diario promedio de 625 soles.',
    parts: [
      'a) ¿Hay evidencia de que el ingreso diario promedio sea menor del que afirma el presente dueño? Utilice un nivel de significación del 1%.',
      'b) Calcula e interpreta un intervalo de confianza al 95%.',
    ],
    blocks: [
      p('Datos: se conoce la desviación estándar poblacional, por eso se usa Z.'),
      eq([text('n = 30; '), xbar(), text(' = 625; σ = 75; '), mu0(), text(' = 675')]),
      p('Como se desea comprobar si el ingreso promedio es menor que el afirmado, la prueba es de cola izquierda.'),
      eq([h0(), text(': μ ≥ 675')]),
      eq([h1(), text(': μ < 675')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('σ'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('625 - 675')], frac(text('75'), rad(text('30')))),
        text(' = -3,65'),
      ]),
      eq([sub('Z', 'crítico'), text(' = -2,326')]),
      p('Como -3,65 es menor que -2,326, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 95%, se usa el valor crítico 1,96.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('σ'), rad(text('n')))]),
      eq([
        text('IC = 625 ± 1,96'),
        frac(text('75'), rad(text('30'))),
        text(' = 625 ± 26,84'),
      ]),
      eq([sub('IC', '95%'), text(' = (598,16; 651,84)')]),
    ],
    response:
      'Sí hay evidencia de que el ingreso diario promedio es menor que 675 soles. Con 95% de confianza, el ingreso diario promedio está entre 598,16 y 651,84 soles.',
  },
  {
    statement:
      'Una empresa eléctrica fabrica focos que tienen una duración que está distribuida aproximadamente en forma normal con media de 800 horas y una desviación estándar de 40.',
    parts: [
      'a) Pruebe la hipótesis de que la duración promedio sea verdaderamente 800 horas si en una muestra aleatoria de 300 focos tiene una duración promedio de 788 horas. Utilice un nivel de significancia de 0.04.',
      'b) Calcula e interpreta un intervalo de confianza al 99%',
    ],
    blocks: [
      p('Datos: se conoce la desviación estándar poblacional y la muestra es grande, por eso se usa Z.'),
      eq([text('n = 300; '), xbar(), text(' = 788; σ = 40; '), mu0(), text(' = 800')]),
      p('Como se pregunta si la duración promedio es verdaderamente 800 horas, la prueba es bilateral.'),
      eq([h0(), text(': μ = 800')]),
      eq([h1(), text(': μ ≠ 800')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('σ'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('788 - 800')], frac(text('40'), rad(text('300')))),
        text(' = -5,20'),
      ]),
      eq([sub('Z', 'crítico'), text(' = ±2,054')]),
      p('Como -5,20 queda fuera del intervalo de aceptación, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 99%, se usa el valor crítico 2,576.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('σ'), rad(text('n')))]),
      eq([
        text('IC = 788 ± 2,576'),
        frac(text('40'), rad(text('300'))),
        text(' = 788 ± 5,95'),
      ]),
      eq([sub('IC', '99%'), text(' = (782,05; 793,95)')]),
    ],
    response:
      'Se rechaza que la duración promedio sea verdaderamente 800 horas. Con 99% de confianza, la duración promedio está entre 782,05 y 793,95 horas.',
  },
  {
    statement:
      'Pruebe la hipótesis de que el contenido promedio en recipientes de un lubricante en particular es de 10 litros si los contenidos de una muestra aleatoria de 10 recipientes son 10.2, 9.7, 10.1, 10.3, 10.1, 9.8, 9.9, 10.4, 10.3 y 9.8 litros. Utilice un nivel de significancia de 0.01. Calcula e interpreta un intervalo de confianza al 95%',
    parts: [],
    blocks: [
      p('Datos: se tiene una muestra pequeña y la desviación estándar poblacional no se conoce; por eso se calcula la desviación estándar muestral y se usa t de Student con 9 grados de libertad.'),
      eq([text('n = 10; '), mu0(), text(' = 10')]),
      eq([xbar(), text(' = '), frac(text('∑x'), text('n')), text(' = '), frac(text('100,6'), text('10')), text(' = 10,06')]),
      eq([
        text('s = '),
        rad(frac([text('∑'), sup(delim([sub('x', 'i'), text(' - '), xbar()]), '2')], text('n - 1'))),
        text(' = 0,2459'),
      ]),
      p('Como se prueba si el contenido promedio es de 10 litros, la prueba es bilateral.'),
      eq([h0(), text(': μ = 10')]),
      eq([h1(), text(': μ ≠ 10')]),
      eq([text('t = '), frac([xbar(), text(' - '), mu0()], frac(text('s'), rad(text('n'))))]),
      eq([
        text('t = '),
        frac([text('10,06 - 10')], frac(text('0,2459'), rad(text('10')))),
        text(' = 0,77'),
      ]),
      eq([sub('t', 'crítico'), text(' = ±3,250')]),
      p('Como 0,77 se encuentra dentro del intervalo de aceptación, no se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 95%, se usa t crítico igual a 2,262.'),
      eq([text('IC = '), xbar(), text(' ± t'), frac(text('s'), rad(text('n')))]),
      eq([
        text('IC = 10,06 ± 2,262'),
        frac(text('0,2459'), rad(text('10'))),
        text(' = 10,06 ± 0,18'),
      ]),
      eq([sub('IC', '95%'), text(' = (9,88; 10,24)')]),
    ],
    response:
      'No se rechaza que el contenido promedio sea de 10 litros. Con 95% de confianza, el contenido promedio está entre 9,88 y 10,24 litros.',
  },
  {
    statement:
      'Una agencia de publicidad tiene un registro de datos sobre el tiempo (en minutos) de los anuncios publicitarios por cada 20 minutos en los programas principales de TV. Una muestra aleatoria de 35 de estos registros proporcionó un tiempo medio de publicidad de 3 minutos por cada 20 minutos de publicidad. Suponiendo que el tiempo de anuncios en minutos sigue una distribución normal con una desviación estándar de 1.2 minutos. Determine:',
    parts: [
      'a) Compruebe que el tiempo medio de publicidad sea mayor a los 2.5 minutos. Con una significancia del 3%.',
      'b) Un intervalo de confianza del 99% para el tiempo medio de anuncios publicitarios en los programas principales cada 20 minutos.',
      'c) Un intervalo de confianza del 90% para el tiempo medio de anuncios publicitarios en los programas principales cada 20 minutos.',
      'd) De qué tamaño debe tomarse una muestra, para tener un 95% de confianza y un margen de error de 0,5 minutos en la estimación.',
    ],
    blocks: [
      p('Datos: se conoce la desviación estándar poblacional, por eso se usa Z.'),
      eq([text('n = 35; '), xbar(), text(' = 3; σ = 1,2; '), mu0(), text(' = 2,5')]),
      p('Para el apartado a), se quiere comprobar si el tiempo promedio es mayor que 2,5 minutos; por eso la prueba es de cola derecha.'),
      eq([h0(), text(': μ ≤ 2,5')]),
      eq([h1(), text(': μ > 2,5')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('σ'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('3 - 2,5')], frac(text('1,2'), rad(text('35')))),
        text(' = 2,47'),
      ]),
      eq([sub('Z', 'crítico'), text(' = 1,881')]),
      p('Como 2,47 es mayor que 1,881, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 99%, se usa el valor crítico 2,576.'),
      eq([sub('IC', '99%'), text(' = '), xbar(), text(' ± Z'), frac(text('σ'), rad(text('n')))]),
      eq([
        text('IC'),
        sub('IC', '99%'),
        text(' = 3 ± 2,576'),
        frac(text('1,2'), rad(text('35'))),
        text(' = 3 ± 0,52'),
      ]),
      eq([sub('IC', '99%'), text(' = (2,48; 3,52)')]),
      p('Para el intervalo de confianza al 90%, se usa el valor crítico 1,645.'),
      eq([
        text('IC'),
        sub('IC', '90%'),
        text(' = 3 ± 1,645'),
        frac(text('1,2'), rad(text('35'))),
        text(' = 3 ± 0,33'),
      ]),
      eq([sub('IC', '90%'), text(' = (2,67; 3,33)')]),
      p('Para el tamaño de muestra, se usa 95% de confianza, desviación estándar 1,2 y error máximo 0,5 minutos.'),
      eq([text('n = '), sup(delim([frac([text('Zσ')], text('E'))]), '2')]),
      eq([
        text('n = '),
        sup(delim([frac([text('(1,96)(1,2)')], text('0,5'))]), '2'),
        text(' = 22,13'),
      ]),
      eq([text('n = 23')]),
    ],
    response:
      'Sí se comprueba que el tiempo medio de publicidad es mayor a 2,5 minutos. El intervalo al 99% es de 2,48 a 3,52 minutos, el intervalo al 90% es de 2,67 a 3,33 minutos y se necesita una muestra mínima de 23 registros.',
  },
  {
    statement:
      'El gerente de una distribuidora indica que el número promedio de llamadas por semana a los clientes es de 23 llamadas. Se pidió al personal de ventas que presentara informes semanales con los clientes llamados durante una semana. En una muestra de 36 informes semanales se determinó un promedio de 22.4 llamadas a clientes por semana, y una desviación estándar de 5 llamadas.',
    parts: [
      'a) Compruebe la afirmación del gerente con una significancia del 5%.',
      'b) Determinar un intervalo de confianza del 98% para el número promedio de llamadas semanales a clientes.',
    ],
    blocks: [
      p('Datos: la muestra es grande, por eso se usa aproximación normal con la desviación estándar muestral.'),
      eq([text('n = 36; '), xbar(), text(' = 22,4; s = 5; '), mu0(), text(' = 23')]),
      p('Como se comprueba una afirmación de igualdad, la prueba es bilateral.'),
      eq([h0(), text(': μ = 23')]),
      eq([h1(), text(': μ ≠ 23')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('s'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('22,4 - 23')], frac(text('5'), rad(text('36')))),
        text(' = -0,72'),
      ]),
      eq([sub('Z', 'crítico'), text(' = ±1,96')]),
      p('Como -0,72 se encuentra dentro del intervalo de aceptación, no se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 98%, se usa el valor crítico 2,326.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('s'), rad(text('n')))]),
      eq([
        text('IC = 22,4 ± 2,326'),
        frac(text('5'), rad(text('36'))),
        text(' = 22,4 ± 1,94'),
      ]),
      eq([sub('IC', '98%'), text(' = (20,46; 24,34)')]),
    ],
    response:
      'No se rechaza la afirmación del gerente de que el promedio sea de 23 llamadas semanales. Con 98% de confianza, el promedio está entre 20,46 y 24,34 llamadas.',
  },
  {
    statement:
      'Una empresa fabrica focos cuya duración tiene una distribución aproximadamente normal con desviación estándar poblacional de 60 horas. Suponga que una muestra de 20 focos tiene una duración promedio de 780 horas.',
    parts: [
      'a) Demuestre que la duración mínima promedio de los focos es de por lo menos 750 horas, con una significancia del 3%.',
      'b) Calcule e interprete un intervalo de confianza del 96% para la duración promedio de todos los focos producidos por esta empresa.',
    ],
    blocks: [
      p('Datos: se conoce la desviación estándar poblacional, por eso se utiliza Z.'),
      eq([text('n = 20; '), xbar(), text(' = 780; σ = 60; '), mu0(), text(' = 750')]),
      p('Para demostrar que la duración promedio es por lo menos 750 horas, se plantea una prueba de cola derecha.'),
      eq([h0(), text(': μ ≤ 750')]),
      eq([h1(), text(': μ > 750')]),
      eq([text('Z = '), frac([xbar(), text(' - '), mu0()], frac(text('σ'), rad(text('n'))))]),
      eq([
        text('Z = '),
        frac([text('780 - 750')], frac(text('60'), rad(text('20')))),
        text(' = 2,24'),
      ]),
      eq([sub('Z', 'crítico'), text(' = 1,881')]),
      p('Como 2,24 es mayor que 1,881, se rechaza la hipótesis nula.'),
      p('Para el intervalo de confianza al 96%, se usa el valor crítico 2,054.'),
      eq([text('IC = '), xbar(), text(' ± Z'), frac(text('σ'), rad(text('n')))]),
      eq([
        text('IC = 780 ± 2,054'),
        frac(text('60'), rad(text('20'))),
        text(' = 780 ± 27,55'),
      ]),
      eq([sub('IC', '96%'), text(' = (752,45; 807,55)')]),
    ],
    response:
      'Sí hay evidencia para afirmar que la duración promedio es por lo menos 750 horas. Con 96% de confianza, la duración promedio está entre 752,45 y 807,55 horas.',
  },
];

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function withSpaceAttr(tagName, value) {
  const textValue = String(value);
  const preserve = /^\s|\s$/.test(textValue);
  const escaped = xmlEscape(textValue);
  return preserve
    ? `<${tagName} xml:space="preserve">${escaped}</${tagName}>`
    : `<${tagName}>${escaped}</${tagName}>`;
}

function run(value, options = {}) {
  const { bold = false, italic = false, size = 22 } = options;
  const props = [
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>',
    `<w:sz w:val="${size}"/>`,
    `<w:szCs w:val="${size}"/>`,
  ];
  if (bold) props.push('<w:b/>');
  if (italic) props.push('<w:i/>');
  return `<w:r><w:rPr>${props.join('')}</w:rPr>${withSpaceAttr('w:t', value)}</w:r>`;
}

function paragraph(runs, options = {}) {
  const {
    align = 'both',
    before = 0,
    after = 80,
    line = 260,
    indent = 0,
  } = options;
  const pPr = [
    `<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>`,
  ];
  if (align) pPr.push(`<w:jc w:val="${align}"/>`);
  if (indent) pPr.push(`<w:ind w:left="${indent}"/>`);
  return `<w:p><w:pPr>${pPr.join('')}</w:pPr>${runs.join('')}</w:p>`;
}

function p(value) {
  return paragraph([run(value)], { after: 60 });
}

function blankParagraph(after = 100) {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="${after}" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
}

function text(value) {
  return `<m:r>${withSpaceAttr('m:t', value)}</m:r>`;
}

function wrap(content) {
  return Array.isArray(content) ? content.join('') : content;
}

function sub(baseValue, subValue) {
  const base = baseValue ? text(baseValue) : text('');
  const subscript = subValue ? text(subValue) : text('');
  return `<m:sSub><m:e>${base}</m:e><m:sub>${subscript}</m:sub></m:sSub>`;
}

function sup(baseContent, supValue) {
  return `<m:sSup><m:e>${wrap(baseContent)}</m:e><m:sup>${text(supValue)}</m:sup></m:sSup>`;
}

function frac(num, den) {
  return `<m:f><m:num>${wrap(num)}</m:num><m:den>${wrap(den)}</m:den></m:f>`;
}

function rad(content) {
  return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:e>${wrap(content)}</m:e></m:rad>`;
}

function delim(content, begin = '(', end = ')') {
  return `<m:d><m:dPr><m:begChr m:val="${xmlEscape(begin)}"/><m:endChr m:val="${xmlEscape(end)}"/></m:dPr><m:e>${wrap(content)}</m:e></m:d>`;
}

function xbar() {
  return text('x̄');
}

function mu0() {
  return sub('μ', '0');
}

function h0() {
  return sub('H', '0');
}

function h1() {
  return sub('H', '1');
}

function eq(content, options = {}) {
  const { align = 'center', before = 0, after = 60 } = options;
  return [
    '<w:p>',
    '<w:pPr>',
    `<w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>`,
    `<w:jc w:val="${align}"/>`,
    '</w:pPr>',
    '<m:oMathPara>',
    `<m:oMath>${wrap(content)}</m:oMath>`,
    '</m:oMathPara>',
    '</w:p>',
  ].join('');
}

function exerciseBlock(exercise, index) {
  const body = [];
  body.push(
    paragraph([run(`${index + 1}. ${exercise.statement}`, { bold: true })], {
      before: index === 0 ? 80 : 120,
      after: 50,
      align: 'both',
    })
  );

  exercise.parts.forEach((part) => {
    body.push(
      paragraph([run(part, { bold: true })], {
        indent: 360,
        after: 40,
        align: 'both',
      })
    );
  });

  body.push(paragraph([run('Desarrollo:', { bold: true })], { align: 'left', after: 40 }));
  exercise.blocks.forEach((block) => body.push(block));
  body.push(
    paragraph([run('Respuesta:', { bold: true }), run(` ${exercise.response}`)], {
      before: 20,
      after: 110,
    })
  );
  return body.join('');
}

function buildDocumentXml(originalDocumentXml) {
  const openTagMatch = originalDocumentXml.match(/^([\s\S]*?<w:body>)/);
  const sectPrMatch = originalDocumentXml.match(/(<w:sectPr[\s\S]*<\/w:sectPr>)/);

  if (!openTagMatch || !sectPrMatch) {
    throw new Error('No se pudo recuperar la estructura base del documento.');
  }

  const body = [];
  body.push(
    paragraph(
      [run('Evidencia: Estimación y Prueba de hipótesis para una media', { bold: true, size: 32 })],
      { align: 'center', after: 100, line: 280 }
    )
  );

  exercises.forEach((exercise, index) => {
    body.push(exerciseBlock(exercise, index));
  });
  body.push(blankParagraph(80));

  return `${openTagMatch[1]}${body.join('')}${sectPrMatch[1]}</w:body></w:document>`;
}

function updateStyles(stylesXml) {
  let updated = stylesXml.replace(
    /<w:rFonts[^>]*w:asciiTheme="minorHAnsi"[^>]*\/>/,
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
  );

  updated = updated.replace(/<w:sz w:val="\d+"\/>/, '<w:sz w:val="22"/>');
  updated = updated.replace(/<w:szCs w:val="\d+"\/>/, '<w:szCs w:val="22"/>');

  updated = updated.replace(
    /<w:lang w:val="[^"]+" w:eastAsia="[^"]+" w:bidi="[^"]+"\/>/,
    '<w:lang w:val="es-PE" w:eastAsia="es-PE" w:bidi="ar-SA"/>'
  );

  return updated;
}

async function main() {
  const inputBuffer = fs.readFileSync(baseDocxPath);
  const zip = await JSZip.loadAsync(inputBuffer);

  const documentXml = await zip.file('word/document.xml').async('string');
  const stylesXml = await zip.file('word/styles.xml').async('string');

  zip.file('word/document.xml', buildDocumentXml(documentXml));
  zip.file('word/styles.xml', updateStyles(stylesXml));

  const outputBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });

  fs.writeFileSync(outputDocxPath, outputBuffer);
  console.log(outputDocxPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
