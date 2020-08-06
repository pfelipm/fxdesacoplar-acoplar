![Banner des(acoplar)](https://user-images.githubusercontent.com/12829262/89408938-84324100-d721-11ea-85f0-89b0a2e95b10.png)

*   [(Des)acoplando las filas de un intervalo de datos](#desacoplando-las-filas-de-un-intervalo-de-datos)
*   [Función DESACOPLAR()](#usando-la-funci%C3%B3n-desacoplar)
*   [Función ACOPLAR()](#usando-la-funci%C3%B3n-acoplar)
*   [Modo de uso](#modo-de-uso)
*   [Mirando bajo el capó (implementación](#mirando-bajo-el-cap%C3%B3-implementaci%C3%B3n-gear) [⚙️](#mirando-bajo-el-cap%C3%B3-gear-implementaci%C3%B3n))
*   [Mejoras](#mejoras)
*   [Licencia](#licencia)

# (Des)acoplando las filas de un intervalo de datos

Este repositorio contiene el código Apps Script necesario para implementar las funciones personalizadas para hojas de cálculo de Google `DESACOPLAR` y `ACOPLAR`. Encontrarás una motivación más detallada acerca de su utilidad en este artículo introductorio.

### **Separar filas con respuestas múltiples**

La función `DESACOPLAR()` recorre todas las filas de un intervalo de datos, que se facilita como parámetro de entrada, generando tantas copias consecutivas de cada una de dichas filas como sean necesarias para separar los datos delimitados por la secuencia de caracteres facilitada que se encuentren almacenados en las columnas indicadas por el usuario. Resulta ideal (y de hecho creo con esta finalidad) para para facilitar el tratamiento estadístico de las respuestas a un formulario cuando alguna de sus preguntas admite múltiples opciones (casillas de verificación), que en ese caso aparecen separadas por la secuencia delimitadora `,` . Tras ser desacopladas, las respuestas (filas) con opciones múltiples se repiten en el intervalo resultante para cada combinación posible de los valores múltiples únicos de las columnas especificadas.

Veamos un ejemplo en el que se muestran inicialmente las respuestas recibidas en la hoja de cálculo asociada a un formulario de Google utilizado en un proceso de inscripción a una serie de actividades de formación:

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom, Edpuzzle</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Hidalgo Iglesias, Pedro</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom, Edpuzzle, Sites</td><td>Mañana, Tarde</td><td>Online</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle, Sites</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Classroom, Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Mañana, Tarde</td><td>Presencial</td></tr></tbody></table>

Como se puede apreciar, las columnas **Curso** y **Turno** contiene valores múltiples, separados por _coma espacio_, habituales cuando se utilizan preguntas a las que se puede responder marcando casillas de verificación. Veamos ahora cuál sería el resultado cuando se aplica la función `DESACOPLAR()` sobre el intervalo anterior y las mencionadas columnas:

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Prieto González, Isabel</td><td>Edpuzzle</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Hidalgo Iglesias, Pedro</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom</td><td>Mañana</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Edpuzzle</td><td>Mañana</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Edpuzzle</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Classroom</td><td>Mañana</td><td>Online</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Tarde</td><td>Presencial</td></tr></tbody></table>

Ahora solo vemos valores únicos en las columnas **Curso** y **Turno** de cada fila. Para lograrlo, se han generado tantas respuestas a partir de cada fila (respuesta) original como han sido necesarias para acoger todas las combinaciones posibles de los valores de las columnas que inicialmente contenían múltiples opciones, realizando por tanto algo así como un proceso de descoplándodo.

### **Unificar filas con respuestas múltiples**

La función `ACOPLAR` realiza un proceso complementario al anterior, **aunque no necesariamente simétrico** ⚠️. En este caso, la función recibe también un intervalo de datos pero ahora, en lugar de indicar las columnas cuyos datos múltiples deben procesarse, se debe facilitar la columna o columnas **clave** cuya combinación caracteriza de manera individual cada una de las respuestas recibidas (entidades o elementos únicos, hablando en términos generales). La función reagrupará las filas de manera que los distintos valores de aquellas columnas no identificadas como de tipo clave se consolidarán en una cadena de texto única, utilizando como separador la secuencia delimitadora que se especifique, en cada respuesta diferenciada.

En el caso de nuestro ejemplo, si aplicamos `ACOPLAR` sobre la tabla anterior, indicando como columna clave **Nombre**, el resultado será de nuevo el inicial:

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom, Edpuzzle</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Hidalgo Iglesias, Pedro</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom, Edpuzzle, Sites</td><td>Mañana, Tarde</td><td>Online</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle, Sites</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Classroom, Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Mañana, Tarde</td><td>Presencial</td></tr></tbody></table>

`ACOPLAR` evita duplicados, ignorando los valores múltiples repetidos correspondientes a una misma entidad (filas con el mismo campo o combinación de campos clave concatenados).

# Usando la función DESACOPLAR()

```
=DESACOPLAR( intervalo ; [encabezado] ; [separador] ; columna ; [otras_columnas]  )
```

*   `intervalo`: Rango de datos de entrada.
*   `encabezado`: Un valor `VERDADERO` o `FALSO` que indica si el intervalo de datos especificado dispone de una fila de encabezado con etiquetas para cada columna. En este caso se reproducirá también en el intervalo de datos devuelto como resultado. De omitirse se considera `VERDADERO.`
*   `separador`: Secuencia de caracteres que separa los valores múltiples. También es opcional, si se omite se usará `,` (coma espacio).
*   `columna`: Indicador numérico de la posición de la columna que contiene datos múltiples que deben desacoplarse, contando desde la izquierda y comenzado por 1.
*   `otras_columnas`: Columnas adicionales, opcionales, con datos múltiples a desacoplar. Pueden indicarse tantas como se deseen, siempre separadas por `;` (punto y coma).

Ejemplo:

```
=DESACOPLAR( A1:D4 ; ; ; 2 ; 3 )
```

![fx DESACOPLAR - Hojas de cálculo de Google](https://user-images.githubusercontent.com/12829262/89433912-32021780-d743-11ea-9913-11334a60be59.gif)

# Usando la función ACOPLAR()

```
=ACOPLAR( intervalo ; [encabezado] ; [separador] ; columna ; [otras_columnas]  )
```

*   `intervalo`: Rango de datos de entrada.
*   `encabezado`: Un valor `VERDADERO` o `FALSO` que indica si el intervalo de datos especificado dispone de una fila de encabezado con etiquetas para cada columna. En este caso se reproducirá también en el intervalo de datos devuelto como resultado. De omitirse se considera `VERDADERO.`
*   `separador`: Secuencia de caracteres que se utilizará para separar los valores múltiples. También es opcional, si se omite se usará `,` (coma espacio).
*   `columna`: Indicador numérico de la posición de la columna clave que identifica entidades (filas) únicas, contando desde la izquierda y comenzado por 1.
*   `otras_columnas`: Columnas adicionales, opcionales, que también son de tipo clave. Pueden indicarse tantas como se deseen, siempre separadas por `;` (punto y coma). Cuando se especifican varias columnas clave se combinarán todas ellas para determinar qué filas deben diferenciarse.

Ejemplo:

```
=ACOPLAR( A1:D16 ; ; ; 1 )
```

![fx ACOPLAR # demo - Hojas de cálculo de Google](https://user-images.githubusercontent.com/12829262/89435797-a2119d00-d745-11ea-91a0-f97e3f8e3c83.gif)

# **Modo de uso**

Dos posibilidades distintas:

1.  Abre el editor GAS de tu hoja de cálculo (`Herramientas` :fast\_forward: `Editor de secuencias de comandos`), pega el código que encontrarás dentro del archivo `Código.gs` de este repositorio y guarda los cambios. Debes asegurarte de que se esté utilizando el nuevo motor GAS JavaScript V8 (`Ejecutar` :fast\_forward: `Habilitar ... V8`).
2.  Hazte una copia de esto :point\_right: [fx (Des)acoplar # demo](https://docs.google.com/spreadsheets/d/1_d391kb-1X1jKSEvvmJyxREJeGRLiJFsYFmlY8M7LV0/template/preview) :point\_left:, elimina su contenido y edita a tu gusto.

Las funciones, `DESACOPLAR` y `ACOPLAR` estarán en breve disponibles en mi complemento para hojas de cálculo [HdC+](https://gsuite.google.com/marketplace/app/hdc+/410659432888), junto con otras nuevas características que tengo previsto implementar próximamente.

![Selección_091](https://user-images.githubusercontent.com/12829262/86293166-64739e80-bbf2-11ea-8030-2e5f5c37fcaa.png)

# Mirando bajo el capó (implementación :gear:)

Como de costumbre, repasemos algunas cosillas relativas a la implementación. Si solo estás interesado en cómo usar estas funciones en tus propias hojas de cálculo, puedes omitir este apartado sin remordimientos.

`DESACOPLAR()` y `ACOPLAR()` son sendas funciones personalizadas para hojas de cálculo de Google creadas usando Apps Script y, como tales, tienen una estructura y un _modus operandi_ particulares que ya comenté con cierto detenimiento hablando de la implementación de otra función que he desarrollado recientemente, `MEDIAMOVIL()`. En particular, puedes revisar [esta sección](https://github.com/pfelipm/mediamovil/blob/master/README.md#mirando-bajo-el-cap%C3%B3-gear-implementaci%C3%B3n) de su documentación si no estás muy familiarizado con el modo en que:

*   Se disponen los elementos de la ayuda contextual por medio de [JSDoc](https://jsdoc.app/about-getting-started.html).
*   Se gestionan los parámetros de entrada, en general, y los opcionales, en particular.
*   Se realiza el control de errores y se emiten mensajes informativos para el usuario por medio de excepciones controladas, en su caso.

No obstante en esta ocasión también hay otros aspectos que me parece relevante comentar. Hablaremos en lo que sigue de:

*   Parámetros opcionales, de número indeterminado, y el operador de propagación.
*   Conjuntos JavaScript.
*   Funciones de tipo IIFE recursivas.

### Las interioridades de DESACOPLAR()

Comencemos por el bloque que se encarga del control de los parámetros de entrada de `DESACOPLAR` (el de `ACOPLAR` es prácticamente idéntico).

```javascript
function DESACOPLAR(intervalo, encabezado, separador, columna, ...masColumnas) {

  // Control de parámetros inicial

  if (typeof intervalo == 'undefined' || !Array.isArray(intervalo)) throw 'No se ha indicado un intervalo.';
  if (typeof encabezado != 'boolean') encabezado = true;
  if (intervalo.length == 1 && encabezado) throw 'El intervalo es demasiado pequeño, añade más filas.';
  separador = separador || ', ';
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto.';
```

Como hemos visto en apartados anteriores de este documento, `encabezado` y `separador` son parámetros opcionales de la función (de `columna` y `...masColumnas` hablaremos en un momento porque esa es otra guerra). Entonces ¿por qué no hemos usado la sintaxis ES6 habitual en estos casos? Algo como esto:

```javascript
function DESACOPLAR(intervalo, encabezado = 'true', separador = ', ', columna, ...masColumnas) {
```

Pues no lo hemos hecho porque este tipo de declaraciones esconde una trampa, y la trampa es que no es posible utilizar el punto y coma para "pasar" al siguiente parámetro, obviando su declaración explícita, de modo que se adopte el valor indicado por defecto. ¿Y eso por qué? Porque en ese caso lo que se le pasa realmente a la función es una cadena vacía. Y una cadena vacía no es lo mismo que nada.

Por ejemplo, de haber declarado `DESACOPLAR()` del modo indicado justo arriba, al hacer:

```
=DESACOPLAR( A1:D4 ; ; ; 2 )
```

Nos encontraríamos con esto:

<table><tbody><tr><td><strong>Parámetro</strong></td><td><strong>Valor</strong></td></tr><tr><td>intervalo</td><td>Matriz con el contenido de las celdas del rango A1:D4</td></tr><tr><td>encabezado</td><td>'' (cadena vacía)&nbsp;<span>:warning:</span>&nbsp;</td></tr><tr><td>separador</td><td>'' (cadena vacía)&nbsp;<span>:warning:</span>&nbsp;</td></tr><tr><td>columna</td><td>2</td></tr></tbody></table>

Y evidentemente no es lo que queremos.

Hay que ser muy cauto a la hora de utilizar la declaración de parámetros opcionales que nos ofrece ES6. En general, si tras un parámetro declarado de este modo pueden aparecer otros, será necesario introducir controles adicionales sobre tipos (`typeof`) y / o valores para asegurarnos de que no se nos cuela nada que no debería. Incluso es posible que tengamos que recurrir al modo en que se hacían las cosas antes de ES6, como aquí, para salvar este problemilla con las cadenas vacías:

```javascript
 separador = separador || ', ';
```

¿Y qué pasa con las columnas?

```javascript
function DESACOPLAR(intervalo, encabezado, separador, columna, ...masColumnas) {

  // Control de parámetros inicial

  ...

  let columnas = typeof columna != 'undefined' ? [columna, ...masColumnas] : [...masColumnas];
  if (columnas.length == 0) throw 'No se han indicado columnas a descoplar.';
  if (columnas.some(col => typeof col != 'number' || col < 1)) throw 'Las columnas deben indicarse mediante números enteros';
  if (Math.max(...columnas) > intervalo[0].length) throw 'Al menos una columna está fuera del intervalo.';
```

Al diseñar estas funciones me pareció buena idea permitir que el usuario pudiera especificar un número indefinido de columnas. Esto lo conseguimos utilizando el (bendito) **operador de propagación** de ES6 (`...`), que aquí viene a significar algo así como "y todo lo que venga detrás". En este caso, los valores cardinales del resto de columnas pasadas como parámetro se reciben dentro del vector `masColumnas`, que viene a ser algo así como un coche escoba para **el** **resto** de parámetros. Comodísimo, oiga. Y decía hace un momento lo de _bendito_ porque antes de ES6 teníamos con andarnos con [saltos mortales hacia atrás](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Functions/arguments) com el objeto `arguments` para resolver esto de manera más o menos satisfactoria.

Si analizas el código responsable de controlar la corrección de los parámetros que indican las columnas con datos múltiples, comprobarás que lo que se hace no es otra cosa que construir un vector numérico (de columnas) que integra tanto el parámetro `columna` como el vector `masColumnas`, verificando que todos sus elementos son numeritos dentro de un rango aceptable. Ah, y si te fijas realmente `columna` no es obligatorio, la única exigencia es que al menos se haya indicado una.

En mi opinión, dominar el operador de propagación y saber emplear las denominadas _asignaciones desestructurantes_ (¡menudo palabro!) son aspectos fundamentales para hablar un JavaScript elegante. Si estas cosas te suenan un poco (o un mucho) a chino, te recomiendo una leída atenta (y probablemente reiterada) a este [excelente artículo](https://codeburst.io/a-simple-guide-to-destructuring-and-es6-spread-operator-e02212af5831).

Avancemos un poco, justo debajo del bloque de gestión de parámetros nos encontramos con esto:

```javascript
// Se construye un conjunto (set) para evitar automáticamente duplicados en columnas con valores múltiples

let colSet = new Set();
columnas.forEach(col => colSet.add(col - 1));
```

JavaScript dispone de dos estructuras de datos extremadamente interesantes: los **mapas** (_map_) y los **conjuntos** (_set_). Se trata en ambos casos de colecciones iterables similares a los vectores, pero con ciertas particularidades que las hacen preferibles a estos últimos en determinadas circunstancias. Por ejemplo, lo bueno que tienen los **conjuntos** es que **evitan por su propia naturaleza la inserción de datos duplicados**, y además lo hacen mediante una estrategia interna basada en _tablas hash_ que resulta extremadamente eficiente, hablamos de un coste asintótico tipo **O(1)**, probablemente mucho más de lo que una implementación basada en vectores, en la que se comprobara la existencia de cada elemento antes de su inserción, alcanzaría. Por cierto, si quieres saber más sobre vectores y conjuntos, no dejes de leer [esto](https://medium.com/front-end-weekly/es6-set-vs-array-what-and-when-efc055655e1a).

El caso es que verás que en el código de estas dos funciones personalizadas se hace un uso insistente de los conjuntos. En el fragmento de código anterior, por ejemplo, se utiliza uno para eliminar posibles elementos duplicados en la indicación de las columnas con valores múltiples por parte de un usuario posiblemente despistado: simplemente se van metiendo los parámetros que identifican las columnas en él (restando 1 por aquello de que los arrays JavaScript comienzan en 0, como ya sabemos). Así de fácil. Este conjunto (`colSet`) será utilizado más abajo en el meollo del trabajo que realiza la función.

Ya solo queda contemplar la posibilidad de que exista una fila de encabezado en el intervalo de datos a procesar, que se colocará en su sitio justo antes de que la función devuelva el intervalo ya desacoplado. Sí, otra vez el operador de desestructuración, en este caso para [concatenar vectores](https://twitter.com/pfelipm/status/1279056400476524545).

```javascript
// Listos para comenzar

if (encabezado) encabezado = intervalo.shift();

/* .......................................*/ 
/* Aquí el resto del código de la función */
/* .......................................*/ 

// Si hay fila de encabezados, colocar en 1ª posición en la matriz de resultados

return encabezado.map ? [encabezado, ...intervaloDesacoplado] : intervaloDesacoplado;
```

Ya solo nos queda pegarle un vistazo a ese "resto del código de la función". La estrategia que se sigue para realizar el proceso de desacoplamiento se desarrolla de acuerdo con esta secuencia de cuatro pasos:

1.  Se recorre una a una cada fila del intervalo (líneas 48 - 121 en el código fuente de la función).
2.  Se genera una estructura matricial que contiene los valores múltiples únicos de las columnas que deben desacoplarse (52 - 61).
3.  A partir de la estructura matricial anterior se genera una nueva en la que se realizan todas las combinaciones posibles entre los valores extraídos de cada una de las columnas (63 - 94). Esta es la parte más densa del código o probablemente la menos comprensible, de entrada, ya que se ha implementado con una función recursiva, que se invoca a sí misma al ser declarada en plan IIFE (_immediately invoked function expression_). Por cierto, información jugosa sobre las IIFE [aquí](https://gustavodohara.com/blogangular/todos-los-misterios-iife-immediately-invoked-function-expressions/).
4.  Por último, a partir de los valores de la fila original se generan tantas copias como combinaciones posibles se hayan generado en (3), completando datos con los procedentes de las columnas que no se han desacoplado y se guardan en la matriz resultado que devolverá la función (101 - 119).

Para entender mejor lo que sigue, permíteme retomar el ejemplo con el que se iniciaba este documento, aunque con ligeras modificaciones en los datos para que lo que sigue resulte más clarificador. Supongamos que nuestra función está procesando esta fila y le hemos indicado que debe desacoplar los datos de las columnas 2 (**Curso**) y 3 (**Turno**):

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom, Edpuzzle</td><td>Mañana, Tarde</td><td>Presencial</td></tr></tbody></table>

Hagamos ahora zoom :mag: sobre el **paso** **\[2\]**. Para cada fila se construye un vector cuyos elementos son a su vez vectores que contienen los valores múltiples, descartando duplicados, contenidos en las columnas a desacoplar indicadas por el usuario. El contenido de cada celda se trocea con la secuencia de caracteres delimitadora utilizando el método `.spli`t y se añade a un conjunto (`opcionesSet`) para evitar valores duplicados. Finalmente, el conjunto se transforma en vector expandiéndolo mediante el operador de propagación (`...`)

```javascript
// Enumerar los valores únicos en cada columna que se ha indicado contiene datos múltiples

let opciones = [];
for (let col of colSet) {

  // Eliminar opciones duplicadas, si las hay, en cada columna gracias al uso de un nuevo conjunto

  let opcionesSet = new Set();
  String(fila[col]).split(separador).forEach(opcion => opcionesSet.add(opcion)); // split solo funciona con string, convertimos números
  opciones.push([...opcionesSet]); // también opciones.push(Array.from(opcionesSet))

}              
```

Al terminar, partiendo de nuestro ejemplo, nos encontraríamos esto:

```javascript
opciones = [  [ 'Classroom' , 'EdPuzzle' ] , [ 'Mañana , 'Tarde' ]  ]
```

A continuación viene la parte más complicada, en el **paso** **\[3\]**. Básicamente, el código de este bloque masticará el vector `opciones` anterior y devolverá esto:

```javascript
combinaciones = [  [ 'Classroom' , 'Mañana' ] ,  [ 'Classroom' , 'Tarde' ] , [ 'EdPuzzle' , 'Mañana' ] ,  [ 'EdPuzzle' , 'Tarde' ]  ]
```

Para conseguirlo, se declara y ejecuta en el mismo momento la función `combinar()` . Aunque las IIFE pueden resultar enigmáticas, en este caso me pareció buena idea utilizar una de ellas para evitar la declaración de una función externa a `desacoplar()`, poco adecuado desde el punto de vista de la organización del código.

```
let combinaciones = (function combinar(vector) {

/* Aquí el resto del código de la función */

})(opciones);
```

La función `combinar()` emplea una estrategia recursiva para reducir la complejidad del problema de combinar los elementos, en número indeterminado, de n vectores. El _caso base_ se da cuando hemos alcanzado la última columna a combinar, esto es, el último elemento del vector `opciones`. Recuerda que cada uno de estos elementos es a su vez un vector que contiene los valores múltiples extraídos de la columna correspondiente de la fila que estamos procesando. En este caso final, la función simplemente devuelve el vector de valores múltiples de la columna, siguiendo con nuestro ejemplo sería `[ 'Mañana' , 'Tarde' ]`.

```javascript
if (vector.length == 1) {

  // Fin del proceso recursivo

  return vector[0];
}
```

De no ser así nos encontraremos en el _caso general_, donde se realiza la reducción de complejidad del problema:

```javascript
else {

  // El resultado se calcula recursivamente

  let resultado = [];
  let subvector = vector.splice(0, 1)[0];
  let subresultado = combinar(vector);

  // Composición de resultados en la secuencia recursiva >> generación de vector de combinaciones

  subvector.forEach(e1 => subresultado.forEach(e2 => resultado.push([e1, ...e2])));

  return resultado;

}
```

Se _descabeza_ el vector `opciones` y se invoca de nueva `combinar()` con los elementos que quedan. A medida que se va deshaciendo la recursión, partiendo del caso base y de "abajo a arriba", se va montando el vector `resultado`, generando todas las posibles combinaciones de los elementos devueltos por la última llamada a la función recursiva y los que forman parte del elemento descabezado en cada etapa de la recursión por medio de sendos `.forEach`, de nuevo acompañados por nuestro ineludible operador (`...`), que ahora usamos para concatenar los elementos de `subvector` y `subresultado`.

![(Des)Acoplar # diagrama recursión](https://user-images.githubusercontent.com/12829262/89544729-c7b0ac00-d802-11ea-849b-c734c9f1116c.png)

Sí, las secuencias recursivas en ocasiones resuelven problemas complejos sin esfuerzo aparente. Y aunque siempre pueden transformarse en iterativas, lo que normalmente se traduce en algoritmos más eficientes, resultan tan naturales y elegantes que en este caso me vas a permitir que no lo haga.

Finalmente, en el **paso \[4\]** ya solo hay que duplicar cada fila tantas veces como sea necesario para acomodar las combinaciones de las columnas con valores múltiples, guardar las filas generadas en la matriz `intervaloDesacoplado` y finalmente devolverla como resultado, claro está, con o sin su fila de encabezados.

```javascript
  // Ahora hay que generar las filas repetidas para cada combinación de datos múltiples
  // Ej:
  //     ENTRADA: combinaciones = [ [a, 1], [a, 2], [b, 1], [b, 2] ]
  //     SALIDA:  respuestaDesacoplada = [ [Pablo, a, 1, Tarde], [Pablo, a, 2, Tarde], [Pablo, b, 1, Tarde], [Pablo, b, 2, Tarde] ]

  let respuestaDesacoplada = combinaciones.map(combinacion => {

    let colOpciones = 0;
    let filaDesacoplada = [];
    fila.forEach((valor, columna) => {

      // Tomar columna de la fila original o combinación de datos generada anteriormente
      // correspondiente a cada una de las columnas con valores múltiples

      if (!colSet.has(columna)) filaDesacoplada.push(valor);
      else filaDesacoplada.push(combinacion[colOpciones++]);

    });
    return filaDesacoplada;
  });

  // Se desestructura (...) respuestaDesacoplada dado que combinaciones.map es [[]]

  intervaloDesacoplado.push(...respuestaDesacoplada);

});

// Si hay fila de encabezados, colocar en 1ª posición en la matriz de resultados

return encabezado.map ? [encabezado, ...intervaloDesacoplado] : intervaloDesacoplado; 

}
```

Y hasta aquí llega la función `DESACOPLAR()`.

### Las interioridades de ACOPLAR()

La implementación de `ACOPLAR()`, por su parte, creo que es un poco más sencilla.

El bloque de control de parámetros es prácticamente idéntico, aunque en este caso la columna o columnas que facilita el usuario son las que determinarán qué filas son identificadas como entidades únicas y diferenciadas, algo así como la _clave principal_ del intervalo de datos. Los valores en el resto de columnas, no especificadas de manera explícita como parámetros al invocar la función, serán los que se combinarán para generar una sola fila canónica, con valores múltiples (sin repetición) separados por la secuencia de caracteres de separación escogida. En esta ocasión nos vendrá bien tener también a mano este último grupo de columnas, así que llevaremos la cuenta de ambos grupos mediante los conjuntos `colSet` y `colNoClaveSet`.

```javascript
// Se construye un conjunto (set) para evitar automáticamente duplicados en columnas CLAVE

let colSet = new Set();
columnas.forEach(col => colSet.add(col - 1));

// ...y en este conjunto se identifican las columnas susceptibles de contener valores que deben concatenarse

let colNoClaveSet = new Set();
for (let col = 0; col < intervalo[0].length; col++) {

  if (!colSet.has(col)) colNoClaveSet.add(col);

}
```

Tomaremos las mismas precauciones por lo que hace a la existencia (o no) de una fila de encabezado:

```javascript
if (encabezado) encabezado = intervalo.shift();
```

La estrategia que sigue `ACOPLAR()` es la siguiente:

1.  Se recorren todas las filas del intervalo de datos para extraer las claves principales. El intervalo de datos resultado contendrá tantas filas diferenciadas como claves únicas identificadas (líneas 179 - 191 en el código fuente de la función).
2.  Para cada clave única (193 - 224):
    1.  Se obtienen todas las filas en las que está presente dicha clave.
    2.  Se identifican los valores distintos presentes en cada columna que no está designada como clave (207 - 215).
    3.  Se construye la fila canónica concatenando los valores encontrados usando la secuencia de caracteres de separación y se almacena en la matriz resultado que devolverá la función (217 - 222).

**Paso \[1\]**. A destacar el uso de un carácter delimitador (`/`)cuando se concatenan los valores de las distintas columnas de tipo clave para evitar falsos positivos. Las claves de cada entidad se almacenan en un conjunto (sí, otra vez) para evitar nuevamente valores duplicados.

```javascript
// Listos para comenzar

if (encabezado) encabezado = intervalo.shift();

let intervaloAcoplado = [];

// 1ª pasada: recorremos el intervalo fila a fila para identificar entidades (concatenación de columnas clave) únicas

let entidadesClave = new Set();
intervalo.forEach(fila => {

  let clave = '';                
  // Se utiliza delimitador de campo (/) para evitar confusiones (Ej: claves: col1 = 'pablo', col3 = '1' / col1 = 'pablo1', col3 = '')
  for (let col of colSet) {clave += '/' + String(fila[col]);}
  entidadesClave.add(clave);

});
```

**Paso \[2-i\]**. Usamos el método `.filter` sobre el vector de filas del intervalo para procesar de manera agrupada las que presenten una misma clave única.

```javascript
// 2ª pasada: obtener filas para cada clave única, combinar columnas no-clave y generar filas resultado

for (let clave of entidadesClave) {

  let filasEntidad = intervalo.filter(fila => {

    let claveActual = '';
    for (let col of colSet) {claveActual += '/' + String(fila[col]);}
    return clave == claveActual;

  });
```

**Paso \[2-ii\]**. Se crea un vector con tantos conjuntos (una vez más se evitan así valores duplicados) como columnas de tipo no clave para contener los valores múltiples únicos en ellas asociados a cada entidad (fila) diferenciada, que como sabemos está caracterizada por una clave única.

```javascript
  // Acoplar todas las filas de cada entidad, concatenando valores en columnas no-clave con separador indicado

  let filaAcoplada = filasEntidad[0];  // Se toma la 1ª fila del grupo como base
  let noClaveSets = [];
  for (let col = 0; col < colNoClaveSet.size; col++) {noClaveSets.push(new Set())}; // Vector de sets para recoger valores múltiples  
  filasEntidad.forEach(fila => {

    let conjunto = 0;  
    for (let col of colNoClaveSet) {noClaveSets[conjunto++].add(String(fila[col]));}

  });
```

**Paso \[2-iii\]**. Ahora se construye la fila canónica que representa a todas las que contienen la misma clave única integrando las columnas clave con los valores múltiples extraídos de las que no lo son, concatenados utilizando la secuencia de caracteres delimitadora. Y, finalmente, tras repetir el proceso con la totalidad de claves únicas, la función devuelve la matriz resultado.

```javascript
  // Set >> Vector >> Cadena única con separador

  conjunto = 0;
  for (let col of colNoClaveSet) {filaAcoplada[col] = [...noClaveSets[conjunto++]].join(separador);}

  intervaloAcoplado.push(filaAcoplada);

}

return encabezado.map ? [encabezado, ...intervaloAcoplado] : intervaloAcoplado;
```

Aunque el proceso secuencial de las funciones `DESCOPLAR()` y `ACOPLAR()` no deja de ser lógico, lo cierto es que las estrategias utilizadas para manipular vectores bidimensionales y conjuntos, que utilizan frecuentemente el operador de propagación para transformar elementos iterables en vectores y realizar operaciones de concatenación, pueden resultar un tanto desconcertantes inicialmente. Si es así, te sugiero que trates de seguir nuevamente las secuencias descritas paso a paso, pero esta vez con papel y un bolígrafo en la mano para no perderte.

# Mejoras

Ambas funciones evitan los valores duplicados tanto en el proceso de combinación como en el de desagrupación de filas. Tal vez podría resultar de utilidad introducir un parámetro adicional en ambas para permitir la existencia de esos duplicados, aunque personalmente no acabo de apreciar en qué circunstancias eso sería conveniente.

Queda por tanto como ejercicio de ampliación, que seguramente requerirá del uso de vectores en lugar de los conjuntos actuales en aquellos fragmentos del código que se ocupan de la gestión de los valores múltiples en la columnas seleccionadas por el usuario.

# **Licencia**

© 2020 Pablo Felip Monferrer ([@pfelipm](https://twitter.com/pfelipm)). Se distribuye bajo licencia MIT.
