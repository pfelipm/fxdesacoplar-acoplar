![Banner des(acoplar)](https://user-images.githubusercontent.com/12829262/89408938-84324100-d721-11ea-85f0-89b0a2e95b10.png)

*   [(Des)acoplando las filas de un intervalo de datos](#desacoplando-las-filas-de-un-intervalo-de-datos)
*   [Función DESACOPLAR()](#funci%C3%B3n-desacoplar)
*   [Función ACOPLAR()](#funci%C3%B3n-acoplar)
*   [Modo de uso](#modo-de-uso)
*   [Mirando bajo el capó ⚙️ (implementación](#mirando-bajo-el-cap%C3%B3-gear-implementaci%C3%B3n))
*   [Mejoras](#mejoras)
*   [Licencia](#licencia)

# (Des)acoplando las filas de un intervalo de datos

Este repositorio contiene el código Apps Script necesario para implementar las funciones personalizadas para hojas de cálculo de Google `DESACOPLAR` y `ACOPLAR`. Encontrarás una motivación más detallada acerca de su utilidad en este artículo introductorio.

**Separar filas con respuestas múltiples**

La función `DESACOPLAR` recorre todas las filas de un intervalo de datos, que se facilita como parámetro de entrada, generando tantas copias consecutivas de cada una de dichas filas como sean necesarias para separar los datos delimitados por la secuencia de caracteres facilitada que se encuentren almacenados en las columnas indicadas por el usuario. Resulta ideal (y de hecho creo con esta finalidad) para para facilitar el tratamiento estadístico de las respuestas a un formulario cuando alguna de sus preguntas admite múltiples opciones (casillas de verificación), que en ese caso aparecen separadas por la secuencia delimitadora `,` . Tras ser desacopladas, las respuestas (filas) con opciones múltiples se repiten en el intervalo resultante para cada combinación posible de los valores múltiples únicos de las columnas especificadas.

Veamos un ejemplo en el que se muestran inicialmente las respuestas recibidas en la hoja de cálculo asociada a un formulario de Google utilizado en un proceso de inscripción a una serie de actividades de formación:

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom, Edpuzzle</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Hidalgo Iglesias, Pedro</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom, Edpuzzle, Sites</td><td>Mañana, Tarde</td><td>Online</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle, Sites</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Classroom, Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Mañana, Tarde</td><td>Presencial</td></tr></tbody></table>

Como se puede apreciar, las columnas **Curso** y **Turno** contiene valores múltiples, separados por _coma espacio_, habituales cuando se utilizan preguntas a las que se puede responder marcando casillas de verificación. Veamos ahora cuál sería el resultado cuando se aplica la función `DESACOPLAR` sobre el intervalo anterior y las mencionadas columnas:

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Prieto González, Isabel</td><td>Edpuzzle</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Hidalgo Iglesias, Pedro</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom</td><td>Mañana</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Edpuzzle</td><td>Mañana</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Edpuzzle</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Classroom</td><td>Mañana</td><td>Online</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Tarde</td><td>Presencial</td></tr></tbody></table>

Ahora solo vemos valores únicos en las columnas **Curso** y **Turno** de cada fila. Para lograrlo, se han generado tantas respuestas a partir de cada fila (respuesta) original como han sido necesarias para acoger todas las combinaciones posibles de los valores de las columnas que inicialmente contenían múltiples opciones, realizando por tanto algo así como un proceso de descoplándodo.

**Unificar filas con respuestas múltiples**

La función `ACOPLAR` realiza un proceso complementario al anterior, **aunque no necesariamente simétrico** ⚠️. En este caso, la función recibe también un intervalo de datos pero ahora, en lugar de indicar las columnas cuyos datos múltiples deben procesarse, se debe facilitar la columna o columnas **clave** cuya combinación caracteriza de manera individual cada una de las respuestas recibidas (entidades o elementos únicos, hablando en términos generales). La función reagrupará las filas de manera que los distintos valores de aquellas columnas no identificadas como de tipo clave se consolidarán en una cadena de texto única, utilizando como separador la secuencia delimitadora que se especifique, en cada respuesta diferenciada.

En el caso de nuestro ejemplo, si aplicamos `ACOPLAR` sobre la tabla anterior, indicando como columna clave **Nombre**, el resultado será de nuevo el inicial:

<table><tbody><tr><td><strong>Nombre</strong></td><td><strong>Curso</strong></td><td><strong>Turno</strong></td><td><strong>Modalidad</strong></td></tr><tr><td>Prieto González, Isabel</td><td>Classroom, Edpuzzle</td><td>Mañana</td><td>Presencial</td></tr><tr><td>Hidalgo Iglesias, Pedro</td><td>Sites</td><td>Tarde</td><td>Online</td></tr><tr><td>Sánchez Santana, María</td><td>Classroom, Edpuzzle, Sites</td><td>Mañana, Tarde</td><td>Online</td></tr><tr><td>Moya González, Manuel</td><td>Edpuzzle, Sites</td><td>Tarde</td><td>Presencial</td></tr><tr><td>Carmona Navarro, Juan Carlos</td><td>Classroom, Sites</td><td>Mañana</td><td>Online</td></tr><tr><td>Medina Márquez, Gloria</td><td>Classroom</td><td>Mañana, Tarde</td><td>Presencial</td></tr></tbody></table>

`ACOPLAR` evita duplicados, ignorando los valores múltiples repetidos correspondientes a una misma entidad (filas con el mismo campo o campos clave).

# Función DESACOPLAR()

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

# Función ACOPLAR()

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

# Mirando bajo el capó :gear: (implementación)

# Mejoras

# **Licencia**

© 2020 Pablo Felip Monferrer ([@pfelipm](https://twitter.com/pfelipm)). Se distribuye bajo licencia MIT.
